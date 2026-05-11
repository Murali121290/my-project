import re
import copy
import zipfile
import tempfile
import shutil
import os
from datetime import datetime
from pathlib import Path
from lxml import etree

from manuscript_core.extractor import _classify_paragraph, _is_reference_heading, mask_quotes

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def qn(tag):
    prefix, localname = tag.split(':')
    if prefix == 'w':
        return f"{{{W_NS}}}{localname}"
    return tag

_rev_id_counter = 1

def _get_time_str():
    return datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')


def _get_para_style(p_element) -> str:
    """Extract the paragraph style name from a <w:p> element."""
    pPr = p_element.find(qn('w:pPr'))
    if pPr is None:
        return ""
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        return ""
    return pStyle.get(qn('w:val')) or ""


def _get_para_text(p_element) -> str:
    """Concatenate all w:t text within a paragraph element."""
    parts = []
    for t in p_element.iter(qn('w:t')):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _fix_paragraph(p_element, fixes, para_mask: list[bool] | None = None):
    """
    Applies a list of dicts: {"pattern": re.Pattern, "replacement": str}
    to the runs in a paragraph using Track Changes.
    Skips matches that fall inside quoted regions (para_mask).
    """
    global _rev_id_counter

    changed = True
    while changed:
        changed = False

        # Build char offset map: for each run, its starting offset in the
        # paragraph's concatenated text. Needed for quote-mask checking.
        run_elements = p_element.findall(qn('w:r'))
        char_offsets: dict[int, int] = {}
        offset = 0
        for idx, r in enumerate(run_elements):
            char_offsets[idx] = offset
            for t in r.findall(qn('w:t')):
                offset += len(t.text or "")

        run_texts = []
        for r in run_elements:
            run_texts.append(''.join(t.text or "" for t in r.findall(qn('w:t'))))
        para_text = "".join(run_texts)

        for run_idx, r_elem in enumerate(run_elements):
            t_nodes = r_elem.findall(qn('w:t'))
            if not t_nodes:
                continue

            original_text = run_texts[run_idx]
            if not original_text:
                continue

            run_offset = char_offsets.get(run_idx, 0)
            run_end = run_offset + len(original_text)

            matched_fix = None
            match = None
            for fix in fixes:
                for m in fix["pattern"].finditer(para_text):
                    if m.start() >= run_offset and m.end() <= run_end:
                        if para_mask:
                            if any(para_mask[i] for i in range(m.start(), min(m.end(), len(para_mask)))):
                                continue
                        match = m
                        matched_fix = fix
                        break
                if match:
                    break

            if not match:
                continue

            changed = True
            rPr = r_elem.find(qn('w:rPr'))

            for t in t_nodes:
                r_elem.remove(t)

            insert_idx = p_element.index(r_elem)

            start = match.start() - run_offset
            end = match.end() - run_offset
            pre_text = original_text[:start]
            del_text = match.group(0)
            replacement = matched_fix["replacement"]
            if callable(replacement):
                ins_text = replacement(match)
            else:
                ins_text = replacement
            post_text = original_text[end:]

            # 1. Pre-text run
            if pre_text:
                new_t = etree.Element(qn('w:t'))
                new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                new_t.text = pre_text
                r_elem.append(new_t)
                insert_idx += 1
            else:
                p_element.remove(r_elem)

            # 2. Deletion
            del_node = etree.Element(qn('w:del'))
            del_node.set(qn('w:id'), str(_rev_id_counter))
            del_node.set(qn('w:author'), 'AI Consistency Checker')
            del_node.set(qn('w:date'), _get_time_str())
            _rev_id_counter += 1

            del_r = etree.Element(qn('w:r'))
            if rPr is not None: del_r.append(copy.deepcopy(rPr))
            del_t = etree.Element(qn('w:delText'))
            del_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            del_t.text = del_text
            del_r.append(del_t)
            del_node.append(del_r)

            p_element.insert(insert_idx, del_node)
            insert_idx += 1

            # 3. Insertion
            ins_node = etree.Element(qn('w:ins'))
            ins_node.set(qn('w:id'), str(_rev_id_counter))
            ins_node.set(qn('w:author'), 'AI Consistency Checker')
            ins_node.set(qn('w:date'), _get_time_str())
            _rev_id_counter += 1

            ins_r = etree.Element(qn('w:r'))
            if rPr is not None: ins_r.append(copy.deepcopy(rPr))
            ins_t = etree.Element(qn('w:t'))
            ins_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            ins_t.text = ins_text
            ins_r.append(ins_t)
            ins_node.append(ins_r)

            p_element.insert(insert_idx, ins_node)
            insert_idx += 1

            # 4. Post-text run
            if post_text:
                post_r = etree.Element(qn('w:r'))
                if rPr is not None: post_r.append(copy.deepcopy(rPr))
                post_t = etree.Element(qn('w:t'))
                post_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                post_t.text = post_text
                post_r.append(post_t)
                p_element.insert(insert_idx, post_r)

            break  # re-scan paragraph from start


def _fix_paragraph_targeted(p_element, fix: dict, allowed_start: int):
    """Apply a single fix only at the exact char offset `allowed_start` within the paragraph."""
    global _rev_id_counter

    run_elements = p_element.findall(qn('w:r'))
    char_offsets: dict[int, int] = {}
    offset = 0
    for idx, r in enumerate(run_elements):
        char_offsets[idx] = offset
        for t in r.findall(qn('w:t')):
            offset += len(t.text or "")

    run_texts = []
    for r in run_elements:
        run_texts.append(''.join(t.text or "" for t in r.findall(qn('w:t'))))
    para_text = "".join(run_texts)

    for run_idx, r_elem in enumerate(run_elements):
        t_nodes = r_elem.findall(qn('w:t'))
        if not t_nodes:
            continue
        original_text = run_texts[run_idx]
        if not original_text:
            continue
        run_offset = char_offsets.get(run_idx, 0)
        run_end = run_offset + len(original_text)

        for m in fix["pattern"].finditer(para_text):
            if m.start() != allowed_start:
                continue
            if not (m.start() >= run_offset and m.end() <= run_end):
                continue

            rPr = r_elem.find(qn('w:rPr'))
            for t in t_nodes:
                r_elem.remove(t)
            insert_idx = p_element.index(r_elem)

            start = m.start() - run_offset
            end = m.end() - run_offset
            pre_text = original_text[:start]
            del_text = m.group(0)
            replacement = fix["replacement"]
            ins_text = replacement(m) if callable(replacement) else replacement
            post_text = original_text[end:]

            if pre_text:
                new_t = etree.Element(qn('w:t'))
                new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                new_t.text = pre_text
                r_elem.append(new_t)
                insert_idx += 1
            else:
                p_element.remove(r_elem)

            del_node = etree.Element(qn('w:del'))
            del_node.set(qn('w:id'), str(_rev_id_counter))
            del_node.set(qn('w:author'), 'AI Consistency Checker')
            del_node.set(qn('w:date'), _get_time_str())
            _rev_id_counter += 1
            del_r = etree.Element(qn('w:r'))
            if rPr is not None: del_r.append(copy.deepcopy(rPr))
            del_t = etree.Element(qn('w:delText'))
            del_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            del_t.text = del_text
            del_r.append(del_t)
            del_node.append(del_r)
            p_element.insert(insert_idx, del_node)
            insert_idx += 1

            ins_node = etree.Element(qn('w:ins'))
            ins_node.set(qn('w:id'), str(_rev_id_counter))
            ins_node.set(qn('w:author'), 'AI Consistency Checker')
            ins_node.set(qn('w:date'), _get_time_str())
            _rev_id_counter += 1
            ins_r = etree.Element(qn('w:r'))
            if rPr is not None: ins_r.append(copy.deepcopy(rPr))
            ins_t = etree.Element(qn('w:t'))
            ins_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            ins_t.text = ins_text
            ins_r.append(ins_t)
            ins_node.append(ins_r)
            p_element.insert(insert_idx, ins_node)
            insert_idx += 1

            if post_text:
                post_r = etree.Element(qn('w:r'))
                if rPr is not None: post_r.append(copy.deepcopy(rPr))
                post_t = etree.Element(qn('w:t'))
                post_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                post_t.text = post_text
                post_r.append(post_t)
                p_element.insert(insert_idx, post_r)
            return


def apply_fixes_targeted(input_docx: Path, output_docx: Path, targeted_fixes: list[dict]):
    """Apply fixes only at specific paragraph + char-offset locations (editor per-occurrence mode).

    Each entry in targeted_fixes:
      {"para_index": int, "match_start": int, "surface": str, "replacement": str,
       "source": str, "region": str}
    """
    temp_dir = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(input_docx, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        xml_files = [
            temp_dir / "word" / "document.xml",
            temp_dir / "word" / "footnotes.xml",
            temp_dir / "word" / "endnotes.xml",
        ]

        for xml_file in xml_files:
            if not xml_file.exists():
                continue

            tree = etree.parse(str(xml_file))
            root = tree.getroot()

            para_counter = 0
            for p in root.findall(f".//{qn('w:p')}"):
                source = "body"
                ancestors = [a.tag for a in p.iterancestors()]
                if qn("w:tbl") in ancestors:
                    source = "table"
                elif qn("w:txbxContent") in ancestors or any(a.endswith("}textbox") for a in ancestors):
                    source = "textbox"

                # Find targeted fixes for this paragraph
                para_fixes = [
                    tf for tf in targeted_fixes
                    if tf["para_index"] == para_counter and tf.get("source", "body") == source
                ]

                if para_fixes:
                    for tf in para_fixes:
                        pat_str = r'\b' + re.escape(tf["surface"]) + r'\b'
                        fix = {
                            "pattern": re.compile(pat_str),
                            "replacement": tf["replacement"],
                        }
                        _fix_paragraph_targeted(p, fix, tf["match_start"])

                para_counter += 1

            tree.write(str(xml_file), xml_declaration=True, encoding="UTF-8", standalone="yes")

        with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for folder_name, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = Path(folder_name) / filename
                    arcname = file_path.relative_to(temp_dir)
                    zip_ref.write(file_path, arcname)
    finally:
        shutil.rmtree(temp_dir)


def apply_fixes_to_docx(input_docx: Path, output_docx: Path, fixes: list[dict], selected_rule_ids: list[str] | None = None):
    """
    Unzips the docx, parses document.xml, footnotes.xml, endnotes.xml,
    applies track changes, and rezips it.
    Respects exclusion zones (references, epigraphs, extracts, captions)
    and skips text inside double-quoted spans.

    Args:
        input_docx: Path to input DOCX file
        output_docx: Path to output DOCX file
        fixes: List of fix dicts with pattern, replacement, etc.
        selected_rule_ids: Optional list of rule IDs to filter fixes by
    """
    # Filter fixes by selected rule IDs if provided
    if selected_rule_ids:
        fixes = [f for f in fixes if f.get("rule_id") in selected_rule_ids]
    temp_dir = Path(tempfile.mkdtemp())

    try:
        with zipfile.ZipFile(input_docx, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        xml_files = [
            temp_dir / "word" / "document.xml",
            temp_dir / "word" / "footnotes.xml",
            temp_dir / "word" / "endnotes.xml"
        ]

        for xml_file in xml_files:
            if not xml_file.exists():
                continue

            tree = etree.parse(str(xml_file))
            root = tree.getroot()

            in_reference_block = False
            current_region = "front"

            for p in root.findall(f".//{qn('w:p')}"):
                # Determine source type for source_filter support
                source = "body"
                ancestors = [a.tag for a in p.iterancestors()]
                if qn("w:tbl") in ancestors:
                    source = "table"
                elif qn("w:txbxContent") in ancestors or any(a.endswith("}textbox") for a in ancestors):
                    source = "textbox"

                style_name = _get_para_style(p)
                para_text = _get_para_text(p)

                # Reference-block heading detection (same logic as extract_segments)
                if in_reference_block:
                    if "<ref-close>" in para_text.lower():
                        in_reference_block = False
                    elif re.match(r"^\s*(?:figure|table|box)\b", para_text, re.IGNORECASE):
                        in_reference_block = False

                # Region mapping using identical flags to the extractor
                p_text_lower = para_text.lower()
                if "<front>" in p_text_lower:
                    current_region = "front"
                elif "<body>" in p_text_lower:
                    current_region = "body"
                elif "<ref-open>" in p_text_lower:
                    current_region = "references"

                if _is_reference_heading(para_text) or "<ref-open>" in p_text_lower:
                    in_reference_block = True

                # Skip excluded paragraphs entirely
                excluded, _ = _classify_paragraph(style_name, para_text)
                if excluded or in_reference_block:
                    continue

                applicable_fixes = [
                    f for f in fixes 
                    if f.get("source_filter", source) == source
                    and f.get("region", current_region) == current_region
                ]
                if not applicable_fixes:
                    continue

                para_mask = mask_quotes(para_text)
                _fix_paragraph(p, applicable_fixes, para_mask=para_mask)

            tree.write(str(xml_file), xml_declaration=True, encoding="UTF-8", standalone="yes")

        # Re-zip
        with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for folder_name, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = Path(folder_name) / filename
                    arcname = file_path.relative_to(temp_dir)
                    zip_ref.write(file_path, arcname)

    finally:
        shutil.rmtree(temp_dir)


def build_fixes_from_selection(
    selected_patterns: list[dict],
    data: dict,
    selected_rule_ids: list[str] | None = None,
    exclude_elements: list[str] | None = None,
) -> list[dict]:
    """
    Reads the data-driven replacements natively embedded in the findings JSON.
    Filters them based on the exact (element, pattern) rows selected by the UI.

    Args:
        selected_patterns: List of selected (element, pattern) dicts from UI
        data: Findings data with all detected issues
        selected_rule_ids: Optional list of rule IDs to filter by
        exclude_elements: Optional list of elements to exclude (e.g., ["Figure", "Table", "Box"])

    Returns:
        List of fix dicts with pattern, replacement, rule_id, etc.
    """
    import re
    import logging
    from manuscript_core.ia_mapping import RULE_ID_TO_IA

    logger = logging.getLogger(__name__)

    if exclude_elements is None:
        exclude_elements = []

    # 1. Reverse lookup: Which rule_ids belong to the selected UI rows?
    target_rules = set()
    matched_patterns = []
    unmatched_patterns = []

    for sel in selected_patterns:
        elem = str(sel.get("element") or "").strip()
        pat = str(sel.get("pattern") or "").strip()

        if not elem or not pat:
            continue

        # Skip excluded elements (e.g., Figure, Table, Box for highlighting-only rules)
        if elem in exclude_elements:
            continue

        pattern_found = False
        # Scan IA template configuration
        for rule_id, (mapped_element, parent, mapped_pattern) in RULE_ID_TO_IA.items():
            if mapped_element == elem and mapped_pattern == pat:
                target_rules.add(rule_id)
                pattern_found = True

        if pattern_found:
            matched_patterns.append(f"{elem} / {pat}")
        else:
            unmatched_patterns.append(f"{elem} / {pat}")

    if unmatched_patterns:
        logger.warning(f"Unmatched patterns in template: {unmatched_patterns}")

    # If selected_rule_ids provided, filter further
    if selected_rule_ids:
        target_rules = target_rules & set(selected_rule_ids)

    logger.info(f"Found {len(target_rules)} matching rule_ids for {len(matched_patterns)} patterns")

    # 2. Build absolute literal replacement commands from the original JSON
    fixes = []
    findings = data.get('findings', [])

    for finding in findings:
        if finding.get('rule_id') in target_rules and finding.get('replacement'):

            # Use the search_pattern the detector used, or fallback to exact string
            pat_str = finding.get('search_pattern')
            if not pat_str:
                pat_str = r'\b' + re.escape(finding['surface']) + r'\b'

            fixes.append({
                "pattern": re.compile(pat_str),
                "replacement": finding['replacement'],
                "region": finding.get('region', 'body'),
                "source_filter": finding.get('source', 'body'),
                "rule_id": finding.get('rule_id'),  # Track rule_id for filtering
            })

    logger.info(f"Built {len(fixes)} fixes from {len(target_rules)} matching rules")
    return fixes


def build_highlight_texts_from_selection(
    selected_patterns: list,
    data: dict,
    selected_rule_ids: list = None,
    exclude_elements: list = None,
) -> list:
    """Collect TE findings with no replacement — these get yellow highlights instead of track changes."""
    import logging
    from manuscript_core.ia_mapping import RULE_ID_TO_IA

    logger = logging.getLogger(__name__)
    exclude_elements = exclude_elements or []
    target_rules = set()

    for sp in selected_patterns:
        elem = str(sp.get("element") or "").strip()
        pat = str(sp.get("pattern") or "").strip()
        if not elem or not pat or elem in exclude_elements:
            continue
        for rule_id, (mapped_element, _parent, mapped_pattern) in RULE_ID_TO_IA.items():
            if mapped_element == elem and mapped_pattern == pat:
                target_rules.add(rule_id)

    if selected_rule_ids:
        target_rules = target_rules & set(selected_rule_ids)

    findings = data.get("findings", [])
    highlight_texts = []
    seen = set()

    for finding in findings:
        if finding.get("rule_id") in target_rules and not finding.get("replacement"):
            pat_str = finding.get("search_pattern")
            if not pat_str:
                pat_str = r'\b' + re.escape(finding.get("surface", "")) + r'\b'
            key = (pat_str, finding.get("region", "body"), finding.get("source", "body"))
            if key not in seen:
                seen.add(key)
                highlight_texts.append({
                    "pattern":       re.compile(pat_str, re.IGNORECASE),
                    "region":        finding.get("region", "body"),
                    "source_filter": finding.get("source", "body"),
                    "rule_id":       finding.get("rule_id"),
                    "surface":       finding.get("surface", ""),
                })

    logger.info(f"Built {len(highlight_texts)} TE highlight entries (no-replacement findings)")
    return highlight_texts


def _highlight_pattern_in_para(para, pattern, doc):
    """Split runs at pattern matches and apply yellow highlight to matched segments."""
    from copy import deepcopy
    from docx.oxml.ns import qn as docx_qn
    from docx.oxml import OxmlElement

    runs_snapshot = list(para.runs)
    for run in runs_snapshot:
        text = run.text or ""
        if not text:
            continue
        matches = list(pattern.finditer(text))
        if not matches:
            continue

        # Entire run is the match — just highlight it directly
        if len(matches) == 1 and matches[0].start() == 0 and matches[0].end() == len(text):
            run.font.highlight_color = _TE_HIGHLIGHT_COLOR
            continue

        # Build segments: [(text_segment, is_match), ...]
        segments = []
        pos = 0
        for m in matches:
            if m.start() > pos:
                segments.append((text[pos:m.start()], False))
            segments.append((m.group(), True))
            pos = m.end()
        if pos < len(text):
            segments.append((text[pos:], False))

        run_elem = run._element
        parent = run_elem.getparent()
        idx = list(parent).index(run_elem)
        orig_rPr = run_elem.find(docx_qn('w:rPr'))

        new_elems = []
        for seg_text, is_match in segments:
            if not seg_text:
                continue
            new_r = OxmlElement('w:r')
            if orig_rPr is not None:
                new_r.append(deepcopy(orig_rPr))
            t = OxmlElement('w:t')
            if seg_text != seg_text.strip():
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = seg_text
            new_r.append(t)
            if is_match:
                rPr = new_r.find(docx_qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    new_r.insert(0, rPr)
                highlight = OxmlElement('w:highlight')
                highlight.set(docx_qn('w:val'), 'yellow')
                rPr.append(highlight)
            new_elems.append(new_r)

        parent.remove(run_elem)
        for i, elem in enumerate(new_elems):
            parent.insert(idx + i, elem)


def apply_te_highlights_to_docx(input_path: str, output_path: str, highlight_texts: list):
    """Apply yellow highlight to TE findings that have no replacement."""
    from docx import Document

    doc = Document(input_path)
    current_region = "body"

    for para in doc.paragraphs:
        text = para.text
        if "<front>" in text:
            current_region = "front"
        elif "<body>" in text:
            current_region = "body"
        elif "<ref-open>" in text:
            current_region = "references"

        applicable = [
            h for h in highlight_texts
            if h.get("region", "body") == current_region
        ]
        if not applicable:
            continue

        for hl in applicable:
            _highlight_pattern_in_para(para, hl["pattern"], doc)

    doc.save(output_path)


try:
    from docx.enum.text import WD_COLOR_INDEX
    _TE_HIGHLIGHT_COLOR = WD_COLOR_INDEX.YELLOW
except Exception:
    _TE_HIGHLIGHT_COLOR = 7  # YELLOW fallback value

