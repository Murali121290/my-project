"""
pipeline/step8_content_controls.py â€” SDT (content control) grouping and Nested Regex Tagging.

Creates two kinds of Word Structured Document Tags:

BLOCK SDTs (Captions):
  Figure Caption SDT:
    caption para â†’ any intermediate paras â†’ source/credit paras
  Table Caption SDT:  caption para only
  TableGroup SDT:     Table Caption SDT â†’ w:tbl â†’ abbreviation/footnote/source paras
  - Inner Captions: Dynamically parses internal strings (e.g. "Figure 12.1") and slices 
    inner tags exclusively (Figure, ChapNo, SeqNo).

INLINE SDTs (Citations):
  FigureRef outer inline SDT maps multiple nested inline SDTs (Figure, ChapNo, SeqNo).
  TableRef outer inline SDT maps multiple nested inline SDTs (Table, ChapNo, SeqNo).

Replaces legacy `fldChar` methodology with a high-fidelity Regex `lxml`-teardown engine mimicking VBA operations perfectly natively.
"""

import re
from copy import deepcopy
from lxml import etree
from docx import Document
from docx_pipeline.config import (
    CAPTION_BOUNDARY_STYLES, CAPTION_GROUP_STYLES, SDT_TAG_TABLEGROUP, BX_STYLE_RE
)
from docx_pipeline.utils.sdt_builder import make_block_sdt, make_inline_sdt
from docx_pipeline.utils.report import ReportLogger

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_ = "{%s}" % W

# --- Regex Patterns explicitly mapped to VBA User rules ---
RE_FIG_SINGLE = re.compile(r"(Figure|Fig\.?)\s+([0-9]{1,3})[\.-]([0-9]{1,3})", re.IGNORECASE)
RE_TBL_SINGLE = re.compile(r"(Table)\s+([0-9]{1,3})[\.-]([0-9]{1,3})", re.IGNORECASE)

RE_FIG_MULTI_AND = re.compile(r"(Figures?|Figs?\.?)\s+([0-9]{1,3})[\.-]([0-9]{1,3})\s+and\s+([0-9]{1,3})[\.-]([0-9]{1,3})", re.IGNORECASE)
RE_TBL_MULTI_AND = re.compile(r"(Tables?)\s+([0-9]{1,3})[\.-]([0-9]{1,3})\s+and\s+([0-9]{1,3})[\.-]([0-9]{1,3})", re.IGNORECASE)

RE_FIG_MULTI_TO = re.compile(r"(Figures?|Figs?\.?)\s+([0-9]{1,3})[\.-]([0-9]{1,3})\s+to\s+([0-9]{1,3})[\.-]([0-9]{1,3})", re.IGNORECASE)
RE_TBL_MULTI_TO = re.compile(r"(Tables?)\s+([0-9]{1,3})[\.-]([0-9]{1,3})\s+to\s+([0-9]{1,3})[\.-]([0-9]{1,3})", re.IGNORECASE)

_BX_SUFFIX_RE = re.compile(r'^BX\d+[_-](.+)$', re.IGNORECASE)


def _xml_style(el: etree._Element) -> str | None:
    pStyle = el.find(f".//{W_}pStyle")
    if pStyle is not None:
        return pStyle.get(W_ + "val")
    return None


def _find_child_index(parent: etree._Element, child: etree._Element) -> int:
    """Find the index of a child element in its parent using identity comparison."""
    for i, elem in enumerate(parent):
        if elem is child:
            return i
    raise ValueError(f"Child element not found in parent")

def _xml_text(el: etree._Element) -> str:
    return "".join(t.text or "" for t in el.iter(W_ + "t"))


# â”€â”€ Core XML Slicing Engine â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def split_runs(p_el: etree._Element, split_indices: list[int]):
    """
    Tears apart `<w:r>` tags exactly at the string index boundaries designated by regex matches.
    """
    if not split_indices: return
    indices = sorted(list(set(split_indices)))
    
    current_char_pos = 0
    runs = list(p_el.findall(f".//{W_}r"))
    
    for r in runs:
        t_els = r.findall(f".//{W_}t")
        if not t_els:
            continue
            
        run_text = "".join(t.text or "" for t in t_els)
        run_len = len(run_text)
        
        splits_in_run = [idx for idx in indices if current_char_pos < idx < current_char_pos + run_len]
        
        if splits_in_run:
            local_splits = [idx - current_char_pos for idx in splits_in_run]
            pieces = []
            prev = 0
            for ls in local_splits:
                pieces.append(run_text[prev:ls])
                prev = ls
            pieces.append(run_text[prev:])
            
            parent = r.getparent()
            r_index = list(parent).index(r)
            parent.remove(r)
            
            for piece in pieces:
                new_r = deepcopy(r)
                for old_t in new_r.findall(f".//{W_}t"):
                    new_r.remove(old_t)
                new_t = etree.SubElement(new_r, f"{W_}t")
                if piece.startswith(" ") or piece.endswith(" "):
                    new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                new_t.text = piece
                parent.insert(r_index, new_r)
                r_index += 1
                
        current_char_pos += run_len

def apply_character_style(runs_or_elements: list[etree._Element], style_name: str):
    """Pushes Word Character Style XML strictly onto <w:r> targets."""
    for el in runs_or_elements:
        if el.tag == f"{W_}r":
            rPr = el.find(f".//{W_}rPr")
            if rPr is None:
                rPr = etree.Element(f"{W_}rPr")
                el.insert(0, rPr)
            rStyle = rPr.find(f".//{W_}rStyle")
            if rStyle is None:
                rStyle = etree.Element(f"{W_}rStyle")
                rPr.append(rStyle)
            rStyle.set(f"{W_}val", style_name)

def wrap_range_in_sdt(p_el: etree._Element, start_idx: int, end_idx: int, alias: str, tag: str, style: str = None) -> etree._Element:
    """
    Sweeps direct structural children of `p_el`. Gathers anything bound inside `[start, end]`.
    Rips them out and encapsulates them seamlessly into an Inline SDT marker.
    """
    current_char = 0
    wrap_elements = []
    
    for child in list(p_el):
        t_els = child.findall(f".//{W_}t")
        text = "".join(t.text or "" for t in t_els)
        child_len = len(text)
        
        if child_len > 0:
            if current_char >= start_idx and current_char < end_idx:
                wrap_elements.append(child)
        else:
            if current_char > start_idx and current_char < end_idx:
                wrap_elements.append(child)
                
        current_char += child_len
        
    if wrap_elements:
        if style:
            apply_character_style(wrap_elements, style)
            
        first_el = wrap_elements[0]
        parent = first_el.getparent()
        insert_index = list(parent).index(first_el)
        
        for w in wrap_elements:
            parent.remove(w)
            
        new_sdt = make_inline_sdt(alias, tag, wrap_elements)
        parent.insert(insert_index, new_sdt)
        return new_sdt
    return None

# â”€â”€ Dynamic Regex Execution â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _process_paragraph_regex(p_el: etree._Element, logger: ReportLogger, is_caption: bool, caption_type: str = None) -> int:
    """Invokes mapping rules corresponding to citations. Applies nested SDTs structurally via split engine."""
    text = _xml_text(p_el)
    if not text.strip():
        return 0
        
    flags = 0
    patterns = []
    
    if is_caption:
        if caption_type == "Figure":
            patterns.append((RE_FIG_SINGLE, "Figure Caption", "Figure", "FIG-NUM"))
        elif caption_type == "Table":
            patterns.append((RE_TBL_SINGLE, "Table Caption", "Table", "TN"))
    else:
        patterns.extend([
            (RE_FIG_SINGLE, "FigureRef", "Figure", "FigureCitation"),
            (RE_TBL_SINGLE, "TableRef", "Table", "TableCitation"),
            (RE_FIG_MULTI_AND, "FigureRef", "Figure", "FigureCitation"),
            (RE_TBL_MULTI_AND, "TableRef", "Table", "TableCitation"),
            (RE_FIG_MULTI_TO, "FigureRef", "Figure", "FigureCitation"),
            (RE_TBL_MULTI_TO, "TableRef", "Table", "TableCitation"),
        ])
        
    all_matches = []
    for reg, outer_tag, inner_tag, style in patterns:
        for m in reg.finditer(text):
             all_matches.append({
                 "bounds": (m.start(), m.end()),
                 "groups": m.groups(),
                 "spans": [m.span(i) for i in range(1, len(m.groups())+1)],
                 "outer": outer_tag,
                 "inner": inner_tag,
                 "style": style,
                 "do_outer_wrap": not is_caption  
             })

    if not all_matches:
        return 0

    split_indices = set()
    for match in all_matches:
        split_indices.add(match["bounds"][0])
        split_indices.add(match["bounds"][1])
        for sf, sl in match["spans"]:
            if sf != -1: 
                split_indices.add(sf)
                split_indices.add(sl)
                
    # 1. Crack XML structure exactly at text boundaries
    split_runs(p_el, list(split_indices))
    
    # 2. Iterate matches backwards implicitly avoiding layout mutation index collision
    for match in sorted(all_matches, key=lambda x: x["bounds"][0], reverse=True):
        groups = match["groups"]
        spans = match["spans"]
        style = match["style"]
        
        # INNERS
        wrap_range_in_sdt(p_el, spans[0][0], spans[0][1], match["inner"], match["inner"], style)
        wrap_range_in_sdt(p_el, spans[1][0], spans[1][1], "ChapNo", "ChapNo", style)
        wrap_range_in_sdt(p_el, spans[2][0], spans[2][1], "SeqNo", "SeqNo", style)
        
        # Multi-references have extended scopes
        if len(groups) == 5:
            wrap_range_in_sdt(p_el, spans[3][0], spans[3][1], "ChapNo1", "ChapNo1", style)
            wrap_range_in_sdt(p_el, spans[4][0], spans[4][1], "SeqNo1", "SeqNo1", style)
            
        # OUTER
        if match["do_outer_wrap"]:
            wrap_range_in_sdt(p_el, match["bounds"][0], match["bounds"][1], match["outer"], match["outer"], None)
            
        flags += 1

    return flags


# â”€â”€ Block SDT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _collect_group(body: etree._Element, anchor_idx: int, is_table: bool) -> list:
    children = list(body)
    group    = [children[anchor_idx]]
    i        = anchor_idx + 1

    if is_table:
        if i < len(children) and children[i].tag == W_ + "tbl":
            group.append(children[i])
            i += 1

    while i < len(children):
        el    = children[i]
        sname = _xml_style(el) if el.tag == W_ + "p" else None

        if sname in CAPTION_BOUNDARY_STYLES:
            break
        if el.tag == W_ + "tbl":
            break

        if el.tag == W_ + "p":
            text = _xml_text(el).strip()
            if sname in CAPTION_GROUP_STYLES or not text:
                group.append(el)
                i += 1
                continue
            break
        i += 1
    return group

def _process_block_sdts(body: etree._Element, logger: ReportLogger) -> int:
    count = 0
    i = 0
    children = list(body)

    while i < len(children):
        el    = children[i]
        sname = _xml_style(el) if el.tag == W_ + "p" else None

        if sname not in ("FIG-LEG", "FGC", "T1", "TT"):
            i += 1
            continue

        is_table = (sname in ("T1", "TT"))
        cap_type = "Table" if is_table else "Figure"

        # Apply Caption Regex first dynamically creating inner SDTs
        _process_paragraph_regex(el, logger, is_caption=True, caption_type=cap_type)

        group      = _collect_group(body, i, is_table)
        insert_pos = list(body).index(group[0])

        for g_el in group:
            body.remove(g_el)

        if is_table:
            caption_sdt = make_block_sdt("Table Caption", "Table Caption", [group[0]])
            sdt = make_block_sdt(SDT_TAG_TABLEGROUP, SDT_TAG_TABLEGROUP,
                                 [caption_sdt] + group[1:])
            logger.info(f"Block Macro Engine Assigned SDT [Table Caption] to 1 Elements.")
            logger.info(f"Block Macro Engine Assigned SDT [TableGroup] to {len(group)} Elements.")
        else:
            alias = "Figure Caption"
            sdt = make_block_sdt(alias, alias, group)
            logger.info(f"Block Macro Engine Assigned SDT [{alias}] to {len(group)} Elements.")

        body.insert(insert_pos, sdt)
        count += 1
        children = list(body)
        i = insert_pos + 1

    return count

def _process_inline_sdts(body: etree._Element, logger: ReportLogger) -> int:
    count = 0
    for para in body.iter(W_ + "p"):
        sname = _xml_style(para)
        if sname in ("FIG-LEG", "FGC", "T1", "TT"):
            continue
            
        flags = _process_paragraph_regex(para, logger, is_caption=False)
        count += flags
        
    return count

_CAPTION_CHAR_STYLES = {
    "FIG-NUM":        "FFFF00",   # yellow        â€” figure number in captions
    "TN":             "92D050",   # brittengreen  â€” table number in captions
    "FigureCitation": "FFFF00",   # yellow        â€” figure reference in body text
    "TableCitation":  "92D050",   # brittengreen  â€” table reference in body text
}

def _ensure_caption_char_styles(doc: Document) -> None:
    """
    Creates FIG-NUM (yellow shading) and TN (brittengreen shading) character
    styles in the document if they are not already present.
    """
    existing = {s.name for s in doc.styles}
    styles_el = doc.styles.element
    for style_name, fill_hex in _CAPTION_CHAR_STYLES.items():
        if style_name in existing:
            continue
        style_el = etree.SubElement(styles_el, f"{W_}style")
        style_el.set(f"{W_}type", "character")
        style_el.set(f"{W_}customStyle", "1")
        style_el.set(f"{W_}styleId", style_name)
        name_el = etree.SubElement(style_el, f"{W_}name")
        name_el.set(f"{W_}val", style_name)
        rPr_el = etree.SubElement(style_el, f"{W_}rPr")
        shd_el = etree.SubElement(rPr_el, f"{W_}shd")
        shd_el.set(f"{W_}val",   "clear")
        shd_el.set(f"{W_}color", "auto")
        shd_el.set(f"{W_}fill",  fill_hex)


def _wrap_bx_inner_sdts(elems: list, digit: str) -> list:
    """
    For each paragraph in elems whose style has a BX sub-style suffix
    (e.g. BX1-TTL, BX1-TXT, BX2-H1), wrap it in an inner block SDT
    named NBX{digit}-{SUFFIX}. Tables and other non-paragraph elements
    pass through unchanged.
    """
    result = []
    for elem in elems:
        local_tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if local_tag == 'p':
            style = _xml_style(elem)
            m = _BX_SUFFIX_RE.match(style) if style else None
            if m:
                suffix = m.group(1).upper()
                inner_alias = f"NBX{digit}-{suffix}"
                result.append(make_block_sdt(inner_alias, inner_alias, [elem]))
                continue
        result.append(elem)
    return result


def _wrap_bx_groups(body: etree._Element, logger: ReportLogger) -> int:
    """
    Wrap consecutive BX<n>-* paragraphs (and any tables between them) into block SDTs.
    Groups are identified by matching paragraph styles to BX_STYLE_RE pattern.
    Tables appearing within a BX group are absorbed into the group.
    """
    children = list(body)
    groups = []  # list of (digit_str, [element, ...])

    current_digit = None
    current_elems = []

    for elem in children:
        local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

        if local == 'p':
            style = _xml_style(elem)
            m = BX_STYLE_RE.match(style) if style else None

            if m:
                digit = m.group(1)
                if digit != current_digit:
                    # Different BX number â€” close previous group
                    if current_elems:
                        groups.append((current_digit, current_elems))
                    current_digit = digit
                    current_elems = []
                current_elems.append(elem)
            else:
                # Non-BX paragraph â€” close current group
                if current_elems:
                    groups.append((current_digit, current_elems))
                    current_digit = None
                    current_elems = []

        elif local in ('tbl', 'sdt'):
            if current_digit is not None:
                current_elems.append(elem)

    # Handle group that reaches end of document
    if current_elems:
        groups.append((current_digit, current_elems))

    # Wrap each group in reverse order so indices stay valid
    count = 0
    for digit, elems in reversed(groups):
        first_idx = _find_child_index(body, elems[0])
        for e in elems:
            body.remove(e)
        inner_elems = _wrap_bx_inner_sdts(elems, digit)
        alias = f"NBX{digit}"
        sdt = make_block_sdt(alias, alias, inner_elems)
        body.insert(first_idx, sdt)
        count += 1

    if count > 0:
        logger.info(f"Block Macro Engine Assigned SDT [BX Groups] to {count} Groups.")

    return count


def run(doc_path: str, logger: ReportLogger) -> str:
    logger.set_step("8-content-controls")

    doc  = Document(doc_path)
    body = doc.element.body

    _ensure_caption_char_styles(doc)
    block_count  = _process_block_sdts(body, logger)
    inline_count = _process_inline_sdts(body, logger)
    bx_count     = _wrap_bx_groups(body, logger)

    doc.save(doc_path)
    logger.info(
        f"VBA Macro Upgrade SDT tagging complete: {block_count} Block SDT(s), "
        f"{inline_count} Inline Regex matches wrapped, {bx_count} BX Group(s) wrapped.")
    return doc_path

