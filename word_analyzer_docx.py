
import os
import re
import datetime
import subprocess
from pathlib import Path
from collections import defaultdict
from typing import List, Tuple, Dict, Any, Set, Optional
import jinja2
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

import importlib.util as _importlib_util
HAS_PDFPLUMBER: bool = _importlib_util.find_spec("pdfplumber") is not None

def _normalize_for_match(text: str) -> str:
    """Strip all non-alphanumeric chars and lowercase for robust matching."""
    return re.sub(r'\W+', '', text.lower())

# ------------------------------
# 1. HTML Templates & Helpers
# (Templates moved to templates/ directory and loaded via Jinja2)
# ------------------------------

def escape_html(s: str) -> str:
    if not isinstance(s, str): return str(s)
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            .replace("\n", "<br>"))


# ------------------------------
# 2. Citation Analyzer Class (Same Logic)
# ------------------------------

class CitationAnalyzer:
    def __init__(self):
        self.supported_types = ["Figure", "Table", "Box", "Exhibit", "Appendix", "Case Study"]
        self.regex_patterns = self._setup_regex_patterns()

    def _setup_regex_patterns(self) -> Dict[str, re.Pattern]:
        patterns = {}
        patterns['single'] = re.compile(
            r'(?:\(|\b)(Figures?|Figs?\.?|Tables?|Tabs?\.?|Box(?:es)?|BX|Exhibits?|Appendix|Appendices|Case\s+Stud(?:y|ies))\.?\s*([0-9]+(?:[.\-][0-9]+)*)([A-Za-z]?)(?:\)|\b)',
            re.IGNORECASE
        )
        patterns['range'] = re.compile(
            r'(?:\(|\b)(Figures?|Figs?\.?|Tables?|Tabs?\.?|Boxes?|Exhibits?|Appendices?|Case\s+Studies?)\.?\s+([0-9]+(?:[\.\-][0-9]+)*)([A-Za-z]?)(?:\s+(?:to|through)\s+|\s*[\u2013\u2014]\s*|\s+-\s+)([0-9]+(?:[\.\-][0-9]+)*)([A-Za-z]?)(?:\)|\b)',
            re.IGNORECASE
        )
        patterns['and'] = re.compile(
            r'(?:\(|\b)(Figures?|Figs?\.?|Tables?|Tabs?\.?|Boxes?|Exhibits?|Appendices?|Case\s+Studies?)\.?\s+([0-9]+(?:[\.\-][0-9]+)*)([A-Za-z]?)\s+(?:and|&)\s*([0-9]+(?:[\.\-][0-9]+)*)([A-Za-z]?)(?:\)|\b)',
            re.IGNORECASE
        )
        return patterns

    def normalize_for_regex(self, text: str) -> str:
        # Keep en-dash/em-dash distinct from hyphen so range detection stays correct.
        # Only normalize non-breaking space to regular space.
        text = text.replace('\xa0', ' ')
        return text

    def normalize_type(self, label: str) -> str:
        if not label:
            return "Figure"
        lbl = label.lower()
        if lbl.startswith('fig'):
            return "Figure"
        if lbl.startswith('tab'):
            return "Table"
        if lbl.startswith('box') or lbl.startswith('bx'):
            return "Box"
        if lbl.startswith('exhibit'):
            return "Exhibit"
        if lbl.startswith('appendix'):
            return "Appendix"
        if lbl.startswith('case'):
            return "Case Study"
        return "Figure"

    def normalize_fig_number(self, fig_ref: str) -> str:
        if not fig_ref:
            return ""
        fig_ref = fig_ref.strip()
        fig_ref = fig_ref.replace('--', '-').replace('\u2013', '-').replace('\u2014', '-')
        for ch in ['[', ']', '°']:
            fig_ref = fig_ref.replace(ch, '')
        m = re.search(r'([0-9]+(?:[.\-][0-9]+)*)([A-Za-z]?)', fig_ref)
        if m:
            base = m.group(1).replace('-', '.')
            suffix = m.group(2)
            if base.endswith('.'):
                base = base[:-1]
            return base + suffix
        return fig_ref

    def is_caption_paragraph(self, text: str, style_name: str = "") -> bool:
        # 1) Check explicit style match (if provided)
        if style_name:
            s_low = style_name.strip().lower()
            # User-requested styles: FIG-LEG, FGC, T1, TT, FigureLegend, TableCaption, etc.
            if s_low in ['fig-leg', 'fgc', 't1', 'tt', 'figurelegend', 'tablecaption', 'cs-ttl','nbx1-num','nbx1-ttl','nbx2-num','nbx2-ttl', 'exhibitcaption']:
                return True

        t_norm = self.normalize_for_regex(text.strip())
        t_norm = re.sub(r'^(?:<[^>]*>\s*)+', '', t_norm)  # local: ignore leading tags for caption match
        if not t_norm:
            return False
        if len(t_norm.splitlines()) > 7:
            return False

        # Match label + number, followed by optional caption text
        # e.g., "Figure 12.1. Text..." or "Figure 12.2"
        match = re.match(r'(?i)^(figure|fig\.|table|tab\.|box|exhibit|appendix|case\s+study)\s+([0-9]+(?:[.\-][0-9]+)*[a-zA-Z]?)(.*)', t_norm)
        if match:
            remainder = match.group(3).strip()
            
            # If the remainder doesn't contain any alphanumeric characters (e.g., it's empty or just "."), 
            # it lacks actual caption text. We return False so it gets treated as a citation, 
            # which correctly triggers a "Missing Caption" error in the dashboard.
            if not re.search(r'[A-Za-z0-9]', remainder):
                return False
                
            # Remove leading punctuation and spaces to check the first actual alphanumeric character
            first_word_char = re.sub(r'^[\W_]+', '', remainder)
            
            # If the text after the number starts with a lowercase letter, 
            # it's likely a body text sentence referencing the figure (e.g., " shows that...")
            if first_word_char and first_word_char[0].islower():
                return False
                
            return True
            
        return False

    def analyze_document_citations(self, document_content: List[Tuple[str, int, bool]]) -> Dict[str, Any]:
        dict_types = {t: {"Caption": {}, "Citation": {}, "CaptionPage": {}, "CitationPage": {}} for t in self.supported_types}

        for text, page_no, is_caption in document_content:
            txt = self.normalize_for_regex(text)
            txt = re.sub(r'^(?:<[^>]*>\s*)+', '', txt)  # local: ignore leading tags for citation match

            # Determine the boundary of the caption label so subsequent matches in the 
            # same paragraph are treated as citations rather than captions.
            label_boundary = 60
            if is_caption:
                match_boundary = re.search(r'[\.\:\-\u2013\u2014]\s', txt)
                if match_boundary and match_boundary.start() < 60:
                    label_boundary = match_boundary.end() + 5

            for m in self.regex_patterns['range'].finditer(txt):
                is_match_caption = is_caption and m.start() <= label_boundary
                label = self.normalize_type(m.group(1))
                start_num = self.normalize_fig_number(m.group(2))
                end_num = self.normalize_fig_number(m.group(4))
                try:
                    sp = start_num.split('.')
                    ep = end_num.split('.')
                    if int(sp[0]) == int(ep[0]) and len(sp) > 1 and len(ep) > 1:
                        start_minor = int(sp[1])
                        end_minor = int(ep[1])
                        for n in range(start_minor, end_minor + 1):
                            item_id = f"{label} {sp[0]}.{n}"
                            self._store(dict_types, label, item_id, page_no, is_match_caption)
                    else:
                        self._store(dict_types, label, f"{label} {start_num}", page_no, is_match_caption)
                        self._store(dict_types, label, f"{label} {end_num}", page_no, is_match_caption)
                except Exception:
                    self._store(dict_types, label, f"{label} {start_num}", page_no, is_match_caption)
                    self._store(dict_types, label, f"{label} {end_num}", page_no, is_match_caption)

            for m in self.regex_patterns['and'].finditer(txt):
                is_match_caption = is_caption and m.start() <= label_boundary
                label = self.normalize_type(m.group(1))
                first_num = self.normalize_fig_number(m.group(2))
                second_num = self.normalize_fig_number(m.group(4))
                self._store(dict_types, label, f"{label} {first_num}", page_no, is_match_caption)
                self._store(dict_types, label, f"{label} {second_num}", page_no, is_match_caption)

            for m in self.regex_patterns['single'].finditer(txt):
                is_match_caption = is_caption and m.start() <= label_boundary
                label = self.normalize_type(m.group(1))
                main_no = m.group(2)
                suffix = m.group(3) or ""
                item_id = f"{label} {self.normalize_fig_number(main_no + suffix)}"
                self._store(dict_types, label, item_id, page_no, is_match_caption)

        return dict_types

    def _store(self, dict_types, label, item_id, page_no, is_caption):
        tdict = dict_types.get(label)
        if tdict is None:
            return
        if is_caption:
            if item_id not in tdict['Caption']:
                tdict['Caption'][item_id] = True
                tdict['CaptionPage'][item_id] = page_no
        else:
            if item_id not in tdict['Citation']:
                tdict['Citation'][item_id] = True
                tdict['CitationPage'][item_id] = page_no


# ------------------------------
# Box Tag Linker
# ------------------------------
_NBX_TYPE_DEF_RE  = re.compile(
    r'<BX_TYPE>\s*Box\s+(\d+[\.\-]\d+)[^<]*<BX_TTL>(.*)', re.IGNORECASE
)
_NBX_TTL_SAME_RE  = re.compile(                         # same-line legacy: <BXN.N> <NBX-TTL>Title
    r'<BX(\d+[\.\-]\d+)>\s*<NBX-TTL>(.*)', re.IGNORECASE
)
_NBX_TTL_NEXT_RE  = re.compile(r'^<NBX-TTL>(.*)', re.IGNORECASE)  # next-line (chap12)
_BX_TAG_ONLY_RE   = re.compile(r'^<BX(\d+[\.\-]\d+)>$', re.IGNORECASE)  # standalone tag
_BX_TAG_RE        = re.compile(r'<BX(\d+[\.\-]\d+)>',   re.IGNORECASE)
_BX_PLAIN_DEF_RE  = re.compile(                         # ch007: "Box 7-1\u2003Title" at para start
    r'^Box\s+(\d+[\.\-]\d+)\u2003\s*(\S.*)', re.IGNORECASE  # em-space required — reliable definition signal
)
_BX_PLAIN_SPC_RE  = re.compile(                         # fallback: "Box 7-1 Text" with regular space
    r'^Box\s+(\d+[\.\-]\d+)\s+(\S.*)', re.IGNORECASE
)
_BX_CAPTION_WORD_LIMIT = 15                             # ≤ 15 words → caption title; > 15 words → body prose
_BX_TEXT_RE       = re.compile(r'\bBox(?:es)?\s+(\d+[\.\-]\d+)\b', re.IGNORECASE)


class BoxTagLinker:
    def __init__(self, chapter_number: Optional[str] = None):
        # Normalize to the leading digit(s) only: "Chapter 12" → "12", "7" → "7"
        raw = str(chapter_number).strip() if chapter_number else ""
        _m = re.search(r'\d+', raw)
        self.chapter_number: Optional[str] = _m.group() if _m else (raw or None)
        self.citations:     Dict[str, List[int]] = {}
        self.definitions:   Dict[str, Dict[str, Any]] = {}
        self.cross_chapter: Dict[str, List[int]] = {}
        self.errors:        List[str] = []

    def _norm(self, raw: str) -> str:
        return raw.replace('.', '-')

    def _is_cross_chapter(self, norm_id: str) -> bool:
        if not self.chapter_number:
            return False
        first_part = norm_id.split('-')[0]
        return first_part != self.chapter_number

    def scan(self, paragraphs: list) -> None:
        # paragraphs: list of (text, page_no, is_caption, ...)
        texts = [(item[0].strip(), item[1]) for item in paragraphs]
        skip_next = False
        for line_no, (s, page_no) in enumerate(texts):
            if skip_next:
                skip_next = False
                continue
            # Priority 1: <BX_TYPE>Box N-N <BX_TTL>Title  (PPD same-line format)
            m = _NBX_TYPE_DEF_RE.search(s)
            if m:
                self._store_def(self._norm(m.group(1)), m.group(2).strip(), line_no, page_no)
                continue
            # Priority 2: <BXN.N> standalone paragraph → next line is <NBX-TTL>Title (chap12)
            m = _BX_TAG_ONLY_RE.match(s)
            if m:
                next_s = texts[line_no + 1][0] if line_no + 1 < len(texts) else ""
                nm = _NBX_TTL_NEXT_RE.match(next_s)
                if nm:
                    self._store_def(self._norm(m.group(1)), nm.group(1).strip(), line_no, page_no)
                    skip_next = True
                    continue
                # No NBX-TTL follows → body-text placeholder citation
                self._store_citation(self._norm(m.group(1)), line_no)
                continue
            # Priority 3: <BXN.N> <NBX-TTL>Title same-line (legacy)
            m = _NBX_TTL_SAME_RE.search(s)
            if m:
                self._store_def(self._norm(m.group(1)), m.group(2).strip(), line_no, page_no)
                continue
            # Priority 4: Box N-N\u2003Title at para start — em-space is definitive (ch007)
            m = _BX_PLAIN_DEF_RE.match(s)
            if m:
                self._store_def(self._norm(m.group(1)), m.group(2).strip(), line_no, page_no)
                continue
            # Priority 4b: Box N-N[space]Text — short paragraph = caption title; long = body prose
            m = _BX_PLAIN_SPC_RE.match(s)
            if m:
                nid  = self._norm(m.group(1))
                rest = m.group(2).strip()
                if len(rest.split()) <= _BX_CAPTION_WORD_LIMIT:
                    self._store_def(nid, rest, line_no, page_no)
                else:
                    self._store_citation(nid, line_no)
                continue
            # Priority 5 & 6: citations — tag then plain-text
            for m in _BX_TAG_RE.finditer(s):
                self._store_citation(self._norm(m.group(1)), line_no)
            for m in _BX_TEXT_RE.finditer(s):
                self._store_citation(self._norm(m.group(1)), line_no)

    def _store_def(self, nid: str, caption: str, line_no: int, page_no: int) -> None:
        if nid in self.definitions:
            prev = self.definitions[nid]['line']
            self.errors.append(f"Duplicate box definition: Box {nid} (lines {prev} and {line_no})")
        else:
            self.definitions[nid] = {"caption": caption, "line": line_no, "page": page_no}

    def _store_citation(self, nid: str, line_no: int) -> None:
        if self._is_cross_chapter(nid):
            if nid not in self.cross_chapter:
                self.cross_chapter[nid] = []
            self.cross_chapter[nid].append(line_no)
        else:
            if nid not in self.citations:
                self.citations[nid] = []
            self.citations[nid].append(line_no)

    def validate(self):
        for nid in self.citations:
            if nid not in self.definitions:
                self.errors.append(f"Missing caption for Box {nid}")
        for nid in self.definitions:
            if nid not in self.citations:
                self.errors.append(f"Orphan box — no in-chapter reference: Box {nid}")

    def results(self):
        def sort_key(x):
            return [int(p) for p in x.replace('-', '.').split('.') if p.isdigit()]
        all_ids = sorted(set(self.citations) | set(self.definitions), key=sort_key)
        rows = []
        for nid in all_ids:
            cit  = nid in self.citations
            defn = self.definitions.get(nid)
            status = "Matched" if (cit and defn) else ("Missing Caption" if cit else "Orphan Box")
            rows.append({
                "box_id": f"Box {nid}", "citation_found": cit,
                "caption_found": bool(defn),
                "caption_text": defn["caption"] if defn else "", "status": status,
            })
        for nid in sorted(self.cross_chapter, key=sort_key):
            rows.append({
                "box_id": f"Box {nid}", "citation_found": True,
                "caption_found": False, "caption_text": "", "status": "Cross-Chapter Ref",
            })
        return rows

    def build_html(self):
        rows = self.results()
        if not rows and not self.errors:
            return ""
        icon = {
            "Matched": "✅ Matched", "Missing Caption": "⚠️ Missing Caption",
            "Orphan Box": "⚠️ Orphan Box", "Cross-Chapter Ref": "ℹ️ Cross-Chapter Ref",
        }
        lines = [
            '<h3>Box Citation ↔ Caption Mapping</h3>',
            '<table><thead><tr><th>Box ID</th><th>Cited in Text</th>'
            '<th>Caption Found</th><th>Status</th><th>Caption Title</th></tr></thead><tbody>',
        ]
        for r in rows:
            if r["status"] == "Matched":
                continue
            lines.append(
                f'<tr><td>{r["box_id"]}</td>'
                f'<td>{"Yes" if r["citation_found"] else "No"}</td>'
                f'<td>{"Yes" if r["caption_found"] else "No"}</td>'
                f'<td>{icon.get(r["status"], r["status"])}</td>'
                f'<td>{r["caption_text"]}</td></tr>'
            )
        lines.append('</tbody></table>')
        if self.errors:
            lines.append('<ul style="color:red">')
            lines += [f'<li>{e}</li>' for e in self.errors]
            lines.append('</ul>')
        return "\n".join(lines)


def build_element_mapping_html(
    dict_types: Dict[str, Any],
    type_key: str,
    chapter_number: str = "",
) -> str:
    # Mapping table docstring omitted due to strange byte issues
    data: Dict[str, Any] = dict_types.get(type_key) or {}
    if not data:
        return ""

    captions:  Dict[str, Any] = data.get("Caption",     {})
    citations: Dict[str, Any] = data.get("Citation",    {})
    cap_pages: Dict[str, Any] = data.get("CaptionPage", {})
    cit_pages: Dict[str, Any] = data.get("CitationPage",{})

    _cm = re.search(r'\d+', chapter_number) if chapter_number else None
    ch_digit = _cm.group() if _cm else ""

    def _id_prefix(label: str) -> str:
        m = re.search(r'(\d+)[.\-]\d+', label)
        return m.group(1) if m else ""

    def _norm(label: str) -> str:
        return re.sub(r'[.\-]', '-', label.strip().lower())

    all_ids: List[str] = sorted(
        set(captions.keys()) | set(citations.keys()),
        key=lambda x: [int(d) for d in re.findall(r'\d+', str(x))]
    )

    icon = {
        "Matched":           "✅ Matched",
        "Missing Caption":   "⚠️ Missing Caption",
        "Orphan":            "⚠️ Missing citation",
        "Cross-Chapter Ref": "ℹ️ Cross-Chapter Ref",
    }

    rows: List[Dict[str, Any]] = []
    cross_ids: List[str] = []
    for label in all_ids:
        prefix = _id_prefix(label)
        if ch_digit and prefix and prefix != ch_digit:
            cross_ids.append(label)
            continue
        norm = _norm(label)
        cap_found = any(_norm(k) == norm for k in captions)
        cit_found = any(_norm(k) == norm for k in citations)
        status = ("Matched"         if cap_found and cit_found else
                  "Missing Caption" if cit_found               else "Orphan")
        rows.append({
            "label":     label,
            "cit_found": cit_found,
            "cap_found": cap_found,
            "status":    status,
            "cit_page":  cit_pages.get(label, ""),
            "cap_page":  cap_pages.get(label, ""),
        })
    for label in cross_ids:
        rows.append({
            "label":     label,
            "cit_found": label in citations,
            "cap_found": label in captions,
            "status":    "Cross-Chapter Ref",
            "cit_page":  cit_pages.get(label, ""),
            "cap_page":  cap_pages.get(label, ""),
        })

    if not rows:
        return ""

    lines = [
        f'<h3>{type_key} Citation \u2194 Caption Mapping</h3>',
        '<table><thead><tr>'
        f'<th>{type_key} ID</th>'
        '<th>Cited in Text</th><th>Caption Found</th>'
        '<th>Status</th><th>Citation Page</th><th>Caption Page</th>'
        '</tr></thead><tbody>',
    ]
    for r in rows:
        if r["status"] == "Matched":
            continue
        lines.append(
            f'<tr>'
            f'<td>{r["label"]}</td>'
            f'<td>{"Yes" if r["cit_found"] else "No"}</td>'
            f'<td>{"Yes" if r["cap_found"] else "No"}</td>'
            f'<td>{icon.get(r["status"], r["status"])}</td>'
            f'<td>{r["cit_page"]}</td>'
            f'<td>{r["cap_page"]}</td>'
            f'</tr>'
        )
    lines.append('</tbody></table>')
    return "\n".join(lines)


def build_detailed_summary_table(
    dict_types: dict,
    figure_count: int,
    table_count: int,
    footnote_count: int,
    endnote_count: int,
    fmt_content: str,
    spec_content: str,
    comment_content: str,
    ref_count: int = 0,
    unnumbered_counts: dict = None,
    chapter_number: str = "",
    box_linker: Optional["BoxTagLinker"] = None,
) -> str:
    # (Implementation identical to word_analyzer.py, omitted for brevity but logic is same)
    # Re-using the logic from the original file since it's pure string manipulation
    def count_items(section_html: str, token: str) -> int:
        return section_html.lower().count(token.lower())

    def extract_num(item: str) -> str:
        parts = item.strip().split()
        return parts[-1] if parts else item

    def format_num_list(items: List[str]) -> str:
        def _num_key(s: str):
            return [int(x) for x in re.findall(r'\d+', s)]
        nums = sorted(set(extract_num(i) for i in items), key=_num_key)
        if len(nums) == 1:
            return nums[0]
        return ", ".join(nums[:-1]) + f" and {nums[-1]}"

    def _chap_digit(chapter_number: str) -> str:
        m = re.search(r'\d+', chapter_number)
        return m.group() if m else ""

    def _item_chap(item: str) -> str:
        m = re.search(r'(\d+)', item.strip())
        return m.group(1) if m else ""

    def _format_action(verb: str, kind: str, items: List[str],
                       chapter_number: str = "") -> str:
        if not items:
            return ""
        n = len(items)
        cn = _chap_digit(chapter_number) if chapter_number else ""
        if cn:
            this_ch  = [i for i in items if _item_chap(i) == cn]
            other_ch = [i for i in items if _item_chap(i) != cn]
            parts = []
            if this_ch:
                parts.append(f"Chapter ({cn}): {format_num_list(this_ch)}.")
            if other_ch:
                parts.append(f"Other chapters: {format_num_list(other_ch)}.")
            split_text = "<br>&nbsp;&nbsp;".join(parts)
            return f"Missing {n} {kind}(s): {verb}:<br>&nbsp;&nbsp;{split_text}"
        else:
            return f"Missing {n} {kind}(s): {verb} {format_num_list(items)}."

    def build_action_text(miss_cap_items: List[str],
                          miss_cit_items: List[str], chapter_number: str = "") -> str:
        cap_icon = "<i class='fas fa-times-circle' style='color:#e74c3c;'></i> "
        cit_icon = "<i class='fas fa-exclamation-triangle' style='color:#f39c12;'></i> "
        parts = []
        cap_text = _format_action("Provide captions for", "caption", miss_cap_items, chapter_number)
        cit_text = _format_action("Insert citations for", "citation", miss_cit_items, chapter_number)
        if cap_text:
            parts.append(cap_icon + cap_text)
        if cit_text:
            parts.append(cit_icon + cit_text)
        return "<br>".join(parts) if parts else "No action required"

    def build_progress_row(title: str, cap_cnt: int, cit_cnt: int, miss_cap: int, miss_cit: int,
                           action_text: str = "No action required",
                           miss_cap_items: Optional[List[str]] = None,
                           miss_cit_items: Optional[List[str]] = None,
                           chapter_number: str = "",
                           tab_target: str = "citations") -> str:
        miss_cap_items = miss_cap_items or []
        miss_cit_items = miss_cit_items or []
        total = max(cap_cnt, cit_cnt)
        complete_pct = round(((total - miss_cap - miss_cit) / total * 100), 1) if total else 0
        html = (
            f"<tr class='summary-table-row'>\n"
            f"  <td style='cursor:pointer;' onclick=\"showTabFromRow('{tab_target}', this.closest('tr'))\"><strong>{title}</strong></td>\n"
            f"  <td>{total}</td>\n"
            f"  <td>\n"
            f"    <div style='display:flex;align-items:center;gap:10px;'>\n"
            f"      <div style='width:100px;height:20px;background:#f0f0f0;border-radius:10px;overflow:hidden;'>\n"
            f"        <div style='width:0%;height:100%;background:linear-gradient(90deg,#27ae60,#2ecc71);transition:width 1s ease-in-out;' data-w='{complete_pct}'></div>\n"
            f"      </div>\n"
            f"      <span style='font-size:12px;color:#27ae60;'>{complete_pct}% Complete</span>\n"
            f"    </div>\n"
            f"  </td>\n"
            f"  <td><i class='fas fa-check-circle' style='color:#27ae60;'></i> {cit_cnt} citation(s)</td>\n"
            f"  <td><i class='fas fa-check-circle' style='color:#27ae60;'></i> {cap_cnt} caption(s)</td>\n"
            f"  <td>{action_text}</td>\n"
            f"</tr>\n"
        )
        return html

    # def build_critical_issues_block(fig_miss_cap, fig_miss_cit, tab_miss_cap, tab_miss_cit, fmt_count):
    #     html = """
    #     <div style='background:#fff3cd;border:1px solid #ffeaa7;border-radius:10px;padding:20px;margin-top:20px;'>
    #       <h3 style='color:#856404;margin-bottom:10px;cursor:pointer;user-select:none;'
    #           onclick="var ul=this.nextElementSibling;ul.style.display=ul.style.display==='none'?'block':'none';this.querySelector('.ci-arrow').textContent=ul.style.display==='none'?'▶':'▼';">
    #         <i class='fas fa-exclamation-triangle'></i> Critical Issues Requiring Attention
    #         <span class='ci-arrow' style='float:right;font-size:14px;'>▼</span>
    #       </h3>
    #       <ul style='margin:0;padding-left:20px;color:#856404;'>
    #     """
    #     if (fig_miss_cit + tab_miss_cit) > 0:
    #         html += f"<li><strong>{fig_miss_cit + tab_miss_cit} Missing Citations:</strong> Check missing citations in Citations tab</li>"
    #     if (fig_miss_cap + tab_miss_cap) > 0:
    #         html += f"<li><strong>{fig_miss_cap + tab_miss_cap} Missing Captions:</strong> Check missing captions in Citations tab</li>"
    #     if fmt_count > 0:
    #         html += f"<li><strong>{fmt_count} Formatting Issues:</strong> See Formatting tab</li>"
    #     html += "</ul></div>"
    #     return html

    fmt_count = count_items(fmt_content, "<tr><td>")
    spec_count = count_items(spec_content, "<tr><td>")
    comment_count_val = count_items(comment_content, "<tr><td>")

    global_stats = {
        "fmt_issues": fmt_count,
        "missing_citations": 0,
        "missing_captions": 0,
        "fig_missing_cap": 0,
        "fig_missing_cit": 0,
        "tab_missing_cap": 0,
        "tab_missing_cit": 0,
        "box_missing_cap": 0,
        "box_missing_cit": 0
    }

    fig_cap = fig_cit = fig_miss_cap = fig_miss_cit = 0
    tab_cap = tab_cit = tab_miss_cap = tab_miss_cit = 0
    fig_miss_cap_items: List[str] = []
    fig_miss_cit_items: List[str] = []
    tab_miss_cap_items: List[str] = []
    tab_miss_cit_items: List[str] = []

    def normalize_ref(ref: str) -> str:
        return ref.replace("-", ".").strip().lower()

    for type_key in dict_types.keys():
        if type_key == "Figure":
            fig_cap = len(dict_types[type_key]["Caption"])
            fig_cit = len(dict_types[type_key]["Citation"])
            for k in dict_types[type_key]["Citation"]:
                norm = normalize_ref(k)
                if not any(normalize_ref(x) == norm for x in dict_types[type_key]["Caption"]):
                    fig_miss_cap += 1
                    fig_miss_cap_items.append(str(k))  # type: ignore[arg-type]
            for k in dict_types[type_key]["Caption"]:
                norm = normalize_ref(k)
                if not any(normalize_ref(x) == norm for x in dict_types[type_key]["Citation"]):
                    fig_miss_cit += 1
                    fig_miss_cit_items.append(str(k))  # type: ignore[arg-type]
        elif type_key == "Table":
            tab_cap = len(dict_types[type_key]["Caption"])
            tab_cit = len(dict_types[type_key]["Citation"])
            for k in dict_types[type_key]["Citation"]:
                norm = normalize_ref(k)
                if not any(normalize_ref(x) == norm for x in dict_types[type_key]["Caption"]):
                    tab_miss_cap += 1
                    tab_miss_cap_items.append(str(k))  # type: ignore[arg-type]
            for k in dict_types[type_key]["Caption"]:
                norm = normalize_ref(k)
                if not any(normalize_ref(x) == norm for x in dict_types[type_key]["Citation"]):
                    tab_miss_cit += 1
                    tab_miss_cit_items.append(str(k))  # type: ignore[arg-type]

    html = """
    <div class='header'>
      <div class='section-title'><i class='fas fa-chart-pie'></i> Analysis Summary</div>
      <table style='margin-bottom:20px;width:100%;border-collapse:collapse;'>
        <thead>
          <tr>
            <th>Element Type</th>
            <th>Totals</th>
            <th>Status Overview</th>
            <th>Citations Status</th>
            <th>Captions Status</th>
            <th>Recommended Actions</th>
          </tr>
        </thead><tbody>
    """

    html += build_progress_row("Figures", fig_cap, fig_cit, fig_miss_cap, fig_miss_cit,
                               build_action_text(fig_miss_cap_items, fig_miss_cit_items, chapter_number),
                               fig_miss_cap_items, fig_miss_cit_items, chapter_number)
    html += build_progress_row("Tables", tab_cap, tab_cit, tab_miss_cap, tab_miss_cit,
                               build_action_text(tab_miss_cap_items, tab_miss_cit_items, chapter_number),
                               tab_miss_cap_items, tab_miss_cit_items, chapter_number)

    global_stats["missing_captions"] += fig_miss_cap + tab_miss_cap
    global_stats["missing_citations"] += fig_miss_cit + tab_miss_cit
    global_stats["fig_missing_cap"] = fig_miss_cap
    global_stats["fig_missing_cit"] = fig_miss_cit
    global_stats["tab_missing_cap"] = tab_miss_cap
    global_stats["tab_missing_cit"] = tab_miss_cit

    # Box row — use BoxTagLinker data when available (accurate caption detection via tags)
    if box_linker is not None:
        bx_cit_cnt = len(box_linker.citations)
        bx_cap_cnt = len(box_linker.definitions)
        bx_miss_cap = [nid for nid in box_linker.citations if nid not in box_linker.definitions]
        bx_miss_cit = [nid for nid in box_linker.definitions if nid not in box_linker.citations]
        global_stats["missing_captions"] += len(bx_miss_cap)
        global_stats["missing_citations"] += len(bx_miss_cit)
        global_stats["box_missing_cap"] = len(bx_miss_cap)
        global_stats["box_missing_cit"] = len(bx_miss_cit)
        html += build_progress_row(
            "Boxes", bx_cap_cnt, bx_cit_cnt,
            len(bx_miss_cap), len(bx_miss_cit),
            build_action_text(bx_miss_cap, bx_miss_cit, chapter_number),
            bx_miss_cap, bx_miss_cit, chapter_number
        )

    # Additional element types (Exhibit, Appendix, Case Study) — Box handled above via BoxTagLinker
    other_types: List[str] = [str(k) for k in dict_types.keys() if k not in ("Figure", "Table")]  # type: ignore[misc]
    for type_key in other_types:
        if box_linker is not None and type_key == "Box":
            continue  # already rendered above via BoxTagLinker
        o_cap = len(dict_types[type_key]["Caption"])
        o_cit = len(dict_types[type_key]["Citation"])
        o_miss_cap = 0
        o_miss_cit = 0
        o_miss_cap_items: List[str] = []
        o_miss_cit_items: List[str] = []
        for k in dict_types[type_key]["Citation"]:
            norm = normalize_ref(k)
            if not any(normalize_ref(x) == norm for x in dict_types[type_key]["Caption"]):
                o_miss_cap += 1
                o_miss_cap_items.append(str(k))  # type: ignore[arg-type]
        for k in dict_types[type_key]["Caption"]:
            norm = normalize_ref(k)
            if not any(normalize_ref(x) == norm for x in dict_types[type_key]["Citation"]):
                o_miss_cit += 1
                o_miss_cit_items.append(str(k))  # type: ignore[arg-type]
        
        global_stats["missing_captions"] += o_miss_cap
        global_stats["missing_citations"] += o_miss_cit

        if o_cap > 0 or o_cit > 0:
            html += build_progress_row(str(type_key) + "s", o_cap, o_cit, o_miss_cap, o_miss_cit,
                                       build_action_text(o_miss_cap_items, o_miss_cit_items, chapter_number),
                                       o_miss_cap_items, o_miss_cit_items, chapter_number)

    html += f"""
    <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('special-chars', this.closest('tr'))"><strong>Special Characters</strong></td><td>{spec_count}</td>
        <td colspan='3'><a href='javascript:void(0);' onclick="showTab('special-chars');"
        style='color:#667eea;text-decoration:underline;'>Review multilingual symbols</a></td>
        <td>Review unusual characters</td></tr>

    <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('formatting', this.closest('tr'))"><strong>Formatting Issues</strong></td><td>{fmt_count}</td>
        <td colspan='3'><a href='javascript:void(0);' onclick="showTab('formatting');"
        style='color:#f39c12;text-decoration:underline;'>View formatting issues</a></td>
        <td>Review formatting anomalies</td></tr>

    <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('comments', this.closest('tr'))"><strong>Comments</strong></td><td>{comment_count_val}</td>
        <td colspan='3'><a href='javascript:void(0);' onclick="showTab('comments');"
        style='color:#3498db;text-decoration:underline;'>Review editor comments</a></td>
        <td>Review highlighted feedback</td></tr>

    <tr class='summary-table-row'><td><strong>Notes</strong></td><td>{footnote_count + endnote_count}</td>
        <td colspan='3'>{footnote_count} Footnotes, {endnote_count} Endnotes</td>
        <td>No action required</td></tr>
    """

    if figure_count > 0:
        html += f"""
        <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('media', this.closest('tr'))"><strong>Images</strong></td><td>{figure_count}</td>
        <td colspan='3'><a href='javascript:void(0);' onclick="showTab('media');"
        style='color:#27ae60;text-decoration:underline;'><i class='fas fa-check-circle'></i> {figure_count} image(s) detected</a></td>
        <td>No action required</td></tr>
        """
    else:
        html += """
        <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('media', this.closest('tr'))"><strong>Images</strong></td><td>0</td>
        <td colspan='3'><span style='color:#e67e22;'><i class='fas fa-exclamation-triangle'></i> No images detected</span></td>
        <td>Check for missing image elements</td></tr>
        """

    # Reference count row
    if ref_count > 0:
        html += f"""
        <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('media', this.closest('tr'))"><strong>References</strong></td><td>{ref_count}</td>
        <td colspan='3'><span style='color:#27ae60;'><i class='fas fa-check-circle'></i> {ref_count} reference(s) detected</span></td>
        <td>No action required</td></tr>
        """

    # Unnumbered elements row
    if unnumbered_counts:
        u_figs         = unnumbered_counts.get("unnumbered_images",    0)
        u_tabs         = unnumbered_counts.get("unnumbered_tables",    0)
        u_boxes        = unnumbered_counts.get("unnumbered_boxes",     0)
        u_callouts     = unnumbered_counts.get("callouts",             0)
        u_eq_omml      = unnumbered_counts.get("equations_omml",       0)
        u_eq_mt        = unnumbered_counts.get("equations_mathtype",   0)
        u_placeholders = unnumbered_counts.get("image_placeholders",   0)
        u_total = u_figs + u_tabs + u_boxes + u_callouts + u_eq_omml + u_eq_mt + u_placeholders
        if u_total > 0:
            detail = ", ".join(filter(None, [
                f"{u_figs} fig(s)"                       if u_figs          else "",
                f"{u_tabs} table(s)"                     if u_tabs          else "",
                f"{u_boxes} box(es)"                     if u_boxes         else "",
                f"{u_callouts} callout(s)"               if u_callouts      else "",
                f"{u_eq_omml} OMML eq(s)"                if u_eq_omml       else "",
                f"{u_eq_mt} MathType eq(s)"              if u_eq_mt         else "",
                f"{u_placeholders} image placeholder(s)" if u_placeholders  else "",
            ]))
            html += f"""
            <tr class='summary-table-row'><td style='cursor:pointer;' onclick="showTabFromRow('unnumbered', this.closest('tr'))"><strong>Unnumbered Elements</strong></td><td>{u_total}</td>
            <td colspan='3'><span style='color:#e67e22;'><i class='fas fa-exclamation-triangle'></i> {detail}</span></td>
            <td>Add numbers/captions where needed</td></tr>
            """

    html += "</tbody></table>"
    # html += build_critical_issues_block(fig_miss_cap, fig_miss_cit, tab_miss_cap, tab_miss_cit, fmt_count)

    html += f"""
<style>
  .summary-row-active td {{ background:#f0f4ff !important; border-left:4px solid #667eea; }}
  .summary-table-row td:first-child:hover {{ background:#eef2ff; }}
</style>
<script>
(function(){{
  document.addEventListener('DOMContentLoaded', function(){{
    document.querySelectorAll('[data-w]').forEach(function(bar){{
      var w = bar.getAttribute('data-w');
      setTimeout(function(){{ bar.style.width = w + '%'; }}, 150);
    }});
  }});
}})();
</script>"""

    html += "</div>"

    return html, global_stats

def build_comments_html(comments: List[Tuple]):
    if not comments:
        return "<p>No comments found.</p>"
    html = "<table><thead><tr><th>#</th><th>Page</th><th>Author</th><th>Comment</th></tr></thead><tbody>"
    for i, (author, text, page) in enumerate(comments, start=1):
        html += f"<tr><td>{i}</td><td>{page}</td><td>{escape_html(author)}</td><td>{escape_html(text)}</td></tr>"
    html += "</tbody></table>"
    return html


def build_unnumbered_tab_html(unnumbered_counts: dict) -> str:
    if not unnumbered_counts:
        return "<p>No unnumbered elements data available.</p>"
    rows = [
        ("Figures",          unnumbered_counts.get("unnumbered_images",  0), "Images with no numbered Figure caption"),
        ("Tables",           unnumbered_counts.get("unnumbered_tables",  0), "Tables with no numbered Table caption"),
        ("Boxes",            unnumbered_counts.get("unnumbered_boxes",   0), "Box-style paragraphs with no numbered Box caption"),
        ("Callouts",         unnumbered_counts.get("callouts",           0), "Vague cross-references (e.g. 'see figure above', page refs)"),
        ("OMML Equations",   unnumbered_counts.get("equations_omml",     0), "Inline OMML math without numbered labels"),
        ("MathType Equations", unnumbered_counts.get("equations_mathtype", 0), "MathType math without numbered labels"),
    ]
    total = sum(r[1] for r in rows)
    html = "<table><thead><tr><th>Element Type</th><th>Count</th><th>Notes</th></tr></thead><tbody>"
    if total == 0:
        html += "<tr><td colspan='3'>No unnumbered elements found.</td></tr>"
    else:
        for label, count, notes in rows:
            if count > 0:
                html += f"<tr><td>{label}</td><td>{count}</td><td>{notes}</td></tr>"
    html += "</tbody></table>"
    return html


def build_export_highlight_html(paragraphs_full):
    highlights = []
    for t, p, is_cap, is_high in paragraphs_full:
        if is_high:
            highlights.append((t, p))
    if not highlights:
        return "<p>No highlighted paragraphs found.</p>"
    html = "<table><thead><tr><th>Highlighted Text</th><th>Page</th></tr></thead><tbody>"
    for t, p in highlights:
        html += f"<tr><td>{escape_html(t)}</td><td>{p}</td></tr>"
    html += "</tbody></table>"
    return html


# ------------------------------
# 3. New docx-based Implementations
# ------------------------------

def get_xml_comments(doc):
    """Parses word/comments.xml to extract comments."""
    comments = []
    try:
        # Access the comments part
        # doc.part.package.parts is a list of Part objects. We need to find the one with rel 'comments'
        for part in doc.part.package.parts:
            if part.partname.endswith('comments.xml'):
                comments_xml = part.blob
                root = etree.fromstring(comments_xml)
                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }
                for comment in root.findall('.//w:comment', namespaces):
                    author = comment.get(qn('w:author'), 'Unknown')
                    text_nodes = comment.findall('.//w:t', namespaces)
                    text = "".join([t.text for t in text_nodes if t.text])
                    # Page finding is hard in XML without layout engine, fallback to "See Context"
                    comments.append((author, text, "See Context"))
    except Exception:
        pass
    return comments

def get_xml_note_count(doc, note_type='footnotes'):
    """
    Counts note references in the main document body.
    This matches doc.Footnotes.Count in Word (visual count).
    """
    count = 0
    try:
        # Determine tag
        if note_type == 'footnotes':
            tag = qn('w:footnoteReference')
        else:
            tag = qn('w:endnoteReference')
            
        # Search in the main document element (Body)
        # Note: This counts footnotes in the main text.
        # If footnotes are in textboxes/headers, they might be missed here unless we scan those parts too.
        # But usually 'doc.Footnotes.Count' primarily reflects main story.
        elements = doc.element.findall(f'.//{tag}')
        count = len(elements)
        #print(f"[DEBUG] Found {count} {note_type} references in document body.")
        
    except Exception as e:
        #print(f"[DEBUG] Error counting {note_type}: {e}")
        pass
    return count

def extract_with_docx(doc_path: str, doc: Optional[Any] = None):
    """
    Robust extraction using python-docx + lxml.
    Returns: paragraphs, comments, img_count, footnotes, endnotes
    """
    if doc is None:
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"{doc_path} not found")
        doc = Document(doc_path)

    analyzer = CitationAnalyzer()

    # 1. Paragraphs (Text, Page, Caption, Highlighted)
    paragraphs = []

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue

        try:
            s_name = p.style.name
        except:
            s_name = ""
        is_caption = analyzer.is_caption_paragraph(text, style_name=s_name)

        # Check highlighting: if ANY run is highlighted
        is_highlighted = False
        for run in p.runs:
            if run.font.highlight_color:
                is_highlighted = True
                break

        paragraphs.append((text, i // 40 + 1, is_caption, is_highlighted))
        
    # 1b. Scan first row of each table for embedded captions (e.g. "Table 7-3 Title" as first cell)
    # python-docx excludes table cell paragraphs from doc.paragraphs, so we walk them separately.
    for tbl_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue
        first_row = table.rows[0]
        found_caption = False
        for cell in first_row.cells:
            for cp in cell.paragraphs:
                cell_text = cp.text.strip()
                if not cell_text:
                    continue
                try:
                    s_name = cp.style.name or ""
                except Exception:
                    s_name = ""
                # Strip leading markup tags before caption check
                clean_text = re.sub(r'^(?:<[^>]*>\s*)+', '', cell_text)
                if analyzer.is_caption_paragraph(clean_text, style_name=str(s_name)):
                    # Use last body paragraph's page as proxy — far closer than tbl_idx math
                    page_est = paragraphs[-1][1] if paragraphs else 1
                    paragraphs.append((cell_text, page_est, True, False))
                    found_caption = True
                    break
            if found_caption:
                break

    # 2. Comments (XML)
    comments = get_xml_comments(doc)

    # 3. Images (RELS)
    img_count = 0
    for rel in doc.part.rels.values():
         if "image" in rel.reltype:
             img_count += 1
             
    # 4. Footnotes/Endnotes (XML)
    footnotes = get_xml_note_count(doc, 'footnotes')
    endnotes = get_xml_note_count(doc, 'endnotes')
    
    return paragraphs, comments, img_count, footnotes, endnotes

def remove_tags_keep_formatting_docx(doc_path: str, doc: Optional[Any] = None):
    """
    Removes <tags> using regex on run text, preserving other formatting.
    """
    should_save = False
    if doc is None:
        if not os.path.exists(doc_path):
            return
        doc = Document(doc_path)
        should_save = True
        
    tag_cleaner = re.compile(r'<[^>]+>')
    
    modified = False
    
    for p in doc.paragraphs:
        for run in p.runs:
            if '<' in run.text and '>' in run.text:
                new_text = tag_cleaner.sub('', run.text)
                if new_text != run.text:
                    run.text = new_text
                    modified = True
                    
    # Also clean tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if '<' in run.text:
                             new_text = tag_cleaner.sub('', run.text)
                             if new_text != run.text:
                                 run.text = new_text
                                 modified = True

    if modified and should_save:
        doc.save(doc_path)
        
    return doc_path


def extract_unnumbered_image_markups(doc_path: str, doc: Optional[Any] = None, text_page_map: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """
    Scans paragraph text for image placeholder markup (e.g. <UNFIG 5-1>,
    <ch007_csimage001>, <insert Photo>) and returns a list of dicts:
        [{"markup": "<UNFIG 5-1>", "page": 12}, ...]
    Must be called BEFORE remove_tags_keep_formatting_docx() so tags are still present.
    """
    if not os.path.exists(doc_path):
        return []
    if doc is None:
        doc = Document(doc_path)
    if text_page_map is None:
        _, text_page_map = build_text_page_map(doc_path)

    results = []
    for para_idx, p in enumerate(doc.paragraphs):
        text = p.text
        matches = _UNIMG_MARKUP_RE.findall(text)
        if matches:
            page = _page_from_map(text_page_map, text, para_idx // 40 + 1)
            for m in matches:
                results.append({"markup": m.strip(), "page": page})
    return results


def build_unnumbered_image_markups_html(items: List[Dict[str, Any]]) -> str:
    """
    Renders a list from extract_unnumbered_image_markups() as an HTML table.
    Columns: # | Markup Text | Page
    """
    if not items:
        return "<p>No unnumbered image placeholders found.</p>"

    rows = []
    for i, item in enumerate(items, 1):
        markup_escaped = item["markup"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        rows.append(
            f'<tr><td style="text-align:center">{i}</td>'
            f'<td><code>{markup_escaped}</code></td>'
            f'<td style="text-align:center">{item["page"]}</td></tr>'
        )

    rows_html = "\n".join(rows)
    return (
        '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">'
        "<thead><tr>"
        '<th style="text-align:center">#</th>'
        "<th>Markup Text</th>"
        '<th style="text-align:center">Page</th>'
        "</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        "</table>"
    )


# ------------------------------
# PDF Conversion & Page Lookup
# ------------------------------

def convert_to_pdf(docx_path: str) -> str:
    """
    Convert a DOCX file to PDF. Returns the PDF path on success, or "" on failure.
    Uses LibreOffice subprocess, which works on Linux without Windows/Word dependencies.
    """
    import shutil
    pdf_path = str(Path(docx_path).with_suffix(".pdf"))
    
    # OPTIMIZATION: Check if the batch processor already created the PDF
    if os.path.exists(pdf_path):
        return pdf_path

    lo_cmd = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo_cmd and os.name == "nt" and os.path.exists(r"C:\Program Files\LibreOffice\program\soffice.exe"):
        lo_cmd = r"C:\Program Files\LibreOffice\program\soffice.exe"

    if not lo_cmd:
        print("DEBUG: LibreOffice not found in PATH. Conversion will fail.")
        return ""

    # Strategy 1: LibreOffice subprocess (Linux / macOS / Windows with LO installed)
    try:
        out_dir = str(Path(docx_path).parent)
        result = subprocess.run(
            [lo_cmd, "--headless", "--convert-to", "pdf", "--outdir", out_dir, os.path.abspath(docx_path)],
            timeout=60, capture_output=True
        )
        if result.returncode == 0 and os.path.exists(pdf_path):
            return pdf_path
    except Exception:
        pass

    return ""



def build_text_page_map(docx_path: str) -> Tuple[int, List[str]]:
    """
    Convert DOCX to PDF then build a list of normalized text for each page.
    Returns (total_pages, list_of_page_texts).
    Falls back to (0, []) if pdfplumber is unavailable or conversion fails.
    """
    if not HAS_PDFPLUMBER:
        return 0, []
    pdf_path = convert_to_pdf(docx_path)
    if not pdf_path:
        return 0, []
        
    page_texts: List[str] = []
    total_pages = 0
    try:
        import pdfplumber as _pdfplumber  # local import avoids unbound warning
        with _pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            for page in pdf.pages:
                raw = page.extract_text() or ""
                page_texts.append(_normalize_for_match(raw))
    except Exception:
        pass
    return total_pages, page_texts


def _page_from_map(text_page_map: Optional[List[str]], text: str, fallback: int) -> int:
    """Look up real page number by searching normalized text; fall back to estimated value."""
    if text_page_map:
        search_str = _normalize_for_match(text)[:60]
        if search_str:
            for page_idx, p_text in enumerate(text_page_map):
                if search_str in p_text:
                    return page_idx + 1
    return fallback


def generate_formatting_html(doc_path: str, used_word: bool = False,
                             text_page_map: Optional[List[str]] = None,
                             doc: Optional[Any] = None) -> str:
    """
    Scans for Strikethrough, Hidden, Section Breaks using python-docx.
    Ignores `used_word` flag as we are strictly python-docx now.
    """
    if doc is None:
        doc = Document(doc_path)
    rows = []
    
    # 1. Strikethrough & Hidden (Run level)
    for i, p in enumerate(doc.paragraphs):
        page = _page_from_map(text_page_map, p.text, i // 40 + 1)
        for run in p.runs:
            if run.font.strike or run.font.double_strike:
                rows.append(("Formatting", page, "Strikethrough", escape_html(run.text[:50])))
            # Hidden text (w:vanish)
            # python-docx exposes run.font.hidden
            if run.font.hidden:
                rows.append(("Formatting", page, "Hidden", escape_html(run.text[:50])))
                
    # 2. Section Breaks
    for i, section in enumerate(doc.sections):
        rows.append(("Formatting", "N/A", "Section Break", f"Section {i+1}"))

    html = "<table><thead><tr><th>Type</th><th>Page</th><th>Category</th><th>Details</th></tr></thead><tbody>"
    if rows:
        for r in rows:
            html += f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td></tr>"
    else:
        html += "<tr><td colspan='4'>No significant formatting issues found.</td></tr>"
    html += "</tbody></table>"
    return html

def generate_multilingual_html(doc_path: str,
                               text_page_map: Optional[List[str]] = None,
                               doc: Optional[Any] = None) -> str:
    """
    highlights multilingual chars and keywords using python-docx.
    Saves document if changes made.
    Returns HTML summary.
    """
    should_save = False
    if doc is None:
        doc = Document(doc_path)
        should_save = True
    modified = False
    page_map = defaultdict(set)
    
    keywords = [
        "Refer", "Insert", "Pick-up", "pickup", "See",
        "COMP", "AU", "AQ", "SPU", "Compositor",
        "Ph", "Photo", "video", "images"
    ]
    keyword_pattern = re.compile(r'\b(' + '|'.join(re.escape(k) for k in keywords) + r')\b\s+(\S+)', re.IGNORECASE)
    
    multilingual_ranges = [
        ("Chinese",      0x4E00, 0x9FFF),
        ("Greek",        0x0370, 0x03FF),
        ("Cyrillic",     0x0400, 0x04FF),
        ("Hebrew",       0x0590, 0x05FF),
        ("Arabic",       0x0600, 0x06FF),
        ("Devanagari",   0x0900, 0x097F),
        ("Japanese",     0x3040, 0x309F), 
        ("Korean",       0xAC00, 0xD7AF),
        ("Thai",         0x0E00, 0x0E7F),
    ]

    from docx.enum.text import WD_COLOR_INDEX

    for i, p in enumerate(doc.paragraphs):
        text = p.text
        if not text: continue
        page = _page_from_map(text_page_map, text, i // 40 + 1)

        # 1. Keywords
        for match in keyword_pattern.finditer(text):
            # Applying highlighting to specific sub-range in python-docx is hard 
            # because text is split across runs randomly.
            # Strategy: If keyword found, highlight the WHOLE RUN(s) containing it? 
            # Or simplified: verify if we can just highlight the paragraph for attention?
            # For strict correctness, we'd need to split runs. 
            # For this dashboard tool, we'll try to find the run containing the text and highlight it.
            for run in p.runs:
                if match.group(0) in run.text:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    modified = True
                    
        # 2. Multilingual
        for char in text:
            code = ord(char)
            for lang, low, high in multilingual_ranges:
                if low <= code <= high:
                    page_map[lang].add(page)
                    # Highlight runs containing this char
                    for run in p.runs:
                         if char in run.text:
                             run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                             modified = True
                    break

    if modified and should_save:
        doc.save(doc_path)

    html = "<table><thead><tr><th>Language/Type</th><th>Page</th></tr></thead><tbody>"
    for lang, pages in page_map.items():
        for p in sorted(pages):
            html += f"<tr><td>{lang}</td><td>{p}</td></tr>"
    if not page_map:
        html += "<tr><td colspan='2'>No multilingual characters found</td></tr>"
    html += "</tbody></table>"
    return html

# ------------------------------
# 4. Document Metadata & Extended Counts
# ------------------------------

_REFS_HEADING_RE = re.compile(
    r'^\s*('
    r'references?'                      # Reference / References
    r'|bibliographys?'                  # Bibliography / Bibliographys
    r'|notes\s+and\s+bibliography'      # Notes and Bibliography
    r'|works\s+citeds?'                 # Works Cited / Works Citeds
    r'|cited\s+works?'                  # Cited Work / Cited Works
    r'|literature\s+cited'              # Literature Cited (biology style)
    r'|reference\s+list'                # Reference List
    r'|selected\s+bibliography'         # Selected Bibliography
    r'|further\s+reading'               # Further Reading
    r')\s*:?\s*$',
    re.IGNORECASE
)
_MARKUP_TAG_RE   = re.compile(r'<[^>]+>')
_FIGURE_LEGENDS_RE = re.compile(r'^\s*(figure\s+legends?|list\s+of\s+(figures?|tables?|illustrations?))\s*:?\s*$', re.IGNORECASE)
_AUTHOR_STYLE_RE = re.compile(r'author|by.?line|^a[0-9]$', re.IGNORECASE)
_TITLE_STYLE_RE  = re.compile(r'heading\s*1|chapter\s*title|^ct$|^title$', re.IGNORECASE)
_CALLOUT_RE      = re.compile(r'\b(see\s+(figure|fig\.?|table|tab\.?|box)\s+(above|below|following|on\s+page))\b', re.IGNORECASE)
_PAGE_REF_RE     = re.compile(r'\(p\.?\s*\d+\)', re.IGNORECASE)
_UNIMG_MARKUP_RE = re.compile(
    # optional prefix like ch007_ before the keyword (e.g. <ch007_csimage001>, <ch007_unfigure002>)
    r'<\s*(?:[a-zA-Z0-9]+_)*'
    r'(?:unfig(?:ure)?'                          # <UNFIG 5-1>, <ch007_unfigure002
    r'|csimage'                                   # <ch007_csimage001>, <csimage...>
    r'|coimage'                                   # <ch007_COimageXXX>, <coimage...>
    r'|insert\s+(?:photo|unf(?:ig(?:ure)?)?'     # <insert Photo>, <Insert UNF Here>, <insert unfigure>
    r'|fig(?:ure)?|here)'
    r'|icon\s+here'                              # <ICON HERE>
    r'|unf\b)'                                   # standalone <UNF ...>
    r'[^>]*>?',
    re.IGNORECASE
)


def extract_chapter_metadata(doc_path: str, doc: Optional[Any] = None):
    """
    Returns (chapter_title: str, authors: str) extracted from the chapter opener.
    Supports production markup tags (<ct>=title, <cau>/<au>=author, <cn>=chapter num)
    and plain-text heuristics for documents without markup.
    """
    if not os.path.exists(doc_path):
        return "", ""
    if doc is None:
        doc = Document(doc_path)
    chapter_number = ""
    chapter_title = ""
    authors = []
    found_title = False

    _TAG_PREFIX_RE   = re.compile(r'^<([^>]+)>(.*)', re.DOTALL)
    _CHAP_NUM_RE     = re.compile(r'^(?:(?:chapter|ch\.?)\s*)?\d+\s*$', re.IGNORECASE)

    _CHAP_INLINE_RE  = re.compile(r'^((?:chapter|ch\.?)\s*\d+)\s+(.+)$', re.IGNORECASE)
    _TITLE_TAGS      = {'ct', 'chapter-title', 'chaptertitle', 'chap-title', 'chaptitle'}
    _AUTHOR_TAGS     = {'cau', 'au', 'author', 'byline', 'by-line', 'contrib'}
    _CHAP_NUM_TAGS   = {'cn', 'chapternum', 'chnum', 'cn1'}

    for p in doc.paragraphs[:4]:
        raw = p.text.strip()
        if not raw:
            continue
        try:
            s_name = p.style.name.strip()
        except:
            s_name = ""

        tag_match = _TAG_PREFIX_RE.match(raw)
        tag_name  = tag_match.group(1).lower().strip() if tag_match else ""
        text      = _MARKUP_TAG_RE.sub('', raw).strip()
        if not text:
            continue

        # --- Markup-tag based (highest priority) ---
        if tag_name in _TITLE_TAGS:
            chapter_title = text
            found_title = True
            continue
        if tag_name in _AUTHOR_TAGS:
            authors.append(text)
            continue
        if tag_name in _CHAP_NUM_TAGS:
            chapter_number = text
            continue
        # Any other markup tag after we already have title/authors → stop scanning
        if tag_name and (found_title or authors):
            break
        # Any other markup tag before title → skip this line
        if tag_name:
            continue

        # --- Style-based fallback ---
        if _TITLE_STYLE_RE.search(s_name):
            chapter_title = text
            found_title = True
            continue
        if _AUTHOR_STYLE_RE.search(s_name):
            authors.append(text)
            continue

        # --- Plain-text heuristics (no markup, Normal style) ---
        # Case 3: "Chapter 12  Title on same line" → split into number + title
        inline_m = _CHAP_INLINE_RE.match(text)
        if inline_m:
            chapter_number = inline_m.group(1).strip()
            chapter_title  = inline_m.group(2).strip()
            found_title = True
            continue
        # Case 2: standalone "Chapter 12" or bare "7" line → capture as chapter number
        if _CHAP_NUM_RE.match(text):
            chapter_number = text
            continue

        if not found_title:
            # Short line → chapter title (case-insensitive)
            words = text.split()
            if chapter_number and len(words) <= 15:
                chapter_title = text
                found_title = True
            elif len(words) <= 15:
                chapter_title = text
                found_title = True
        else:
            # Short line matching a person-name pattern → author(s)
            words = text.split()
            if (len(words) <= 12
                    and re.match(r'^[A-Za-z][\w.]+ [A-Za-z]', text, re.IGNORECASE)
                    and not re.search(r'[.?!]$', text)):
                authors.append(text)
            elif len(words) > 20:
                break  # hit body text

    return chapter_number, chapter_title, "; ".join(authors) if authors else ""


def count_references_and_body_wc(doc_path: str, doc: Optional[Any] = None):
    """
    Returns (ref_count: int, body_wc: int, total_wc: int).
    ref_count  — paragraphs in the References/Bibliography section.
    body_wc    — words in body text only (ignores everything after References).
    total_wc   — words in the entire document including references, tables, and figures.
    """
    if not os.path.exists(doc_path):
        return 0, 0, 0
    if doc is None:
        doc = Document(doc_path)
    analyzer = CitationAnalyzer()

    # Collect table paragraph texts so we can exclude them
    table_para_texts: Set[int] = set()
    total_wc = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for tp in cell.paragraphs:
                    table_para_texts.add(id(tp))
                    total_wc += len(tp.text.split())

    body_wc = 0
    ref_count = 0
    in_refs = False

    # Caption starters used to detect figure/table legends section inside refs
    _CAPTION_START_RE = re.compile(r'^\s*(figure|fig\.?|table|tab\.?|box|exhibit|appendix)\s', re.IGNORECASE)

    for p in doc.paragraphs:
        raw_text = p.text.strip()
        if not raw_text:
            continue
            
        total_wc += len(raw_text.split())
        
        # Strip inline markup tags (e.g. <REF1>, <CE:AUTHOR>) from visible text
        text = _MARKUP_TAG_RE.sub('', raw_text).strip()
        if not text:
            continue
        try:
            s_name = p.style.name.strip()
        except:
            s_name = ""

        # Detect References/Bibliography boundary
        if _REFS_HEADING_RE.match(text) and ("heading" in s_name.lower() or len(text.split()) <= 3):
            in_refs = True
            continue

        if in_refs:
            # Stop counting if we hit a "Figure Legends" / "List of Figures" header
            if _FIGURE_LEGENDS_RE.match(text):
                in_refs = True # We continue ignoring everything after references
                continue
            # Stop counting if we hit a new section heading (non-reference paragraph)
            if "heading" in s_name.lower() and not re.match(r'^\d', text):
                in_refs = True # Continue ignoring everything after references
                continue
            # Stop counting when the first figure/table caption appears after references
            if _CAPTION_START_RE.match(text):
                in_refs = True # Continue ignoring everything after references
                continue
            # Skip table/figure footnote lines (source notes, abbreviation keys, etc.)
            if re.match(r'^\s*(source|note[s]?|adapted\s+from|information\s+(based|from)|'
                        r'abbreviation|\*not\s+a\s+U\.S\.|IM,|IV,|PO,)', text, re.IGNORECASE):
                continue
            # Only count paragraphs that look like reference entries:
            # must contain a year (APA/Vancouver/numbered refs always have one)
            # OR start with a digit (numbered reference style)
            if not (re.search(r'(?<!\d)(?:19|20)\d{2}(?!\d)', text) or re.match(r'^\d+[\.\t\s]', text)):
                continue
            ref_count += 1
        else:
            # Not in refs, so count as body text
            # Do NOT explicitly exclude tables or captions in the body per new instruction.
            # All tables and figures are assumed to be placed after references.
            body_wc += len(raw_text.split())

    return ref_count, body_wc, total_wc


def count_equations(doc: Any) -> Dict[str, int]:
    """
    Count equations in a python-docx Document.
    Returns:
        omml     — native Word OMML equations (<m:oMath> elements)
        mathtype — MathType / Equation Editor OLE objects
    """
    _M_NS        = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    _W_NS        = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    _O_NS        = 'urn:schemas-microsoft-com:office:office'
    _omath_tag   = f'{{{_M_NS}}}oMath'
    _obj_tag     = f'{{{_W_NS}}}object'
    _ole_tag     = f'{{{_O_NS}}}OLEObject'
    _progid_attr = f'{{{_O_NS}}}ProgID'

    # Search entire document XML tree so equations inside table cells,
    # text boxes, and headers are included (not just doc.paragraphs)
    root = doc.element

    # OMML: every <m:oMath> = one equation block (display or inline)
    omml_count = len(root.findall(f'.//{_omath_tag}'))

    # MathType / Equation Editor OLE objects
    mathtype_count = 0
    for obj in root.findall(f'.//{_obj_tag}'):
        for ole in obj.findall(f'.//{_ole_tag}'):
            prog_id = ole.get(_progid_attr, '')
            if 'Equation' in prog_id or 'MathType' in prog_id:
                mathtype_count += 1

    return {"omml": omml_count, "mathtype": mathtype_count}


def count_unnumbered_elements(doc_path: str, dtypes: dict, doc: Optional[Any] = None):
    """
    Returns a dict with counts of unnumbered/uncaptioned elements:
      unnumbered_images  — images with no numbered Figure caption
      unnumbered_tables  — doc.tables with no numbered Table caption
      unnumbered_boxes   — box-style paragraphs with no numbered Box caption
      callouts           — vague references like "see figure above"
    """
    if not os.path.exists(doc_path):
        return {}
    if doc is None:
        doc = Document(doc_path)

    # Images: total rels - numbered figures
    img_count = sum(1 for rel in doc.part.rels.values() if "image" in rel.reltype)
    numbered_figs = len(dtypes.get("Figure", {}).get("Caption", {}))
    unnumbered_images = max(0, img_count - numbered_figs)

    # Tables: doc.tables count - numbered table captions
    numbered_tabs = len(dtypes.get("Table", {}).get("Caption", {}))
    unnumbered_tables = max(0, len(doc.tables) - numbered_tabs)

    # Boxes: find paragraphs whose style name contains 'nbx', 'box', or 'sidebar'
    # (case-insensitive). These are considered box-style elements.
    # Subtract how many numbered Box captions exist in dtypes —
    # the remainder are boxes that lack a proper numbered caption.
    numbered_boxes = len(dtypes.get("Box", {}).get("Caption", {}))
    box_style_re = re.compile(r'nbx|box|sidebar', re.IGNORECASE)
    box_para_count = sum(1 for p in doc.paragraphs
                         if p.text.strip() and box_style_re.search(getattr(p.style, 'name', '') or ''))
    unnumbered_boxes = max(0, box_para_count - numbered_boxes)

    # Callouts: scan every paragraph's text for vague cross-references such as
    # "see figure above" or "on page X" using _CALLOUT_RE and _PAGE_REF_RE.
    # Each matching paragraph increments the callout counter — these are flagged
    # because they rely on relative position rather than a numbered label.
    callout_count = 0
    for p in doc.paragraphs:
        t = p.text
        if _CALLOUT_RE.search(t) or _PAGE_REF_RE.search(t):
            callout_count += 1

    eq = count_equations(doc)
    return {
        "unnumbered_images":  unnumbered_images,
        "unnumbered_tables":  unnumbered_tables,
        "unnumbered_boxes":   unnumbered_boxes,
        "callouts":           callout_count,
        "equations_omml":     eq["omml"],
        "equations_mathtype": eq["mathtype"],
    }


def build_combined_dashboard_html(chapters_data: list, css: str, js: str, logo_b64: str) -> str:
    """
    Builds a single Combined_Dashboard.html from a list of per-chapter data dicts.
    Each chapter gets its own section with scoped tab IDs (ch0_, ch1_, ...) so
    tabs in different chapters don't interfere with each other.
    """
    # Load Jinja2 templates from the 'templates' directory relative to the current working directory
    template_dir = os.path.join(os.getcwd(), 'templates')
    
    # If the default templates dir is not found, fallback to script directory
    if not os.path.exists(template_dir):
        template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
        
    env = jinja2.Environment(loader=jinja2.FileSystemLoader(template_dir))
    
    try:
        template = env.get_template('word_analyzer_dashboard.html')
        css_content = env.get_template('word_analyzer_styles.css').render()
        js_content = env.get_template('word_analyzer_scripts.js').render()
    except jinja2.exceptions.TemplateNotFound as e:
        # Fallback inline templates if files are missing during execution
        raise FileNotFoundError(f"Template not found: {e}. Ensure templates/ contains the HTML/CSS/JS dashboard files.")

    # Render the main dashboard template with context
    return template.render(
        chapters_data=chapters_data,
        css_content=css_content,
        js_content=js_content,
        logo_b64=logo_b64
    )


# ------------------------------
# 5. Exports
# ------------------------------
__all__ = [
    "CitationAnalyzer",
    "BoxTagLinker",
    "build_element_mapping_html",
    "extract_with_docx",
    "generate_formatting_html",
    "generate_multilingual_html",
    "build_comments_html",
    "build_detailed_summary_table",
    "build_combined_dashboard_html",
    "build_export_highlight_html",
    "remove_tags_keep_formatting_docx",
    "extract_chapter_metadata",
    "count_references_and_body_wc",
    "count_unnumbered_elements",
    "extract_unnumbered_image_markups",
    "build_unnumbered_image_markups_html",
    "convert_to_pdf",
    "build_text_page_map",
    "_page_from_map",
    "HAS_PDFPLUMBER",
]


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python word_analyzer_docx.py <path_to_docx>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        sys.exit(1)
        
    print(f"Analyzing {file_path}...")
    try:
        doc = Document(file_path)
        paras, comments, imgs, footnotes, endnotes = extract_with_docx(file_path, doc=doc)
        print(f"Extraction Success:")
        print(f" - Paragraphs: {len(paras)}")
        print(f" - Comments: {len(comments)}")
        print(f" - Images: {imgs}")
        print(f" - Footnotes: {footnotes}")
        print(f" - Endnotes: {endnotes}")
        
        print("\nChecking Formatting...")
        fmt_html = generate_formatting_html(file_path, doc=doc)
        print("Formatting HTML generated (length: {} chars)".format(len(fmt_html)))
        
        print("\nChecking Multilingual...")
        multi_html = generate_multilingual_html(file_path, doc=doc)
        print("Multilingual HTML generated (length: {} chars)".format(len(multi_html)))
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
