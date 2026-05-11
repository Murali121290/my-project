"""
Extract structured text from a .docx file.

Returns a list of "segments" â€” each segment is a contiguous block of text with
metadata about its location (chapter, paragraph index, approximate page, source
kind like body/footnote/endnote) and an `excluded` flag marking it as a zone we
should NOT match rules against (quoted text, extracts, reference lists,
captions).

Design:
- Uses python-docx for paragraphs, tables, and styles.
- Uses lxml to reach into w:footnotes.xml and w:endnotes.xml, which python-docx
  doesn't expose directly.
- Approximates page numbers using `w:lastRenderedPageBreak` markers written by
  Word on save; falls back to a words-per-page estimate.
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

WORDS_PER_PAGE_FALLBACK = 300

# Paragraph style names (lowercased) we treat as "extract" zones.
EXTRACT_STYLE_KEYWORDS = (
    "quote",
    "intense quote",
    "extract",
    "blockquote",
    "epigraph",
    "epigraph text",
    "epi",
    "epi-text",
    "ext-mid",
    "ext-only",
    "py",
    "ref-n",
    "ref-u",
    "verse",
    "poetry",
    "block quote",
    "long quote",
    "displayed quote",
    "pull quote",
)
CAPTION_STYLE_KEYWORDS = (
    "caption",
    "fig-leg",
    "t1",
    "tt",
    "fgc",
    "figurelegend",
    "tablecaption",
)
REFERENCE_STYLE_KEYWORDS = ("bibliography", "references", "workcited", "literaturecited", "ref-list")
HEADING_STYLE_PREFIX = "heading"

# Headings that mark the start of a reference list (case-insensitive contains).
REFERENCE_HEADING_PATTERNS = [
    re.compile(r"^\s*references?\s*$", re.IGNORECASE),
    re.compile(r"^\s*bibliography\s*$", re.IGNORECASE),
    re.compile(r"^\s*works\s+cited\s*$", re.IGNORECASE),
    re.compile(r"^\s*literature\s+cited\s*$", re.IGNORECASE),
    re.compile(r"^\s*cited\s+works\s*$", re.IGNORECASE),
    re.compile(r"^\s*sources?\s+cited\s*$", re.IGNORECASE),
    re.compile(r"^\s*notes?\s+and\s+references?\s*$", re.IGNORECASE),
    re.compile(r"^\s*further\s+reading\s*$", re.IGNORECASE),
    re.compile(r"^\s*selected\s+references?\s*$", re.IGNORECASE),
    re.compile(r"^\s*references?\s+cited\s*$", re.IGNORECASE),
    re.compile(r"^\s*endnotes?\s*$", re.IGNORECASE),
]

# Straight + smart quote pairs. We mask text *between* them.
QUOTE_PAIRS = [
    ('"', '"'),
    ("\u201c", "\u201d"),  # â€œ â€
    ("\u2018", "\u2019"),  # â€˜ â€™  (apostrophe risk â€” handled below)
    ("'", "'"),
]


@dataclass
class Segment:
    """A contiguous span of text with origin metadata."""

    chapter_index: int
    chapter_name: str
    source: str  # "body", "footnote", "endnote", "table"
    para_index: int  # stable index within the chapter (body paragraphs)
    text: str
    style: str = ""
    page: int = 1
    excluded: bool = False
    exclude_reason: str = ""
    # Per-character mask into `text`: True where that char is inside a
    # quote/extract/caption/ref zone and must NOT be matched.
    mask: list[bool] = field(default_factory=list)
    region: str = "body"

    def to_dict(self) -> dict:
        return {
            "chapter_index": self.chapter_index,
            "chapter_name": self.chapter_name,
            "source": self.source,
            "para_index": self.para_index,
            "text": self.text,
            "style": self.style,
            "page": self.page,
            "excluded": self.excluded,
            "exclude_reason": self.exclude_reason,
            "region": self.region,
        }


# ---------------------------------------------------------------------------
# Footnote / endnote extraction
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _read_notes(docx_path: Path, part_name: str) -> dict[str, str]:
    """Return {note_id: concatenated_text} for footnotes or endnotes.

    part_name is 'footnotes' or 'endnotes'.
    """
    notes: dict[str, str] = {}
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            target = f"word/{part_name}.xml"
            if target not in z.namelist():
                return notes
            xml_bytes = z.read(target)
    except (zipfile.BadZipFile, KeyError):
        return notes

    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return notes

    ns = {"w": W_NS}
    tag = "footnote" if part_name == "footnotes" else "endnote"
    for note in root.findall(f"w:{tag}", ns):
        note_id = note.get(f"{{{W_NS}}}id")
        if note_id is None:
            continue
        # Skip separator/continuation auto-notes (type="separator" etc.).
        note_type = note.get(f"{{{W_NS}}}type")
        if note_type and note_type != "normal":
            continue
        texts = [t.text or "" for t in note.iter(f"{{{W_NS}}}t")]
        combined = "".join(texts).strip()
        if combined:
            notes[note_id] = combined
    return notes


# ---------------------------------------------------------------------------
# Exclusion masking
# ---------------------------------------------------------------------------


def mask_quotes(text: str) -> list[bool]:
    """Return a per-char boolean list; True where char is inside quotes.

    Handles ASCII double quotes and curly double quotes as paired delimiters.
    Single quotes / apostrophes are skipped (too many false positives from
    contractions and possessives).
    """
    mask = [False] * len(text)
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        # Straight double quote
        if c == '"':
            j = text.find('"', i + 1)
            if j == -1:
                break
            for k in range(i, j + 1):
                mask[k] = True
            i = j + 1
            continue
        # Left curly double quote â€” find matching right curly
        if c == "\u201c":
            j = text.find("\u201d", i + 1)
            if j == -1:
                # No closing â€” bail, leave unmasked
                i += 1
                continue
            for k in range(i, j + 1):
                mask[k] = True
            i = j + 1
            continue
        i += 1
    return mask


# Citation patterns to mask inside body paragraphs.
# We cover the four most common academic citation styles:
#   1. Numeric  [1]  [1,2]  [1-3]  [1â€“3]
#   2. Author-year  (Smith, 2020)  (Smith et al., 2020)  (Smith & Jones 2020)
#   3. Superscript anchors written as plain text  ^1  ^[1]
#   4. Footnote/endnote markers written as (1)  (a)  (i) etc. (short bracketed tokens)
_CITATION_PATTERNS = [
    # [1], [1,2], [1-3], [1â€“3]
    re.compile(r"\[\d[\d,\s\-\u2013]*\]"),
    # (Author, Year) â€” 1 to 3 authors, optional et al., optional p./pp.
    re.compile(
        r"\("
        r"[A-Z][\w\-']+"
        r"(?:\s+et\s+al\.?)?"
        r"(?:\s*[,;&]\s*[A-Z][\w\-']+(?:\s+et\s+al\.?)?)?"
        r"[,\s]+"
        r"(?:pp?\.\s*)?\d{4}[a-z]?"
        r"(?:\s*,\s*pp?\.\s*\d+(?:[\-\u2013]\d+)?)?"
        r"\)",
        re.IGNORECASE,
    ),
    # Short numeric parenthetical (1) (12) (a) (i) â€” only single token inside
    re.compile(r"\(\d{1,3}\)"),
]


def mask_citations(text: str) -> list[bool]:
    """Return a per-char boolean list; True where char is inside a citation span."""
    mask = [False] * len(text)
    for pat in _CITATION_PATTERNS:
        for m in pat.finditer(text):
            for k in range(m.start(), m.end()):
                mask[k] = True
    return mask


_MARKUP_TAG_RE = re.compile(r'</?[A-Za-z][A-Za-z0-9._-]*>', re.IGNORECASE)


def mask_markup_tags(text: str) -> list[bool]:
    """Return a per-char boolean list; True where char is inside a structural markup tag.

    Covers simple tags (<front>, <PT>, <CN>), closing tags (</BXOBJ>),
    and compound tags with digits/dots (<FIG3.1>, <TAB3.1>).
    """
    mask = [False] * len(text)
    for m in _MARKUP_TAG_RE.finditer(text):
        for k in range(m.start(), m.end()):
            mask[k] = True
    return mask


def build_segment_mask(text: str) -> list[bool]:
    """Combine quote, citation, and markup-tag masking into one per-char boolean list."""
    q_mask = mask_quotes(text)
    c_mask = mask_citations(text)
    t_mask = mask_markup_tags(text)
    return [q or c or t for q, c, t in zip(q_mask, c_mask, t_mask)]


def is_caption_paragraph(text: str, style_name: str = "") -> bool:
    if style_name:
        s_low = style_name.strip().lower()
        if s_low in ['fig-leg', 'fgc', 't1', 'tt', 'figurelegend', 'tablecaption', 'cs-ttl','nbx1-num','nbx1-ttl','nbx2-num','nbx2-ttl', 'exhibitcaption']:
            return True

    t_norm_orig = text.strip()
    has_title_tag = bool(re.search(r'<(TITLE|CAPTION|BX_TTL|BX-TTL|CS-TTL)>', t_norm_orig, re.IGNORECASE))
    has_type_tag  = bool(re.search(r'^<(FIG|TAB|BX|CS|EX|APP)>', t_norm_orig.strip(), re.IGNORECASE))

    t_norm = re.sub(r'^(?:<[^>]*>\s*)+', '', t_norm_orig)
    if not t_norm:
        return False
    if len(t_norm.splitlines()) > 7:
        return False

    match = re.match(r'(?i)^(figure|fig\.|table|tab\.|box|exhibit|appendix|case\s+study)\s+([0-9]+(?:[.\-][0-9]+)*[a-zA-Z]?)(.*)', t_norm)
    if match:
        if has_title_tag or has_type_tag:
            return True
        remainder = match.group(3).strip()
        if not re.search(r'[A-Za-z0-9]', remainder):
            return False
        # If the remainder starts with a separator (. : - â€” â€“), the
        # lowercase check is skipped: "Table 4.2. pH dependenceâ€¦" is a
        # valid caption even though "pH" starts with a lowercase letter.
        has_separator = bool(re.match(r'^[.\:\-\u2013\u2014]\s', remainder))
        if not has_separator:
            first_word_char = re.sub(r'^[\W_]+', '', remainder)
            if first_word_char and first_word_char[0].islower():
                return False
        return True

    return False


def _classify_paragraph(style_name: str, text: str) -> tuple[bool, str]:
    """Decide if an individual paragraph should be fully masked (excluded).

    Returns (excluded, reason). Heading detection and reference-list
    detection is handled at a higher level â€” this only catches styles.
    """
    style_lower = (style_name or "").lower()
    if any(k in style_lower for k in EXTRACT_STYLE_KEYWORDS):
        return True, "extract"
    if any(k in style_lower for k in REFERENCE_STYLE_KEYWORDS):
        return True, "reference"
        
    if is_caption_paragraph(text, style_name) or any(k in style_lower for k in CAPTION_STYLE_KEYWORDS):
        return True, "caption"

    # Content-pattern epigraph: short block starting with an open quote and
    # ending with an attribution line (â€” Name or â€“ Name).
    lines = [l for l in text.splitlines() if l.strip()]
    if 1 <= len(lines) <= 4:
        first = lines[0].lstrip()
        last = lines[-1].strip()
        starts_with_quote = first.startswith(("â€œ", "â€˜", '"', "'"))
        has_attribution = bool(re.match(r'^[â€”â€“\-]\s*\S', last))
        if starts_with_quote and has_attribution:
            return True, "extract"

    return False, ""


def _is_reference_heading(text: str) -> bool:
    for pat in REFERENCE_HEADING_PATTERNS:
        if pat.match(text.strip()):
            return True
    return False


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------


def _iter_body_paragraphs(doc) -> Iterable:
    """Yield paragraphs in document order, including those inside tables and textboxes."""
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    
    for p_element in body.iter(qn("w:p")):
        source = "body"
        # Determine source by checking ancestors
        ancestors = [a.tag for a in p_element.iterancestors()]
        if qn("w:tbl") in ancestors:
            source = "table"
        elif qn("w:txbxContent") in ancestors or any(a.endswith("}textbox") for a in ancestors):
            source = "textbox"
            
        p_obj = Paragraph(p_element, doc._body)
        yield (source, p_obj)


def _count_rendered_page_breaks(paragraph) -> int:
    """Count w:lastRenderedPageBreak elements within a paragraph."""
    return len(paragraph._element.findall(".//" + qn("w:lastRenderedPageBreak")))


def extract_segments(
    docx_path: str | Path,
    chapter_index: int,
    chapter_name: str,
) -> list[Segment]:
    """Return flat list of Segments from a single chapter .docx."""
    docx_path = Path(docx_path)
    doc = Document(str(docx_path))

    footnotes = _read_notes(docx_path, "footnotes")
    endnotes = _read_notes(docx_path, "endnotes")

    segments: list[Segment] = []
    current_page = 1
    running_word_count = 0
    in_reference_block = False
    in_extract_tag_block = False
    current_region = "front"
    para_idx = 0

    for source, p in _iter_body_paragraphs(doc):
        text = (p.text or "").rstrip()
        text_lower = text.lower()
        style_name = p.style.name if p.style else ""

        # Update region state based on explicitly typed markup tags
        if "<front>" in text_lower:
            current_region = "front"
        elif "<body>" in text_lower:
            current_region = "body"
        elif "<ref-open>" in text_lower:
            current_region = "references"
            
        # Track tagged extract/quote blocks
        if re.search(r'<(quote|extract|epigraph)>', text_lower):
            in_extract_tag_block = True
        if re.search(r'</(quote|extract|epigraph)>', text_lower):
            in_extract_tag_block = False

        # Page tracking: bump page for each lastRenderedPageBreak.
        breaks = _count_rendered_page_breaks(p)
        if breaks:
            current_page += breaks

        # Stop reference block if we hit <ref-close> or a caption-like paragraph
        if in_reference_block:
            if "<ref-close>" in text.lower():
                in_reference_block = False
                current_region = "body"
            elif re.match(r"^\s*(?:figure|table|box)\b", text, re.IGNORECASE):
                in_reference_block = False

        # Reference heading detection â€” everything after this is a reference zone.
        if _is_reference_heading(text) or "<ref-open>" in text.lower():
            in_reference_block = True

        if not text.strip():
            para_idx += 1
            continue

        excluded, reason = _classify_paragraph(style_name, text)
        if in_reference_block:
            excluded = True
            reason = "reference"
        elif in_extract_tag_block or re.search(r'</?(quote|extract|epigraph)>', text_lower):
            excluded = True
            reason = "extract"
        
        # Ignored Tags (metadata - skip entirely)
        if re.search(r'</?(cn|ct|fig(?:[0-9.-]*)?|tags)>', text_lower):
            excluded = True
            reason = "ignored_tag"

        # Heading paragraphs still get scanned (titles can have issues), but we
        # don't mark them excluded. Footnote markers inside body text stay too.
        # Table cells get their own region ("table") instead of inheriting document region
        seg_region = "table" if source == "table" else current_region

        seg = Segment(
            chapter_index=chapter_index,
            chapter_name=chapter_name,
            source=source,
            para_index=para_idx,
            text=text,
            style=style_name,
            page=current_page,
            excluded=excluded,
            exclude_reason=reason,
            region=seg_region,
        )

        # Quote + citation masking overlays the segment even if not fully excluded.
        if excluded:
            seg.mask = [True] * len(text)
        else:
            seg.mask = build_segment_mask(text)

        segments.append(seg)

        # Rough fallback page estimate if Word didn't leave break markers.
        running_word_count += len(text.split())
        if breaks == 0 and running_word_count >= WORDS_PER_PAGE_FALLBACK:
            current_page += running_word_count // WORDS_PER_PAGE_FALLBACK
            running_word_count %= WORDS_PER_PAGE_FALLBACK

        para_idx += 1

    # Footnotes and endnotes â€” these get their own segments, NOT excluded by
    # default, but the quote mask still applies.
    for note_id, note_text in footnotes.items():
        seg = Segment(
            chapter_index=chapter_index,
            chapter_name=chapter_name,
            source="footnote",
            para_index=10_000 + int(note_id),
            text=note_text,
            style="Footnote Text",
            page=current_page,  # best we can do â€” footnote page isn't easily derivable
            excluded=False,
            exclude_reason="",
        )
        seg.mask = build_segment_mask(note_text)
        segments.append(seg)

    for note_id, note_text in endnotes.items():
        seg = Segment(
            chapter_index=chapter_index,
            chapter_name=chapter_name,
            source="endnote",
            para_index=20_000 + int(note_id),
            text=note_text,
            style="Endnote Text",
            page=current_page,
            excluded=False,
            exclude_reason="",
        )
        seg.mask = build_segment_mask(note_text)
        segments.append(seg)

    return segments

