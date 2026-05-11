"""
pipeline/step7_char_styles.py â€” Apply named character styles.

For every run with direct bold/italic/underline/caps/script formatting:
  - Detect the combination of font properties (as a frozenset)
  - Map to a named character style via CHAR_STYLE_MAP_COMPREHENSIVE
  - Apply the character style by injecting <w:rStyle> XML directly and
    stripping conflicting direct-format tags (no run.style= which
    requires a live parent.part reference)

If a target character style is missing from the document (e.g. removed by
step 2 as "unused"), it is AUTO-CREATED with the correct font properties
from its name â€” mirroring VBA's "If Not StyleExists â†’ Styles.Add + SetFontProperty".

Runs inside SEMANTIC_BOLD_STYLES paragraphs are SKIPPED entirely.
Their direct formatting is preserved as-is (italic species names inside
headings remain italic).

Sweep scope: main body, tables, headers, footers, footnotes, endnotes.
"""

from __future__ import annotations

from lxml import etree
from docx import Document
from docx.oxml.ns import qn

from docx_pipeline.config import CHAR_STYLE_MAP_COMPREHENSIVE, SEMANTIC_BOLD_STYLES, STYLE_PROPERTY_MAP
from docx_pipeline.utils.report import ReportLogger

W_ = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

SEMANTIC_LOWER = {s.lower() for s in SEMANTIC_BOLD_STYLES}

_PARA_PART_SUFFIXES = (
    "document.main+xml",
    "header+xml",
    "footer+xml",
    "footnotes+xml",
    "endnotes+xml",
)

# Direct-format XML tags to strip once a character style takes over
_FORMAT_TAGS = ["b", "bCs", "i", "iCs", "u", "caps", "smallCaps", "vertAlign"]


# â”€â”€ Auto-create helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _shading_hex_for_props(props: frozenset) -> str | None:
    """Return hex fill colour for a property set, mirroring VBA ApplyBookcolor."""
    b  = 'bold'            in props
    i  = 'italic'          in props
    su = 'superscript'     in props
    sb = 'subscript'       in props
    s1 = 'singleunderline' in props
    d2 = 'doubleunderline' in props
    ac = 'allcaps'         in props
    sc = 'smallcaps'       in props
    if b and i:   return "FFCCCC"   # Light Red    â€” bold italic
    if b:         return "CCE0FF"   # Light Blue   â€” bold (any)
    if i and su:  return "FFFFCC"   # Light Yellow â€” italic superscript
    if i and s1:  return "E8CCFF"   # Light Violet â€” italic single underline
    if i:         return "CCFFCC"   # Light Green  â€” italic (any other)
    if d2:        return "CCFFFF"   # Light Cyan   â€” double underline
    if s1:        return "CCFFCC"   # Light Green  â€” single underline
    if su:        return "CCFFFF"   # Light Cyan   â€” superscript
    if sb:        return "FFCCFF"   # Light Pink   â€” subscript
    if ac and sc: return "E0E0E0"   # Light Gray   â€” smallcaps
    if ac:        return "F0F0F0"   # Lighter Gray â€” allcaps
    return None


def _make_char_style(doc: Document, style_name: str) -> str:
    """
    Create a character style named `style_name` in the document with font
    properties sourced from STYLE_PROPERTY_MAP (reverse of CHAR_STYLE_MAP_COMPREHENSIVE).
    Returns the styleId.  Mirrors VBA SetFontProperty4CStyle.
    """
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_UNDERLINE

    new_style = doc.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
    new_style.base_style = None
    f = new_style.font

    props = STYLE_PROPERTY_MAP.get(style_name, frozenset())
    if 'bold'            in props: f.bold        = True
    if 'italic'          in props: f.italic       = True
    if 'singleunderline' in props: f.underline    = True
    if 'doubleunderline' in props: f.underline    = WD_UNDERLINE.DOUBLE
    if 'superscript'     in props: f.superscript  = True
    if 'subscript'       in props: f.subscript    = True
    if 'allcaps'         in props: f.all_caps     = True
    if 'smallcaps'       in props: f.small_caps   = True

    fill = _shading_hex_for_props(props)
    if fill:
        from lxml import etree
        rPr_el = new_style.element.find(qn("w:rPr"))
        if rPr_el is None:
            rPr_el = etree.SubElement(new_style.element, qn("w:rPr"))
        shd = etree.SubElement(rPr_el, qn("w:shd"))
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  fill)

    return new_style.element.get(qn("w:styleId"), style_name)


def _build_style_id_map(doc: Document) -> dict[str, str]:
    """Return {styleName: styleId} for every style currently in doc.styles."""
    result: dict[str, str] = {}
    for style in doc.styles:
        el = style.element
        sid = el.get(qn("w:styleId"), style.name)
        result[style.name] = sid
    return result


def _get_or_create_style_id(doc: Document, style_name: str,
                             style_id_map: dict[str, str],
                             logger: ReportLogger) -> str | None:
    """Return styleId for style_name, auto-creating it if missing."""
    sid = style_id_map.get(style_name)
    if sid is not None:
        return sid
    try:
        sid = _make_char_style(doc, style_name)
        style_id_map[style_name] = sid
        logger.info(f'Auto-created missing char style "{style_name}".')
        return sid
    except Exception as exc:
        logger.warning(f'Could not create char style "{style_name}": {exc}')
        return None


# â”€â”€ XML-level property detection â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _get_props(r_el: etree._Element) -> frozenset:
    """
    Detect direct-format properties from the run XML element.
    Returns a frozenset of token strings (e.g. frozenset({'bold', 'italic'})).
    Works without a python-docx parent reference.
    """
    props: set[str] = set()
    rPr = r_el.find(W_ + "rPr")
    if rPr is None:
        return frozenset()

    def _on(tag: str) -> bool:
        el = rPr.find(W_ + tag)
        if el is None:
            return False
        val = el.get(W_ + "val", "true").lower()
        return val not in ("0", "false")

    if _on("b"):
        props.add("bold")
    if _on("i"):
        props.add("italic")

    u_el = rPr.find(W_ + "u")
    if u_el is not None:
        uval = u_el.get(W_ + "val", "none").lower()
        if uval == "single":
            props.add("singleunderline")
        elif uval == "double":
            props.add("doubleunderline")

    if _on("caps"):
        props.add("allcaps")
    if _on("smallCaps"):
        props.add("smallcaps")

    va_el = rPr.find(W_ + "vertAlign")
    if va_el is not None:
        va = va_el.get(W_ + "val", "").lower()
        if va == "superscript":
            props.add("superscript")
        elif va == "subscript":
            props.add("subscript")

    return frozenset(props)


# â”€â”€ Apply character style via raw XML (no parent.part needed) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _apply_char_style_xml(r_el: etree._Element, style_id: str) -> None:
    """
    Inject <w:rStyle w:val="{style_id}"/> into the run's <w:rPr>
    and strip the direct-format tags the character style now owns.
    Pure lxml â€” no python-docx API calls that need parent.part.
    """
    rPr = r_el.find(W_ + "rPr")
    if rPr is None:
        rPr = etree.Element(W_ + "rPr")
        r_el.insert(0, rPr)

    # Remove any pre-existing rStyle
    existing = rPr.find(W_ + "rStyle")
    if existing is not None:
        rPr.remove(existing)

    # Build new rStyle and insert at position 0 (must be first child of rPr)
    rStyle = etree.Element(W_ + "rStyle")
    rStyle.set(W_ + "val", style_id)
    rPr.insert(0, rStyle)

    # Strip direct-format tags superseded by the character style
    for tag in _FORMAT_TAGS:
        el = rPr.find(W_ + tag)
        if el is not None:
            rPr.remove(el)


# â”€â”€ Paragraph / part processors â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _process_paragraph_el(p_el: etree._Element,
                           doc: Document,
                           style_id_map: dict[str, str],
                           logger: ReportLogger,
                           styled: list, skipped: list) -> None:
    # Determine paragraph style from XML (no python-docx Paragraph wrapper)
    pStyle_el = p_el.find(W_ + "pPr/" + W_ + "pStyle")
    style_val = (pStyle_el.get(W_ + "val") if pStyle_el is not None else "") or ""

    if style_val.lower() in SEMANTIC_LOWER:
        # Skip â€” preserve all formatting inside heading/semantic paragraphs
        skipped[0] += 1
        return

    for r_el in p_el.findall(W_ + "r"):
        props = _get_props(r_el)
        if not props:
            continue

        style_name = CHAR_STYLE_MAP_COMPREHENSIVE.get(props)
        if style_name is None:
            continue

        style_id = _get_or_create_style_id(doc, style_name, style_id_map, logger)
        if style_id is None:
            continue

        _apply_char_style_xml(r_el, style_id)
        styled[0] += 1


def _sweep_part(part, doc: Document, style_id_map: dict[str, str],
                logger: ReportLogger, styled: list, skipped: list) -> None:
    try:
        element = part.element
    except AttributeError:
        return
    for p_el in element.iter(W_ + "p"):
        _process_paragraph_el(p_el, doc, style_id_map, logger, styled, skipped)


# â”€â”€ Entry point â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("7-char-styles")

    style_id_map = _build_style_id_map(doc)
    styled  = [0]
    skipped = [0]

    for part in doc.part.package.parts:
        ct = part.content_type
        if any(ct.endswith(sfx) for sfx in _PARA_PART_SUFFIXES):
            _sweep_part(part, doc, style_id_map, logger, styled, skipped)

    logger.info(
        f"Character styles applied: {styled[0]} run(s) styled across all "
        f"document parts (footnotes/endnotes included). "
        f"{skipped[0]} semantic/heading paragraph(s) skipped â€” "
        f"their run formatting preserved."
    )
    return doc

