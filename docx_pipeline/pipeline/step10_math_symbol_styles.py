"""
pipeline/step10_math_symbol_styles.py â€” Apply Math/Symbol character styles.

Migrates the VBA Math() / Symbol() macros.

After step9 converts Unicode math/symbol characters to HTML entity strings,
this step:
  1. Finds every run whose text contains a known entity string.
  2. Splits the run at entity boundaries so each entity is isolated in its
     own run (surrounding text stays in separate runs with original formatting).
  3. Applies the correct character style (Math, MathI, SymbolBI, â€¦) to each
     entity run, based on the original run's direct formatting
     (bold Ã— italic Ã— superscript Ã— subscript).
  4. Converts the entity string back to the original Unicode character.

Result: only the specific math/symbol character carries the named style;
surrounding text in the same original run is unaffected.
"""

from __future__ import annotations

from copy import deepcopy

from lxml import etree
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn as _qn

from docx_pipeline.config import (
    ENTITY_TO_CHAR,
    SYMBOL_ENTITIES, MATH_ENTITIES,
    MATH_STYLE_MAP, SYMBOL_STYLE_MAP,
    MATH_STYLE_SHADING, SYMBOL_STYLE_SHADING,
)
from docx_pipeline.utils.report import ReportLogger

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_ = "{%s}" % W
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_ALL_ENTITIES: frozenset[str] = SYMBOL_ENTITIES | MATH_ENTITIES

_PARA_PART_SUFFIXES = (
    "document.main+xml",
    "header+xml",
    "footer+xml",
    "footnotes+xml",
    "endnotes+xml",
)


# â”€â”€ Helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _get_formatting(r_el: etree._Element) -> tuple[bool, bool, bool, bool]:
    """Return (bold, italic, superscript, subscript) from direct run XML."""
    rPr = r_el.find(f"{W_}rPr")
    if rPr is None:
        return (False, False, False, False)

    def _flag(tag: str) -> bool:
        el = rPr.find(f"{W_}{tag}")
        if el is None:
            return False
        return el.get(f"{W_}val", "true") not in ("0", "false", "FALSE")

    bold   = _flag("b")
    italic = _flag("i")
    vert   = rPr.find(f"{W_}vertAlign")
    sup    = vert is not None and vert.get(f"{W_}val") == "superscript"
    sub    = vert is not None and vert.get(f"{W_}val") == "subscript"
    return (bool(bold), bool(italic), bool(sup), bool(sub))


def _ensure_style(doc: Document, style_name: str, fmt: tuple) -> str:
    """Return the styleId of an existing or newly created character style."""
    try:
        return doc.styles[style_name].element.get(_qn("w:styleId"))
    except KeyError:
        bold, italic, sup, sub = fmt
        s = doc.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
        f = s.font
        if bold:   f.bold        = True
        if italic: f.italic      = True
        if sup:    f.superscript = True
        if sub:    f.subscript   = True
        shading_map = SYMBOL_STYLE_SHADING if style_name.startswith("Symbol") else MATH_STYLE_SHADING
        fill = shading_map.get(fmt)
        if fill:
            rPr_el = s.element.find(_qn("w:rPr"))
            if rPr_el is None:
                rPr_el = etree.SubElement(s.element, _qn("w:rPr"))
            shd = etree.SubElement(rPr_el, _qn("w:shd"))
            shd.set(_qn("w:val"),   "clear")
            shd.set(_qn("w:color"), "auto")
            shd.set(_qn("w:fill"),  fill)
        return s.element.get(_qn("w:styleId"))


def _apply_style(r_el: etree._Element, style_id: str) -> None:
    """Inject <w:rStyle> and strip conflicting direct formatting."""
    rPr = r_el.find(f"{W_}rPr")
    if rPr is None:
        rPr = etree.Element(f"{W_}rPr")
        r_el.insert(0, rPr)
    for old in rPr.findall(f"{W_}rStyle"):
        rPr.remove(old)
    rStyle = etree.Element(f"{W_}rStyle")
    rStyle.set(f"{W_}val", style_id)
    rPr.insert(0, rStyle)
    for tag in ("b", "bCs", "i", "iCs", "vertAlign"):
        for el in rPr.findall(f"{W_}{tag}"):
            rPr.remove(el)


def _make_text_run(template_r: etree._Element, text: str) -> etree._Element:
    """Clone a run's rPr into a new run with the given text."""
    new_r = etree.Element(f"{W_}r")
    rPr = template_r.find(f"{W_}rPr")
    if rPr is not None:
        new_r.append(deepcopy(rPr))
    new_t = etree.SubElement(new_r, f"{W_}t")
    new_t.text = text
    if text != text.strip() or len(text) == 1:
        new_t.set(XML_SPACE, "preserve")
    return new_r


# â”€â”€ Core: split run at entity boundaries and style entity segments â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _process_run(
    r_el: etree._Element,
    doc: Document,
    style_cache: dict[str, str],
) -> int:
    """
    Process one run.  Splits it at entity boundaries; styles each entity
    segment.  Returns the number of entity segments styled.
    """
    full_text = "".join(t.text or "" for t in r_el.findall(f"{W_}t"))
    if not full_text:
        return 0

    # Find all entity occurrences (position, end, entity_string)
    hits: list[tuple[int, int, str]] = []
    for entity in _ALL_ENTITIES:
        start = 0
        while True:
            idx = full_text.find(entity, start)
            if idx == -1:
                break
            hits.append((idx, idx + len(entity), entity))
            start = idx + len(entity)

    if not hits:
        return 0

    hits.sort()

    # Build segments: list of (text, entity_or_None)
    segments: list[tuple[str, str | None]] = []
    pos = 0
    for s, e, entity in hits:
        if pos < s:
            segments.append((full_text[pos:s], None))
        segments.append((full_text[s:e], entity))
        pos = e
    if pos < len(full_text):
        segments.append((full_text[pos:], None))

    fmt = _get_formatting(r_el)

    # Fast path: entire run is a single entity â€” modify in-place
    if len(segments) == 1 and segments[0][1] is not None:
        entity = segments[0][1]
        has_symbol = entity in SYMBOL_ENTITIES
        style_name = (SYMBOL_STYLE_MAP if has_symbol else MATH_STYLE_MAP).get(fmt)
        if style_name:
            if style_name not in style_cache:
                style_cache[style_name] = _ensure_style(doc, style_name, fmt)
            _apply_style(r_el, style_cache[style_name])
        char = ENTITY_TO_CHAR.get(entity, entity)
        for t_el in r_el.findall(f"{W_}t"):
            if t_el.text and entity in t_el.text:
                t_el.text = t_el.text.replace(entity, char)
                if t_el.text != t_el.text.strip() or len(t_el.text) == 1:
                    t_el.set(XML_SPACE, "preserve")
        return 1

    # Split path: entity is embedded inside larger text â€” create new runs
    parent = r_el.getparent()
    if parent is None:
        return 0

    insert_pos = list(parent).index(r_el)
    new_runs: list[etree._Element] = []
    entity_count = 0

    for seg_text, entity in segments:
        if not seg_text:
            continue
        if entity is None:
            # Plain text â€” keep original formatting
            new_runs.append(_make_text_run(r_el, seg_text))
        else:
            # Entity segment â€” restore char and apply style
            char = ENTITY_TO_CHAR.get(entity, entity)
            new_r = _make_text_run(r_el, char)
            has_symbol = entity in SYMBOL_ENTITIES
            style_name = (SYMBOL_STYLE_MAP if has_symbol else MATH_STYLE_MAP).get(fmt)
            if style_name:
                if style_name not in style_cache:
                    style_cache[style_name] = _ensure_style(doc, style_name, fmt)
                _apply_style(new_r, style_cache[style_name])
            new_runs.append(new_r)
            entity_count += 1

    parent.remove(r_el)
    for i, new_r in enumerate(new_runs):
        parent.insert(insert_pos + i, new_r)

    return entity_count


# â”€â”€ Sweep â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def _sweep_part(part, doc: Document, style_cache: dict, count: list) -> None:
    try:
        element = part.element
    except AttributeError:
        return
    # Snapshot run list before modification to avoid iterator invalidation
    for r_el in list(element.iter(W_ + "r")):
        count[0] += _process_run(r_el, doc, style_cache)


# â”€â”€ Entry point â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("10-math-symbol-styles")

    style_cache: dict[str, str] = {}
    count = [0]

    for part in doc.part.package.parts:
        ct = part.content_type
        if any(ct.endswith(sfx) for sfx in _PARA_PART_SUFFIXES):
            _sweep_part(part, doc, style_cache, count)

    if style_cache:
        logger.info(
            f"Auto-created character styles: {', '.join(sorted(style_cache))}"
        )
    logger.info(
        f"Math/Symbol styles applied: {count[0]} character(s) styled and "
        f"restored to Unicode."
    )
    return doc

