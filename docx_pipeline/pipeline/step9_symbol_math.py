"""
pipeline/step9_symbol_math.py â€” Convert mathematical and Greek symbol characters
to their HTML-entity / Unicode equivalents.

Replicates VBA SymbolMathV4.Symbol2Unicode + Other_ent_unicode + Math module.

Logic:
  For every text run in the document (body + headers + footers +
  footnotes + endnotes), scan the plain-text content for any character
  that is listed in SYMBOL_MATH_MAP.  Replace that character with the
  corresponding HTML numeric entity string (e.g. &#x00D7;).

Note: This deliberately runs *after* step 7 so that character styles
applied to math runs are already in place before the text is mutated.
"""

from __future__ import annotations

from lxml import etree
from docx import Document
from docx.oxml.ns import qn

from docx_pipeline.config import SYMBOL_MATH_MAP
from docx_pipeline.utils.report import ReportLogger

W_ = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_PARA_PART_SUFFIXES = (
    "document.main+xml",
    "header+xml",
    "footer+xml",
    "footnotes+xml",
    "endnotes+xml",
)


def _replace_in_text(text: str) -> tuple[str, int]:
    """Return (replaced_text, count_of_substitutions)."""
    count = 0
    for char, entity in SYMBOL_MATH_MAP.items():
        if char in text:
            n = text.count(char)
            text = text.replace(char, entity)
            count += n
    return text, count


def _sweep_part(part, logger: ReportLogger, replaced: list) -> None:
    try:
        element = part.element
    except AttributeError:
        return

    for t_el in element.iter(W_ + "t"):
        if t_el.text:
            new_text, n = _replace_in_text(t_el.text)
            if n:
                t_el.text = new_text
                replaced[0] += n


def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("9-symbol-math")

    replaced = [0]

    for part in doc.part.package.parts:
        ct = part.content_type
        if any(ct.endswith(sfx) for sfx in _PARA_PART_SUFFIXES):
            _sweep_part(part, logger, replaced)

    logger.info(
        f"Symbol/Math conversion: {replaced[0]} character(s) replaced with "
        f"HTML entities."
    )
    return doc

