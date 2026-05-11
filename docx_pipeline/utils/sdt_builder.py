"""
utils/sdt_builder.py â€” Factory functions for Word SDT (content control) XML.
"""

from lxml import etree

W  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_ = "{%s}" % W


def _make_sdtPr(alias: str, tag: str) -> etree._Element:
    """Build w:sdtPr with alias and tag child elements."""
    sdtPr = etree.Element(W_ + "sdtPr")

    al = etree.SubElement(sdtPr, W_ + "alias")
    al.set(W_ + "val", alias)

    tg = etree.SubElement(sdtPr, W_ + "tag")
    tg.set(W_ + "val", tag)

    return sdtPr


def make_block_sdt(alias: str, tag: str,
                   elements: list) -> etree._Element:
    """
    Wrap a list of block-level elements (w:p, w:tbl) in a
    w:sdt block content control.

    Callers must have already detached elements from the body
    before passing them here.
    """
    sdt = etree.Element(W_ + "sdt")
    sdt.append(_make_sdtPr(alias, tag))

    content = etree.SubElement(sdt, W_ + "sdtContent")
    for el in elements:
        content.append(el)

    return sdt


def make_inline_sdt(alias: str, tag: str,
                    runs: list) -> etree._Element:
    """
    Wrap a list of inline elements (w:r, w:fldChar etc.) in a
    w:sdt inline content control (lives inside a w:p).
    """
    sdt = etree.Element(W_ + "sdt")
    sdt.append(_make_sdtPr(alias, tag))

    content = etree.SubElement(sdt, W_ + "sdtContent")
    for r in runs:
        content.append(r)

    return sdt

