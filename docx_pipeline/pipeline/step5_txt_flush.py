"""
pipeline/step5_txt_flush.py â€” Heading â†’ TXT-FLUSH validation (flag-only).

After every heading paragraph, the next non-empty paragraph
must use the TXT_FLUSH_STYLE style. Violations are flagged only â€”
no auto-correction is applied.
"""

from docx import Document
from docx_pipeline.config import HEADING_STYLES, TXT_FLUSH_STYLE
from docx_pipeline.utils.report import ReportLogger


def _is_heading(para) -> bool:
    sname = para.style.name if para.style else ""
    return sname in HEADING_STYLES


def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("5-txt-flush")

    paras = doc.paragraphs
    flags = 0

    for i, para in enumerate(paras):
        if not _is_heading(para):
            continue

        heading_snippet = para.text[:50].strip()

        # Find next non-empty paragraph
        next_para = None
        for j in range(i + 1, len(paras)):
            if paras[j].text.strip():
                next_para = paras[j]
                break

        if next_para is None:
            # Heading at end of document
            continue

        next_style = next_para.style.name if next_para.style else ""
        if next_style not in TXT_FLUSH_STYLE:
            logger.flag(
                f'After heading "{heading_snippet}" '
                f'(para {i}): expected {TXT_FLUSH_STYLE} '
                f'but found "{next_style}" â€” '
                f'"{next_para.text[:40].strip()}"')
            flags += 1

    if flags == 0:
        logger.info("TXT-FLUSH check: all headings followed correctly.")
    else:
        logger.warning(
            f"TXT-FLUSH check: {flags} violation(s) flagged for review.")

    return doc

