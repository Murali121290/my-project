"""
pipeline/step4_heading_validation.py â€” Heading level validation.

Checks:
  1. Validates heading hierarchies mapped by text markup (e.g. <H1>).
  2. Validates heading hierarchies mapped by Word styling (e.g. "H1").
  3. Validates that a Heading is not directly followed by explicitly invalid body styles ("TXT", "TX").

Violations are logged to output AND explicitly appended as native Word Comments exactly where they occur.
"""

import re
from docx import Document
from docx_pipeline.utils.report import ReportLogger
from docx_pipeline.utils.docx_helpers import add_comment_to_paragraph

# Explicit styles defined by the macros
HEADING_STYLES_ARRAY = ["H1", "H2", "H3", "H4", "H5", "H6"]
INVALID_BODY_STYLES = ["TXT", "TX"]

# Regex corresponding to: regEx.Pattern = "<H(\d+)>"
MARKUP_REGEX = re.compile(r"<H(\d+)>", re.IGNORECASE)


def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("4-heading-validation")

    paras = doc.paragraphs
    flags = 0

    # Validation 1: Markup Hierarchy Tracking
    last_level_markup = 0
    # Validation 2: Style Hierarchy Tracking
    last_level_style = 0

    for i in range(len(paras)):
        para = paras[i]
        text = para.text.strip()
        style_name = para.style.name if para.style else ""

        # ------------- RULE 1: Markup Regex Hierarchy -------------
        match = MARKUP_REGEX.search(text)
        if match:
            current_level_markup = int(match.group(1))
            if current_level_markup > last_level_markup + 1:
                comment = "CE/TE: The heading levels are not in the correct hierarchy "
                add_comment_to_paragraph(doc, para, comment)
                logger.flag(f"Markup Hierarchy Error: H{last_level_markup} jumped to H{current_level_markup} at paragraph {i}")
                flags += 1
            last_level_markup = current_level_markup

        # ------------- RULE 2: Style Hierarchy -------------
        # Get 1-based index (e.g. "H1" -> 1)
        current_level_style = -1
        for idx, hs in enumerate(HEADING_STYLES_ARRAY):
            if style_name.lower() == hs.lower():
                current_level_style = idx + 1
                break

        if current_level_style > 0:
            if current_level_style > last_level_style + 1:
                comment = f"Hierarchy error: Heading jumps from level {last_level_style} to level {current_level_style}."
                add_comment_to_paragraph(doc, para, comment)
                logger.flag(comment + f" (Paragraph {i})")
                flags += 1
            last_level_style = current_level_style

        # ------------- RULE 3: Heading Followed By Invalid Style -------------
        if current_level_style > 0 and i < len(paras) - 1:
            next_para = paras[i + 1]
            next_style_name = next_para.style.name if next_para.style else ""
            
            for invalid_style in INVALID_BODY_STYLES:
                if next_style_name.lower() == invalid_style.lower():
                    comment = f"CE/TE: Paragraph styled as '{next_style_name}' follows a heading. Consider revising the style."
                    add_comment_to_paragraph(doc, next_para, comment)
                    logger.flag(f"Invalid Next Style at para {i+1}: {next_style_name} following {style_name}")
                    flags += 1
                    break

    if flags == 0:
        logger.info("Heading validation complete: No heading hierarchy or style sequence errors found.")
    else:
        logger.warning(f"Heading validation complete: {flags} hierarchy error(s) flagged and commented natively.")

    return doc

