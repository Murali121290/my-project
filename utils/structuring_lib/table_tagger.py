"""
Table tagging and styling module (deprecated).
Functionality moved to styler.py. Kept for backwards compatibility.
"""

import logging
from typing import Literal
from docx import Document
from .logger_config import get_logger

logger = get_logger(__name__)


def tag_tables(doc: Document, mode: Literal["style", "tag"] = "style") -> None:
    """
    Apply styles or tags to table cells (deprecated).
    
    Args:
        doc: python-docx Document object
        mode: "style" to apply Word styles, "tag" to prefix text with [TAG]
    
    Note:
        This function is deprecated. Use styler.tag_tables() instead.
    """
    logger.warning("table_tagger.tag_tables is deprecated. Use styler.tag_tables() instead.")
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if not text:
                        continue
                    
                    if text.isupper():
                        tag = "TABLE_HEADER"
                        style = "Table Header"
                    else:
                        tag = "TABLE_BODY"
                        style = "Table Body"
                    
                    if mode == "style":
                        try:
                            para.style = style
                        except KeyError:
                            logger.warning(f"Style '{style}' not found")
                    else:
                        if para.runs:
                            para.runs[0].text = f"[{tag}] " + para.runs[0].text
