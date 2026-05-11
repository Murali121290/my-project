# annotator.py
"""
Document annotation module.
Identifies and tags document elements (headings, sections, lists, etc).
"""

import re
import logging
from typing import List, Dict, Any, Optional
from docx.document import Document
from docx.text.paragraph import Paragraph
from .rules_loader import get_rules_loader
from .logger_config import get_logger

logger = get_logger(__name__)
rules_loader = get_rules_loader()
EXPLICIT_STYLE_RE = re.compile(
    r"^(?:<\s*(/?[A-Za-z0-9.\-]+)\s*>|\[STYLE:([A-Za-z0-9.\-]+)\])",
    re.IGNORECASE,
)


# =========================
# Helpers
# =========================

def is_list_paragraph(paragraph: Paragraph) -> bool:
    """
    Detect if paragraph is a list item (has numbering or bullets).
    
    Args:
        paragraph: python-docx Paragraph object
    
    Returns:
        True if paragraph has list formatting
    """
    try:
        p = paragraph._p
        if p.pPr is not None and p.pPr.numPr is not None:
            return True
        # Check if list formatting is inherited from the style
        if paragraph.style is not None and hasattr(paragraph.style, '_element'):
            from docx.oxml.ns import qn
            style_pPr = paragraph.style._element.find(qn('w:pPr'))
            if style_pPr is not None and style_pPr.find(qn('w:numPr')) is not None:
                return True
        return False
    except Exception as e:
        logger.debug(f"Error checking list formatting: {e}")
        return False


def get_word_list_type(para: Paragraph, doc: Document) -> Optional[str]:
    """
    Determine if a Word-formatted list paragraph is bulleted or numbered
    by querying the document's numbering.xml part.
    """
    try:
        p = para._p
        if p.pPr is None or p.pPr.numPr is None:
            return None
            
        num_id = p.pPr.numPr.numId.val
        ilvl = p.pPr.numPr.ilvl.val if p.pPr.numPr.ilvl is not None else 0
        
        numbering_part = doc.part.numbering_part
        if not numbering_part:
            return None
            
        # 1. Map numId to abstractNumId
        num_element = numbering_part.element.find(f'.//w:num[@w:numId="{num_id}"]', numbering_part.element.nsmap)
        if num_element is None:
            return None
            
        abstract_num_id_element = num_element.find('.//w:abstractNumId', numbering_part.element.nsmap)
        if abstract_num_id_element is None:
            return None
        abstract_num_id = abstract_num_id_element.get(f'{{{numbering_part.element.nsmap["w"]}}}val')
        
        # 2. Find abstractNum definition and the specific level
        abstract_num = numbering_part.element.find(f'.//w:abstractNum[@w:abstractNumId="{abstract_num_id}"]', numbering_part.element.nsmap)
        if abstract_num is None:
            return None
            
        lvl = abstract_num.find(f'.//w:lvl[@w:ilvl="{ilvl}"]', numbering_part.element.nsmap)
        if lvl is None:
            return None
            
        # 3. Read numFmt (bullet, decimal, lowerRoman, etc.)
        num_fmt = lvl.find('.//w:numFmt', numbering_part.element.nsmap)
        if num_fmt is not None:
            fmt_val = num_fmt.get(f'{{{numbering_part.element.nsmap["w"]}}}val')
            if fmt_val == 'bullet':
                return 'bullet'
            elif fmt_val in ['decimal', 'lowerLetter', 'upperLetter']:
                return 'number'
            elif fmt_val in ['lowerRoman', 'upperRoman']:
                return 'roman'
    except Exception as e:
        logger.debug(f"Error determining list type from XML: {e}")
        
    return None


def is_references_end(text: str) -> bool:
    """
    Detect end of references section.
    
    Args:
        text: Paragraph text to check
    
    Returns:
        True if text indicates end of references
    """
    if not text:
        return False
    
    text = text.strip()
    patterns = [
        r"^(?:FIGURE|Figure)\s+\d+",
        r"^(?:TABLE|Table)\s+\d+",
        r"^(?:BOX|Box)\s+\d+",
        r"^[A-Z][A-Z\s]{3,}$",
    ]
    
    return any(re.match(pattern, text) for pattern in patterns)


def detect_list_kind(text: str, style_name: str = "", is_word_list: bool = False, para: Optional[Paragraph] = None, doc: Optional[Document] = None) -> Optional[str]:
    """
    Detect the list type for a paragraph.

    Returns one of: "bullet", "number", "roman", or None.
    """
    list_config = rules_loader.get_list_patterns()
    bullet_pattern = list_config.get("bullet_pattern", r"^[•\-\–]\s+")
    number_pattern = list_config.get("number_pattern", r"^\d+\.\s+")
    roman_pattern = list_config.get("roman_pattern", r"^(?i:[ivxlcdm]+)[\.)]\s+")

    text = (text or "").strip()
    style_name = style_name or ""
    lowered_style = style_name.lower()

    if re.match(bullet_pattern, text) or "bullet" in lowered_style:
        return "bullet"
    if re.match(number_pattern, text) or re.match(r"^\s*\d+[\.\)]\s+", text) or "number" in lowered_style:
        return "number"
    if re.match(roman_pattern, text) or "roman" in lowered_style:
        return "roman"
    
    if is_word_list and para is not None and doc is not None:
        xml_type = get_word_list_type(para, doc)
        if xml_type:
            return xml_type
        # Fallback to number if xml parsing fails but it is a word list
        return "number"
        
    return None


def parse_leading_style_hint(text: str) -> tuple[Optional[str], Optional[str], str]:
    """
    Parse a leading explicit tag or [STYLE:] hint.

    Returns:
        Tuple of (full_match, token, stripped_text)
    """
    text = text or ""
    match = EXPLICIT_STYLE_RE.match(text.strip())
    if not match:
        return None, None, text.strip()

    token = match.group(1) or match.group(2)
    stripped = text.strip()[len(match.group(0)):].strip()
    return match.group(0), token, stripped


def normalize_style_token(token: Optional[str], context_kind: Optional[str] = None) -> Optional[str]:
    """
    Normalize incoming source styles and explicit tags to canonical styles.
    """
    if not token:
        return None

    token = token.strip()
    cfg = rules_loader.get_normalization_config()
    explicit_map = cfg.get("explicit_tag_map", {})
    source_map = cfg.get("source_style_map", {})

    if token in explicit_map:
        return explicit_map[token]
    if token in source_map:
        return source_map[token]

    if re.match(r"^FIG\d+(?:\.\d+)?$", token, re.IGNORECASE):
        return "PMI"

    box_cfg = rules_loader.get_box_config()
    open_patterns = box_cfg.get("open_patterns", [])
    close_patterns = box_cfg.get("close_patterns", [])
    if any(re.match(pattern, token, re.IGNORECASE) for pattern in open_patterns):
        return box_cfg.get("open_style", "PMI")
    if any(re.match(pattern, token, re.IGNORECASE) for pattern in close_patterns):
        return box_cfg.get("close_style", "PMI")

    if context_kind == "box" and token == "TITLE":
        return box_cfg.get("title_style", "NBX1-TTL")

    return token


def classify_explicit_context(token: Optional[str]) -> Optional[str]:
    """
    Map explicit tag tokens to bounded inheritance contexts.
    """
    if not token:
        return None

    box_cfg = rules_loader.get_box_config()
    if any(re.match(pattern, token, re.IGNORECASE) for pattern in box_cfg.get("open_patterns", [])):
        return "objective" if token.upper() == "BXOBJ" else "box"

    keyterm_cfg = rules_loader.get_keyterm_config()
    if token.upper() in {style.upper() for style in keyterm_cfg.get("explicit_styles", [])}:
        return "keyterm"

    return None


def is_explicit_context_closer(token: Optional[str], context_kind: Optional[str]) -> bool:
    """Check whether a token closes the current explicit context."""
    if not token or not context_kind:
        return False

    if context_kind == "box":
        return any(
            re.match(pattern, token, re.IGNORECASE)
            for pattern in rules_loader.get_box_config().get("close_patterns", [])
        )
    if context_kind == "objective":
        return token.upper() in {"/BXOBJ", "/BX"}
    if context_kind == "keyterm":
        return token.upper() in {"/KT"}
    return False


def get_general_list_tag_style(list_kind: str) -> tuple[str, str]:
    """Resolve general list tag/style from rules.yaml."""
    list_config = rules_loader.get_list_patterns()
    key_map = {
        "bullet": ("general_bulleted", ("BL-MID", "BL-MID")),
        "number": ("general_numbered", ("NL-MID", "NL-MID")),
        "roman": ("general_roman", ("RL-MID", "RL-MID")),
    }
    config_key, default = key_map.get(list_kind, ("general", ("BL-MID", "BL-MID")))
    config = list_config.get(config_key, {})
    return config.get("tag", default[0]), config.get("style", default[1])


def get_reference_entry_tag_style(text: str, list_kind: Optional[str] = None, para: Optional[Paragraph] = None) -> tuple[str, str]:
    """
    Classify a reference entry as numbered or unnumbered/author-year.

    Args:
        text: Paragraph text inside the references block
        list_kind: The type of list formatting applied (if any)
        para: The Paragraph object

    Returns:
        Tuple of (tag, style)
    """
    text = (text or "").strip()
    
    # 1. Use the explicit numbered list detection logic
    is_numbered = bool(re.match(r'^\[?\d+\]?[\.\)\t\s]', text))
    if not is_numbered and para is not None:
        try:
            from docx.oxml.ns import qn
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:numPr')) is not None:
                is_numbered = True
        except Exception:
            pass
            
    if is_numbered or list_kind == "number":
        return "REF-N", "REF-N"

    author_year_patterns = [
        r"^[A-Z][A-Za-z'`\-]+(?:,\s*(?:[A-Z]\.\s*)+).*\(\d{4}[a-z]?\)",
        r"^[A-Z][A-Za-z'`\-]+(?:\s+et\s+al\.)?.*\(\d{4}[a-z]?\)",
        r"^[A-Z][A-Za-z'`\-]+(?:,\s*[A-Z][A-Za-z'`\-]+)*(?:\s*&\s*[A-Z][A-Za-z'`\-]+)?.*\b\d{4}[a-z]?\b",
    ]
    if any(re.match(pattern, text) for pattern in author_year_patterns):
        return "REF-U", "REF-U"

    return "REF-U", "REF-U"


# =========================
# Core annotator
# =========================

def annotate_document(doc: Document) -> List[Dict[str, Any]]:
    """
    Annotate all paragraphs in a document with tags and styles.
    
    Args:
        doc: python-docx Document object
    
    Returns:
        List of annotation dictionaries with keys:
        - para: Paragraph object
        - tag: String tag (e.g., "CHAPTER_TITLE", "BODY_TEXT")
        - style: Word style name (e.g., "Heading 1", "Normal")
    
    Raises:
        ValueError: If document is invalid
    """
    if not doc or not hasattr(doc, 'paragraphs'):
        raise ValueError("Invalid Document object")
    
    annotations: List[Dict[str, Any]] = []
    current_block: Optional[str] = None
    block_item_count: int = 0
    explicit_context_kind: Optional[str] = None
    in_chapter_preamble: bool = False
    previous_tag: Optional[str] = None
    
    logger.info(f"Annotating document with {len(doc.paragraphs)} paragraphs")
    
    for para_idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        tag = "TXT"
        style = "TXT"
        
        try:
            if not text:
                empty_style = rules_loader.rules.get("defaults", {}).get("empty_paragraph", {}).get("style", "TXT")
                annotations.append({"para": para, "tag": "EMPTY", "style": empty_style})
                continue
            
            # ===== PRIORITY 0: EXPLICIT TAGS <TAG> =====
            full_hint, explicit_token, stripped_text = parse_leading_style_hint(text)
            explicit_tag_found = False
            
            if explicit_token:
                normalized_explicit_style = normalize_style_token(explicit_token, explicit_context_kind)
                if normalized_explicit_style:
                    tag = normalized_explicit_style
                    style = normalized_explicit_style
                text = stripped_text
                explicit_tag_found = True
                logger.debug(f"Para {para_idx}: Found explicit tag/style '{explicit_token}' -> '{style}'")

                if is_explicit_context_closer(explicit_token, explicit_context_kind):
                    explicit_context_kind = None
                else:
                    new_context = classify_explicit_context(explicit_token)
                    if new_context:
                        explicit_context_kind = new_context

                if text in rules_loader.get_block_start_markers():
                    current_block = rules_loader.get_block_start_markers()[text]

                # Override: if text is "OBJECTIVES", set to OBJ1 regardless of explicit tag
                objectives_synonyms = ["objectives", "learning objectives", "learningobjectives", "lesson objectives"]
                if text.lower() in objectives_synonyms:
                    tag = "OBJ1"
                    style = "OBJ1"
                    logger.debug(f"Para {para_idx}: Overriding to OBJ1 because text is '{text}'")

            # ===== BLOCK START (Based on text match) =====
            block_markers = rules_loader.get_block_start_markers()
            # This logic below is for exact text matches from 'blocks: start_markers'
            if not explicit_tag_found and text in block_markers:
                current_block = block_markers[text]
                block_item_count = 0
                # Default behavior for block marker headings, can be overridden by regex rules below
                tag = current_block + "_HEADING"
                style = "H2" 
                logger.debug(f"Para {para_idx}: Detected block start: {current_block}")
            
            # ===== BLOCK END =====
            elif current_block == "REFERENCES_BLOCK" and is_references_end(text):
                current_block = None
                logger.debug(f"Para {para_idx}: References block ended")
            
            # ===== WORD / MANUAL LISTS =====
            style_name = para.style.name if para.style else ""
            list_kind = detect_list_kind(text, style_name, is_list_paragraph(para), para, doc)
            if not explicit_tag_found and list_kind is not None:
                if current_block:
                    block_item_count += 1
                if current_block == "LEARNING_OBJECTIVES_BLOCK" or explicit_context_kind == "objective":
                    if list_kind == "bullet":
                        tag = "OBJ-BL-MID"
                        style = "OBJ-BL-MID"
                    else:
                        tag = "OBJ-NL-MID"
                        style = "OBJ-NL-MID"
                elif current_block == "REFERENCES_BLOCK":
                    tag, style = get_reference_entry_tag_style(text, list_kind, para)
                else:
                    tag, style = get_general_list_tag_style(list_kind)
            
            # ===== PARAGRAPH RULES =====
            else:
                rule_matched = explicit_tag_found # If we found explicit tag, we skip regex search

                if explicit_tag_found and style in {"H1", "H2", "H3", "H4", "TXT", "TXT-FLUSH"}:
                    rule_matched = False
                
                # Priority 0.5: Detect "OBJECTIVES" heading (case-insensitive)
                objectives_synonyms = ["objectives", "learning objectives", "learningobjectives", "lesson objectives"]
                if text.lower() in objectives_synonyms:
                    tag = "OBJ1"
                    style = "OBJ1"
                    rule_matched = True
                    logger.debug(f"Para {para_idx}: Detected learning objectives heading '{text}'")

                # Priority 1: Force CT after CN
                # Only if not explicit (Explicit tags override strict sequencing if present)
                if not explicit_tag_found and previous_tag == "CN":
                    tag = "CT"
                    style = "CT"
                    rule_matched = True
                    logger.debug(f"Para {para_idx}: Forced CT due to previous CN")

                # Priority 2: Regex Rules
                if not rule_matched:
                    paragraph_rules = rules_loader.get_paragraphs()
                    for rule in paragraph_rules:
                        try:
                            if re.match(rule["pattern"], text):
                                tag = rule["tag"]
                                style = rule["style"]
                                rule_matched = True
                                
                                # Special Logic: Check if this matched rule is actually starting a block
                                if text in block_markers:
                                     current_block = block_markers[text]
                                     logger.debug(f"Para {para_idx}: Matched block start rule '{tag}', set block to {current_block}")

                                logger.debug(f"Para {para_idx}: Matched rule '{tag}'")
                                break
                        except re.error as e:
                            logger.error(f"Invalid regex in rule '{rule.get('tag', 'unknown')}': {e}")
                
                # ===== BLOCK RESET LOGIC (Moved out to run for ALL tags) =====
                # If we hit a new section heading, exit the current block
                # This now applies to Explicit Tags, Regex matches, or Forced tags
                if tag in ("H1", "H2", "H3", "H4", "CT", "CN", "OBJ1", "REFH1"):
                    if tag == "OBJ1":
                        current_block = "LEARNING_OBJECTIVES_BLOCK"
                        block_item_count = 0
                    elif tag == "REFH1":
                        current_block = "REFERENCES_BLOCK"
                        block_item_count = 0
                    else:
                        is_current_block_starter = False
                        if current_block:
                            if text in block_markers and block_markers[text] == current_block:
                                is_current_block_starter = True
                        
                        if not is_current_block_starter:
                            current_block = None
                            logger.debug(f"Para {para_idx}: Block reset due to heading '{tag}'")

                # Priority 3: Block Context Text
                if not rule_matched and current_block:
                    block_item_count += 1
                    # Map block names to shorter text tags if needed
                    if current_block == "LEARNING_OBJECTIVES_BLOCK":
                        if block_item_count == 1:
                            tag = "OBJ-TXT-FIRST"
                            style = "OBJ-TXT-FIRST"
                        else:
                            tag = "OBJ-TXT"
                            style = "OBJ-TXT"
                    elif current_block == "REFERENCES_BLOCK":
                        tag, style = get_reference_entry_tag_style(text, list_kind, para)
                    else:
                        tag = current_block + "_TXT"
                        style = tag
                elif not rule_matched and explicit_context_kind == "objective":
                    tag = "OBJ-TXT"
                    style = "OBJ-TXT"
                elif (
                    not rule_matched
                    and explicit_context_kind == "box"
                    and text
                    and style == "TXT"
                ):
                    box_cfg = rules_loader.get_box_config()
                    if style_name == "TITLE" or re.match(r"^Box\s+\d+", text):
                        tag = box_cfg.get("title_style", "NBX1-TTL")
                        style = tag
                    else:
                        tag = box_cfg.get("body_style", "NBX-TXT")
                        style = tag
                elif not rule_matched and explicit_context_kind == "keyterm":
                    tag = "KT"
                    style = "KT"
            
            # Priority 3.5: Epigraph Context Logic
            if tag in ("CN", "CT", "CAU"):
                in_chapter_preamble = True
            elif tag in ("H1", "H2", "H3", "OBJ1", "ABS"):
                in_chapter_preamble = False
            
            if in_chapter_preamble and tag in ("TXT", "EPI-ATT"): 
                # If we matched EPI-ATT via regex, keep it. 
                # If TXT, check heuristics.
                if tag == "TXT":
                    if text.startswith("“") or text.startswith('"'):
                        tag = "EPI"
                        style = "EPI"
                    elif re.match(r"^[—–-]", text):
                         tag = "EPI-ATT"
                         style = "EPI-ATT"
                    else:
                         # Default text in preamble (after Title/Author, before H1) is likely Epigraph
                         tag = "EPI"
                         style = "EPI"
            
            # Priority 4: TXT-FLUSH Logic
            if tag == "TXT" and previous_tag in ("H1", "H2", "H3"):
                tag = "TXT-FLUSH"
                style = "TXT-FLUSH"
                logger.debug(f"Para {para_idx}: Changed TXT to TXT-FLUSH (previous: {previous_tag})")

            # Update history
            if tag != "EMPTY":
                previous_tag = tag

            annotations.append({"para": para, "tag": tag, "style": style})
        
        except Exception as e:
            logger.error(f"Error annotating paragraph {para_idx}: {e}", exc_info=True)
            annotations.append({"para": para, "tag": "TXT", "style": "TXT"})
    
    logger.info(f"Document annotation complete: {len(annotations)} paragraphs annotated")
    return annotations

