from datetime import datetime
from docx.oxml.shared import OxmlElement, qn
from docx.oxml import parse_xml
import difflib
import logging
logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler()) # Ensure it doesn't print if not configured.

from docx import Document # Add this import
from docx.text.paragraph import Paragraph # Add this import

# XML Namespaces
nsmap = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

def create_element(name):
    return OxmlElement(name)

def get_current_iso_time():
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

import random

def get_unique_id():
    # Return a random integer as string, strictly within 32-bit signed int range (except 0)
    # Word IDs usually should be unique.
    return str(random.randint(1, 2147483647))

def add_tracked_text(paragraph, text, style=None, author="S4c", date=None, color=None, doc=None, superscript=False):
    """
    Appends a new run with `text` inside a <w:ins> element to `paragraph`.
    Simulates "Track Changes" insertion.

    Args:
        paragraph: The paragraph to append to
        text: The text to insert
        style: Style name (string) or Style object
        author: Author name for track changes
        date: ISO format date string
        color: Color value for text
        doc: Document object (used to resolve style names to style IDs)
        superscript: If True, makes the text superscript
    """
    if not text:
        return None

    if date is None:
        date = get_current_iso_time()

    # Create <w:ins>
    ins = create_element('w:ins')
    ins.set(qn('w:id'), get_unique_id())
    ins.set(qn('w:author'), author)
    ins.set(qn('w:date'), date)

    # Create Run
    run = create_element('w:r')

    # Add Text
    t = create_element('w:t')
    t.text = text
    # Preserve whitespace if needed
    if text.strip() != text or ' ' in text:
         t.set(qn('xml:space'), 'preserve')

    # Add Style/Properties
    if style or color or superscript:
        rPr = create_element('w:rPr')

        if style:
            rStyle = create_element('w:rStyle')
            # Resolve style name to style ID using doc.styles if available
            style_id = None

            if hasattr(style, 'style_id'):
                # It's already a Style object
                style_id = style.style_id
            elif doc and doc.styles:
                # Try to look up the style by name in doc.styles
                try:
                    style_obj = doc.styles[str(style)]
                    style_id = style_obj.style_id
                except:
                    # Fallback: use the style string as-is
                    style_id = str(style)
            else:
                # No doc provided, use the style string as-is
                style_id = str(style)

            rStyle.set(qn('w:val'), style_id)
            rPr.append(rStyle)

        if color:
            c = create_element('w:color')
            c.set(qn('w:val'), color)
            rPr.append(c)

        if superscript:
            vertAlign = create_element('w:vertAlign')
            vertAlign.set(qn('w:val'), 'superscript')
            rPr.append(vertAlign)

        run.append(rPr)

    run.append(t)

    # Append run to ins
    ins.append(run)

    # Append to paragraph
    paragraph._element.append(ins)
    return ins

def wrap_paragraph_content_in_del(paragraph, author="RefBot", date=None):
    """
    Moves ALL existing children (runs, hyperlinks) of a paragraph into a <w:del> tag.
    Excludes pPr (properties).
    """
    if date is None:
        date = get_current_iso_time()
        
    p = paragraph._element
    
    # Create del container
    del_tag = create_element('w:del')
    del_tag.set(qn('w:id'), get_unique_id())
    del_tag.set(qn('w:author'), author)
    del_tag.set(qn('w:date'), date)
    
    # Identify children to move (runs, hyperlinks, etc.)
    # Exclude pPr
    children_to_move = []
    for child in p:
        if child.tag.endswith('pPr'):
            continue
        children_to_move.append(child)
        
    if not children_to_move:
        return

    # Move children
    # Must remove from p first, then append to del
    # But wait, python-docx elements are proxies. We need to work with lxml elements directly mostly.
    
    for child in children_to_move:
        p.remove(child)
        del_tag.append(child)

    # Rename <w:t> to <w:delText> inside all moved runs (OOXML spec compliance)
    for r_elem in del_tag:
        for t_elem in r_elem.findall(qn('w:t')):
            t_elem.tag = qn('w:delText')

    # Append del to p
    p.append(del_tag)

def delete_tracked_run(paragraph, run, author="RefBot", date=None):
    """
    Wraps an existing `run` element in a <w:del> tag.
    Simulates "Track Changes" deletion.
    """
    if date is None:
        date = get_current_iso_time()
        
    p = paragraph._element
    r = run._element
    parent = r.getparent()
    
    if parent.tag == qn('w:ins'):
        # If it is inside an insertion, deleting an insertion just removes it
        parent.remove(r)
        if len(parent) == 0:
            parent.getparent().remove(parent)
        return True
        
    # Create del
    del_tag = create_element('w:del')
    del_tag.set(qn('w:id'), get_unique_id())
    del_tag.set(qn('w:author'), author)
    del_tag.set(qn('w:date'), date)
    
    # Replace r with del_tag in its parent
    r.addprevious(del_tag)
    del_tag.append(r)
    
    # Text inside w:del should be w:delText, not w:t
    for child in r:
        if child.tag == qn('w:t'):
            child.tag = qn('w:delText')
    
    return True

def add_tracked_deletion(paragraph, text, author="RefBot", date=None, doc=None):
    """
    Appends a NEW <w:del> element containing `text` to `paragraph`.
    Useful when reconstructing a paragraph from diffs (inserting 'deleted' history).
    
    Args:
        paragraph: The paragraph to append to
        text: The text to delete
        author: Author name for track changes
        date: ISO format date string
        doc: Document object (accepted for consistency with add_tracked_text, not used for deletions)
    """
    if not text:
        return None
        
    if date is None:
        date = get_current_iso_time()
    
    # Create <w:del>
    del_tag = create_element('w:del')
    del_tag.set(qn('w:id'), get_unique_id())
    del_tag.set(qn('w:author'), author)
    del_tag.set(qn('w:date'), date)
    
    # Create Run inside del
    # Note: Del contains Run contains Text
    run = create_element('w:r')
    t = create_element('w:delText')
    t.text = text
    if text.strip() != text or ' ' in text:
         t.set(qn('xml:space'), 'preserve')
    
    run.append(t)
    del_tag.append(run)
    
    # Append to paragraph
    paragraph._element.append(del_tag)
    return del_tag

def add_tracked_run(paragraph, text, style=None, author="RefBot", date=None, color=None, doc=None):
    return add_tracked_text(paragraph, text, style, author, date, color, doc)

def apply_tracked_changes_to_paragraph(
    paragraph: Paragraph,
    original_text: str, # The text content of this paragraph BEFORE changes (for diffing)
    revised_text: str,  # The final text AFTER changes (what should appear in the document)
    author: str = "RefBot",
    date: datetime = None,
    doc: Document = None # Required for style lookup in add_tracked_text
):
    """
    Applies changes to a paragraph using Word's track changes functionality (insertions and deletions).
    Compares original_text with revised_text and marks differences in the paragraph's XML.

    Args:
        paragraph: The python-docx Paragraph object to modify.
        original_text: The string content of the paragraph *before* any AI processing or stripping.
                       Used as the baseline for diffing.
        revised_text: The new, AI-processed string content that should appear in the paragraph.
        author: The author name for track changes.
        date: The timestamp for track changes.
        doc: The Document object, required for accurate style resolution in tracked insertions.
    """
    if date is None:
        date = get_current_iso_time()
    if doc is None and paragraph.document:
        doc = paragraph.document
    if doc is None:
        logger.warning(f"Document object not provided for paragraph track changes. Styles might not resolve correctly.")

    p_element = paragraph._element
    
    # Remove all existing children from the paragraph's element to prepare for reconstruction
    # Keep pPr (paragraph properties) intact
    for child in list(p_element.iterchildren()): # Use list() to avoid modification during iteration
        if child.tag != qn('w:pPr'): 
            p_element.remove(child)

    # Perform difflib comparison
    matcher = difflib.SequenceMatcher(None, original_text, revised_text)
    
    # Apply changes based on diff operations
    for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
        original_sub_text = original_text[i1:i2]
        revised_sub_text = revised_text[j1:j2]

        if opcode == 'equal':
            # For consistency, even equal parts are inserted as new if we cleared existing runs.
            # This results in the entire paragraph being marked as an insertion.
            if revised_sub_text:
                add_tracked_text(paragraph, revised_sub_text, author=author, date=date, doc=doc)
        elif opcode == 'replace':
            # Mark original part as deleted, insert new part as inserted
            if original_sub_text:
                add_tracked_deletion(paragraph, original_sub_text, author=author, date=date, doc=doc)
            if revised_sub_text:
                add_tracked_text(paragraph, revised_sub_text, author=author, date=date, doc=doc)
        elif opcode == 'delete':
            # Mark original part as deleted
            if original_sub_text:
                add_tracked_deletion(paragraph, original_sub_text, author=author, date=date, doc=doc)
        elif opcode == 'insert':
            # Insert new part as inserted
            if revised_sub_text:
                add_tracked_text(paragraph, revised_sub_text, author=author, date=date, doc=doc)
