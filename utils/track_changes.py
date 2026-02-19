
from datetime import datetime
from docx.oxml.shared import OxmlElement, qn
from docx.oxml import parse_xml

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

def add_tracked_text(paragraph, text, style=None, author="RefBot", date=None, color=None, doc=None):
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
    if style or color:
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
        
    # Append del to p
    p.append(del_tag)

def delete_tracked_run(paragraph, run, author="RefBot", date=None):
    """
    Wraps an existing `run` element in a <w:del> tag.
    Simulates "Track Changes" deletion.
    """
    if date is None:
        date = get_current_iso_time()
        
    # Create <w:del>
    # Note: Structure is <w:p> ... <w:del><w:r>...</w:r></w:del> ... </w:p>
    # We need to replace the run in the parent with the del wrapping it.
    
    p = paragraph._element
    r = run._element
    
    if r.getparent() != p:
        # Run might be inside a hyperlink or other structure?
        # If so, we skip complex nesting for now or handle carefully.
        return False
        
    # Create del
    del_tag = create_element('w:del')
    del_tag.set(qn('w:id'), get_unique_id())
    del_tag.set(qn('w:author'), author)
    del_tag.set(qn('w:date'), date)
    
    # Replace r with del_tag in p
    # Insert del_tag before r
    r.addprevious(del_tag)
    # Move r inside del_tag
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
