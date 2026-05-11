"""
pipeline/step2_unused_styles.py â€” Remove unused custom styles.

Collects every style actually used by paragraphs, runs, and tables
across all document parts (main, headers, footers, footnotes, endnotes),
then removes declared custom styles that are not in that set.
Built-in Word styles (where w:customStyle="1" is not set) are never removed.
"""

from docx import Document
from docx.oxml.ns import qn
from docx_pipeline.utils.report import ReportLogger


def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("2-unused-styles")

    # 1. Collect all styleIds explicitly in use across all XML parts of the document
    used_style_ids = set()

    for part in doc.part.package.parts:
        # Check all XML parts for style referencing tags
        if part.content_type.endswith('+xml'):
            try:
                element = part.element
                for node_name in ('w:pStyle', 'w:rStyle', 'w:tblStyle'):
                    for style_node in element.iter(qn(node_name)):
                        style_id = style_node.get(qn('w:val'))
                        if style_id:
                            used_style_ids.add(style_id)
            except Exception:
                pass # Ignore parts without parsable elements

    # 2. Build dependency tree to protect styles that used styles are based-on or linked to
    styles_el = doc.styles.element
    
    style_elements = {}
    for style_el in styles_el.findall(qn('w:style')):
        style_id = style_el.get(qn('w:styleId'))
        if style_id:
            style_elements[style_id] = style_el
            
    # Recursive function to trace all indirect style usages
    def trace_dependencies(sid, visited):
        if sid in visited or sid not in style_elements:
            return
        visited.add(sid)
        el = style_elements[sid]
        
        for dep_tag in ('w:basedOn', 'w:next', 'w:link'):
            dep_node = el.find(qn(dep_tag))
            if dep_node is not None:
                dep_id = dep_node.get(qn('w:val'))
                if dep_id:
                    trace_dependencies(dep_id, visited)
                
    # Seed the visited set with all explicitly used styles and traverse
    explicitly_used = list(used_style_ids)
    for sid in explicitly_used:
        trace_dependencies(sid, used_style_ids)

    # 3. Walk styles XML and remove unused custom styles
    removed = []
    old_style_count = len(style_elements)
    
    for style_el in list(styles_el):
        if style_el.tag != qn("w:style"):
            continue
            
        style_id = style_el.get(qn("w:styleId"))
        if not style_id:
            continue
            
        # As in the VBA macro (.BuiltIn = False), we only delete CUSTOM styles.
        # Custom styles have the attribute w:customStyle="1".
        # If this is missing or "0", it's a built-in standard style.
        is_custom = style_el.get(qn("w:customStyle")) == "1"
        if not is_custom:
            continue
            
        name_str = style_id
        name_el = style_el.find(qn("w:name"))
        if name_el is not None:
            name_str = name_el.get(qn("w:val"), style_id)

        if style_id not in used_style_ids:
            styles_el.remove(style_el)
            removed.append(name_str)

    new_style_count = len(styles_el.findall(qn('w:style')))

    # 4. Log the results mapped to what the macro outputs
    if removed:
        logger.info(
            f"Removed {len(removed)} unused styles "
            f"(Old Style Count: {old_style_count}, New Style Count: {new_style_count}): "
            + ", ".join(removed)
        )
    else:
        logger.info(
            f"No unused styles removed. "
            f"Old Style Count: {old_style_count}, New Style Count: {new_style_count}"
        )

    return doc

