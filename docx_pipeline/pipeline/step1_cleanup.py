"""
pipeline/step1_cleanup.py â€” Basic DOCX cleanup.

Actions:
  - Remove fully empty paragraphs (configurable: keep max 1 trailing blank)
  - Strip leading/trailing whitespace from run text
  - Normalise non-breaking spaces to regular spaces
  - Remove duplicate consecutive blank paragraphs
"""

from docx import Document
from docx.oxml.ns import qn
from docx_pipeline.utils.report import ReportLogger

NBSP = "\u00a0"


def run(doc: Document, logger: ReportLogger) -> Document:
    logger.set_step("1-cleanup")

    paras       = doc.paragraphs
    removed     = 0
    normalised  = 0
    prev_blank  = False

    for para in paras:
        # 1. Equivalent to Macro Array("^m", "^p")
        # Remove manual page breaks (<w:br w:type="page"/>)
        for r in para.runs:
            for br in r._element.findall(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    br.getparent().remove(br)
                    normalised += 1

        # 2. Equivalent to Macro:
        # Array("  ", " ") -> Replace double space with single space
        # Array("^t^t", "^t") -> Replace double tab with single tab
        prev_char_is_space = False
        prev_char_is_tab = False
        
        for r in para.runs:
            original = r.text
            if not original:
                continue
                
            # Replace NBSP and ascii page breaks
            text = original.replace(NBSP, " ").replace("\x0c", "")
            
            while "  " in text:
                text = text.replace("  ", " ")
                
            while "\t\t" in text:
                text = text.replace("\t\t", "\t")
                
            # Handle boundary spaces/tabs between runs
            if prev_char_is_space and text.startswith(" "):
                text = text[1:]
                
            if prev_char_is_tab and text.startswith("\t"):
                text = text[1:]
                
            if text:
                prev_char_is_space = text.endswith(" ")
                prev_char_is_tab = text.endswith("\t")
            else:
                pass # preserve state if run became empty

            if text != original:
                r.text = text
                normalised += 1

        # 3. Equivalent to Macro:
        # Array("^p ", "^p") and Array("^p^t", "^p") -> Remove space/tab at the start of paragraph
        for r in para.runs:
            if r.text:
                lstripped = r.text.lstrip(" \t")
                if lstripped != r.text:
                    r.text = lstripped
                    normalised += 1
                if r.text:
                    break

        # 4. Equivalent to Macro:
        # Array(" ^p", "^p") -> Remove space at the end of paragraph
        for r in reversed(para.runs):
            if r.text:
                rstripped = r.text.rstrip(" ") # Macro specifically removes space before ^p
                if rstripped != r.text:
                    r.text = rstripped
                    normalised += 1
                if r.text:
                    break

        is_blank = not para.text.strip()

        # 5. Equivalent to Macro:
        # Array("^p^p", "^p") -> Remove duplicate consecutive blanks
        if is_blank and prev_blank:
            p_el = para._element
            p_el.getparent().remove(p_el)
            removed += 1
            continue

        prev_blank = is_blank

    logger.info(f"Cleanup: {removed} empty paras removed, "
                f"{normalised} run modifications applied.")
    return doc

