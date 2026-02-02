"""
Bias Detection Scanner Module
Adapted from bias.py for integration with app_server.py
"""
import os
import csv
import re
import math
import copy
import zipfile
import subprocess
import shutil
from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell
from openpyxl import Workbook

try:
    import pdfplumber
except ImportError as e:
    print(f"DEBUG: pdfplumber Import Failed: {e}")
    pdfplumber = None

WORDS_PER_PAGE = 160


def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object.
    """
    if isinstance(parent, DocumentObject):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, Table):
        # Tables don't have direct block content usually, they have rows/cells.
        # But for recursive robustness:
        return
    else:
        raise ValueError("Unknown parent type")

    for child in parent_elm.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)


def count_page_breaks(paragraph):
    """
    Counts explicit page breaks (w:br type=page) and rendered page breaks (w:lastRenderedPageBreak)
    within a paragraph.
    """
    xml = paragraph._element
    
    # Rendered Page Breaks (calculated by Word)
    rendered_breaks = xml.xpath('.//w:lastRenderedPageBreak')
    
    # Explicit Page Breaks (hard breaks)
    explicit_breaks = xml.xpath('.//w:br[@w:type="page"]')
    
    return len(rendered_breaks) + len(explicit_breaks)


def load_bias_terms(csv_path="bias_terms.csv"):
    """Load bias terms from CSV file"""
    categories = set()
    term_category = {}
    if not os.path.exists(csv_path):
        return {}, []

    with open(csv_path, encoding="utf-8-sig", errors="replace") as f:
        # Handle potential whitespace in headers
        reader_raw = csv.reader(f)
        try:
            headers = next(reader_raw)
            headers = [h.strip() for h in headers]
        except StopIteration:
            return {}, []  # Empty file

        reader = csv.DictReader(f, fieldnames=headers)
        
        for row in reader:
            term = row.get("term", "").strip()
            category = row.get("category", "").strip()
            
            if term:
                # Normalize internal spaces (e.g. "term  one" -> "term one")
                term = re.sub(r'\s+', ' ', term)
                term_category[term] = category
                if category:
                    categories.add(category)
    return term_category, sorted(list(categories))


def estimate_page(word_index):
    return max(1, math.ceil(word_index / WORDS_PER_PAGE))


def apply_highlights(para, highlights):
    """
    Applies highlights to a paragraph by modifying runs.
    highlights: list of (start, end, color_index) tuples.
    """
    if not highlights:
        return
    
    # Resolving color map
    text_len = len(para.text)
    colors = [None] * text_len
    
    # Apply based on priority (Last one wins, so sort Yellow to be last if we want it on top)
    # Priority: Yellow (7) > Bright Green (4)
    # Sort so low priority comes first
    highlights.sort(key=lambda x: 1 if x[2] == WD_COLOR_INDEX.YELLOW else 0)
    
    for start, end, color in highlights:
        # Clamp to bounds
        s = max(0, start)
        e = min(text_len, end)
        for k in range(s, e):
            colors[k] = color
            
    # Iterate runs and apply/split
    current_idx = 0
    i = 0
    # Use a while loop as we might add runs
    while i < len(para.runs):
        run = para.runs[i]
        run_text = run.text
        if not run_text:
            i += 1
            continue
            
        run_len = len(run_text)
        
        # Extract colors for this run
        run_colors = colors[current_idx : current_idx + run_len]
        
        # Group contiguous colors
        from itertools import groupby
        segments = []
        for color, group in groupby(run_colors):
            segments.append((color, len(list(group))))
            
        if not segments:
             current_idx += run_len
             i += 1
             continue

        # If uniform color
        if len(segments) == 1:
            color = segments[0][0]
            if color:
                run.font.highlight_color = color
            i += 1
        else:
            # Split run
            current_run_element = run._element
            parent = current_run_element.getparent()
            
            # 1. Reuse existing run for the first segment
            first_color, first_len = segments[0]
            run.text = run_text[:first_len]
            if first_color:
                run.font.highlight_color = first_color
            else:
                run.font.highlight_color = None
                
            # 2. Insert new runs for subsequent segments
            remaining_text = run_text[first_len:]
            seg_offset = 0
            
            insert_point_index = parent.index(current_run_element)
            
            for color, length in segments[1:]:
                seg_text = remaining_text[seg_offset : seg_offset + length]
                
                # Clone
                new_run_element = copy.deepcopy(current_run_element)
                new_run = type(run)(new_run_element, parent) 
                
                # Reset text and specific style
                new_run.text = seg_text
                if color:
                    new_run.font.highlight_color = color
                else:
                    new_run.font.highlight_color = None 
                
                # Insert
                insert_point_index += 1
                parent.insert(insert_point_index, new_run_element)
                
                seg_offset += length
            
            # Skip the newly added runs
            i += len(segments)
            
        current_idx += run_len


def get_context_range(text, start_idx, end_idx, word_limit=5):
    # Find start index
    count = 0
    s = start_idx
    while s > 0 and count < word_limit:
        s -= 1
        if text[s].isspace() and (s+1 < len(text) and not text[s+1].isspace()):
             count += 1
    
    if s > 0 or (s == 0 and text[s].isspace()):
         s += 1
         
    # Find end index
    count = 0
    e = end_idx
    while e < len(text) and count < word_limit:
        if text[e].isspace() and (e-1 >= 0 and not text[e-1].isspace()):
             count += 1
        e += 1
        
    return s, e


def convert_to_pdf(docx_path):
    """
    Converts DOCX to PDF using LibreOffice (soffice).
    Returns path to PDF if successful, else None.
    """
    pdf_out_dir = os.path.join("outputs", "pdf")
    os.makedirs(pdf_out_dir, exist_ok=True)
    
    # Check for soffice
    soffice_cmd = "soffice"
    if os.name == 'nt':
        # Check standard installation paths if not in PATH
        if not shutil.which("soffice"):
            possible_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            for p in possible_paths:
                if os.path.exists(p):
                    soffice_cmd = p
                    break
    
    print(f"DEBUG: Using soffice command: '{soffice_cmd}'")
    
    # Use absolute paths for LibreOffice to avoid CWD ambiguity
    abs_docx_path = os.path.abspath(docx_path)
    abs_pdf_out_dir = os.path.abspath(pdf_out_dir)

    try:
        subprocess.run([
            soffice_cmd,
            "--headless",
            "--convert-to", "pdf",
            abs_docx_path,
            "--outdir", abs_pdf_out_dir
        ], check=True, capture_output=True)
        
        filename = os.path.basename(docx_path)
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        abs_pdf_path = os.path.join(abs_pdf_out_dir, pdf_name)
        
        if os.path.exists(abs_pdf_path):
            print(f"DEBUG: PDF Generated at {abs_pdf_path}")
            return abs_pdf_path
        else:
             print(f"DEBUG: PDF Missing at {abs_pdf_path}")
             if os.path.exists(abs_pdf_out_dir):
                 print(f"DEBUG: Dir contents: {os.listdir(abs_pdf_out_dir)}")
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"DEBUG: PDF Conversion Error: {e}")
        if hasattr(e, 'stderr'):
            print(f"DEBUG: Stderr: {e.stderr}")
        return None
    return None


def find_page_number_in_pdf(pdf_obj, sentence_text):
    """
    Searches for sentence text in PDF pages using the open pdf_obj.
    Returns 1-based page number or None.
    """
    if not pdf_obj:
        return None
        
    try:
        # Aggressive cleaning: remove all non-alphanumeric chars and lowercase
        clean_text = re.sub(r'[\W_]+', '', sentence_text).lower()
        search_query = clean_text[:60]  # Use significant prefix
        
        for i, page in enumerate(pdf_obj.pages):
            text = page.extract_text() or ""
            # Normalize PDF text same way
            clean_page = re.sub(r'[\W_]+', '', text).lower()
            
            if search_query in clean_page:
                return i + 1
    except Exception:
        pass
    return None


def scan_paragraph(para, term_category_map, report_rows, filename, page_info, pdf_obj=None):
    """
    page_info: dict {'current_xml_page': int, 'current_word_count': int, 'use_xml': bool}
    """
    text = para.text
    if not text.strip():
        return

    # Find sentences offsets
    sentences = []
    for match in re.finditer(r'[^.!?]+[.!?]*', text):
        s_text = match.group()
        if s_text.strip():
            sentences.append((match.start(), match.end(), s_text))
    
    highlights = []
    
    base_page_num = 1
    if page_info['use_xml']:
        base_page_num = page_info['current_xml_page']
    else:
        base_page_num = estimate_page(page_info['current_word_count'])

    for start, end, sentence in sentences:
        sentence_lower = sentence.lower()
        
        # Resolve Page Number specific to this sentence
        final_page_num = base_page_num
        
        if pdf_obj:
             # Try to find exact page in PDF
             pdf_page = find_page_number_in_pdf(pdf_obj, sentence)
             if pdf_page:
                 final_page_num = pdf_page
        
        # Find terms
        matched_terms = []
        for term, category in term_category_map.items():
            
            # Determine Case Sensitivity
            is_case_sensitive = term.isupper()
            
            if is_case_sensitive:
                if term in sentence:
                     for m in re.finditer(rf"\b{re.escape(term)}\b", sentence):
                         matched_terms.append((m.start(), m.end(), term, category))
            else:
                if term.lower() in sentence_lower:
                    for m in re.finditer(rf"\b{re.escape(term)}\b", sentence, re.IGNORECASE):
                         matched_terms.append((m.start(), m.end(), term, category))

        if matched_terms:
            # Deduplicate by span
            unique_matches = {}
            for m in matched_terms:
                span = (m[0], m[1])
                if span not in unique_matches:
                    unique_matches[span] = m
            
            deduped_terms = sorted(list(unique_matches.values()), key=lambda x: x[0])

            for t_start, t_end, term, category in deduped_terms:
                # Calculate Context Range for this term
                c_start, c_end = get_context_range(sentence, t_start, t_end, 5)
                
                abs_c_start = start + c_start
                abs_c_end = start + c_end
                
                # 1. Highlight Context Range TURQUOISE
                highlights.append((abs_c_start, abs_c_end, WD_COLOR_INDEX.TURQUOISE))
            
                abs_start = start + t_start
                abs_end = start + t_end
                
                # 2. Highlight Terms YELLOW
                highlights.append((abs_start, abs_end, WD_COLOR_INDEX.YELLOW))
                
                report_rows.append([
                    filename,
                    final_page_num,  # PDF or Hybrid Page Number
                    category,
                    term,
                    sentence.strip()
                ])

    if highlights:
        apply_highlights(para, highlights)


def scan_recursive(parent, term_category_map, report_rows, filename, page_info, pdf_obj=None):
    """
    Recursively scans blocks.
    page_info: dict mutable reference
    """
    for block in iter_block_items(parent):
        if isinstance(block, Paragraph):
            scan_paragraph(block, term_category_map, report_rows, filename, page_info, pdf_obj)
            
            # Update XML page count
            page_info['current_xml_page'] += count_page_breaks(block)
            
            # Update Word count
            text = block.text
            if text.strip():
                 # Simple word count approximation
                 words = len(re.findall(r'\w+', text))
                 page_info['current_word_count'] += words

        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    scan_recursive(cell, term_category_map, report_rows, filename, page_info, pdf_obj)


def scan_docx(filepath, term_category_map, output_dir):
    """
    Scan a DOCX file for bias terms and return paths to output files.
    
    Args:
        filepath: Path to input DOCX file
        term_category_map: Dictionary mapping terms to categories
        output_dir: Directory to save output files
        
    Returns:
        tuple: (highlighted_docx_path, report_rows)
    """
    doc = Document(filepath)
    filename = os.path.basename(filepath)
    
    # Try converting to PDF for exact page numbers
    pdf_path = convert_to_pdf(filepath)
    pdf_obj = None

    if pdf_path and pdfplumber:
        try:
            pdf_obj = pdfplumber.open(pdf_path)
            print(f"PDF Opened: {pdf_path} with {len(pdf_obj.pages)} pages")
        except Exception as e:
            print(f"Failed to open PDF: {e}")
    else:
        print("PDF Conversion Failed or LibreOffice not found. Falling back to Hybrid Estimation.")

    # Check for pagination tags to decide strategy
    body_xml = doc.element.body
    
    soft_breaks = len(body_xml.xpath('.//w:lastRenderedPageBreak'))
    hard_breaks = len(body_xml.xpath('.//w:br[@w:type="page"]'))
    
    use_xml_paging = (soft_breaks > 0 or hard_breaks > 0)
    
    # State object
    page_info = {
        'current_xml_page': 1,
        'current_word_count': 0,
        'use_xml': use_xml_paging
    }

    report_rows = []
    
    try:
        scan_recursive(doc, term_category_map, report_rows, filename, page_info, pdf_obj)
        
        if hasattr(doc.part, "footnotes_part") and doc.part.footnotes_part:
            for footnote in doc.part.footnotes_part.footnotes:
                 for para in footnote.paragraphs:
                     scan_paragraph(para, term_category_map, report_rows, filename, page_info, None)
    finally:
        if pdf_obj:
            pdf_obj.close()

    # Save highlighted document
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, filename)
    doc.save(out_path)
    
    return out_path, report_rows


def write_excel(rows, output_path):
    """Write bias report to Excel file"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Bias Report"
    headers = ["Filename", "Page Number", "Category", "Term", "Sentence"]
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(output_path)
    return output_path


def create_zip(word_dir, excel_path, zip_path):
    """Create ZIP file containing Word files and Excel report"""
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        # Add Word files
        if os.path.exists(word_dir):
            for file in os.listdir(word_dir):
                file_path = os.path.join(word_dir, file)
                if os.path.isfile(file_path) and not file.startswith("~$"):
                    zipf.write(file_path, arcname=f"word/{file}")
        
        # Add Excel report
        if os.path.exists(excel_path):
            zipf.write(excel_path, arcname="bias_report.xlsx")
    
    return zip_path


def cleanup_pdf_files():
    """
    Remove all PDF files from the outputs/pdf directory.
    Called after ZIP file creation to clean up temporary PDF files.
    """
    pdf_dir = os.path.join("outputs", "pdf")
    
    if not os.path.exists(pdf_dir):
        return
    
    try:
        for filename in os.listdir(pdf_dir):
            if filename.lower().endswith('.pdf'):
                file_path = os.path.join(pdf_dir, filename)
                try:
                    os.remove(file_path)
                    print(f"DEBUG: Removed PDF file: {file_path}")
                except Exception as e:
                    print(f"DEBUG: Failed to remove PDF {file_path}: {e}")
    except Exception as e:
        print(f"DEBUG: Error accessing PDF directory {pdf_dir}: {e}")
