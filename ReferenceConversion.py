
import os
import json
import logging
import re
from typing import Optional, List, Dict
from pathlib import Path
from gemini_ref_converter import convert_to_granular_style
from docx import Document

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def process_conversion_with_gemini(input_docx: Path, output_dir: Optional[Path] = None, target_style: str = "APA") -> Dict[str, Path]:
    """
    Process a document to convert references using Gemini AI.
    This is a standalone process, separate from the main structuring logic.
    """
    if output_dir is None:
        output_dir = input_docx.parent
    
    output_docx_path = output_dir / f"{input_docx.stem}_Converted.docx"
    log_file_path = output_dir / f"{input_docx.stem}_conversion_log.txt"
    
    doc = Document(input_docx)
    log_lines = []
    
    log_lines.append(f"Starting AI Reference Conversion for: {input_docx.name}")
    log_lines.append(f"Target Style: {target_style}")
    log_lines.append("-" * 50)
    
    count = 0
    converted_count = 0
    
    in_ref_section = False
    for para in doc.paragraphs:
        raw_text = para.text.strip()
        if not raw_text:
            continue
            
        # Tag Detection (Strict)
        raw_lower = raw_text.lower()
        if '<ref-open>' in raw_lower:
            in_ref_section = True
            log_lines.append("Reference section started (<ref-open>).")
            continue
        if '<ref-close>' in raw_lower:
            in_ref_section = False
            log_lines.append("Reference section ended (<ref-close>).")
            continue
            
        if not in_ref_section:
            continue
            
        # Basic heuristic to identify candidate references within the section
        if len(raw_text) > 10:
            log_lines.append(f"\nOriginal: {raw_text}")
            
            # Determine style for this specific reference if Auto is requested
            working_style = target_style
            if working_style == "Auto":
                if re.match(r'^\[?\d+\]?', raw_text):
                    working_style = "AMA"
                else:
                    working_style = "APA"
            
            gemini_result_str = convert_to_granular_style(raw_text, working_style)
            
            if gemini_result_str:
                try:
                     # Gemini returns a JSON string
                     data = json.loads(gemini_result_str)
                     metadata = data.get('metadata', {})
                     
                     from ReferencesStructing import generate_apa_citation, generate_ama_citation
                     
                     # Map the new bib_ fields back to the internal item format for generators
                     # This ensures we keep using the same styling logic.
                     item = {
                        'type': 'book' if metadata.get('bib_book') else 'journal-article',
                        'title': [metadata.get('bib_title', '')],
                        'container-title': [metadata.get('bib_journal', '') or metadata.get('bib_book', '')],
                        'DOI': metadata.get('bib_doi', ''),
                        'URL': metadata.get('bib_url', ''),
                        'volume': metadata.get('bib_volume', ''),
                        'issue': metadata.get('bib_issue', ''),
                        'page': f"{metadata.get('bib_fpage', '')}-{metadata.get('bib_lpage', '')}".strip('-'),
                        'year': metadata.get('bib_year', ''),
                        'publisher': metadata.get('bib_publisher', ''),
                        'author': []
                     }
                     
                     # Handle individual author fields (surname/fname)
                     # The schema seems to have single fields for bib_surname/bib_fname.
                     # If Gemini provides multiple, we might need to check if it's comma separated or what.
                     # Assuming for now single or comma-sep.
                     surnames = metadata.get('bib_surname')
                     fnames = metadata.get('bib_fname')
                     
                     if surnames:
                         s_list = surnames.split(',')
                         f_list = fnames.split(',') if fnames else []
                         for i, s in enumerate(s_list):
                             f = f_list[i] if i < len(f_list) else ""
                             item['author'].append({'family': s.strip(), 'given': f.strip()})
                     
                     new_text_segments = []
                     if working_style == "APA":
                         new_text_segments = generate_apa_citation(item)
                     else:
                         new_text_segments = generate_ama_citation(item)
                         
                     # Flatten segments to string
                     if isinstance(new_text_segments, list):
                         flat_text = "".join([t[0] for t in new_text_segments])
                     else:
                         flat_text = str(new_text_segments)

                     para.text = flat_text
                     log_lines.append(f"Converted ({working_style}): {flat_text}")
                     converted_count += 1
                     
                except Exception as e:
                    log_lines.append(f"Error parsing/formatting Gemini result: {e}")
            else:
                log_lines.append("Gemini returned no result.")
            
            count += 1
            
    doc.save(output_docx_path)
    
    with open(log_file_path, "w", encoding="utf-8") as f:
        f.writelines([l + "\n" for l in log_lines])
        
    return {'output_docx': output_docx_path, 'log_file': log_file_path}

if __name__ == "__main__":
    # Test
    pass
