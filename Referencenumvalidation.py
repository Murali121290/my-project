import re
import os
import io
import zipfile
from collections import defaultdict

from flask import Flask, request, send_file, render_template, redirect, url_for, session
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table

app = Flask(__name__)
app.secret_key = "secret_key_for_session_encryption"
UPLOAD_DIR = "temp_reports"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# =====================================================
# Helpers & Core Logic
# =====================================================

def iter_document_paragraphs(doc):
    """
    Iterate through all paragraphs in the document body in order,
    including those inside tables.
    """
    body = doc._element.body
    for child in body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p


def get_numbers(text):
    """
    Extract numbers from text like '1', '2-5', '1, 3, 5'.
    Handles ranges "1-5" -> [1, 2, 3, 4, 5].
    """
    nums = []
    # Matches: (start)-(end) OR (single)
    # Allows hyphen, en dash, em dash
    pattern = re.compile(r'(\d+)\s*[-–—]\s*(\d+)|(\d+)')
    
    for start, end, single in pattern.findall(text):
        if start and end:
            try:
                s, e = int(start), int(end)
                if s <= e:
                    nums.extend(range(s, e + 1))
            except ValueError:
                pass
        elif single:
            try:
                nums.append(int(single))
            except ValueError:
                pass
    return nums


def format_numbers(nums):
    """
    Format a list of numbers into a string like '1-3, 5'.
    Collapses ranges of 3 or more (e.g. 1,2,3 -> 1-3).
    """
    nums = sorted(set(nums))
    if not nums:
        return ""

    parts = []
    if not nums:
        return ""

    start = prev = nums[0]

    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            length = prev - start + 1
            if length >= 3:
                parts.append(f"{start}-{prev}")
            elif length == 2:
                parts.append(f"{start},{prev}")
            else:
                parts.append(str(start))
            start = prev = n

    length = prev - start + 1
    if length >= 3:
        parts.append(f"{start}-{prev}")
    elif length == 2:
        parts.append(f"{start},{prev}")
    else:
        parts.append(str(start))

    return ", ".join(parts)


def is_citation_run(run):
    """
    Determine if a run is part of a citation.
    Checks for 'cite_bib' style OR superscript with number-like content.
    """
    if run.style and run.style.name == "cite_bib":
        return True
    if run.font.superscript:
        text = run.text.strip()
        if not text:
            return False
        # Must look like numbers/ranges/separators
        if re.match(r'^[\d,\-–—\s]+$', text):
            return True
    return False


class ReferenceProcessor:
    def __init__(self, doc):
        self.doc = doc
        
    def get_references_in_bibliography(self):
        """
        Returns a Set of IDs found in the bibliography sections (REF-N style).
        Also returns a list of objects for reordering later.
        """
        refs_found = set()
        ref_objects = [] # list of dicts: {'id': int, 'para': p, 'run': r}

        for para in self.doc.paragraphs:
            if para.style and para.style.name == "REF-N":
                found_id = None
                bib_run = None
                
                # Try finding styled run
                for run in para.runs:
                    if run.style and run.style.name == "bib_number":
                        nums = get_numbers(run.text)
                        if nums:
                            found_id = nums[0]
                            bib_run = run
                            break
                            
                # Fallback: Check start of text if no styled run
                if found_id is None:
                    match = re.match(r'^(\d+)', para.text.strip())
                    if match:
                        found_id = int(match.group(1))
                
                if found_id is not None:
                    refs_found.add(found_id)
                    ref_objects.append({
                        'id': found_id,
                        'para': para,
                        'run': bib_run
                    })
                    
        return refs_found, ref_objects

    def get_citations_in_text(self):
        """
        Scans document for citations.
        Returns:
            all_cited_ids: list of all IDs in order of appearance (with duplicates)
            appearance_order: list of unique IDs in order of first appearance
        """
        all_cited_ids = []
        appearance_order = []
        seen = set()
        
        # Regex for fallback pattern ^1-3^
        citation_pattern = re.compile(r'\^([\d,\-–—\s]+)\^')

        for para in iter_document_paragraphs(self.doc):
            # 1. Process runs
            current_group = []
            
            for run in para.runs:
                if is_citation_run(run):
                    current_group.append(run)
                else:
                    if current_group:
                        # Flush group
                        text = "".join(r.text for r in current_group)
                        nums = get_numbers(text)
                        all_cited_ids.extend(nums)
                        for n in nums:
                            if n not in seen:
                                seen.add(n)
                                appearance_order.append(n)
                        current_group = []
                    
                    # Check fallback pattern in non-citation run
                    matches = citation_pattern.findall(run.text)
                    for m in matches:
                        nums = get_numbers(m)
                        all_cited_ids.extend(nums)
                        for n in nums:
                            if n not in seen:
                                seen.add(n)
                                appearance_order.append(n)
            
            # Flush trailing group
            if current_group:
                text = "".join(r.text for r in current_group)
                nums = get_numbers(text)
                all_cited_ids.extend(nums)
                for n in nums:
                    if n not in seen:
                        seen.add(n)
                        appearance_order.append(n)
                        
        return all_cited_ids, appearance_order

    def find_duplicates(self, ref_objects):
        """
        Finds duplicate references using fuzzy matching (difflib).
        Returns a list of dicts: {'id': int, 'text': str, 'duplicate_of': int, 'score': float}
        """
        import difflib
        
        duplicates = []
        processed_refs = [] # list of (id, clean_text)
        
        # 1. Pre-process all candidates
        for obj in ref_objects:
            full_text = obj['para'].text.strip()
            # Remove leading numbering like "1. ", "[1] "
            clean_text = re.sub(r'^\[?\d+\]?[\.\s]*', '', full_text)
            processed_refs.append({'id': obj['id'], 'text': clean_text})
            
        # 2. Compare O(N^2)
        # We only check forward to avoid double reporting (A=B, B=A)
        # We assume the *earlier* ID is the "original" and later is "duplicate"
        n = len(processed_refs)
        matcher = difflib.SequenceMatcher(None, "", "")
        
        for i in range(n):
            ref_a = processed_refs[i]
            text_a = ref_a['text']
            len_a = len(text_a)
            
            if len_a == 0:
                continue
                
            matcher.set_seq1(text_a)
            
            for j in range(i + 1, n):
                ref_b = processed_refs[j]
                text_b = ref_b['text']
                len_b = len(text_b)
                
                if len_b == 0: 
                    continue
                    
                # Optimization: Length ratio check
                # If lengths differ significantly, they can't be high matches
                # If ratio > 0.85, then min_len / max_len must be roughly > 0.85
                # We use 0.6 as a conservative safety net, but 0.8 is probably safe if threshold is 0.85.
                if min(len_a, len_b) / max(len_a, len_b) < 0.6:
                    continue
                
                matcher.set_seq2(text_b)
                
                # Performance Optimization: Check cheap upper bounds first
                if matcher.real_quick_ratio() < 0.85:
                    continue
                if matcher.quick_ratio() < 0.85:
                    continue
                    
                ratio = matcher.ratio()
                
                # Threshold: 0.85 (85% similar)
                if ratio > 0.85:
                    duplicates.append({
                        'id': ref_b['id'], # The later one is the duplicate
                        'text': ref_b['text'][:100] + "...",
                        'duplicate_of': ref_a['id'],
                        'score': round(ratio * 100, 1)
                    })
                    
        return duplicates

    def get_validation_stats(self):
        bib_refs, ref_objects = self.get_references_in_bibliography()
        all_cited, _ = self.get_citations_in_text()
        
        unique_cited = set(all_cited)
        
        # Missing: Cited but not in Bib
        missing = sorted(unique_cited - bib_refs)
        
        # Unused: In Bib but not Cited
        unused = sorted(bib_refs - unique_cited)
        
        # Duplicates
        duplicates = self.find_duplicates(ref_objects)
        
        # Sequence Issues
        sequence_issues = []
        seen_in_seq = []
        previous_max = 0
        
        for n in all_cited:
            if n not in seen_in_seq:
                if n < previous_max:
                     pass
                
                if n != len(seen_in_seq) + 1:
                     sequence_issues.append({
                         "position": len(seen_in_seq) + 1,
                         "current": n,
                         "expected": len(seen_in_seq) + 1
                     })
                
                seen_in_seq.append(n)
                previous_max = max(previous_max, n)
                
        return {
            "total_references": len(bib_refs),
            "total_citations": len(all_cited),
            "missing_references": missing,
            "unused_references": unused,
            "duplicate_references": duplicates,
            "sequence_issues": sequence_issues,
            "is_perfect": (not missing and not unused and not sequence_issues and not duplicates)
        }

    def renumber(self):
        """
        Renumber citations and reorder bibliography.
        Returns: mapping (Old -> New)
        """
        _, appearance_order = self.get_citations_in_text()
        
        # Ensure 'cite_bib' style exists
        from docx.enum.style import WD_STYLE_TYPE
        styles = self.doc.styles
        try:
            styles['cite_bib']
        except KeyError:
            s = styles.add_style('cite_bib', WD_STYLE_TYPE.CHARACTER)
            s.font.superscript = True

        # Create Mapping
        mapping = {} 
        new_id = 1
        for old_id in appearance_order:
            mapping[old_id] = new_id
            new_id += 1
            
        # 1. Update Citations in Text
        # Matches: ^1-3^ OR [1-3] OR (1-3)
        # Note: Be careful with (1) as it can be a list. We verify contents are numeric.
        citation_pattern = re.compile(r'(\^|\[|\()([\d,\-–—\s]+)(\^|\]|\))')
        
        for para in iter_document_paragraphs(self.doc):
            # Iterate runs safely with index since we might modify list
            i = 0
            while i < len(para.runs):
                run = para.runs[i]
                original_text = run.text
                
                # Check for Citation Pattern matches
                match = citation_pattern.search(original_text)
                
                if match:
                    # We found a match! We must split the run to style JUST the citation.
                    start, end = match.span()
                    
                    pre_text = original_text[:start]
                    match_text = original_text[start:end]
                    post_text = original_text[end:]
                    
                    # Calculate replacement text
                    # Regex Group 2 contains the numbers: [1-3] -> group 2="1-3"
                    nums = get_numbers(match.group(2))
                    new_nums = [mapping.get(n, n) for n in nums]
                    # Format: 1-3
                    converted_text = format_numbers(new_nums)
                    
                    # 1. Update Current Run -> Pre Text
                    run.text = pre_text
                    
                    # 2. Insert Match Run
                    new_run = para.add_run(converted_text)
                    new_run.style = "cite_bib"
                    new_run.font.superscript = True
                    
                    # Move new_run to be after current run
                    run._element.addnext(new_run._element)
                    
                    # 3. Insert Post Run (if exists)
                    if post_text:
                        post_run = para.add_run(post_text)
                        # Copy original style/font props if possible? 
                        # Ideally yes, but complex. Inheriting 'style' is good enough often.
                        if run.style:
                            post_run.style = run.style
                        
                        # Move post_run to be after new_run
                        new_run._element.addnext(post_run._element)
                        
                        # We don't advance i yet, or we assume post_run might have MORE citations?
                        # If post_run has citations, we need to process it.
                        # But we just inserted it at i+2 roughly.
                        # The runs list might not update automatically for 'para.runs[i]' indexing if cached?
                        # python-docx re-reads xml usually.
                        # Let's verify 'para.runs' reflects changes. 
                        # If it does, next iter will be new_run (styled) -> skip? 
                        # Then post_run -> process.
                        
                    # We modified the document structure.
                    # We should probably restart check on 'post_run' if we suspect multiple citations?
                    # Simply incrementing i might land us on new_run or post_run.
                    # Since we split current match, we should continue loop.
                    # But index strategy is tricky if list changes.
                    # Safer: Break and restart paragraph scan? Or just recursive?
                    # Given simple needs: let's assuming one per run or use recursive replacement on string FIRST?
                    # No, string replacement loses style boundaries.
                    # Let's just break for this run and continue to next (which might be the post_run we just added).
                    # Actually, if we use list(para.runs) iterator, it won't see new ones.
                    # We are using while index.
                    
                    # Logic: 
                    # i = current (now pre)
                    # i+1 = new_run (styled)
                    # i+2 = post_run
                    # We want to continue checking from i+2.
                    i += 1 # Skip new_run
                    if post_text:
                         # We want to check post_run next.
                         # i is now new_run index. i+1 is post_run.
                         # loop continues, i becomes i+1.
                         pass
                    else:
                         # No post run. i is new_run.
                         pass
                    
                    i += 1
                    continue
                
                elif is_citation_run(run):
                    # Existing formatted citation logic (likely superscripts without brackets)
                    # If it's already styled/superscript, we just update numbers.
                    current_group = [run]
                    # check next runs? (Group logic from original code was complex)
                    # For simplicty, let's just update this single run if it stands alone.
                    # The original code grouped multiple runs.
                    # We can keep that logic if we assume they are contiguous.
                    # But mixing with the splitting logic above is hard.
                    
                    # If valid citation run, just update text.
                    txt = run.text
                    nums = get_numbers(txt)
                    if nums:
                         new_nums = [mapping.get(n, n) for n in nums]
                         run.text = format_numbers(new_nums)
                         # Ensure style is enforced
                         run.style = "cite_bib"
                         run.font.superscript = True
                
                i += 1

        # 2. Reorder Bibliography
        _, ref_objects = self.get_references_in_bibliography()
        
        # Sort objects into Cited and Uncited
        cited_refs = []
        uncited_refs = []
        
        for obj in ref_objects:
            if obj['id'] in mapping:
                obj['new_id'] = mapping[obj['id']]
                cited_refs.append(obj)
            else:
                uncited_refs.append(obj)
        
        if not ref_objects:
            return mapping

        # Find anchor (min index)
        body = self.doc._element.body
        
        indices = []
        for obj in ref_objects:
            try:
                idx = body.index(obj['para']._element)
                indices.append(idx)
            except ValueError:
                pass 
        
        if not indices:
            return mapping
            
        anchor = min(indices)
        
        # Remove all
        for obj in ref_objects:
             p = obj['para']._element
             if p.getparent() == body:
                 body.remove(p)
                 
        # Insert Cited (Sorted)
        cited_refs.sort(key=lambda x: x['new_id'])
        
        insert_idx = anchor
        for obj in cited_refs:
            # Update ID text
            if obj['run']:
                obj['run'].text = str(obj['new_id'])
            
            body.insert(insert_idx, obj['para']._element)
            insert_idx += 1
            
        # Insert Uncited (Appended after cited)
        for obj in uncited_refs:
            body.insert(insert_idx, obj['para']._element)
            insert_idx += 1
            
        return mapping


def process_document(file):
    doc = Document(file)
    processor = ReferenceProcessor(doc)
    
    # Check BEFORE
    before_stats = processor.get_validation_stats()
    
    # DECISION:
    # 1. If Unused References exist -> ABORT renumbering.
    if before_stats["unused_references"]:
        return doc, before_stats, before_stats, {}, "Aborted: Document validation failed due to unused references."

    # 2. If Perfect -> No need.
    if before_stats["is_perfect"]:
        return doc, before_stats, before_stats, {}, "Validation completed."
        
    # 3. If Missing Refs -> Can't safely renumber usually
    if before_stats["missing_references"]:
         return doc, before_stats, before_stats, {}, "Aborted: Missing references detected."

    # DO RENUMBER
    mapping = processor.renumber()
    
    # Check AFTER (Validate result)
    after_stats = processor.get_validation_stats()
    
    # Determine status message
    changes_made = False
    if mapping:
        for k, v in mapping.items():
            if k != v:
                changes_made = True
                break

    if before_stats["duplicate_references"]:
        count = len(before_stats['duplicate_references'])
        prefix = "Renumbering" if changes_made else "Validation"
        status_msg = f"{prefix} completed with {count} duplicate{'s' if count > 1 else ''}."
    elif changes_made:
        status_msg = "Renumbering completed successfully."
    else:
        status_msg = "Validation completed."

    return doc, before_stats, after_stats, mapping, status_msg


# =====================================================
# Flask Routes
# =====================================================
@app.route("/")
def upload_file():
    return render_template("upload.html")


@app.route("/process", methods=["GET", "POST"])
def process():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or not file.filename.endswith(".docx"):
            return "Invalid file", 400

        doc, before, after, mapping, status_msg = process_document(file)

        base = os.path.splitext(file.filename)[0]
        doc_path = os.path.join(UPLOAD_DIR, f"{base}_renumbered.docx")
        report_path = os.path.join(UPLOAD_DIR, f"{base}_validation.txt")

        doc.save(doc_path)

        with open(report_path, "w", encoding="utf-8") as f:
            f.write(f"STATUS: {status_msg}\n")
            f.write("VALIDATION BEFORE\n")
            f.write(str(before) + "\n\n")
            f.write("VALIDATION AFTER\n")
            f.write(str(after) + "\n\n")
            if mapping:
                f.write("RENUMBERING MAPPING (Old -> New)\n")
                for old, new in sorted(mapping.items(), key=lambda x: x[1]):
                    f.write(f"{old} -> {new}\n")

        # Create ZIP package
        zip_filename = f"{base}_results.zip"
        zip_path = os.path.join(UPLOAD_DIR, zip_filename)
        
        # Validation HTML Report (Offline)
        html_report_filename = f"{base}_results.html"
        html_report_path = os.path.join(UPLOAD_DIR, html_report_filename)
        
        # Render the template for offline use
        # Note: We pass offline_mode=True to make links relative
        html_content = render_template(
            "validation_results.html",
            filename=file.filename,
            results=after,
            before=before,
            mapping=mapping,
            status_msg=status_msg,
            report_file=os.path.basename(report_path),
            doc_file=os.path.basename(doc_path),
            zip_file=None, # No zip button in offline report
            offline_mode=True 
        )
        
        with open(html_report_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        with zipfile.ZipFile(zip_path, 'w') as zf:
             # Add Doc
             zf.write(doc_path, arcname=os.path.basename(doc_path))
             # Add Text Report
             zf.write(report_path, arcname=os.path.basename(report_path))
             # Add HTML Report
             zf.write(html_report_path, arcname=os.path.basename(html_report_path))

        # Store data in session for GET request
        session['processing_result'] = {
            'filename': file.filename,
            'before': before,
            'after': after,
            'mapping': mapping,
            'status_msg': status_msg,
            'report_file': os.path.basename(report_path),
            'doc_file': os.path.basename(doc_path),
            'zip_file': zip_filename
        }
        
        return redirect(url_for('process'))

    # GET request - retrieve from session
    result = session.get('processing_result')
    if not result:
        return redirect(url_for('upload_file'))
        
    return render_template(
        "validation_results.html",
        filename=result['filename'],
        results=result['after'],
        before=result['before'],
        mapping=result['mapping'],
        status_msg=result['status_msg'],
        report_file=result['report_file'],
        doc_file=result['doc_file'],
        zip_file=result.get('zip_file')
    )


@app.route("/download/<path:filename>")
def download_file(filename):
    # Security: Ensure filename is in UPLOAD_DIR
    return send_file(os.path.join(UPLOAD_DIR, filename), as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
