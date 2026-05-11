"""Background job tasks using RQ."""
from pathlib import Path
import json
import zipfile
import tempfile
from manuscript_core.analyzer import analyze_manuscript
from manuscript_core.fixer import apply_fixes_to_docx
from manuscript_core.exporters import build_excel

APP_ROOT = Path(__file__).parent.parent
RESULTS_DIR = APP_ROOT / "results"
UPLOAD_DIR = APP_ROOT / "uploads"
OUTPUTS_DIR = APP_ROOT / "outputs"


def analyze_job(job_id: str, chapters: list[dict]) -> dict:
    """Run manuscript analysis in background."""
    try:
        # Analyze
        findings = analyze_manuscript(chapters)

        # Add job_id to findings for later retrieval
        findings["job_id"] = job_id

        # Save JSON results
        out_path = RESULTS_DIR / f"{job_id}.json"
        out_path.write_text(json.dumps(findings, ensure_ascii=False), encoding="utf-8")

        # Save Excel report
        job_out_dir = OUTPUTS_DIR / job_id
        job_out_dir.mkdir(parents=True, exist_ok=True)

        xlsx_bytes = build_excel(findings, job_id)
        (job_out_dir / f"manuscript_consistency_{job_id}.xlsx").write_bytes(xlsx_bytes)

        return {"status": "completed", "job_id": job_id}
    except Exception as e:
        return {"status": "failed", "job_id": job_id, "error": str(e)}


def fix_job(job_id: str, selected_patterns: list[dict], findings: list[dict], chapters: list[dict]) -> dict:
    """Apply fixes in background."""
    try:
        from manuscript_core.fixer import (
            build_fixes_from_selection,
            build_highlight_texts_from_selection,
            apply_te_highlights_to_docx,
        )
        from manuscript_core.figure_table_highlighter import FigureTableHighlighter
        import shutil

        job_dir = UPLOAD_DIR / job_id
        fixes = build_fixes_from_selection(selected_patterns, {"findings": findings})
        highlight_texts = build_highlight_texts_from_selection(selected_patterns, {"findings": findings})

        highlight_elements = set()
        for pat in selected_patterns:
            elem = pat.get("element")
            if elem in ("Figure", "Table", "Box", "Exhibit", "Appendix", "Case Study"):
                highlight_elements.add(elem)

        if not fixes and not highlight_elements and not highlight_texts:
            return {"status": "no_fixes", "job_id": job_id}

        # Apply fixes to each chapter
        temp_dir = Path(tempfile.mkdtemp())
        zip_path = temp_dir / f"Fixed_Manuscript_{job_id}.zip"

        highlighter = FigureTableHighlighter() if highlight_elements else None

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for chapter in chapters:
                orig_file = job_dir / chapter["filename"]
                if orig_file.exists():
                    fixed_name = f"FIXED_{chapter['filename']}"
                    fixed_file = temp_dir / fixed_name

                    if fixes:
                        apply_fixes_to_docx(orig_file, fixed_file, fixes)
                    else:
                        shutil.copy2(orig_file, fixed_file)

                    if highlight_texts:
                        apply_te_highlights_to_docx(str(fixed_file), str(fixed_file), highlight_texts)

                    if highlighter and highlight_elements:
                        highlighter.apply_highlighting_to_docx(
                            str(fixed_file),
                            str(fixed_file),
                            list(highlight_elements)
                        )

                    zip_ref.write(fixed_file, fixed_name)

        # Save zip to outputs
        output_dir = OUTPUTS_DIR / job_id
        output_dir.mkdir(parents=True, exist_ok=True)
        output_zip = output_dir / f"Fixed_Manuscript_{job_id}.zip"
        output_zip.write_bytes(zip_path.read_bytes())

        return {"status": "completed", "job_id": job_id}
    except Exception as e:
        return {"status": "failed", "job_id": job_id, "error": str(e)}

