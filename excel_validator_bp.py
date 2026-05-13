from __future__ import annotations
import os
import shutil
import uuid
from functools import wraps
from pathlib import Path

from flask import (Blueprint, jsonify, redirect, render_template,
                   request, send_file, session, url_for, flash)
from werkzeug.utils import secure_filename

from excel_validator_core import build_merged_comparison_workbook

excel_validator_bp = Blueprint(
    "excel_validator",
    __name__,
    template_folder=str(Path(__file__).parent / "templates" / "excel_validator"),
)

ALLOWED_EXT = {".xlsx", ".csv"}
ALLOWED_ROLES = {"COPYEDIT", "COPYEDITPM", "PM", "PERMISSIONS", "PPD", "POST_PROD", "ADMIN"}
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "S4C-Processed-Documents"


def _auth(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in to continue.")
            return redirect(url_for("login"))
        role = (session.get("role") or "").upper()
        if not session.get("is_admin") and role not in ALLOWED_ROLES:
            flash("You do not have permission to access this page.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return wrapped


@excel_validator_bp.route("/", methods=["GET"])
@_auth
def upload_page():
    return render_template("merge_compare.html", current_role=session.get("role"))


@excel_validator_bp.route("/process", methods=["POST"])
@_auth
def process():
    uploaded = request.files.getlist("files")
    valid = [f for f in uploaded
             if f.filename and Path(f.filename).suffix.lower() in ALLOWED_EXT]

    if len(valid) < 2:
        return jsonify({"error": "Please upload at least 2 .xlsx or .csv files."}), 400

    token = uuid.uuid4().hex
    temp_dir = UPLOAD_DIR / token
    temp_dir.mkdir(parents=True, exist_ok=True)

    file_paths: list[str] = []
    filenames: list[str] = []
    try:
        for f in valid:
            safe = secure_filename(f.filename)
            dest = temp_dir / safe
            f.save(str(dest))
            file_paths.append(str(dest))
            filenames.append(f.filename)

        buf = build_merged_comparison_workbook(file_paths, filenames)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="merged_comparison.xlsx",
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Processing failed: {exc}"}), 500
    finally:
        shutil.rmtree(str(temp_dir), ignore_errors=True)
