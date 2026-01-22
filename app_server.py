from flask import (
    Flask, render_template, render_template_string, request, redirect, flash,
    send_file, url_for, jsonify, session, send_from_directory, g
)
import os
import io
import shutil
import zipfile
import subprocess
import traceback
import uuid
import json
import re
import sqlite3
import socket
import time
import threading
import concurrent.futures

# Windows/Linux Compatibility
try:
    import pythoncom
    import win32com.client as win32
    import pywintypes
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    class pywintypes:
        class com_error(Exception):
            pass
    pythoncom = None
    win32 = None

from pathlib import Path
from threading import Lock
from datetime import datetime, timedelta, timezone
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix
from flask_wtf.csrf import CSRFProtect
from functools import wraps
from waitress import serve
from contextlib import contextmanager
from queue import Queue, Empty
import logging
from logging.handlers import RotatingFileHandler
from collections import defaultdict
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
import difflib
from highlighter.core_highlighter_docx import process_docx
from ReferencesStructing import process_docx_file
from ReferenceAPAValidation import validate_document_multi_style, insert_comments_in_document, generate_report as generate_apa_report, apply_citation_formatting
import tempfile
from io import BytesIO
from extractor import extract_from_file, write_permission_log
# -----------------------
# Configuration
# -----------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "S4C-Processed-Documents")
COMMON_MACRO_FOLDER = os.path.join(BASE_DIR, "S4c-Macros")
DEFAULT_MACRO_NAME = 'CE_Tool.dotm'
REPORT_FOLDER = "reports"
DATABASE = os.path.join(BASE_DIR, "reference_validator.db")
LOG_FILE = os.path.join(BASE_DIR, 'user_activity.log')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(COMMON_MACRO_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)
os.makedirs("logs", exist_ok=True)

ALLOWED_EXTENSIONS = {'.doc', '.docx'}
WORD_START_RETRIES = 3
WORD_LOCK = Lock()
TOKEN_TTL = timedelta(hours=1)

# -----------------------
# Route-Specific Macro Configuration
# -----------------------
ROUTE_MACROS = {
    'language': {
        'name': 'Language Editing',
        'description': 'Language editing and grammar correction tools',
        'icon': 'edit',
        'macros': [
            "LanguageEdit.GrammarCheck_WithErrorHandling",
            "LanguageEdit.SpellCheck_Advanced",
            "LanguageEdit.StyleConsistency_Check",
            "LanguageEdit.ReadabilityAnalysis",
            "LanguageEdit.TerminologyValidation"
        ]
    },
    'technical': {
        'name': 'Technical Editing',
        'description': 'Technical document formatting and validation tools',
        'icon': 'cog',
        'macros': [
            "Referencevalidation.ValidateBWNumCite_WithErrorHandling",
            "ReferenceRenumber.Reorderbasedonseq",
            "Copyduplicate.duplicate4",
            "citationupdateonly.citationupdate",
            "techinal.technicalhighlight"
        ]
    },
    'macro_processing': {
        'name': 'Reference Processing',
        'description': 'Reference validation and citation tools',
        'icon': 'bookmark',
        'macros': [
            "Referencevalidation.ValidateBWNumCite_WithErrorHandling",
            "ReferenceRenumber.Reorderbasedonseq",
            "Copyduplicate.duplicate4",
            "citationupdateonly.citationupdate",
            "Prediting.Preditinghighlight",
            "msrpre.GenerateDashboardReport",
        ]
    },
    'ppd': {
        'name': 'PPD Processing',
        'description': 'PPD final processing tools (from PPD_Final.py)',
        'icon': 'magic',
        'macros': [
            "PPD_HTML.GenerateDocument",
            "PPD_HTML.Generate_HTML_WORDReport",
        ]
    }
}

ROUTE_MACROS['credit_extractor'] = {
    'name': 'Credit / Permission Log',
    'description': 'Caption & credit line extraction for permissions',
    'icon': 'file-text',
    'macros': []
}

# Flask app
app = Flask(__name__)

# Apply ProxyFix for Nginx (handles HTTPS, X-Forwarded-Proto, etc.)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

# Fix for CSRF token missing in multi-worker environment
# Ensure secret key is consistent across workers by storing it in a file if not in env
secret_key_path = os.path.join(BASE_DIR, '.flask_secret_key')
if os.environ.get('SECRET_KEY'):
    app.secret_key = os.environ.get('SECRET_KEY')
elif os.path.exists(secret_key_path):
    with open(secret_key_path, 'rb') as f:
        app.secret_key = f.read()
else:
    # Generate and save a new key so it persists across restarts and workers
    generated_key = os.urandom(24)
    try:
        with open(secret_key_path, 'wb') as f:
            f.write(generated_key)
        app.secret_key = generated_key
    except IOError:
        # Fallback if cannot write to file
        app.secret_key = 'fallback-secret-key-change-this-in-prod'

csrf = CSRFProtect(app)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['COMMON_MACRO_FOLDER'] = COMMON_MACRO_FOLDER
app.config['REPORT_FOLDER'] = REPORT_FOLDER
app.config['DATABASE'] = DATABASE
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

# Token-based download map
download_tokens = {}

ROUTE_PERMISSIONS = {
    'language': ['COPYEDIT', 'ADMIN'],
    'technical': ['COPYEDIT', 'ADMIN'],
    'macro_processing': ['COPYEDIT', 'ADMIN'],
    'ppd': ['COPYEDIT', 'PPD', 'PM', 'ADMIN'],
    'credit_extractor': ['PERMISSIONS', 'PM', 'ADMIN']
}

def get_user_role():
    return session.get('role') or (g.user.get('role') if g.user else None)

def has_role(*roles):
    role = get_user_role()
    return role is not None and role.upper() in [r.upper() for r in roles]

def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if 'user_id' not in session:
                flash("Please log in to continue.")
                return redirect(url_for('login'))

            if not has_role(*allowed_roles) and not session.get('is_admin'):
                flash("You don't have permission to access this page.", "error")
                return redirect(url_for('dashboard'))

            return f(*args, **kwargs)
        return wrapped
    return decorator

def process_credit_extractor_job(job_id, temp_dir, file_paths, original_filenames, user_id, username):
    with app.app_context():
        # temp_dir passed from route
        all_results = []
        
        try:
            for idx, path in enumerate(file_paths, start=1):
                filename = original_filenames[idx-1]
                app.config["PROGRESS_DATA"][job_id].update({
                    "current": idx,
                    "status": f"Processing {filename}"
                })

                all_results.extend(extract_from_file(path))

            if not all_results:
                app.config["PROGRESS_DATA"][job_id]["status"] = "No captions found"
                return

            output_xlsx = os.path.join(temp_dir, "permission_log.xlsx")
            write_permission_log(all_results, output_xlsx)

            # ZIP if multiple files
            if len(original_filenames) > 1:
                zip_path = os.path.join(temp_dir, "permission_logs.zip")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                    z.write(output_xlsx, arcname="permission_log.xlsx")
                final_path = zip_path
                processed_files = ["permission_logs.zip"]
            else:
                final_path = output_xlsx
                processed_files = ["permission_log.xlsx"]

            # Register download token
            token = uuid.uuid4().hex
            download_tokens[token] = {
                "path": temp_dir,
                "expires": _now_utc() + TOKEN_TTL,
                "user": username,
                "route_type": "credit_extractor"
            }

            # DB logging
            with db_pool.get_connection() as db:
                db.execute(
                    '''INSERT INTO macro_processing
                       (user_id, token, original_filenames, processed_filenames, selected_tasks, route_type)
                       VALUES (?, ?, ?, ?, ?, ?)''',
                    (
                        user_id,
                        token,
                        json.dumps(original_filenames),
                        json.dumps(processed_files),
                        json.dumps({"route_type": "credit_extractor"}),
                        "credit_extractor"
                    )
                )
                db.commit()

            app.config["PROGRESS_DATA"][job_id].update({
                "status": "Completed",
                "download_token": token
            })

        except Exception as e:
            app.config["PROGRESS_DATA"][job_id]["status"] = f"Failed: {e}"

@app.route("/credit-extractor", methods=["GET", "POST"])
@csrf.exempt
@role_required(ROUTE_PERMISSIONS.get('credit_extractor', ['ADMIN']))
def credit_extractor():
    if request.method == "POST":
        files = request.files.getlist("files")

        if not files or all(f.filename == "" for f in files):
            return jsonify({"error": "No files selected"}), 400

        job_id = str(int(time.time() * 1000))
        
        # Save files synchronously before threading
        temp_dir = tempfile.mkdtemp()
        saved_paths = []
        original_filenames = []
        
        try:
            for f in files:
                if f.filename:
                    safe_name = secure_filename(f.filename) or f"document_{len(saved_paths)}"
                    path = os.path.join(temp_dir, safe_name)
                    f.save(path)
                    saved_paths.append(path)
                    original_filenames.append(f.filename)
        except Exception as e:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
            return jsonify({"error": f"File save failed: {e}"}), 500

        app.config.setdefault("PROGRESS_DATA", {})
        app.config["PROGRESS_DATA"][job_id] = {
            "total": len(saved_paths),
            "current": 0,
            "status": "Starting"
        }

        threading.Thread(
            target=process_credit_extractor_job,
            args=(job_id, temp_dir, saved_paths, original_filenames, session['user_id'], session['username']),
            daemon=True
        ).start()

        return jsonify({"job_id": job_id})

    return render_template("upload_credit.html")

# -----------------------
# Database Connection Pool
# -----------------------
class DatabasePool:
    def __init__(self, database_path, pool_size=5):
        self.database_path = database_path
        self.pool = Queue(maxsize=pool_size)
        self.lock = threading.Lock()

        for _ in range(pool_size):
            conn = sqlite3.connect(database_path, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            conn.execute("PRAGMA cache_size=10000")
            self.pool.put(conn)

    @contextmanager
    def get_connection(self):
        try:
            conn = self.pool.get(timeout=5)
            yield conn
        except Empty:
            conn = sqlite3.connect(self.database_path, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            yield conn
        finally:
            try:
                self.pool.put(conn, block=False)
            except:
                conn.close()


db_pool = DatabasePool(DATABASE)


# -----------------------
# Optimized Word Processor
# -----------------------
class OptimizedDocumentProcessor:
    def __init__(self):
        self.word = None
        self.docs = []
        self.macro_template_loaded = False

    def __enter__(self):
        if HAS_WIN32COM:
            pythoncom.CoInitialize()
            self.word = self._start_word_optimized()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cleanup()

    def _start_word_optimized(self):
        if not HAS_WIN32COM:
            return None
            
        for attempt in range(WORD_START_RETRIES):
            try:
                subprocess.run(["taskkill", "/f", "/im", "winword.exe"],
                               capture_output=True, check=False)

                word = win32.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False
                word.AutomationSecurity = 1
                word.ScreenUpdating = False
                word.Options.DoNotPromptForConvert = True
                word.Options.ConfirmConversions = False
                return word
            except Exception as e:
                if attempt == WORD_START_RETRIES - 1:
                    raise RuntimeError(f"Failed to start Word: {e}")
                time.sleep(1)

    def _load_macro_template(self):
        if not self.word:
            return False

        if self.macro_template_loaded:
            return True

        try:
            macro_path = os.path.join(COMMON_MACRO_FOLDER, DEFAULT_MACRO_NAME)
            if not os.path.exists(macro_path):
                return False

            for addin in self.word.AddIns:
                try:
                    if hasattr(addin, 'FullName') and addin.FullName.lower().endswith(DEFAULT_MACRO_NAME.lower()):
                        self.macro_template_loaded = True
                        return True
                except:
                    continue

            self.word.AddIns.Add(macro_path, True)
            self.macro_template_loaded = True
            return True

        except Exception as e:
            log_errors([f"Failed to load macro template: {str(e)}"])
            return False

    def process_documents_batch(self, file_paths, selected_tasks, route_type):
        errors = []

        if not self.word:
            return ["Word automation (macros) not supported on Linux."]

        if not self._load_macro_template():
            errors.append("Failed to load macro template")
            return errors

        route_macros = ROUTE_MACROS.get(route_type, {}).get('macros', [])

        for doc_path in file_paths:
            try:
                abs_path = os.path.abspath(doc_path)
                if not os.path.exists(abs_path):
                    errors.append(f"File not found: {abs_path}")
                    continue

                doc = self.word.Documents.Open(abs_path, ReadOnly=False, AddToRecentFiles=False)
                self.docs.append(doc)

                for task_index in selected_tasks:
                    try:
                        idx = int(task_index)
                        if 0 <= idx < len(route_macros):
                            macro_name = route_macros[idx]
                            try:
                                self.word.Run(macro_name)
                            except pywintypes.com_error as ce:
                                errors.append(f"COM error running '{macro_name}': {ce}")
                            except Exception as me:
                                errors.append(f"Macro '{macro_name}' failed: {me}")
                        else:
                            errors.append(f"Invalid task index {idx} for route {route_type}")
                    except ValueError:
                        errors.append(f"Invalid task index: {task_index}")

                try:
                    doc.Save()
                    doc.Close(SaveChanges=False)
                    self.docs.remove(doc)
                except Exception as se:
                    errors.append(f"Failed to save document: {se}")

            except Exception as doc_err:
                errors.append(f"Document processing failed: {doc_err}")

        return errors

    def _cleanup(self):
        for doc in self.docs:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass

        if self.word:
            try:
                self.word.Quit()
            except:
                pass

        try:
            if HAS_WIN32COM:
                pythoncom.CoUninitialize()
        except:
            pass


# -----------------------
# Utility Functions
# -----------------------
def get_ip_address():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip_address = s.getsockname()[0]
        s.close()
        return ip_address
    except Exception:
        return "127.0.0.1"


def log_activity(username, action, details=""):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"{timestamp} - {username} - {action} - {details}\n")


def log_errors(error_list):
    with open(LOG_FILE, "a", encoding="utf-8") as log_file:
        for err in error_list:
            log_file.write(f"{datetime.now().isoformat()} - ERROR - {err}\n")


def allowed_file(filename):
    return any(filename.lower().endswith(ext) for ext in ALLOWED_EXTENSIONS)


def setup_logging():
    if not app.debug:
        file_handler = RotatingFileHandler('logs/s4c.log', maxBytes=10240000, backupCount=10)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
        ))
        file_handler.setLevel(logging.INFO)
        app.logger.addHandler(file_handler)
        app.logger.setLevel(logging.INFO)


def cleanup_expired_tokens():
    current_time = _now_utc()
    expired_tokens = []

    for token, data in list(download_tokens.items()):
        expires = data.get("expires")
        if not expires:
            expired_tokens.append(token)
            continue

        expires = _ensure_utc(expires)

        if current_time > expires:
            expired_tokens.append(token)

    for token in expired_tokens:
        try:
            info = download_tokens.get(token)
            if not info:
                continue

            path = info.get("path")
            if path and os.path.exists(path):
                shutil.rmtree(path, ignore_errors=True)

            log_activity(
                info.get("user", "system"),
                f"TOKEN_EXPIRED_{info.get('route_type', 'UNKNOWN').upper()}",
                token[:8]
            )

            download_tokens.pop(token, None)

        except Exception as e:
            log_errors([f"Token cleanup failed ({token}): {e}"])



def kill_word_processes():
    if os.name != 'nt':
        return
    try:
        subprocess.run(["taskkill", "/f", "/im", "winword.exe"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def save_uploaded_file(file, folder):
    try:
        filename = secure_filename(file.filename)
        file_path = os.path.join(folder, filename)

        with open(file_path, 'wb') as f:
            file.save(f)

        return file_path, None
    except Exception as e:
        return None, str(e)


# -----------------------
# Template Filters
# -----------------------
@app.template_filter('from_json')
def from_json_filter(value):
    try:
        return json.loads(value)
    except (ValueError, TypeError):
        return value


@app.template_filter('format_date')
def format_date_filter(value):
    try:
        if isinstance(value, str):
            dt = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
        else:
            dt = value
        return dt.strftime('%b %d, %Y %I:%M %p')
    except (ValueError, AttributeError):
        return value


# -----------------------
# Database Functions
# -----------------------
def get_db():
    return db_pool.get_connection()


def init_db():
    with app.app_context():
        with db_pool.get_connection() as db:
            # Create tables
            db.execute('''CREATE TABLE IF NOT EXISTS users (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            username TEXT UNIQUE NOT NULL,
                            password TEXT NOT NULL,
                            email TEXT,
                            is_admin BOOLEAN DEFAULT FALSE,
                            role TEXT DEFAULT 'USER',
                            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

            db.execute('''CREATE TABLE IF NOT EXISTS files (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            user_id INTEGER NOT NULL,
                            original_filename TEXT NOT NULL,
                            stored_filename TEXT NOT NULL,
                            report_filename TEXT,
                            upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                            FOREIGN KEY (user_id) REFERENCES users(id))''')

            db.execute('''CREATE TABLE IF NOT EXISTS validation_results (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            file_id INTEGER NOT NULL,
                            total_references INTEGER,
                            total_citations INTEGER,
                            missing_references TEXT,
                            unused_references TEXT,
                            sequence_issues TEXT,
                            FOREIGN KEY (file_id) REFERENCES files(id))''')

            db.execute('''CREATE TABLE IF NOT EXISTS macro_processing (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            user_id INTEGER NOT NULL,
                            token TEXT UNIQUE NOT NULL,
                            original_filenames TEXT NOT NULL,
                            processed_filenames TEXT NOT NULL,
                            selected_tasks TEXT NOT NULL,
                            processing_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                            errors TEXT,
                            route_type TEXT DEFAULT 'general',
                            FOREIGN KEY (user_id) REFERENCES users(id))''')

            # Create indexes for performance
            try:
                db.execute("CREATE INDEX IF NOT EXISTS idx_files_user_id ON files(user_id)")
                db.execute("CREATE INDEX IF NOT EXISTS idx_files_upload_date ON files(upload_date)")
                db.execute("CREATE INDEX IF NOT EXISTS idx_macro_user_id ON macro_processing(user_id)")
                db.execute("CREATE INDEX IF NOT EXISTS idx_macro_route_type ON macro_processing(route_type)")
            except sqlite3.OperationalError as e:
                if "no such column" in str(e):
                    print("Warning: Column doesn't exist yet, skipping index creation")
                else:
                    raise

            # Create default admin
            admin_user = db.execute("SELECT * FROM users WHERE username='admin'").fetchone()
            if not admin_user:
                hashed_password = generate_password_hash("admin123", method='pbkdf2:sha256')
                db.execute("INSERT INTO users (username,password,email,is_admin) VALUES (?,?,?,?)",
                           ('admin', hashed_password, 'admin@example.com', True))
                db.commit()

def migrate_add_role_column():
    """Ensure the 'role' column exists for legacy DBs."""
    try:
        with db_pool.get_connection() as db:
            cur = db.execute("PRAGMA table_info(users)")
            cols = [r["name"] for r in cur.fetchall()]
            if "role" not in cols:
                db.execute("ALTER TABLE users ADD COLUMN role TEXT DEFAULT 'USER'")
                db.commit()
                app.logger.info("Added 'role' column to users table")
    except Exception as e:
        log_errors([f"Migration error adding role column: {e}"])

@app.context_processor
def inject_current_role():
    return {'current_role': get_user_role()}
# -----------------------
# Enhanced Reference Validator
# -----------------------
# -----------------------
# Enhanced Reference Validator (Logic from Referencenumvalidation.py)
# -----------------------

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
        for i in range(n):
            ref_a = processed_refs[i]
            
            # Skip if strict duplicate logic already caught it? 
            # No, let's just do fuzzy for all.
            
            for j in range(i + 1, n):
                ref_b = processed_refs[j]
                
                # Metric: similarity ratio
                # Quick check: length difference shouldn't be too huge
                len_a = len(ref_a['text'])
                len_b = len(ref_b['text'])
                if len_a == 0 or len_b == 0: 
                    continue
                    
                # Optimization: Length ratio check
                if min(len_a, len_b) / max(len_a, len_b) < 0.6:
                    continue
                    
                ratio = difflib.SequenceMatcher(None, ref_a['text'], ref_b['text']).ratio()
                
                # Threshold: 0.85 (85% similar)
                # The user's example is extremely similar, probably > 90%
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
        
        # Create Mapping
        mapping = {} 
        new_id = 1
        for old_id in appearance_order:
            mapping[old_id] = new_id
            new_id += 1
            
        # 1. Update Citations in Text
        citation_pattern = re.compile(r'\^([\d,\-–—\s]+)\^')
        
        for para in iter_document_paragraphs(self.doc):
            current_group = []
            
            for run in para.runs:
                if is_citation_run(run):
                    current_group.append(run)
                else:
                    if current_group:
                        # Replace
                        text = "".join(r.text for r in current_group)
                        nums = get_numbers(text)
                        if nums:
                            new_nums = [mapping.get(n, n) for n in nums]
                            new_text = format_numbers(new_nums)
                            current_group[0].text = new_text
                            for r in current_group[1:]:
                                r.text = ""
                        current_group = []
                    
                    # Pattern replacement
                    def replace_func(m):
                         nums = get_numbers(m.group(1))
                         new_nums = [mapping.get(n, n) for n in nums]
                         return "^" + format_numbers(new_nums) + "^"
                    
                    new_run_text = citation_pattern.sub(replace_func, run.text)
                    if new_run_text != run.text:
                        run.text = new_run_text

            if current_group:
                text = "".join(r.text for r in current_group)
                nums = get_numbers(text)
                if nums:
                    new_nums = [mapping.get(n, n) for n in nums]
                    new_text = format_numbers(new_nums)
                    current_group[0].text = new_text
                    for r in current_group[1:]:
                        r.text = ""

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


def process_document(file_path):
    doc = Document(file_path)
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
        status_msg = "Validation completed." # Fallback if no changes and no duplicates but not 'perfect' initially (e.g. sequence issues resolved to identity?)

    return doc, before_stats, after_stats, mapping, status_msg



# -----------------------
# Authentication (update load_logged_in_user)
# -----------------------
@app.before_request
def load_logged_in_user():
    user_id = session.get('user_id')
    if user_id is None:
        g.user = None
    else:
        with db_pool.get_connection() as db:
            user = db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
            g.user = dict(user) if user else None
            if g.user:
                session['role'] = g.user.get('role', 'USER')



@app.before_request
def require_login():
    if request.endpoint in (
        'login', 'logout', 'static',
        'download_report',  # ✅ add this
        'register', 'reset_database',  # we'll secure this below
        'macro_download'
    ):
        return None
    if not session.get('user_id'):
        flash("Please log in to continue.")
        return redirect(url_for('login'))


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('is_admin'):
            flash("Admin privileges required", "error")
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)

    return decorated_function

# -----------------------
# HTML to Excel (remove images)
# -----------------------
import pandas as pd
from bs4 import BeautifulSoup
import os
from pathlib import Path
from datetime import datetime
import chardet

import chardet  # at top of file with other imports

# -----------------------
# HTML to Excel (remove images)
# -----------------------
def html_to_excel_no_images(html_path, output_dir):
    """
    Converts an HTML file to an .xls file by removing <img> tags and writing
    the resulting HTML to a .xls file so Excel can open it.
    Returns the output file path or None on failure.
    """
    try:
        # read raw bytes and detect encoding
        with open(html_path, "rb") as f:
            raw_data = f.read()

        encoding = None
        try:
            detected = chardet.detect(raw_data)
            encoding = detected.get("encoding") or "utf-8"
        except Exception:
            encoding = "utf-8"

        try:
            html_content = raw_data.decode(encoding, errors="ignore")
        except Exception:
            html_content = raw_data.decode("utf-8", errors="ignore")

        # Remove <img> tags (handles attributes and self-closing)
        html_no_images = re.sub(r"<img\b[^>]*>", "", html_content, flags=re.IGNORECASE)

        # Also remove inline base64 images in style attributes (background-image:url(data:...))
        html_no_images = re.sub(r'url\(\s*data:[^)]+\)', 'url()', html_no_images, flags=re.IGNORECASE)

        # Build a safe output filename
        base = Path(html_path).stem
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = os.path.join(output_dir, f"{base}_{timestamp}.xls")

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_no_images)

        return output_file
    except Exception as e:
        log_errors([f"HTML to Excel conversion failed for {html_path}: {e}"])
        return None


# -----------------------
# Generic Route Handler
# -----------------------
def _process_macro_request(route_type):
    """
    Generic handler for macro routes. Accepts files from form field 'word_files[]'
    and task indices from 'tasks[]'. Processes documents using OptimizedDocumentProcessor
    (protected by WORD_LOCK), then if route_type == 'ppd' converts any produced HTML
    files in the output folder to .xls (images removed).
    Thread-safe access to download_tokens is used via download_tokens_lock.
    """
    word_files = request.files.getlist('word_files[]')
    selected_tasks = request.form.getlist('tasks[]')
    user_id = session.get('user_id')
    username = session.get('username', 'unknown')

    if not word_files or not selected_tasks:
        flash("Please upload files and select at least one task.")
        return redirect(url_for(route_type))

    token = uuid.uuid4().hex
    unique_folder = os.path.join(app.config['UPLOAD_FOLDER'], token)
    os.makedirs(unique_folder, exist_ok=True)

    # Register download token (thread-safe)
    try:
        with download_tokens_lock:
            download_tokens[token] = {
                'path': unique_folder,
                'expires': _now_utc() + TOKEN_TTL,
                'user': username,
                'route_type': route_type
            }
    except NameError:
        # If the lock isn't present for some reason, fall back (but warn)
        download_tokens[token] = {
            'path': unique_folder,
            'expires': _now_utc() + TOKEN_TTL,
            'user': username,
            'route_type': route_type
        }

    word_paths = []
    original_filenames = []

    for f in word_files:
        if f and allowed_file(f.filename):
            filename = secure_filename(f.filename)
            save_path = os.path.join(unique_folder, filename)
            try:
                f.save(save_path)
                word_paths.append(save_path)
                original_filenames.append(filename)
            except Exception as e:
                log_errors([f"Error saving uploaded file {filename}: {str(e)}"])

    if not word_paths:
        flash("No valid Word files uploaded.")
        return redirect(url_for(route_type))

    all_errors = []

    try:
        with WORD_LOCK:
            with OptimizedDocumentProcessor() as processor:
                # reuse processor.process_documents_batch to run macros and collect errors
                try:
                    batch_errors = processor.process_documents_batch(word_paths, selected_tasks, route_type)
                    if batch_errors:
                        all_errors.extend(batch_errors)
                except Exception as e:
                    all_errors.append(f"Batch processing failed: {str(e)}")
                    log_errors([traceback.format_exc()])

                # log processed docs
                for doc_path in word_paths:
                    log_activity(username, f"MACRO_PROCESS_{route_type.upper()}",
                                 details=os.path.basename(doc_path))

    except Exception as e:
        all_errors.append(f"Processing failed: {str(e)}")
        log_errors([traceback.format_exc()])

    # -----------------------
    # PPD-specific processing: Convert HTML outputs to Excel without images
    # This must happen AFTER document processing completes
    # -----------------------
    if route_type.lower() == 'ppd':
        try:
            if os.path.exists(unique_folder):
                html_files = [f for f in os.listdir(unique_folder) if f.lower().endswith(".html")]
            else:
                html_files = []

            # debug prints can be kept or removed
            app.logger.debug(f"PPD: found HTML files -> {html_files}")

            converted_files = []
            for file in html_files:
                html_path = os.path.join(unique_folder, file)
                app.logger.debug(f"PPD: converting {html_path} to Excel (no images)")
                out_xls = html_to_excel_no_images(html_path, unique_folder)
                if out_xls:
                    converted_files.append(os.path.basename(out_xls))
                else:
                    all_errors.append(f"Failed converting {file} to Excel")

            # Optionally: add converted files to processed_filenames list in DB later
        except Exception as e:
            error_msg = f"HTML to Excel conversion failed: {str(e)}"
            all_errors.append(error_msg)
            log_errors([error_msg])

    # -----------------------
    # Store in database
    # -----------------------
    try:
        selected_macro_names = []
        route_macros = ROUTE_MACROS.get(route_type, {}).get('macros', [])
        for task_idx in selected_tasks:
            try:
                idx = int(task_idx)
                if 0 <= idx < len(route_macros):
                    selected_macro_names.append(route_macros[idx])
            except Exception:
                pass

        # Build processed_filenames: include original filenames and any generated files in the folder
        processed_filenames = list(original_filenames)
        try:
            if os.path.exists(unique_folder):
                for root, _, files in os.walk(unique_folder):
                    for fn in files:
                        if fn not in processed_filenames:
                            processed_filenames.append(fn)
        except Exception:
            # If walking the folder fails, we'll still save original filenames
            pass

        with db_pool.get_connection() as db:
            db.execute('''INSERT INTO macro_processing 
                          (user_id, token, original_filenames, processed_filenames, selected_tasks, errors, route_type)
                          VALUES (?, ?, ?, ?, ?, ?, ?)''',
                       (user_id, token,
                        json.dumps(original_filenames),
                        json.dumps(processed_filenames),
                        json.dumps({
                            'route_type': route_type,
                            'task_indices': selected_tasks,
                            'macro_names': selected_macro_names
                        }),
                        json.dumps(all_errors) if all_errors else None,
                        route_type))
            db.commit()
    except Exception as e:
        log_errors([f"Error saving macro processing: {str(e)}"])

    route_name = ROUTE_MACROS.get(route_type, {}).get('name', 'Processing')
    if all_errors:
        flash(f"{route_name} completed with some errors. Check log for details.")
        log_errors(all_errors)
    else:
        flash(f"{route_name} completed successfully!")

    return redirect(url_for(route_type, download_token=token))

# -----------------------
# Routes
# -----------------------
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    else:
        return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'], strict_slashes=False)
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))

    if request.method == "POST":
        username = request.form['username']
        password = request.form['password']

        with db_pool.get_connection() as db:
            user = db.execute("SELECT id, username, password, is_admin FROM users WHERE username=?",
                              (username,)).fetchone()

            if user:
                stored_hash = user['password']
                if stored_hash.startswith('$'):
                    stored_hash = stored_hash[1:]

                if check_password_hash(stored_hash, password):
                    session['user_id'] = user['id']
                    session['username'] = user['username']
                    session['is_admin'] = bool(user['is_admin'])
                    log_activity(username, "LOGIN")
                    flash("Login successful", "success")
                    return redirect(url_for('dashboard'))

        flash("Invalid username or password", "error")

    return render_template('login.html')


@app.route("/register", methods=["GET", "POST"], strict_slashes=False)
def register():
    if request.method == "POST":
        username = request.form['username']
        password = request.form['password']
        email = request.form.get('email', '')

        with db_pool.get_connection() as db:
            try:
                hashed = generate_password_hash(password, method='pbkdf2:sha256')
                db.execute("INSERT INTO users (username,password,email) VALUES (?,?,?)",
                           (username, hashed, email))
                db.commit()
                flash("Registration successful", "success")
                return redirect(url_for('login'))
            except sqlite3.IntegrityError:
                db.rollback()
                flash("Username/email already exists", "error")

    return render_template("register.html")


@app.route('/logout', strict_slashes=False)
def logout():
    user = session.get('username')
    if user:
        log_activity(user, "LOGOUT")
    session.clear()
    flash("Logged out successfully.")
    return redirect(url_for('login'))

def handle_macro_route(route_type, template_name):
    if 'user_id' not in session:
        flash("Please log in to continue.")
        return redirect(url_for('login'))

    if request.method == 'POST':
        return _process_macro_request(route_type)

    download_token = request.args.get('download_token')
    route_config = ROUTE_MACROS.get(route_type, {})

    return render_template(template_name,
                           download_token=download_token,
                           route_config=route_config,
                           macro_names=route_config.get('macros', []))
# -----------------------
# Routes (patched with role_required)
# -----------------------
@app.route('/language', methods=['GET', 'POST'], strict_slashes=False)
@role_required(ROUTE_PERMISSIONS.get('language', ['ADMIN']))
def language():
    return handle_macro_route('language', 'language_edit.html')

@app.route('/macro_processing', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('macro_processing', ['ADMIN']))
def macro_processing():
    return handle_macro_route('macro_processing', 'macro_processing.html')

from jinja2 import Template

@app.route("/ppd", methods=["GET", "POST"])
@csrf.exempt
@role_required(ROUTE_PERMISSIONS.get('ppd', ['ADMIN']))
def ppd():
    if request.method == "GET":
        return render_template("ppd.html")

    # -----------------------
    # Upload Handling
    # -----------------------
    uploaded = request.files.getlist("docfiles")
    if not uploaded:
        return jsonify({"error": "No files uploaded"}), 400

    # Unique job token folder
    token = uuid.uuid4().hex
    unique_folder = os.path.join(app.config['UPLOAD_FOLDER'], token)
    os.makedirs(unique_folder, exist_ok=True)

    # Register token immediately so it's downloadable from history
    download_tokens[token] = {
        'path': unique_folder,
        'expires': datetime.now() + TOKEN_TTL,
        'user': session.get('username'),
        'route_type': 'ppd'
    }

    saved = []
    for f in uploaded:
        fn = os.path.basename(f.filename)

        if not fn.lower().endswith((".doc", ".docx")):
            continue

        # sanitize filename for Windows
        fn = re.sub(r'[<>:"/\\|?*]', "_", fn)

        save_path = os.path.join(unique_folder, fn)
        f.save(save_path)
        saved.append(save_path)

    if not saved:
        return jsonify({"error": "No valid .doc/.docx files uploaded"}), 400

    # Capture username before thread starts
    username = session.get("username") or "Analyst"

    # Job ID + progress tracking
    job_id = str(int(time.time() * 1000))
    app.config.setdefault("PROGRESS_DATA", {})
    app.config["PROGRESS_DATA"][job_id] = {
        "total": len(saved),
        "current": 0,
        "status": "Starting",
        "folder": unique_folder
    }

    # -----------------------
    # Background thread processing
    # -----------------------
    def process_job(username, user_id):
        with app.app_context():
            from word_analyzer_docx import (
                CitationAnalyzer,
                extract_with_word,
                extract_with_docx,
                remove_tags_keep_formatting_docx,
                generate_formatting_html,
                generate_multilingual_html,
                build_comments_html,
                build_export_highlight_html,
                build_detailed_summary_table,
                DASHBOARD_CSS,
                DASHBOARD_JS,
                HTML_WRAPPER,
                HAS_WIN32COM,
            )

            results = []

            for i, path in enumerate(saved, 1):
                fname = os.path.basename(path)
                app.config["PROGRESS_DATA"][job_id].update({
                    "current": i,
                    "status": f"Processing {fname}"
                })

                try:
                    # Extract content
                    if os.name == "nt" and HAS_WIN32COM:
                        paras, comments, imgs, foot, end = extract_with_word(path)
                    else:
                        paras, comments, imgs, foot, end = extract_with_docx(path)

                    remove_tags_keep_formatting_docx(path)

                    analyzer = CitationAnalyzer()
                    doc_data = [(t, p, c) for (t, p, c, _) in paras]
                    dtypes = analyzer.analyze_document_citations(doc_data)
                    table_count = len(dtypes.get("Table", {}).get("Caption", {}))

                    fmt_html = generate_formatting_html(path, used_word=False)
                    spec_html = generate_multilingual_html(path)
                    com_html = build_comments_html(comments)
                    summary_html = build_detailed_summary_table(
                        dtypes, imgs, table_count, foot, end,
                        fmt_html, spec_html, com_html
                    )
                    msr_html = analyzer.build_citation_tables_html(dtypes, fname)
                    exp_html = build_export_highlight_html(paras)

                    wc = sum(len(t.split()) for (t, _, _, _) in paras)

                    # Render HTML without Flask context
                    template = Template(HTML_WRAPPER)
                    html = template.render(
                        doc_name=fname,
                        pages=(len(paras) // 40) + 1,
                        words=wc,
                        ce_pages=(wc // 250) + 1,
                        date=_now_utc().strftime("%d-%m-%Y"),
                        analyst=username,
                        detailed_summary=summary_html,
                        msr_content=msr_html,
                        fmt_content=fmt_html,
                        spec_content=spec_html,
                        comment_content=com_html,
                        export_highlight=exp_html,
                        images=imgs,
                        footnotes=foot,
                        endnotes=end,
                        css=DASHBOARD_CSS,
                        js=DASHBOARD_JS,
                        logo_path="",
                    )

                    out_html = os.path.join(unique_folder, Path(path).stem + "_Dashboard.html")
                    with open(out_html, "w", encoding="utf-8") as f:
                        f.write(html)

                    results.append(out_html)

                    # Convert to XLS (no images)
                    excel_output = html_to_excel_no_images(out_html, unique_folder)
                    if excel_output:
                        results.append(excel_output)
                    else:
                        log_errors([f"Excel conversion FAILED for {out_html}"])

                except Exception as e:
                    app.logger.error(f"Failed processing {fname}: {e}")
                    app.config["PROGRESS_DATA"][job_id]["status"] = f"Failed: {e}"
                    break

            # Create ZIP
            zip_path = os.path.join(unique_folder, "PPD_Results.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                for f in results + saved:
                    z.write(f, arcname=os.path.basename(f))

            app.config["PROGRESS_DATA"][job_id].update({
                "status": "Completed",
                "current": len(saved),
                "zip_path": zip_path
            })

            # --- DB LOGGING START ---
            try:
                with db_pool.get_connection() as db:
                    db.execute(
                        '''INSERT INTO macro_processing 
                           (user_id, token, original_filenames, processed_filenames, selected_tasks, route_type)
                           VALUES (?, ?, ?, ?, ?, ?)''',
                        (user_id,
                         token,
                         json.dumps([os.path.basename(f) for f in saved]),  # Original filenames
                         json.dumps([os.path.basename(f) for f in results]), # Processed filenames
                         json.dumps({
                             'route_type': 'ppd', 
                             'task_indices': []
                         }), 
                         'ppd')
                    )
                    db.commit()
            except Exception as e:
                print(f"DB Logging Error (PPD): {e}")
            # --- DB LOGGING END ---

    current_user_id = session.get('user_id')
    threading.Thread(target=process_job, args=(username, current_user_id), daemon=True).start()
    return jsonify({"job_id": job_id})




@app.route("/progress/<job_id>")
def progress(job_id):
    return jsonify(app.config.get("PROGRESS_DATA", {}).get(job_id, {}))


@app.route("/download_zip/<job_id>")
def download_zip(job_id):
    data = app.config.get("PROGRESS_DATA", {}).get(job_id)
    if not data or "zip_path" not in data:
        return "Not ready", 404

    zip_path = data["zip_path"]
    folder_path = data.get("folder")  # the unique folder used for this job

    if not os.path.exists(zip_path):
        return "ZIP not found", 404

    # Read ZIP into memory
    try:
        with open(zip_path, "rb") as f:
            zip_bytes = f.read()
    except Exception:
        return "Failed reading zip", 500

    # ----- AUTO CLEANUP -----
    try:
        if folder_path and os.path.exists(folder_path):
            shutil.rmtree(folder_path, ignore_errors=True)
        if job_id in app.config.get("PROGRESS_DATA", {}):
            del app.config["PROGRESS_DATA"][job_id]
    except Exception as e:
        app.logger.error(f"Cleanup error for job {job_id}: {e}")
    # -------------------------

    # Return ZIP to client
    return send_file(
        io.BytesIO(zip_bytes),
        mimetype="application/zip",
        as_attachment=True,
        download_name="MSS_Review_Result.zip"
    )


# -----------------------
# File Validation Route
# -----------------------
@app.route("/validate", methods=["GET", "POST"], strict_slashes=False)
def validate_file():
    if 'user_id' not in session:
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({"success": False, "message": "Please log in to continue"})
        return redirect(url_for('login'))

    if request.method == "POST":
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

        # Validate file field
        uploaded_files = request.files.getlist('files')
        if not uploaded_files or not uploaded_files[0].filename:
             # Try 'file' (singular) as fallback or from other forms
             uploaded_files = request.files.getlist('file')

        if not uploaded_files or not uploaded_files[0].filename:
            msg = "No files selected"
            if is_ajax:
                return jsonify({"success": False, "message": msg})
            flash(msg, "error")
            return redirect(request.url)

        results_list = []
        token = uuid.uuid4().hex
        processing_dir = os.path.join(app.config['UPLOAD_FOLDER'], token)
        os.makedirs(processing_dir, exist_ok=True)
        
        processed_file_paths = [] # For ZIP

        try:
            for file in uploaded_files:
                filename = secure_filename(file.filename)
                if not allowed_file(filename):
                    continue

                filepath = os.path.join(processing_dir, filename)
                file.save(filepath)

                # Process
                try:
                     is_report_only = str(request.form.get('report_only')).lower() in ['true', 'on', '1']
                     # If Report Only is selected, we MUST run validation to get the stats
                     run_validation = (str(request.form.get('run_validation')).lower() in ['true', 'on', '1']) or is_report_only
                     
                     run_structuring = str(request.form.get('run_structuring')).lower() in ['true', 'on', '1']
                     run_name_year = str(request.form.get('run_name_year_validation')).lower() in ['true', 'on', '1']

                     current_filepath = filepath
                     status_parts = []
                     
                     # Data collectors for Report/DB
                     before = {'status': 'Skipped'}
                     after = {'status': 'Skipped'}
                     mapping = {}
                     fix_log_content = None
                     apa_report_text = None
                     
                     structured_filename = None
                     renumbered_name = None
                     annotated_name = None

                     # =========================================================
                     # STEP 1: Structuring (Run First as requested)
                     # =========================================================
                     if run_structuring:
                         try:
                             status_parts.append("Structuring Included")
                             # process_docx_file takes input Pth, output Path
                             struct_res = process_docx_file(Path(current_filepath), Path(processing_dir))
                             fixed_docx = struct_res.get('output_docx')
                             fix_log = struct_res.get('log_file')
                             
                             if fix_log and fix_log.exists():
                                 try:
                                     with open(fix_log, "r", encoding="utf-8") as f:
                                         fix_log_content = f.read()
                                 except Exception:
                                     fix_log_content = "Error reading log file."
                                 # We don't add log file to ZIP if we merge it into report
                             
                             if fixed_docx and fixed_docx.exists():
                                 # Structure Step Successful -> Update Chain
                                 current_filepath = str(fixed_docx)
                                 structured_filename = fixed_docx.name
                         except Exception as e:
                             log_errors([f"Structuring failed: {e}"])
                             status_parts.append(f"Structuring Failed: {str(e)}")

                     # =========================================================
                     # STEP 2: Numerical Validation
                     # =========================================================
                     if run_validation:
                         # Runs on current_filepath (which might be the structured one)
                         try:
                             # process_document returns: doc, before, after, mapping, status_msg
                             doc, before, after, mapping, val_msg = process_document(current_filepath)
                             status_parts.append(f"Num Val: {val_msg}")
                             
                             # Check if we need to save the renumbered file
                             is_perfect = before.get('is_perfect', False)
                             
                             # We ALWAYS save if there is a mapping, because we apply formatting (Style/Superscript)
                             # regardless of whether the numbers actually changed.
                             has_citations = bool(mapping)
                             
                             if has_citations or (not is_perfect):
                                 base_n = os.path.splitext(os.path.basename(current_filepath))[0]
                                 renumbered_path = os.path.join(processing_dir, f"{base_n}_Val.docx")
                                 doc.save(renumbered_path)
                                 current_filepath = renumbered_path
                                 renumbered_name = os.path.basename(renumbered_path)
                         except Exception as e:
                             log_errors([f"Numerical Validation Failed: {e}"])
                             status_parts.append(f"Num Val Failed: {str(e)}")

                     # =========================================================
                     # STEP 3: Name & Year Validation
                     # =========================================================
                     if run_name_year:
                         try:
                             # Runs on current_filepath
                             # Note: validate_document_multi_style loads the doc from disk
                             apa_results = validate_document_multi_style(current_filepath)
                             formatted_count = apply_citation_formatting(current_filepath, apa_results)
                             
                             annotated_doc, comment_count = insert_comments_in_document(
                                 current_filepath, 
                                 apa_results, 
                                 apa_results['citation_locations'], 
                                 apa_results['reference_details']
                             )
                             
                             status_parts.append(f"Name/Year: {comment_count} comments")
                             apa_report_text = generate_apa_report(apa_results, filename)
                             
                             # If we made changes or added comments, save
                             if comment_count > 0 or formatted_count > 0:
                                 base_n = os.path.splitext(os.path.basename(current_filepath))[0]
                                 annotated_path = os.path.join(processing_dir, f"{base_n}_NY.docx")
                                 annotated_doc.save(annotated_path)
                                 current_filepath = annotated_path
                                 annotated_name = os.path.basename(annotated_path)

                             # If Validation was SKIPPED, use these stats for DB
                             if not run_validation:
                                 before['total_references'] = apa_results.get('total_references', 0)
                                 before['total_citations'] = apa_results.get('total_citations', 0)
                                 before['missing_references'] = [m.get('reference', 'Unknown') for m in apa_results.get('missing_references', [])]
                                 before['unused_references'] = [u.get('reference', 'Unknown') for u in apa_results.get('unused_references', [])]
                                 after['total_references'] = before['total_references']
                                 after['total_citations'] = before['total_citations']
                         except Exception as e:
                             log_errors([f"Name/Year Failed: {e}"])
                             status_parts.append(f"Name/Year Error: {str(e)}")

                     # =========================================================
                     # FINALIZATION & REPORTING
                     # =========================================================
                     if is_report_only:
                         status_parts.append("(Report Only)")
                     
                     status_msg = " | ".join(status_parts) if status_parts else "Skipped"
                     
                     # Generate Combined Report
                     base_name = os.path.splitext(filename)[0]
                     report_name = f"{base_name}_Process_Report.txt"
                     report_path = os.path.join(processing_dir, report_name)
                     
                     with open(report_path, "w", encoding="utf-8") as f:
                        f.write(f"PROCESS REPORT FOR: {filename}\n")
                        f.write(f"STATUS: {status_msg}\n")
                        f.write("="*60 + "\n\n")
                        
                        if run_structuring:
                            f.write("--- STRUCTURING LOG ---\n")
                            f.write(fix_log_content if fix_log_content else "No log generated or read error.\n")
                            f.write("\n\n")

                        if run_validation:
                            f.write("--- NUMERICAL VALIDATION STATS ---\n")
                            f.write("BEFORE:\n" + str(before) + "\n")
                            f.write("AFTER:\n" + str(after) + "\n")
                            if mapping:
                                f.write("Renumbering Mapping:\n")
                                for o, n in sorted(mapping.items(), key=lambda x: x[1]):
                                     f.write(f"{o} -> {n}\n")
                            f.write("\n\n")
                            
                        if run_name_year:
                             f.write("--- NAME & YEAR VALIDATION REPORT ---\n")
                             f.write(str(apa_report_text) + "\n")

                     # Add Report to Output List
                     processed_file_paths.append(report_path)

                     # Add FINAL Document to Output List (if not report only)
                     final_output_name = None
                     if not is_report_only and os.path.exists(current_filepath):
                         # If current_filepath is NOT the report itself (it shouldn't be)
                         # We rename it to something clean
                         final_output_name = f"{base_name}_Processed.docx"
                         final_output_path = os.path.join(processing_dir, final_output_name)
                         
                         # Avoid overwriting if source == dest (e.g. if we modified in place, though we usually updated paths)
                         if os.path.abspath(current_filepath) != os.path.abspath(final_output_path):
                             shutil.copy2(current_filepath, final_output_path)
                         
                         processed_file_paths.append(final_output_path)

                     results_list.append({
                         'filename': filename,
                         'before': before,
                         'after': after,
                         'mapping': mapping,
                         'status_msg': status_msg,
                         'renumbered_file': renumbered_name,
                         'structured_file': structured_filename,
                         'structuring_log': fix_log_content,
                         'apa_log': apa_report_text,
                         'report_file': report_name,
                         'offline_report': f"{base_name}_Report.html",
                         'final_file': final_output_name
                     })
                     
                     # --- DB LOGGING START ---
                     try:
                         with db_pool.get_connection() as db:
                            # 1. Insert into files
                            cursor = db.execute(
                                'INSERT INTO files (user_id, original_filename, stored_filename, report_filename) VALUES (?, ?, ?, ?)',
                                (session['user_id'], filename, filename, report_name)
                            )
                            file_id = cursor.lastrowid
                            
                            # 2. Insert into validation_results
                            # Convert lists to JSON string or comma-separated string for Text fields
                            missing_str = ", ".join(map(str, before.get('missing_references', [])))
                            unused_str = ", ".join(map(str, before.get('unused_references', [])))
                            # sequence_issues is a list of dicts, let's dump as JSON
                            seq_str = json.dumps(before.get('sequence_issues', []))
                            
                            db.execute(
                                'INSERT INTO validation_results (file_id, total_references, total_citations, missing_references, unused_references, sequence_issues) VALUES (?, ?, ?, ?, ?, ?)',
                                (file_id, 
                                 before.get('total_references', 0),
                                 before.get('total_citations', 0),
                                 missing_str,
                                 unused_str,
                                 seq_str)
                            )
                            db.commit()
                     except Exception as e:
                         # Log error but don't fail the request
                         print(f"DB Logging Error (Validate): {e}")
                     # --- DB LOGGING END ---


                     
                     
                except Exception as e:
                    log_errors([f"Error processing {filename}: {str(e)}"])
                    results_list.append({
                        'filename': filename,
                        'error': str(e),
                        'status_msg': f"Failed: {str(e)}"
                    })
                finally:
                    # Cleanup original file so it doesn't get zipped
                    try:
                        if os.path.exists(filepath):
                            os.remove(filepath)
                    except:
                        pass

            # Create Individual Offline HTML Reports
            if results_list:
                for res in results_list:
                    try:
                        base_name = os.path.splitext(res['filename'])[0]
                        html_report_name = f"{base_name}_Report.html"
                        html_report_path = os.path.join(processing_dir, html_report_name)
                        
                        with app.app_context():
                            rendered_html = render_template(
                                "offline_report.html",
                                results_list=[res], # Pass single result as a list
                                token=None,
                                offline_mode=True,
                                now=datetime.now
                            )
                        with open(html_report_path, "w", encoding="utf-8") as f:
                            f.write(rendered_html)
                        processed_file_paths.append(html_report_path)
                    except Exception as e:
                        log_errors([f"Failed to generate HTML report for {res.get('filename')}: {e}"])

            # Register for download if we have results
            if processed_file_paths:
                download_tokens[token] = {
                    'path': processing_dir,
                    'expires': datetime.now() + timedelta(hours=1),
                    'user': session.get('username', 'unknown'),
                    'route_type': 'validation',
                    'results_list': results_list  # Store results here instead of session
                }
            else:
                 # If no files generated but we have results (e.g. errors), still cache them
                 download_tokens[token] = {
                    'path': processing_dir, # might be empty
                    'expires': datetime.now() + timedelta(hours=1),
                    'user': session.get('username', 'unknown'),
                    'route_type': 'validation',
                    'results_list': results_list
                 }

            # Store only token in session
            session['validation_token'] = token
            
            if is_ajax:
                 return jsonify({"success": True, "redirect": url_for('validate_file')})

            return redirect(url_for('validate_file'))

        except Exception as e:
            log_errors([f"Batch validation failed: {str(e)}"])
            flash(f"An error occurred: {str(e)}", "error")
            return redirect(request.url)

    # GET request
    if request.args.get('new'):
        session.pop('validation_token', None)
        return redirect(url_for('validate_file'))

    token = session.get('validation_token')
    if not token or token not in download_tokens:
        # Show upload page if no results or expired
        return render_template("upload.html")
    
    data = download_tokens[token]
    
    # We serve the result page. 
    # The 'download_token' logic in `macro_download` can handle the ZIP download if we pass the token.
    # But `result.html` expects `zip_file` name. 
    # I will modify `result.html` to use the token for download link: /macro-download?token={{ token }}
    
    return render_template(
        "result.html",
        results_list=data.get('results_list', []),
        token=token,
        offline_mode=False,
        now=datetime.now
    )


import base64

def get_base64_logo():
    logo_path = os.path.join(app.static_folder, "images", "S4c.png")
    try:
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        app.logger.warning(f"Logo not found or failed to load: {e}")
        return ""

# -----------------------
# Dashboard
# -----------------------
@app.route("/dashboard", strict_slashes=False)
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    with db_pool.get_connection() as db:
        if session.get('is_admin'):
            recent_files = db.execute('''SELECT f.*, u.username 
                                       FROM files f 
                                       JOIN users u ON f.user_id = u.id 
                                       ORDER BY f.upload_date DESC LIMIT 5''').fetchall()

            recent_macro = db.execute('''SELECT m.*, u.username 
                                       FROM macro_processing m
                                       JOIN users u ON m.user_id = u.id 
                                       ORDER BY m.processing_date DESC LIMIT 5''').fetchall()

            # Route-specific stats
            route_stats = {}
            for route_type in ROUTE_MACROS.keys():
                count = db.execute("SELECT COUNT(*) FROM macro_processing WHERE route_type = ?",
                                   (route_type,)).fetchone()[0]
                route_stats[route_type] = count

            admin_stats = {
                'total_users': db.execute("SELECT COUNT(*) FROM users").fetchone()[0],
                'total_files': db.execute("SELECT COUNT(*) FROM files").fetchone()[0],
                'total_validations': db.execute("SELECT COUNT(*) FROM validation_results").fetchone()[0],
                'total_macro': db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0],
                'route_stats': route_stats
            }
        else:
            recent_files = db.execute('''SELECT * FROM files 
                                       WHERE user_id=? 
                                       ORDER BY upload_date DESC LIMIT 5''',
                                      (session['user_id'],)).fetchall()

            recent_macro = db.execute('''SELECT * FROM macro_processing 
                                       WHERE user_id=? 
                                       ORDER BY processing_date DESC LIMIT 5''',
                                      (session['user_id'],)).fetchall()

            # User-specific route stats
            route_stats = {}
            for route_type in ROUTE_MACROS.keys():
                count = db.execute("SELECT COUNT(*) FROM macro_processing WHERE user_id = ? AND route_type = ?",
                                   (session['user_id'], route_type)).fetchone()[0]
                route_stats[route_type] = count

            admin_stats = {'route_stats': route_stats}

    return render_template("dashboard.html",
                           recent_files=recent_files,
                           recent_macro=recent_macro,
                           admin_stats=admin_stats,
                           route_macros=ROUTE_MACROS)


# -----------------------
# Download Route
# -----------------------
@app.route('/macro-download', strict_slashes=False)
def macro_download():
    token = request.args.get('token')
    if not token:
        flash("Invalid download request.")
        return redirect(url_for('dashboard'))

    token_data = download_tokens.get(token)
    if not token_data:
        flash("Invalid or expired download token.")
        return redirect(url_for('dashboard'))

    if is_token_expired(token_data):
        cleanup_token_data(token)
        flash("Download token has expired.")
        return redirect(url_for('dashboard'))

    user_folder = token_data['path']
    route_type = token_data.get('route_type', 'general')

    if not os.path.exists(user_folder):
        flash("No files found for this download token.")
        return redirect(url_for('dashboard'))

    try:
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
            for root, _, files in os.walk(user_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    if os.path.getsize(file_path) < 50 * 1024 * 1024:
                        arcname = os.path.relpath(file_path, user_folder)
                        zipf.write(file_path, arcname)

        memory_file.seek(0)

        try:
            shutil.rmtree(user_folder)
            del download_tokens[token]
        except Exception as e:
            log_errors([f"Cleanup error: {str(e)}"])

        route_name = ROUTE_MACROS.get(route_type, {}).get('name', 'Processed')
        zip_filename = f"{route_name.replace(' ', '_')}_Documents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(memory_file,
                         mimetype='application/zip',
                         as_attachment=True,
                         download_name=zip_filename)

    except Exception as e:
        flash(f"Download failed: {str(e)}")
        log_errors([f"Download error for token {token}: {str(e)}"])
        return redirect(url_for('dashboard'))


# -----------------------
# File History
# -----------------------
@app.route('/history', strict_slashes=False)
def file_history():
    if not g.user:
        flash("Please log in to view file history", "error")
        return redirect(url_for('login'))

    page = int(request.args.get('page', 1))
    per_page = 10
    offset = (page - 1) * per_page
    route_filter = request.args.get('route', 'all')

    with get_db() as conn:
        cursor = conn.cursor()

        # Filter logic
        filter_condition = ""
        params = []

        if route_filter != "all":
            if route_filter == "validation":
                filter_condition = "WHERE type = 'validation'"
            else:
                filter_condition = "WHERE type = 'macro' AND route_type = ?"
                params.append(route_filter)

        # Admin vs User-specific
        if session.get("is_admin"):
            user_condition = ""
        else:
            user_condition = "AND user_id = ?" if filter_condition else "WHERE user_id = ?"
            params.append(g.user["id"])

        # Unified query
        query = f"""
            SELECT * FROM (
                SELECT f.id,
                       f.original_filename AS original_filename,
                       f.upload_date AS date,
                       f.report_filename,
                       v.total_references,
                       v.total_citations,
                       u.username,
                       'validation' AS type,
                       '' AS route_type,
                       '' AS token,
                       '' AS selected_tasks,
                       '' AS original_filenames,
                       f.user_id
                FROM files f
                LEFT JOIN validation_results v ON f.id = v.file_id
                JOIN users u ON f.user_id = u.id

                UNION ALL

                SELECT m.id,
                       '' AS original_filename,
                       m.processing_date AS date,
                       '' AS report_filename,
                       0 AS total_references,
                       0 AS total_citations,
                       u.username,
                       'macro' AS type,
                       m.route_type AS route_type, 
                       m.token,
                       m.selected_tasks,
                       m.original_filenames,
                       m.user_id
                FROM macro_processing m
                JOIN users u ON m.user_id = u.id
            ) combined
            {filter_condition}
            {user_condition}
            ORDER BY date DESC
            LIMIT ? OFFSET ?
        """

        params.extend([per_page, offset])
        cursor.execute(query, params)
        history = cursor.fetchall()

        # Count total records for pagination
        count_query = f"""
            SELECT COUNT(*) FROM (
                SELECT f.id, f.user_id, 'validation' AS type
                FROM files f
                UNION ALL
                SELECT m.id, m.user_id, 'macro' AS type
                FROM macro_processing m
            ) combined
            {filter_condition}
            {user_condition}
        """
        cursor.execute(count_query, params[:-2])  # exclude LIMIT/OFFSET
        total_records = cursor.fetchone()[0]

    total_pages = (total_records + per_page - 1) // per_page

    return render_template(
        "file_history.html",
        history=history,
        page=page,
        total_pages=total_pages,
        route_filter=route_filter,
        route_macros=ROUTE_MACROS
    )
# -----------------------
# Admin Routes
# -----------------------
@app.route("/admin", strict_slashes=False)
@admin_required
def admin_dashboard():
    with db_pool.get_connection() as db:
        route_stats = {}
        for route_type in ROUTE_MACROS.keys():
            count = db.execute(
                "SELECT COUNT(*) FROM macro_processing WHERE route_type = ?",
                (route_type,)
            ).fetchone()[0]
            route_stats[route_type] = count

        # totals
        total_users = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total_files = db.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        total_validations = db.execute("SELECT COUNT(*) FROM validation_results").fetchone()[0]
        total_macro = db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0]

        # roles (defensive: handle missing column)
        try:
            role_counts = db.execute(
                "SELECT role, COUNT(*) as count FROM users GROUP BY role"
            ).fetchall()
            role_stats = {
                (r["role"] if r["role"] else "USER"): r["count"] for r in role_counts
            }
        except sqlite3.OperationalError as e:
            log_errors([f"Role stats query failed: {e}"])
            role_stats = {}

        admin_stats = {
            'total_users': total_users,
            'total_files': total_files,
            'total_validations': total_validations,
            'total_macro': total_macro,
            'route_stats': route_stats
        }

    return render_template(
        "admin_dashboard.html",
        admin_stats=admin_stats,
        route_macros=ROUTE_MACROS,
        role_stats=role_stats   # ✅ now passed to template
    )


@app.route("/admin/user/<int:user_id>/change-role", methods=["POST"], strict_slashes=False)
@admin_required
def admin_change_role(user_id):
    new_role = request.form.get('role', '').upper()
    if not new_role:
        flash("No role provided", "error")
        return redirect(url_for('admin_users'))

    if user_id == session.get('user_id'):
        flash("Cannot change your own role", "error")
        return redirect(url_for('admin_users'))

    with db_pool.get_connection() as db:
        user = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash("User not found", "error")
            return redirect(url_for('admin_users'))

        db.execute("UPDATE users SET role=? WHERE id=?", (new_role, user_id))
        db.commit()
        flash("User role updated", "success")
        log_activity(session['username'], 'CHANGE_ROLE', f"user:{user['username']} -> {new_role}")
    return redirect(url_for('admin_users'))
# -----------------------
# Admin User Management
# -----------------------
@app.route("/admin/users", strict_slashes=False)
@admin_required
def admin_users():
    with db_pool.get_connection() as db:
        users = db.execute(
            'SELECT id, username, email, is_admin, role, created_at FROM users ORDER BY created_at DESC').fetchall()
    return render_template("admin_users.html", users=users)


@app.route("/admin/create-user", methods=["GET", "POST"], strict_slashes=False)
@admin_required
def admin_create_user():
    if request.method == "POST":
        username = request.form['username']
        password = request.form['password']
        email = request.form.get('email', '')
        is_admin = 'is_admin' in request.form
        role = request.form.get('role', 'USER').upper()

        with db_pool.get_connection() as db:
            try:
                hashed = generate_password_hash(password, method='pbkdf2:sha256')
                db.execute("INSERT INTO users (username,password,email,is_admin,role) VALUES (?,?,?,?,?)",
                           (username, hashed, email, is_admin, role))
                db.commit()
                flash("User created successfully", "success")
                return redirect(url_for('admin_users'))
            except sqlite3.IntegrityError:
                db.rollback()
                flash("Username/email exists", "error")

    return render_template("admin_create_user.html")


@app.route('/admin/change_password/<int:user_id>', methods=['GET', 'POST'], strict_slashes=False)
@admin_required
def admin_change_password(user_id):
    with db_pool.get_connection() as db:
        user = db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()

        if not user:
            flash("User not found.", "error")
            return redirect(url_for('admin_users'))

        if request.method == 'POST':
            new_password = request.form['new_password']
            hashed = generate_password_hash(new_password)
            db.execute("UPDATE users SET password = ? WHERE id = ?", (hashed, user_id))
            db.commit()
            flash(f"Password updated for {user['username']}.", "success")
            return redirect(url_for('admin_users'))

    return render_template("admin_change_password.html", user=user)


@app.route("/admin/user/<int:user_id>/toggle-admin", methods=["POST"], strict_slashes=False)
@admin_required
def admin_toggle_admin(user_id):
    if user_id == session.get('user_id'):
        flash("Cannot change your own admin status", "error")
        return redirect(url_for('admin_users'))

    with db_pool.get_connection() as db:
        user = db.execute("SELECT is_admin FROM users WHERE id=?", (user_id,)).fetchone()
        if not user:
            flash("User not found", "error")
            return redirect(url_for('admin_users'))

        new_status = not bool(user['is_admin'])
        db.execute("UPDATE users SET is_admin=? WHERE id=?", (new_status, user_id))
        db.commit()
        status_text = "granted" if new_status else "revoked"
        flash(f"Admin privileges {status_text}", "success")

    return redirect(url_for('admin_users'))


@app.route("/admin/user/<int:user_id>/delete", methods=["POST"], strict_slashes=False)
@admin_required
def admin_delete_user(user_id):
    # Prevent admins from deleting themselves
    if user_id == session.get('user_id'):
        flash("Cannot delete your own account", "error")
        return redirect(url_for('admin_users'))

    try:
        with db_pool.get_connection() as db:
            # Check macro history
            macro_count = db.execute(
                "SELECT COUNT(*) FROM macro_processing WHERE user_id=?",
                (user_id,)
            ).fetchone()[0]

            if macro_count > 0:
                flash("Cannot delete user with macro history", "error")
                return redirect(url_for('admin_users'))

            # Check files
            user_files = db.execute(
                "SELECT COUNT(*) FROM files WHERE user_id=?",
                (user_id,)
            ).fetchone()[0]

            if user_files > 0:
                flash("Cannot delete user with files", "error")
                return redirect(url_for('admin_users'))

            # At this point it's safe to delete user
            # Optionally remove any related rows (safety) - will cascade if you used FK cascade, but we'll be explicit
            try:
                db.execute("DELETE FROM validation_results WHERE file_id IN (SELECT id FROM files WHERE user_id=?)", (user_id,))
            except Exception:
                # ignore if validation_results references don't exist
                pass

            try:
                db.execute("DELETE FROM files WHERE user_id=?", (user_id,))
            except Exception:
                # ignore if no files
                pass

            db.execute("DELETE FROM macro_processing WHERE user_id=?", (user_id,))  # should be zero if earlier check passed
            db.execute("DELETE FROM users WHERE id=?", (user_id,))
            db.commit()

            flash("User deleted successfully", "success")
            log_activity(session.get('username', 'system'), "DELETE_USER", f"user_id:{user_id}")
            return redirect(url_for('admin_users'))

    except Exception as e:
        log_errors([f"Error deleting user {user_id}: {e}", traceback.format_exc()])
        flash("An error occurred while deleting the user", "error")
        return redirect(url_for('admin_users'))



@app.route("/admin/files")
@admin_required
def admin_files():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page

    with db_pool.get_connection() as db:
        files = db.execute('''SELECT f.*, u.username, v.total_references, v.total_citations
                           FROM files f
                           JOIN users u ON f.user_id = u.id
                           LEFT JOIN validation_results v ON f.id = v.file_id
                           ORDER BY f.upload_date DESC LIMIT ? OFFSET ?''',
                           (per_page, offset)).fetchall()

        total_count = db.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        total_pages = (total_count + per_page - 1) // per_page

    return render_template("admin_files.html", files=files, page=page, total_pages=total_pages)


@app.route("/admin/file/<int:file_id>/delete", methods=["POST"])
@admin_required
def admin_delete_file(file_id):
    with db_pool.get_connection() as db:
        file = db.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
        if not file:
            flash("File not found", "error")
            return redirect(url_for('admin_files'))

        # Delete the file from storage
        try:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file['stored_filename'])
            if os.path.exists(file_path):
                os.remove(file_path)

            # Delete report file if exists
            if file['report_filename']:
                report_path = os.path.join(REPORT_FOLDER, file['report_filename'])
                if os.path.exists(report_path):
                    os.remove(report_path)
        except Exception as e:
            flash(f"Error deleting file: {str(e)}", "error")
            return redirect(url_for('admin_files'))

        # Delete from database
        db.execute("DELETE FROM validation_results WHERE file_id=?", (file_id,))
        db.execute("DELETE FROM files WHERE id=?", (file_id,))
        db.commit()

        flash("File deleted successfully", "success")
        return redirect(url_for('admin_files'))


@app.route("/admin/stats")
@admin_required
def admin_stats():
    with db_pool.get_connection() as db:
        # Get recent files
        # Get recent files (Combined Validation + Macros)
        recent_files = db.execute('''
            SELECT * FROM (
                SELECT f.id,
                       f.original_filename AS original_filename,
                       f.upload_date AS date,
                       u.username,
                       'validation' AS type,
                       '' AS route_type,
                       '' AS original_filenames
                FROM files f
                JOIN users u ON f.user_id = u.id

                UNION ALL

                SELECT m.id,
                       '' AS original_filename,
                       m.processing_date AS date,
                       u.username,
                       'macro' AS type,
                       m.route_type AS route_type,
                       m.original_filenames
                FROM macro_processing m
                JOIN users u ON m.user_id = u.id
            ) combined
            ORDER BY date DESC LIMIT 20
        ''').fetchall()

        # Month-wise User Activity (Last 6 Months)
        from datetime import datetime
        now = datetime.now()
        month_headers = []
        # Generate last 6 months list [(year, month), ...]
        current_y, current_m = now.year, now.month
        for i in range(6):
            y, m = current_y, current_m - i
            while m <= 0:
                m += 12
                y -= 1
            month_headers.append(f"{y}-{m:02d}")
        
        # Prepare placeholders for SQL IN clause
        placeholders = ','.join(['?'] * len(month_headers))
        
        query = f'''
            SELECT u.username, strftime('%Y-%m', activity_date) as month, COUNT(*) as count
            FROM (
                SELECT user_id, upload_date as activity_date FROM files
                UNION ALL
                SELECT user_id, processing_date as activity_date FROM macro_processing
            ) a
            JOIN users u ON a.user_id = u.id
            WHERE strftime('%Y-%m', activity_date) IN ({placeholders})
            GROUP BY u.username, month
        '''
        
        raw_stats = db.execute(query, month_headers).fetchall()
        
        # Organize data: {username: {'total': 0, 'months': {'2023-01': 0, ...}}}
        user_map = {}
        
        # Initialize users first to ensure we catch those who have 0 activity in this period but exist? 
        # Or just show active ones? The previous query showed ALL users with total count.
        # Let's get ALL users totals first to keep consistency with previous view.
        
        all_users = db.execute('''
            SELECT u.username, 
                   ((SELECT COUNT(*) FROM files f WHERE f.user_id = u.id) + 
                    (SELECT COUNT(*) FROM macro_processing m WHERE m.user_id = u.id)) as total_count
            FROM users u
            ORDER BY total_count DESC
        ''').fetchall()
        
        for u in all_users:
            user_map[u['username']] = {
                'username': u['username'],
                'total': u['total_count'],
                'months': {m: 0 for m in month_headers}
            }
            
        # Fill in monthly data
        for row in raw_stats:
            uname = row['username']
            month = row['month']
            count = row['count']
            if uname in user_map and month in user_map[uname]['months']:
                user_map[uname]['months'][month] = count

        # Convert to list sorted by total count desc
        users_data = sorted(user_map.values(), key=lambda x: x['total'], reverse=True)

        # Get total counts
        total_users = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total_files = db.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        total_validations = db.execute("SELECT COUNT(*) FROM validation_results").fetchone()[0]
        total_macro = db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0]

        # Role stats
        role_counts = db.execute("SELECT role, COUNT(*) as count FROM users GROUP BY role").fetchall()
        role_stats = {r["role"]: r["count"] for r in role_counts}

    return render_template(
        "admin_stats.html",
        recent_files=recent_files,
        users_data=users_data,
        month_headers=month_headers,
        admin_stats={
            'total_users': total_users,
            'total_files': total_files,
            'total_validations': total_validations,
            'total_macro': total_macro
        },
        role_stats=role_stats,
        route_macros=ROUTE_MACROS
    )


@app.route('/doi_finder')
def doi_finder():
    """DOI Correction and Metadata Finder"""
    if 'user_id' not in session:
        flash("Please log in to continue.")
        return redirect(url_for('login'))

    return render_template('doi_finder.html')


@app.route('/api/log-action', methods=['POST'])
def log_action_api():
    """API to log client-side actions like DOI searches"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
        
    try:
        data = request.get_json()
        action_type = data.get('action_type', 'unknown')
        details = data.get('details', {})
        
        user_id = session.get('user_id')
        token = uuid.uuid4().hex  # Generate a dummy token for the existing schema
        
        with db_pool.get_connection() as db:
            db.execute('''INSERT INTO macro_processing 
                          (user_id, token, original_filenames, processed_filenames, selected_tasks, route_type)
                          VALUES (?, ?, ?, ?, ?, ?)''',
                       (user_id, 
                        token, 
                        json.dumps([details.get('query', 'single_lookup')]), # Store query/filename here
                        json.dumps([]), 
                        json.dumps(details), 
                        action_type)) # Use route_type to store the action (e.g., 'doi_finder')
            db.commit()
            
        return jsonify({'status': 'logged'})
    except Exception as e:
        app.logger.error(f"Failed to log action: {e}")
        return jsonify({'error': str(e)}), 500

@app.route("/admin/macro-stats")
@admin_required
def admin_macro_stats():
    with db_pool.get_connection() as db:
        macro_records = db.execute('''SELECT selected_tasks, processing_date, errors, route_type
                                    FROM macro_processing 
                                    ORDER BY processing_date DESC''').fetchall()

    route_stats = {}
    error_stats = {}
    daily_stats = {}

    for record in macro_records:
        route_type = record['route_type'] or 'unknown'

        # Count by route
        if route_type not in route_stats:
            route_stats[route_type] = 0
        route_stats[route_type] += 1

        # Count errors by route
        if record['errors']:
            try:
                error_count = len(json.loads(record['errors']))
                if route_type not in error_stats:
                    error_stats[route_type] = 0
                error_stats[route_type] += error_count
            except:
                pass

        # Daily stats
        date = record['processing_date'][:10]
        if date not in daily_stats:
            daily_stats[date] = {}
        if route_type not in daily_stats[date]:
            daily_stats[date][route_type] = 0
        daily_stats[date][route_type] += 1

    return render_template("admin_macro_stats.html",
                           route_stats=route_stats,
                           error_stats=error_stats,
                           daily_stats=daily_stats,
                           route_macros=ROUTE_MACROS)


@app.route("/admin/macro-history")
@admin_required
def admin_macro_history():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page

    with db_pool.get_connection() as db:
        macro_history = db.execute('''SELECT m.*, u.username
                                    FROM macro_processing m
                                    JOIN users u ON m.user_id = u.id
                                    ORDER BY m.processing_date DESC LIMIT ? OFFSET ?''',
                                   (per_page, offset)).fetchall()

        total_count = db.execute("SELECT COUNT(*) FROM macro_processing").fetchone()[0]
        total_pages = (total_count + per_page - 1) // per_page

    return render_template("admin_macro_history.html",
                           macro_history=macro_history,
                           page=page,
                           total_pages=total_pages,
                           macro_names=ROUTE_MACROS.get('macro_processing', {}).get('macros', []))


# -----------------------
# Report Routes
# -----------------------
@app.route("/report/<filename>")
def download_report(filename):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    safe_filename = secure_filename(filename)
    if safe_filename != filename:
        flash("Invalid filename", "error")
        return redirect(url_for('dashboard'))

    with db_pool.get_connection() as db:
        if session.get('is_admin'):
            file_exists = db.execute('SELECT 1 FROM files WHERE report_filename=?',
                                     (filename,)).fetchone()
        else:
            file_exists = db.execute('SELECT 1 FROM files WHERE report_filename=? AND user_id=?',
                                     (filename, session['user_id'])).fetchone()

        if not file_exists:
            flash("No permission to access this report", "error")
            return redirect(url_for('dashboard'))

    report_path = os.path.join(REPORT_FOLDER, filename)
    if not os.path.exists(report_path):
        flash("Report file not found", "error")
        return redirect(url_for('dashboard'))

    try:
        return send_from_directory(REPORT_FOLDER, filename, as_attachment=True, download_name=f"report_{filename}")
    except FileNotFoundError:
        flash("Report file could not be downloaded", "error")
        return redirect(url_for('dashboard'))


# -----------------------
# Reset Routes
# -----------------------
@app.route('/macro-reset', methods=['POST'])
def macro_reset_application():
    try:
        if os.path.exists(app.config['UPLOAD_FOLDER']):
            shutil.rmtree(app.config['UPLOAD_FOLDER'])
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        download_tokens.clear()
        return jsonify({"success": True, "message": "All files are cleared"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route("/reset-db", methods=["POST"])
@admin_required
def reset_database():
    if os.path.exists(DATABASE):
        os.remove(DATABASE)
    init_db()
    return "Database reset successfully! New admin created: username='admin', password='admin123'"






# -----------------------
# Background Tasks
# -----------------------
def start_background_cleanup():
    def cleanup_worker():
        while True:
            try:
                cleanup_expired_tokens()
                time.sleep(300)  # Run every 5 minutes
            except Exception as e:
                log_errors([f"Background cleanup error: {str(e)}"])

    cleanup_thread = threading.Thread(target=cleanup_worker, daemon=True)
    cleanup_thread.start()


# -----------------------
# Error Handlers
# -----------------------
from werkzeug.exceptions import NotFound

@app.errorhandler(Exception)
def handle_unexpected_error(error):
    if isinstance(error, NotFound):
        return "Not Found", 404

    app.logger.error(f'Unexpected error: {error}')
    if app.debug:
        return str(error), 500
    return 'An unexpected error occurred', 500


# -----------------------
# Application Initialization
# -----------------------
def validate_route_configuration():
    errors = []

    for route_type, config in ROUTE_MACROS.items():
        if not config.get('macros'):
            errors.append(f"Route '{route_type}' has no macros defined")

        if not config.get('name'):
            errors.append(f"Route '{route_type}' has no name defined")

    macro_path = os.path.join(COMMON_MACRO_FOLDER, DEFAULT_MACRO_NAME)
    if not os.path.exists(macro_path):
        errors.append(f"Macro template file not found: {macro_path}")

    if errors:
        for error in errors:
            log_errors([f"Configuration error: {error}"])
        return False

    return True


def initialize_optimized_app():
    if not validate_route_configuration():
        print("Warning: Route configuration validation failed")

    # Initialize DB
    init_db()

    # 🔹 Ensure schema upgrades (e.g. add 'role' column if missing)
    try:
        migrate_add_role_column()
    except Exception as e:
        log_errors([f"Migration failed during startup: {e}"])

    setup_logging()
    start_background_cleanup()

    # populate PPD macros into route configuration on startup (safe guard in case module missing)
    try:
        if hasattr(ppd, 'macro_names') and isinstance(ppd.macro_names, (list, tuple)):
            ROUTE_MACROS['ppd']['macros'] = ppd.macro_names
    except Exception as e:
        log_errors([f"Failed to load PPD macro names: {e}"])

    app.logger.info("Application initialized with route-specific macro processing")

    return app


from datetime import datetime, timezone

# Make sure _now_utc() and is_token_expired() are defined as above

@app.route('/technical', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('technical', ['ADMIN']))
def technical():

    if request.method == 'POST':
        uploaded_files = request.files.getlist("word_files[]")

        run_te = request.form.get("run_technical_editing") == "1"

        if not uploaded_files:
            return render_template("technical_edit.html", error="No files uploaded")

        # 🔹 Create ONE unique folder per job
        token = uuid.uuid4().hex
        unique_folder = os.path.join(UPLOAD_FOLDER, token)
        os.makedirs(unique_folder, exist_ok=True)

        processed_files = []

        for f in uploaded_files:
            filename = secure_filename(f.filename)

            # 🔥 Save input inside unique folder
            input_path = os.path.join(unique_folder, filename)
            f.save(input_path)

            # 🔥 Process output also inside the same unique folder
            output_path = os.path.join(unique_folder, filename)

            if run_te:
                print(f"[TECH] Processing Technical QA: {filename}")
                process_docx(input_path, output_path, skip_validation=True)
            else:
                shutil.copy(input_path, output_path)

            processed_files.append(filename)

        # Save manifest
        with open(os.path.join(unique_folder, "manifest.txt"), "w") as mf:
            mf.write("\n".join(processed_files))

        # Register for download
        download_tokens[token] = {
            "path": unique_folder,
            "expires": _now_utc() + TOKEN_TTL,
            "user": session.get("username"),
            "route_type": "technical"
        }

        # --- DB LOGGING START ---
        try:
            with db_pool.get_connection() as db:
                db.execute(
                    '''INSERT INTO macro_processing 
                       (user_id, token, original_filenames, processed_filenames, selected_tasks, route_type)
                       VALUES (?, ?, ?, ?, ?, ?)''',
                    (session['user_id'],
                     token,
                     json.dumps([f.filename for f in uploaded_files]),  # Original filenames
                     json.dumps(processed_files),                       # Processed filenames
                     json.dumps({
                         'route_type': 'technical', 
                         'run_technical_editing': run_te,
                         'task_indices': ['4'] if run_te else [] # '4' matches Technical QA index in template
                     }), 
                     'technical')
                )
                db.commit()
        except Exception as e:
            print(f"DB Logging Error (Technical): {e}")
        # --- DB LOGGING END ---

        return render_template("technical_edit.html", download_token=token)

    return render_template("technical_edit.html")



# =========================
# UTC-SAFE DATETIME HELPERS
# =========================

def _now_utc():
    """Return timezone-aware UTC datetime."""
    return datetime.now(timezone.utc)


def _ensure_utc(dt):
    """Convert naive datetime to UTC-aware."""
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt


def is_token_expired(token_info):
    """Safe expiration check for tokens."""
    expires = token_info.get("expires")
    if not expires:
        return True

    expires = _ensure_utc(expires)
    return _now_utc() > expires


# -----------------------
# Global Cleanup Helper for ALL Routes
# -----------------------
def cleanup_token_data(token):
    """
    Safely remove:
      - the temp folder
      - the zip file
      - the token entry from download_tokens
    Works for Technical, Macro, Language, and any other route using tokens.
    """
    try:
        token_info = download_tokens.get(token)
        if not token_info:
            return

        folder = token_info.get("path")

        # Remove processed folder
        if folder and os.path.isdir(folder):
            shutil.rmtree(folder)

        # Remove ZIP
        zip_path = folder + ".zip" if folder else None
        if zip_path and os.path.exists(zip_path):
            os.remove(zip_path)

        # Remove token entry
        download_tokens.pop(token, None)

    except Exception as e:
        log_errors([f"CLEANUP ERROR (token={token}): {str(e)}"])


@app.route('/technical/download/<token>')
def technical_download(token):
    info = download_tokens.get(token)
    if not info or is_token_expired(info):
        cleanup_token_data(token)
        flash("Invalid or expired token.")
        return redirect(url_for('dashboard'))

    folder = info["path"]
    zip_path = folder + ".zip"

    try:
        if not os.path.exists(zip_path):
            shutil.make_archive(folder, 'zip', folder)

        mem = io.BytesIO()
        with open(zip_path, 'rb') as f:
            mem.write(f.read())
        mem.seek(0)

    except Exception as e:
        log_errors([f"ZIP read/create failure: {e}"])
        flash("Download failed.")
        return redirect(url_for('dashboard'))

    cleanup_token_data(token)

    return send_file(
        mem,
        as_attachment=True,
        mimetype="application/zip",
        download_name=f"Technical_Documents_{_now_utc().strftime('%Y%m%d_%H%M%S')}.zip"
    )

# -----------------------
# Main Execution
# -----------------------
from waitress import serve

# 🔹 create app globally so waitress-serve can see it
app = initialize_optimized_app()

if __name__ == '__main__':
    print("=== S4C APPLICATION STARTUP ===")
    host_ip = get_ip_address()
    print(f"Your IP address: {host_ip}")

    port = 8081

    print(f"\nAccess URLs:")
    print(f"Local: http://localhost:{port}")
    print(f"Network: http://{host_ip}:{port}")
    print("=================================\n")

    # run with waitress directly if launched via python
    serve(app, host="0.0.0.0", port=port, threads=4)
