from flask import (
    Flask, render_template, request, redirect, flash,
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
import pythoncom
import socket
import win32com.client as win32
import time
import threading
import concurrent.futures
from pathlib import Path
from threading import Lock
from datetime import datetime, timedelta
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from flask_wtf.csrf import CSRFProtect
from functools import wraps
from waitress import serve
import pywintypes
from contextlib import contextmanager
from queue import Queue, Empty
import logging
from logging.handlers import RotatingFileHandler

# -----------------------
# Configuration
# -----------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
USER_DOCUMENTS = str(Path.home() / "Documents")
UPLOAD_FOLDER = os.path.join(USER_DOCUMENTS, "S4C-Processed-Documents")
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
        ]
    },
    'ppd': {
        'name': 'PPD Processing',
        'description': 'PPD final processing tools (from PPD_Final.py)',
        'icon': 'magic',
        'macros': [
            "PPD_HTML.GenerateDocument",
            "PPD_HTML.Generate_HTML_WORDReport"
        ]
    }
}



# Flask app
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(24)
csrf = CSRFProtect(app)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['COMMON_MACRO_FOLDER'] = COMMON_MACRO_FOLDER
app.config['REPORT_FOLDER'] = REPORT_FOLDER
app.config['DATABASE'] = DATABASE
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

# Token-based download map
download_tokens = {}


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
        pythoncom.CoInitialize()
        self.word = self._start_word_optimized()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cleanup()

    def _start_word_optimized(self):
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
    current_time = datetime.now()
    expired = [t for t, data in download_tokens.items() if current_time > data['expires']]

    for token in expired:
        try:
            token_data = download_tokens[token]
            path = token_data['path']
            route_type = token_data.get('route_type', 'unknown')

            if os.path.exists(path):
                shutil.rmtree(path)

            log_activity(token_data.get('user', 'system'),
                         f"TOKEN_CLEANUP_{route_type.upper()}",
                         details=f"Token: {token[:8]}...")

            del download_tokens[token]

        except Exception as e:
            log_errors([f"Error cleaning expired token {token}: {str(e)}"])


def kill_word_processes():
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

@app.context_processor
def inject_current_role():
    return {'current_role': get_user_role()}
# -----------------------
# Enhanced Reference Validator
# -----------------------
class ReferenceValidator:
    def __init__(self, filepath):
        self.filepath = os.path.abspath(filepath)
        self.word = None
        self.doc = None
        self.results = {
            'total_references': 0,
            'total_citations': 0,
            'missing_references': set(),
            'unused_references': set(),
            'sequence_issues': [],
            'citation_sequence': []
        }

    def __enter__(self):
        pythoncom.CoInitialize()
        self.word = win32.Dispatch("Word.Application")
        self.word.Visible = False
        self.word.ScreenUpdating = False
        self.doc = self.word.Documents.Open(self.filepath, ReadOnly=True)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc:
            self.doc.Close(SaveChanges=False)
        if self.word:
            self.word.Quit()
        pythoncom.CoUninitialize()

    def validate(self):
        ref_numbers = self._get_reference_numbers()
        self.results['total_references'] = len(ref_numbers)

        citations = self._get_citations()
        self.results['total_citations'] = len(citations)

        cited_numbers = set()
        for citation in citations:
            cited_numbers.update(citation['numbers'])
            self.results['citation_sequence'].extend(citation['numbers'])

        self.results['missing_references'] = sorted(cited_numbers - ref_numbers)
        self.results['unused_references'] = sorted(ref_numbers - cited_numbers)
        self._check_citation_sequence()
        return self.results

    def _get_reference_numbers(self):
        try:
            style = self.doc.Styles("bib_number")
        except:
            raise ValueError("'bib_number' style not found")

        numbers = set()
        found_ranges = []
        rng = self.doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Style = style
        rng.Find.Text = ""

        while rng.Find.Execute():
            found_ranges.append(rng.Text.strip())
            rng.Collapse(0)

        for text in found_ranges:
            if text:
                numbers.update(self._extract_numbers(text))

        return numbers

    def _get_citations(self):
        try:
            style = self.doc.Styles("cite_bib")
        except:
            raise ValueError("'cite_bib' style not found")

        citations = []
        rng = self.doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Style = style
        rng.Find.Text = ""

        while rng.Find.Execute():
            text = rng.Text.strip()
            if text:
                numbers = self._extract_numbers(text)
                if numbers:
                    citations.append({
                        'text': text,
                        'numbers': numbers,
                        'range_start': rng.Start,
                        'range_end': rng.End
                    })
            rng.Collapse(0)
        return citations

    def _extract_numbers(self, text):
        numbers = []

        # Handle ranges first
        for match in re.finditer(r'(\d+)-(\d+)', text):
            start, end = int(match.group(1)), int(match.group(2))
            numbers.extend(range(start, end + 1))

        # Get individual numbers
        text_no_ranges = re.sub(r'\d+-\d+', '', text)
        numbers.extend([int(match.group()) for match in re.finditer(r'\b\d+\b', text_no_ranges)])

        return numbers

    def _check_citation_sequence(self):
        sequence = self.results['citation_sequence']
        if len(sequence) < 2:
            self.results['sequence_message'] = "Citations are in proper sequence."
            return

        seen = set()
        unique_sequence = []
        for num in sequence:
            if num not in seen:
                unique_sequence.append(num)
                seen.add(num)

        is_ordered = unique_sequence == sorted(unique_sequence)

        if not is_ordered:
            self.results['sequence_issues'].append(sequence)
            self.results['sequence_message'] = "Citations are NOT in sequence."
        else:
            self.results['sequence_message'] = "Citations are in proper sequence."


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

# -----------------------
# Role Permissions
# -----------------------
ROUTE_PERMISSIONS = {
    'language': ['COPYEDIT', 'PM', 'ADMIN'],
    'technical': ['COPYEDIT', 'PM', 'ADMIN'],
    'macro_processing': ['COPYEDIT', 'PPD', 'PM', 'ADMIN'],
    'ppd': ['PPD', 'PM', 'ADMIN']
}


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
                'expires': datetime.now() + TOKEN_TTL,
                'user': username,
                'route_type': route_type
            }
    except NameError:
        # If the lock isn't present for some reason, fall back (but warn)
        download_tokens[token] = {
            'path': unique_folder,
            'expires': datetime.now() + TOKEN_TTL,
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

@app.route('/technical', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('technical', ['ADMIN']))
def technical():
    return handle_macro_route('technical', 'technical_edit.html')

@app.route('/macro_processing', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('macro_processing', ['ADMIN']))
def macro_processing():
    return handle_macro_route('macro_processing', 'macro_processing.html')

@app.route('/ppd', methods=['GET', 'POST'])
@role_required(ROUTE_PERMISSIONS.get('ppd', ['ADMIN']))
def ppd():
    return handle_macro_route('ppd', 'ppd.html')
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
        # Check if it's an AJAX request
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'

        if 'files' not in request.files:
            if is_ajax:
                return jsonify({"success": False, "message": "No files selected"})
            flash("No files selected", "error")
            return redirect(request.url)

        uploaded_files = request.files.getlist('files')
        if len(uploaded_files) == 0 or uploaded_files[0].filename == '':
            if is_ajax:
                return jsonify({"success": False, "message": "No files selected"})
            flash("No files selected", "error")
            return redirect(request.url)

        processed_files = []
        errors = []

        with db_pool.get_connection() as db:
            for file in uploaded_files:
                if file and allowed_file(file.filename):
                    try:
                        filename = secure_filename(file.filename)
                        stored_filename = f"{session['user_id']}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{filename}"
                        filepath = os.path.join(app.config['UPLOAD_FOLDER'], stored_filename)
                        file.save(filepath)

                        with ReferenceValidator(filepath) as validator:
                            results = validator.validate()

                        report_name = f"report_{stored_filename}.html"
                        report_path = os.path.join(REPORT_FOLDER, report_name)
                        with open(report_path, 'w', encoding='utf-8') as f:
                            f.write(render_template('report_template.html',
                                                    filename=filename, results=results, datetime=datetime))

                        cursor = db.cursor()
                        cursor.execute('''INSERT INTO files (user_id, original_filename, stored_filename, report_filename)
                                          VALUES (?,?,?,?)''',
                                       (session['user_id'], filename, stored_filename, report_name))
                        file_id = cursor.lastrowid
                        cursor.execute('''INSERT INTO validation_results
                                          (file_id,total_references,total_citations,missing_references,unused_references,sequence_issues)
                                          VALUES (?,?,?,?,?,?)''',
                                       (file_id, results['total_references'], results['total_citations'],
                                        json.dumps(list(results['missing_references'])),
                                        json.dumps(list(results['unused_references'])),
                                        json.dumps(results['sequence_issues'])))
                        db.commit()

                        processed_files.append({
                            'filename': filename,
                            'report_url': url_for('download_report', filename=report_name)
                        })

                    except Exception as e:
                        errors.append(f"Error processing {file.filename}: {str(e)}")
                        app.logger.error(f"Error processing file {file.filename}: {str(e)}")
                else:
                    errors.append(f"Invalid file type for {file.filename}")

        # Handle response based on request type
        if is_ajax:
            if errors:
                return jsonify({
                    "success": False,
                    "message": f"Processed {len(processed_files)} files with {len(errors)} errors",
                    "errors": errors,
                    "processed_files": processed_files
                })
            else:
                return jsonify({
                    "success": True,
                    "message": f"Successfully processed {len(processed_files)} files",
                    "processed_files": processed_files
                })
        else:
            # For regular form submission
            if processed_files:
                flash(f"Successfully processed {len(processed_files)} files", "success")
            if errors:
                for error in errors:
                    flash(error, "error")
            return redirect(url_for('dashboard'))

    return render_template("upload.html")


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

    if datetime.now() > token_data['expires']:
        flash("Download token has expired.")
        try:
            if os.path.exists(token_data['path']):
                shutil.rmtree(token_data['path'])
            del download_tokens[token]
        except Exception as e:
            log_errors([f"Error cleaning expired token: {str(e)}"])
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
                       '' AS route_type, 
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
        recent_files = db.execute('''SELECT f.*, u.username 
                                   FROM files f 
                                   JOIN users u ON f.user_id = u.id 
                                   ORDER BY f.upload_date DESC LIMIT 10''').fetchall()

        # Get user stats
        user_stats = db.execute('''SELECT u.username, COUNT(f.id) as file_count
                                FROM users u
                                LEFT JOIN files f ON u.id = f.user_id
                                GROUP BY u.id
                                ORDER BY file_count DESC''').fetchall()

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
        user_stats=user_stats,
        admin_stats={
            'total_users': total_users,
            'total_files': total_files,
            'total_validations': total_validations,
            'total_macro': total_macro
        },
        role_stats=role_stats   # ✅ now passed to template
    )


@app.route('/doi_finder')
def doi_finder():
    """DOI Correction and Metadata Finder"""
    if 'user_id' not in session:
        flash("Please log in to continue.")
        return redirect(url_for('login'))

    return render_template('doi_finder.html')

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
@app.errorhandler(Exception)
def handle_unexpected_error(error):
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
