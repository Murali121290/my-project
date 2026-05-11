# S4Carlisle Hub Server - Project Documentation

## 1. Introduction

The **S4Carlisle Hub Server** is a centralized web application designed to automate and streamline document processing workflows. It specializes in academic and technical publishing tasks, offering tools for reference validation, copyediting (macro processing), credit extraction, and XML conversion.

Built with **Flask**, the system handles complex file operations asynchronously, providing real-time feedback and processing large documents efficiently.

## 2. System Architecture

- **Backend Framework**: Python Flask.
- **Task Management**: Threading-based background workers for asynchronous file processing (preventing UI freezes during long operations).
- **Database**:
    - **SQLite**: Default for development and lightweight deployments (`reference_validator.db`).
    - **PostgreSQL**: Supported for production environments (configured via `db_config`).
- **Frontend**: Jinja2 templating engine with HTML5, CSS3, and JavaScript for dynamic interactions.
- **Server**: 
    - **Waitress** (Windows production/dev).
    - **Gunicorn** (Linux production).

## 3. Key Modules & Features

### 3.1 Reference Validation (The "Validation Pipeline")
Located in `app_server.py`, `ReferenceAPAValidation.py`, and `ReferencesStructing.py`.

This is a multi-stage pipeline triggered via the **Upload** page (`/validate`). It processes files in the following order (cumulative workflow):

1.  **Structuring** (`run_structuring`):
    - Parses the document to identify and format the reference section.
    - Generates a structure-fixed version of the document.
2.  **Number Validation** (`run_validation`):
    - Checks citation numbering sequence and integrity.
    - Updates the document from stage 1.
3.  **Name & Year Validation** (`run_name_year`):
    - Validates "Name, Year" citations (APA/Chicago style).
    - Checks against external APIs (Crossref/PubMed) if configured.
    - Inserts comments for missing or invalid references.

**Output**: A concise ZIP file containing:
- `*_Processed.docx`: The final document with all selected validations applied.
- `*_log.txt`: A consolidated log of all actions taken (Structuring logs + Validation messages).

### 3.2 Copyediting (Macro Processing)
*Formerly referred to as Macro Processing.*

These tools automate common editing tasks. On Windows, they leverage Microsoft Word's COM interface for deep integration.

- **Language Editing**: Grammar checks, spell checks, style consistency.
- **Technical Editing**: Reference renumbering, duplicate checking, technical highlighting.
- **PPD Processing**: Final pre-production document preparation (`PPD_Final.py`).

**Configuration**: Defined in `ROUTE_MACROS` dictionary in `app_server.py`.

### 3.3 Credit Extractor
Located in `extractor.py`.
- **Purpose**: Scans documents to extract image captions and credit lines.
- **Output**: Generates a permission log (Excel) to track copyright permissions.

### 3.4 Word to XML
Located in `wordtoxml/`.
- **Purpose**: Converts structured Word documents into XML format for publishing pipelines.

### 3.5 DOI Finder
- A utility tool to batch-search DOIs for a list of references.

## 4. User Roles & Permissions

The application implements Role-Based Access Control (RBAC):

- **ADMIN**: Full access to all modules, user management, and statistics.
- **PM (Project Manager)**: Access to reports, PPD, and Credit Extractor.
- **COPYEDIT**: Access to Language and Technical editing tools.
- **PPD**: Specific access to PPD processing.
- **PERMISSIONS**: Access to Credit Extractor.

Access control is enforced via the `@role_required` decorator in `app_server.py` and strictly checked against `ROUTE_PERMISSIONS`.

## 5. Directory Structure

```text
my-project/
├── app_server.py            # MAIN ENTRY POINT: Flask app and routes
├── ReferenceAPAValidation.py# Logic for APA style validation
├── ReferencesStructing.py   # Logic for reference section structuring
├── extractor.py             # Credit extractor logic
├── templates/               # HTML Templates (UI)
│   ├── upload.html          # Main validation upload page
│   ├── dashboard.html       # User dashboard
│   └── ...
├── static/                  # CSS/JS assets
├── S4C-Processed-Documents/ # Temp storage for uploads & processing
├── logs/                    # Application logs
├── reference_validator.db   # SQLite Database file
└── requirements.txt         # Python dependencies
```

## 6. Deployment & Setup

### Prerequisites
- Python 3.9+
- Microsoft Word (Required for `win32com` macros on Windows)

### Local Run (Windows)
1.  **Activate Venv**: `.\.venv\Scripts\activate`
2.  **Run Server**: `python app_server.py`
3.  **Access**: Open `http://localhost:8081`

### Docker / Linux Support
- The application includes `Dockerfile` and `docker-compose.yml`.
- **Note**: Copyediting tools depending on `win32com` (Word automation) are **disabled** on Linux containers. The Validation Pipeline (python-docx based) works fully on Linux.

## 7. Troubleshooting

- **"Phase 1: Scanning..." logs**: These indicate the Reference Structuring engine is active.
- **Download is empty**: Ensure the job finished successfully (check Status column).
- **Macro fails**: Check if the document is closed in Word before uploading. On server, ensure the Word Process instance is not hung.

---
*Documentation generated by S4Carlisle Hub AI Assistant.*
