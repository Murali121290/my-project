# S4Carlisle Hub Server

A Flask-based web application for document processing, reference validation, and citation analysis. This application supports validating references using PubMed and CrossRef APIs, processing Word documents with macros (on Windows) or python-docx (on Linux), and managing user workflows.

## Features

- **Document Upload & Management**: Upload and manage Word documents (.docx).
- **Reference Validation**: Validate citations against PubMed and CrossRef APIs.
- **Citation Analysis**: Analyze document citations and generate reports.
- **Macro Processing**: Automate Word document tasks (Windows only, requires Word installed).
- **Linux Compatibility**: Runs on Linux servers with reduced functionality (no Word COM automation).
- **User Authentication**: Role-based access control (Admin, User, etc.).
- **Dashboard**: View statistics and recent activities.

## Prerequisites

- Python 3.8+
- Flask
- Gunicorn (for Linux production)
- Nginx (optional, for reverse proxy)
- **Windows Only**: Microsoft Word (for full macro functionality)

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Murali121290/my-project.git
cd my-project
```

### 2. Set up Virtual Environment

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**Windows:**
```powershell
python -m venv venv
.\venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Database Setup

The application uses SQLite. The database will be initialized automatically on the first run.

## Running the Application

### Development (Local)

```bash
python app_server.py
```
Access the app at `http://localhost:8081` (or the IP printed in the console).

### Production (Linux Server)

1. **Install Gunicorn:**
   ```bash
   pip install gunicorn
   ```

2. **Run with Gunicorn:**
   ```bash
   gunicorn --config gunicorn_config.py app_server:app
   ```

3. **Systemd Service:**
   A `hub_app.service` file is provided for setting up a background service.
   - Copy to `/etc/systemd/system/hub_app.service`.
   - Update paths and user in the file.
   - Run:
     ```bash
     sudo systemctl enable hub_app
     sudo systemctl start hub_app
     ```

## Deployment on Linux

This project is configured for deployment on Linux servers (e.g., Ubuntu).

1.  **Clone** the repo to `/var/www/hub-server`.
2.  **Install** dependencies in a virtual environment.
3.  **Configure** Nginx to reverse proxy to port 8000.
4.  **Secure** with SSL (Certbot) if needed.

**Note:** Features requiring Microsoft Word automation (e.g., `.dotm` macros) are disabled on Linux. The application uses `python-docx` for document processing where possible.

## Project Structure

- `app_server.py`: Main Flask application entry point.
- `app_server.py`: Core logic for routing and processing.
- `word_analyzer.py` / `word_analyzer_docx.py`: Document analysis logic.
- `templates/`: HTML templates for the UI.
- `static/`: CSS, JS, and image assets.
- `S4C-Processed-Documents/`: Storage for processed files.
- `requirements.txt`: Python dependencies.
- `gunicorn_config.py`: Gunicorn configuration for production.

## License

[Add License Information Here]
