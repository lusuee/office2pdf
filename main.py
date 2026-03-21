import atexit
import io
import json
import logging
import os
import re
import sys
import threading
import uuid
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from urllib.parse import quote

import pythoncom
import win32com.client as win32
from flask import Flask, request, jsonify, make_response
from waitress import create_server, serve
from werkzeug.exceptions import BadRequest, ClientDisconnected, RequestEntityTooLarge
app = Flask(__name__)
app.name = "office2pdf"

# Allowed upload file extensions
ALLOWED_EXTENSIONS = {"doc", "docx", "xls", "xlsx", "ppt", "pptx"}

# Office export-to-PDF constants
WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32
WD_ALERTS_NONE = 0

# Base directory for runtime files.
# When bundled as an executable, use the executable directory instead of
# PyInstaller's temporary extraction directory.
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(BASE_DIR, "office2pdf.json")


def load_runtime_config() -> dict:
    """Load runtime config from a JSON file in BASE_DIR if it exists."""
    if not os.path.exists(CONFIG_PATH):
        return {}

    with open(CONFIG_PATH, "r", encoding="utf-8") as config_file:
        config = json.load(config_file)

    if not isinstance(config, dict):
        raise ValueError("Config file must contain a JSON object.")

    return config


def get_config_value(config: dict, key: str, env_name: str, default):
    """Return config value using JSON config first, then env var, then default."""
    if key in config and config[key] not in (None, ""):
        return config[key]
    return os.environ.get(env_name, default)


def resolve_runtime_path(path_value: str, default_dir_name: str) -> str:
    """Resolve a configured path relative to BASE_DIR when not absolute."""
    if not path_value:
        return os.path.join(BASE_DIR, default_dir_name)
    if os.path.isabs(path_value):
        return path_value
    return os.path.join(BASE_DIR, path_value)


def parse_size_in_bytes(value, default: int) -> int:
    """Parse size config like 200m, 1g, 512k, or raw bytes."""
    if value in (None, ""):
        return default

    if isinstance(value, int):
        return value

    text = str(value).strip().lower()
    if text.isdigit():
        return int(text)

    units = {
        "k": 1024,
        "kb": 1024,
        "m": 1024 * 1024,
        "mb": 1024 * 1024,
        "g": 1024 * 1024 * 1024,
        "gb": 1024 * 1024 * 1024,
    }

    for suffix, multiplier in units.items():
        if text.endswith(suffix):
            number_text = text[: -len(suffix)].strip()
            if not number_text:
                break
            return int(float(number_text) * multiplier)

    raise ValueError(
        "Invalid max_content_length. Use bytes or values like 200m, 512k, or 1g."
    )


RUNTIME_CONFIG = load_runtime_config()
DEFAULT_HOST = str(get_config_value(RUNTIME_CONFIG, "host", "OFFICE2PDF_HOST", "0.0.0.0"))
DEFAULT_PORT = int(get_config_value(RUNTIME_CONFIG, "port", "OFFICE2PDF_PORT", 8081))
LOG_DIR = resolve_runtime_path(
    str(get_config_value(RUNTIME_CONFIG, "log_dir", "OFFICE2PDF_LOG_DIR", "logs")),
    "logs",
)
UPLOAD_DIR = resolve_runtime_path(
    str(
        get_config_value(
            RUNTIME_CONFIG,
            "upload_dir",
            "OFFICE2PDF_UPLOAD_DIR",
            "uploads",
        )
    ),
    "uploads",
)
MAX_CONTENT_LENGTH = parse_size_in_bytes(
    get_config_value(
        RUNTIME_CONFIG,
        "max_content_length",
        "OFFICE2PDF_MAX_CONTENT_LENGTH",
        "200m",
    ),
    200 * 1024 * 1024,
)

# Thread-local storage: each worker thread has its own COM environment
thread_local = threading.local()

# Limit request body size
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


def ensure_directory(directory: str) -> None:
    """Ensure the given directory exists."""
    try:
        os.makedirs(directory, exist_ok=True)
    except Exception as exc:
        logging.getLogger(app.name).error(
            "Error creating directory %s: %s", directory, exc
        )
        raise


def create_logger() -> logging.Logger:
    """Create a daily rotating logger."""
    ensure_directory(LOG_DIR)

    logger = logging.getLogger(app.name)
    logger.setLevel(logging.DEBUG)

    log_handler = TimedRotatingFileHandler(
        os.path.join(LOG_DIR, f"{app.name}_log"),
        when="midnight",
        interval=1,
        backupCount=30,
        encoding="utf-8",
        utc=False,
    )

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    log_handler.setFormatter(formatter)

    logger.handlers.clear()
    logger.addHandler(log_handler)

    return logger


# Global logger instance used across the module
logger = create_logger()


def get_safe_filename(filename: str) -> str:
    """Return a unique filename while preserving Unicode characters."""
    _, ext = os.path.splitext(filename)
    stem = os.path.splitext(filename)[0].strip()

    # Remove Windows-invalid filename characters but keep Unicode text.
    safe_stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).rstrip(" .")
    safe_ext = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", ext).strip().lower()

    if safe_ext and not safe_ext.startswith("."):
        safe_ext = f".{safe_ext}"

    if not safe_ext:
        safe_ext = ext.lower()

    if not safe_stem:
        safe_stem = "file"

    return f"{uuid.uuid4()}-{safe_stem}{safe_ext}"


def get_office_application(app_type: str):
    """Get a thread-local Office application instance, cached per app type."""
    if app_type not in ("Word", "Excel", "PowerPoint"):
        raise ValueError("Unsupported application type")

    if not hasattr(thread_local, "office_apps") or thread_local.office_apps is None:
        thread_local.office_apps = {}

    office_apps = thread_local.office_apps
    app_instance = office_apps.get(app_type)

    try:
        if app_instance is None:
            raise ValueError("Office application not initialized")
        # Access a property to ensure the instance is still alive
        _ = app_instance.Visible
    except (pythoncom.com_error, ValueError, AttributeError):
        if not getattr(thread_local, "com_initialized", False):
            pythoncom.CoInitialize()
            thread_local.com_initialized = True

        if app_type == "Word":
            app_instance = win32.DispatchEx("Word.Application")
        elif app_type == "Excel":
            app_instance = win32.DispatchEx("Excel.Application")
        elif app_type == "PowerPoint":
            app_instance = win32.DispatchEx("PowerPoint.Application")

        app_instance.Visible = False

        # Best-effort suppression of Office UI during background automation.
        if app_type == "Word":
            app_instance.DisplayAlerts = WD_ALERTS_NONE
            app_instance.ScreenUpdating = False
        elif app_type == "Excel":
            app_instance.DisplayAlerts = False
            app_instance.ScreenUpdating = False

        office_apps[app_type] = app_instance

    return app_instance


def cleanup_office_applications() -> None:
    """Release cached Office application instances for the current thread."""
    office_apps = getattr(thread_local, "office_apps", None)
    if office_apps:
        for app_type, app_instance in list(office_apps.items()):
            try:
                app_instance.Quit()
            except Exception as exc:
                logger.warning("Error quitting %s application: %s", app_type, exc)
        thread_local.office_apps = {}

    if getattr(thread_local, "com_initialized", False):
        try:
            pythoncom.CoUninitialize()
        except Exception as exc:
            logger.warning("Error uninitializing COM: %s", exc)
        finally:
            thread_local.com_initialized = False


def office_to_pdf_stream(file_obj, filename: str, app_type: str):
    """Convert an Office document to PDF and return an in-memory stream.

    :param file_obj: Flask file object.
    :param filename: Original filename.
    :param app_type: Application type ("Word", "Excel", or "PowerPoint").
    """
    upload_base_dir = UPLOAD_DIR
    ensure_directory(upload_base_dir)

    pdf_stream = io.BytesIO()
    doc = None
    input_full_path = None
    pdf_save_path = None

    try:
        # Path for the uploaded source file
        input_filename = get_safe_filename(filename)

        now = datetime.now()
        date_folder = os.path.join(
            upload_base_dir,
            now.strftime("%Y"),
            now.strftime("%m"),
            now.strftime("%d"),
        )
        ensure_directory(date_folder)

        input_save_path = os.path.join(date_folder, input_filename)

        # Save uploaded file to disk
        file_obj.save(input_save_path)

        input_full_path = input_save_path
        logger.info("doc_save_path: %s", input_full_path)

        # Get Office app instance and open the document
        office_app = get_office_application(app_type)

        if app_type == "Word":
            doc = office_app.Documents.Open(
                input_full_path,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
                NoEncodingDialog=True,
            )
        elif app_type == "Excel":
            doc = office_app.Workbooks.Open(
                input_full_path,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
            )
        elif app_type == "PowerPoint":
            doc = office_app.Presentations.Open(
                input_full_path,
                ReadOnly=False,
                Untitled=False,
                WithWindow=False,
            )

        if doc is None:
            logger.error("Failed to open document for %s: %s", app_type, filename)
            raise RuntimeError(f"Failed to open document for {app_type}: {filename}")

        # Filename presented to the client (without UUID)
        pdf_filename = os.path.splitext(filename)[0] + ".pdf"
        # Actual filename on disk remains unique and safe
        save_pdf_filename = get_safe_filename(pdf_filename)

        pdf_save_path = os.path.join(date_folder, save_pdf_filename)
        ensure_directory(os.path.dirname(pdf_save_path))

        logger.info("pdf_save_path: %s", pdf_save_path)
        abs_pdf_path = os.path.abspath(pdf_save_path)

        # Export as PDF
        if app_type == "Word":
            doc.SaveAs(abs_pdf_path, FileFormat=WD_FORMAT_PDF)
        elif app_type == "Excel":
            doc.ExportAsFixedFormat(XL_TYPE_PDF, abs_pdf_path)
        elif app_type == "PowerPoint":
            doc.SaveAs(abs_pdf_path, FileFormat=PP_SAVE_AS_PDF)

        # Read PDF content into memory stream
        with open(pdf_save_path, "rb") as pdf_file:
            pdf_stream.write(pdf_file.read())
        pdf_stream.seek(0)

    except Exception as exc:
        logger.error("Error occurred while processing %s document: %s", app_type, exc)
        raise
    finally:
        # Close the Office document but keep the Office application instance
        if doc is not None:
            try:
                if app_type in ("Word", "Excel"):
                    doc.Close(SaveChanges=False)
                elif app_type == "PowerPoint":
                    doc.Close()
            except Exception as exc:
                logger.warning("Error closing document: %s", exc)

        cleanup_office_applications()

    return pdf_stream, pdf_filename


def allowed_file(filename: str) -> bool:
    """Return True if the filename has an allowed extension."""
    if not filename or "." not in filename:
        return False
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in ALLOWED_EXTENSIONS


@app.route("/convert", methods=["POST"])
def upload_file():
    logger.info("Upload file request received.")
    try:
        files = request.files
    except RequestEntityTooLarge:
        logger.error(
            "Upload rejected because request size exceeds limit: %s bytes",
            MAX_CONTENT_LENGTH,
        )
        return jsonify({"error": "File too large"}), 413
    except ClientDisconnected:
        logger.warning("Client disconnected while uploading request body.")
        return jsonify({"error": "Client disconnected during upload"}), 400
    except BadRequest as exc:
        logger.error("Invalid multipart/form-data request: %s", exc)
        return jsonify({"error": "Invalid multipart form data"}), 400

    if "file" not in files:
        logger.error("No file part in the request.")
        return jsonify({"error": "No file part"}), 400

    file = files["file"]

    if file.filename == "":
        logger.error("No selected file.")
        return jsonify({"error": "No selected file"}), 400

    if file and allowed_file(file.filename):
        filename = file.filename
        ext_lower = filename.lower()

        if ext_lower.endswith((".doc", ".docx")):
            app_type = "Word"
        elif ext_lower.endswith((".xls", ".xlsx")):
            app_type = "Excel"
        else:
            app_type = "PowerPoint"

        try:
            logger.info("Converting file: %s", filename)
            pdf_stream, pdf_filename = office_to_pdf_stream(file, filename, app_type)

            response = make_response(pdf_stream.getvalue())
            response.headers["Content-Disposition"] = (
                f"attachment; filename*=UTF-8''{quote(pdf_filename)}"
            )
            response.headers["Content-Type"] = "application/pdf"

            logger.info("File converted successfully: %s", pdf_filename)
            return response

        except Exception:
            # Detailed errors are already logged; return a generic message
            logger.exception("Error during conversion.")
            return jsonify({"error": "Internal server error"}), 500

    logger.error("Invalid file format.")
    return jsonify({"error": "Invalid file format"}), 400


@app.route("/health", methods=["GET"])
def health():
    """Simple health check endpoint."""
    return jsonify({"status": "ok"}), 200


def create_http_server(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT):
    """Create a stoppable Waitress server instance."""
    return create_server(app, host=host, port=port)


def run_http_server(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> None:
    """Run the HTTP server until the process exits."""
    logger.info("Server starting on %s:%s ...", host, port)
    serve(app, host=host, port=port)


atexit.register(cleanup_office_applications)


if __name__ == "__main__":
    try:
        run_http_server(port=DEFAULT_PORT)
    except Exception as exc:
        logger.error("Server error: %s", exc)
