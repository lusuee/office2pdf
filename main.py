import os
import io
import logging
import threading
import pythoncom
import win32com.client as win32
import uuid

from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from flask import Flask, request, jsonify, make_response
from waitress import serve
from urllib.parse import quote
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.name = "office2pdf"

# 定义允许上传的文件类型
ALLOWED_EXTENSIONS = {"doc", "docx", "xls", "xlsx", "ppt", "pptx"}

# Office PDF Format Constants
WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32

# 基础目录 (确保路径安全)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 创建线程本地存储，确保每个线程有独立的 COM 环境
thread_local = threading.local()


def ensure_directory(directory):
    """确保目录存在"""
    try:
        os.makedirs(directory, exist_ok=True)
    except Exception as e:
        print(f"Error creating directory {directory}: {e}")


def create_logger():
    """创建每日滚动日志记录器"""
    log_dir = os.path.join(BASE_DIR, "logs")
    ensure_directory(log_dir)

    logger = logging.getLogger(app.name)
    logger.setLevel(logging.DEBUG)

    log_handler = TimedRotatingFileHandler(
        os.path.join(log_dir, f"{app.name}_log"),
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


# 在模块导入时创建全局 logger，确保模块中其它位置可以使用
logger = create_logger()


def get_safe_filename(filename):
    """
    生成安全的文件名，使用 UUID 生成唯一标识
    """
    # 生成唯一文件名
    safe = secure_filename(filename)
    return f"{uuid.uuid4()}-{safe}"


def office_to_pdf_stream(file_obj, filename, app_type):
    """
    将 Office 文档转换为 PDF 内存流
    :param file_obj: Flask file object
    :param filename: Original filename
    :param app_type: Application type (Word, Excel, PowerPoint)
    """
    upload_base_dir = os.path.join(BASE_DIR, "uploads")
    ensure_directory(upload_base_dir)

    pdf_stream = io.BytesIO()
    doc = None
    input_full_path = None
    pdf_save_path = None

    try:
        # 生成安全的文件名
        input_filename = get_safe_filename(filename)

        # 构建保存路径
        date_folder = os.path.join(
            upload_base_dir,
            datetime.now().strftime("%Y"),
            datetime.now().strftime("%m"),
            datetime.now().strftime("%d"),
        )
        ensure_directory(date_folder)

        input_save_path = os.path.join(date_folder, input_filename)

        # 优化：直接保存文件流，避免一次性读取到内存
        file_obj.save(input_save_path)

        input_full_path = input_save_path
        logger.info(f"doc_save_path: {input_full_path}")

        # 获取应用程序实例（避免与 Flask `app` 冲突，使用 office_app 作为本地变量）
        office_app = get_office_application(app_type)

        # 打开文档
        if app_type == "Word":
            doc = office_app.Documents.Open(input_full_path)
        elif app_type == "Excel":
            doc = office_app.Workbooks.Open(input_full_path)
        elif app_type == "PowerPoint":
            doc = office_app.Presentations.Open(input_full_path, WithWindow=False)

        if doc is None:
            logger.error(f"Failed to open document for {app_type}: {filename}")
            raise Exception(f"Failed to open document for {app_type}: {filename}")

        # 生成安全的 PDF 文件名
        pdf_filename = os.path.splitext(filename)[0] + ".pdf"
        save_pdf_filename = get_safe_filename(pdf_filename)

        # 保存 PDF 文件
        pdf_save_path = os.path.join(
            date_folder,
            save_pdf_filename,
        )
        # Ensure directory for PDF exists (though likely same as input)
        ensure_directory(os.path.dirname(pdf_save_path))

        abs_pdf_path = os.path.abspath(pdf_save_path)
        logger.info(f"pdf_save_path: {pdf_save_path}")

        # 导出为 PDF
        if app_type == "Word":
            doc.SaveAs(abs_pdf_path, FileFormat=WD_FORMAT_PDF)
        elif app_type == "Excel":
            # 0 is xlTypePDF
            doc.ExportAsFixedFormat(XL_TYPE_PDF, abs_pdf_path)
        elif app_type == "PowerPoint":
            doc.SaveAs(abs_pdf_path, FileFormat=PP_SAVE_AS_PDF)

        # 读取 PDF 到内存流
        with open(pdf_save_path, "rb") as pdf_file:
            pdf_content = pdf_file.read()
            pdf_stream.write(pdf_content)
        pdf_stream.seek(0)

    except Exception as e:
        logger.error(f"Error occurred while processing {app_type} document: {e}")
        raise
    finally:
        # 关闭文档，但不退出应用程序
        if doc is not None:
            try:
                if app_type == "Word":
                    doc.Close(SaveChanges=False)
                elif app_type == "Excel":
                    doc.Close(SaveChanges=False)
                elif app_type == "PowerPoint":
                    doc.Close()
            except Exception as e:
                logger.warning(f"Error closing document: {e}")

        # 清理临时文件
        if input_full_path and os.path.exists(input_full_path):
            try:
                os.remove(input_full_path)
                logger.info(f"Deleted input file: {input_full_path}")
            except Exception as e:
                logger.error(f"Failed to delete input file {input_full_path}: {e}")

        if pdf_save_path and os.path.exists(pdf_save_path):
            try:
                os.remove(pdf_save_path)
                logger.info(f"Deleted PDF file: {pdf_save_path}")
            except Exception as e:
                logger.error(f"Failed to delete PDF file {pdf_save_path}: {e}")

    return pdf_stream, pdf_filename


def get_office_application(app_type):
    """获取线程安全的 Office 应用程序实例"""
    try:
        if not hasattr(thread_local, "office_app") or thread_local.office_app is None:
            raise ValueError("Office application not initialized")

        # 检查应用程序是否仍然可用
        thread_local.office_app.Visible

    except (pythoncom.com_error, ValueError):
        pythoncom.CoInitialize()
        if app_type == "Word":
            thread_local.office_app = win32.DispatchEx("Word.Application")
        elif app_type == "Excel":
            thread_local.office_app = win32.DispatchEx("Excel.Application")
        elif app_type == "PowerPoint":
            thread_local.office_app = win32.DispatchEx("PowerPoint.Application")
        else:
            raise ValueError("Unsupported application type")

        thread_local.office_app.Visible = False

    return thread_local.office_app


@app.route("/convert", methods=["POST"])
def upload_file():
    logger.info("Upload file request received.")
    if "file" not in request.files:
        logger.error("No file part in the request.")
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]

    if file.filename == "":
        logger.error("No selected file.")
        return jsonify({"error": "No selected file"}), 400

    if file and allowed_file(file.filename):
        # 统一转小写判断扩展名，修复大小写敏感 Bug
        filename = file.filename
        ext_lower = filename.lower()

        app_type = (
            "Word"
            if ext_lower.endswith((".doc", ".docx"))
            else "Excel"
            if ext_lower.endswith((".xls", ".xlsx"))
            else "PowerPoint"
        )

        try:
            logger.info(f"Converting file: {filename}")
            pdf_stream, pdf_filename = office_to_pdf_stream(file, filename, app_type)

            response = make_response(pdf_stream.getvalue())
            response.headers["Content-Disposition"] = (
                f"attachment; filename*=UTF-8''{quote(pdf_filename)}"
            )
            response.headers["Content-Type"] = "application/pdf"

            logger.info(f"File converted successfully: {pdf_filename}")
            return response

        except Exception as e:
            logger.error(f"Error during conversion: {str(e)}")
            return jsonify({"error": str(e)}), 500

    logger.error("Invalid file format.")
    return jsonify({"error": "Invalid file format"}), 400


def allowed_file(filename):
    """检查文件扩展名是否合法"""
    if not filename or "." not in filename:
        return False
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in ALLOWED_EXTENSIONS


if __name__ == "__main__":
    logger.info("Server starting...")
    try:
        serve(app, listen="*:8081")
    except Exception as e:
        logger.error(str(e))
