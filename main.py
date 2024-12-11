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

app = Flask(__name__)
app.name = "office2pdf"

# 定义允许上传的文件类型
ALLOWED_EXTENSIONS = {"doc", "docx", "xls", "xlsx", "ppt", "pptx"}

# 创建线程本地存储，确保每个线程有独立的 COM 环境
thread_local = threading.local()


def ensure_directory(directory):
    """确保目录存在"""
    if not os.path.exists(directory):
        try:
            os.makedirs(directory)
        except Exception as e:
            print(f"Error creating directory {directory}: {e}")


def create_logger():
    """创建每日滚动日志记录器"""
    ensure_directory("logs")

    logger = logging.getLogger(app.name)
    logger.setLevel(logging.DEBUG)

    log_handler = TimedRotatingFileHandler(
        os.path.join("logs", f"{app.name}_log"),
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


def get_safe_filename(filename):
    """
    生成安全的文件名，使用 UUID 生成唯一标识
    """
    # 生成唯一文件名
    return f"{uuid.uuid4()}-{filename}"


def save_file_with_date_path(file_path, file_content, base_dir):
    """
    按年月日保存文件
    :param file_path: 保存的文件路径
    :param file_content: 文件内容
    :param base_dir: 基础目录
    :return: 保存后的完整路径
    """
    # 获取当前日期
    current_date = datetime.now()
    date_path = os.path.join(
        base_dir,
        current_date.strftime("%Y"),
        current_date.strftime("%m"),
        current_date.strftime("%d"),
    )

    # 确保目录存在
    ensure_directory(date_path)

    # 保存文件
    with open(file_path, "wb") as f:
        f.write(file_content)

    return file_path


def office_to_pdf_stream(input_stream, filename, app_type):
    """将 Office 文档转换为 PDF 内存流"""
    # 确保 uploads 目录存在
    ensure_directory("uploads")

    pdf_stream = io.BytesIO()
    doc = None

    try:
        # 生成安全的文件名
        input_filename = get_safe_filename(filename)

        # 保存原始上传文件
        input_save_path = os.path.join(
            "uploads",
            datetime.now().strftime("%Y"),
            datetime.now().strftime("%m"),
            datetime.now().strftime("%d"),
            input_filename,
        )
        save_file_with_date_path(input_save_path, input_stream.read(), "uploads")
        input_full_path = os.path.abspath(input_save_path)
        logger.info(f"doc_save_path: {input_full_path}")

        # 获取应用程序实例
        app = get_office_application(app_type)

        # 打开文档
        if app_type == "Word":
            doc = app.Documents.Open(input_full_path)
        elif app_type == "Excel":
            doc = app.Workbooks.Open(input_full_path)
        elif app_type == "PowerPoint":
            doc = app.Presentations.Open(input_full_path)

        if doc is None:
            logger.error(f"Failed to open document for {app_type}: {filename}")
            raise Exception(f"Failed to open document for {app_type}: {filename}")

        # 生成安全的 PDF 文件名
        pdf_filename = os.path.splitext(filename)[0] + ".pdf"
        save_pdf_filename = get_safe_filename(pdf_filename)

        # 保存 PDF 文件
        pdf_save_path = os.path.join(
            "uploads",
            datetime.now().strftime("%Y"),
            datetime.now().strftime("%m"),
            datetime.now().strftime("%d"),
            save_pdf_filename,
        )
        logger.info(f"pdf_save_path: {pdf_save_path}")

        # 导出为 PDF
        doc.SaveAs(os.path.abspath(pdf_save_path), FileFormat=17)

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
            doc.Close()

    return pdf_stream, pdf_filename


def get_office_application(app_type):
    """获取线程安全的 Office 应用程序实例"""
    if not hasattr(thread_local, "office_app"):
        pythoncom.CoInitialize()
        if app_type == "Word":
            thread_local.office_app = win32.Dispatch("Word.Application")
        elif app_type == "Excel":
            thread_local.office_app = win32.Dispatch("Excel.Application")
        elif app_type == "PowerPoint":
            thread_local.office_app = win32.Dispatch("PowerPoint.Application")
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
        # 保留原始文件名
        filename = file.filename

        app_type = (
            "Word"
            if filename.endswith(("doc", "docx"))
            else "Excel" if filename.endswith(("xls", "xlsx")) else "PowerPoint"
        )

        try:
            logger.info(f"Converting file: {filename}")
            pdf_stream, pdf_filename = office_to_pdf_stream(file, filename, app_type)

            response = make_response(pdf_stream.getvalue())
            response.headers["Content-Disposition"] = (
                f"attachment; filename*=UTF-8''{quote(pdf_filename)}; "
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
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


if __name__ == "__main__":
    logger = create_logger()
    logger.info("Server starting...")
    try:
        serve(app, listen="*:8081")
    except Exception as e:
        logger.error(str(e))
