import os
import io
import logging
from logging.handlers import TimedRotatingFileHandler
import pythoncom
import win32com.client as win32
from werkzeug.utils import secure_filename
from flask import Flask, send_file, request, jsonify, make_response
import tempfile


app = Flask(__name__)
app.name = "office2pdf"

# 定义允许上传的文件类型
ALLOWED_EXTENSIONS = {"doc", "docx", "xls", "xlsx", "ppt", "pptx"}


# 创建每日滚动日志记录器
def create_logger():
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


# 创建日志记录器
logger = create_logger()


# 调用本地 Office 应用将文件转换为 PDF 内存流
def office_to_pdf_stream(input_stream, filename, app_type):
    pythoncom.CoInitialize()  # 初始化 COM 库
    pdf_stream = io.BytesIO()

    logger.info(os.path.splitext(filename)[1])
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=os.path.splitext(filename)[1]
        ) as temp_input:
            temp_input.write(input_stream.read())
            temp_input_path = temp_input.name

        logger.info(temp_input_path)

        input_stream.seek(0)  # 重置输入流
        logger.info(f"app_type: {app_type}")

        if app_type == "Word":
            app = win32.Dispatch("Word.Application")
            app.Visible = False
            doc = app.Documents.Open(temp_input_path)

        elif app_type == "Excel":
            app = win32.Dispatch("Excel.Application")
            app.Visible = False
            doc = app.Workbooks.Open(temp_input_path)

        elif app_type == "PowerPoint":
            app = win32.Dispatch("PowerPoint.Application")
            app.Visible = False
            doc = app.Presentations.Open(temp_input_path)

        else:
            raise ValueError("Unsupported application type")

        if doc is None:
            logger.error(f"Failed to open document for {app_type}: {filename}")
            raise Exception(f"Failed to open document for {app_type}: {filename}")

        # 创建临时PDF文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name

        # 导出为 PDF，FileFormat = 17 代表 PDF
        doc.SaveAs(temp_pdf_path, FileFormat=17)
        logger.info(f"Successfully saved PDF in memory")

        logger.info(f"temp_pdf_path: {temp_pdf_path}")
        # 读取PDF到内存流
        with open(temp_pdf_path, "rb") as pdf_file:
            pdf_stream.write(pdf_file.read())
        pdf_stream.seek(0)

    except Exception as e:
        logger.error(f"Error occurred while processing {app_type} document: {e}")
        raise
    finally:
        # 确保关闭文档和应用程序
        if "doc" in locals():
            doc.Close()
        if "app" in locals():
            app.Quit()
        # 清理临时文件
        if "temp_input_path" in locals() and os.path.exists(temp_input_path):
            os.unlink(temp_input_path)
        if "temp_pdf_path" in locals() and os.path.exists(temp_pdf_path):
            os.unlink(temp_pdf_path)

    return pdf_stream


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
        filename = secure_filename(file.filename)

        app_type = (
            "Word"
            if filename.endswith(("doc", "docx"))
            else "Excel" if filename.endswith(("xls", "xlsx")) else "PowerPoint"
        )

        try:
            logger.info(f"Converting file: {filename}")
            pdf_stream = office_to_pdf_stream(file, filename, app_type)

            pdf_filename = filename.rsplit(".", 1)[0] + ".pdf"
            logger.info(f"pdf_filename: {pdf_filename}")
            # 创建响应
            response = make_response(pdf_stream.getvalue())
            response.headers["Content-Disposition"] = (
                f"attachment; filename={pdf_filename}"
            )
            response.headers["Content-Type"] = "application/pdf"

            logger.info(f"File converted successfully: {pdf_filename}")
            return response

        except Exception as e:
            logger.error(f"Error during conversion: {str(e)}")
            return jsonify({"error": str(e)}), 500

    logger.error("Invalid file format.")
    return jsonify({"error": "Invalid file format"}), 400


# 检查文件扩展名是否合法
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def main():
    app.run(port=int(os.environ.get("PORT", 80)))


if __name__ == "__main__":
    main()
