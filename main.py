import os
import logging
import pythoncom
import win32com.client as win32
from werkzeug.utils import secure_filename
from flask import Flask, send_file, request, jsonify


app = Flask(__name__)

# 设置上传文件保存的路径
UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.name = "office2pdf"

# 定义允许上传的文件类型
ALLOWED_EXTENSIONS = {"doc", "docx", "xls", "xlsx", "ppt", "pptx"}

# 创建全局日志记录器
logger = logging.getLogger(app.name)
handler = logging.FileHandler(".\\log\\logfile.log")
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


# 调用本地 Office 应用将文件转换为 PDF
def office_to_pdf(input_file, output_file, app_type):
    pythoncom.CoInitialize()  # 初始化 COM 库
    logger.info(f"app_type: {app_type}")

    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        logger.error(f"Input file does not exist: {input_file}")
        raise FileNotFoundError(f"Input file does not exist: {input_file}")

    try:
        if app_type == "Word":
            app = win32.Dispatch("Word.Application")
            app.Visible = False  # 设置为 False 以避免 UI 错误
            doc = app.Documents.Open(input_file)

        elif app_type == "Excel":
            app = win32.Dispatch("Excel.Application")
            app.Visible = False
            doc = app.Workbooks.Open(input_file)

        elif app_type == "PowerPoint":
            app = win32.Dispatch("PowerPoint.Application")
            app.Visible = False
            doc = app.Presentations.Open(input_file)

        else:
            raise ValueError("Unsupported application type")

        if doc is None:
            logger.error(f"Failed to open document for {app_type}: {input_file}")
            raise Exception(f"Failed to open document for {app_type}: {input_file}")

        logger.info(f"Successfully opened {app_type} document: {input_file}")

        # 导出为 PDF，FileFormat = 17 代表 PDF
        doc.SaveAs(output_file, FileFormat=17)
        logger.info(f"Successfully saved PDF: {output_file}")

    except Exception as e:
        logger.error(f"Error occurred while processing {app_type} document: {e}")
        raise
    finally:
        # 确保关闭文档和应用程序
        if "doc" in locals():
            doc.Close(False)
        if "app" in locals():
            app.Quit()


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
        input_filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        input_filepath = os.path.abspath(input_filepath)
        file.save(input_filepath)

        pdf_filename = filename.rsplit(".", 1)[0] + ".pdf"
        output_filepath = os.path.join(app.config["UPLOAD_FOLDER"], pdf_filename)
        output_filepath = os.path.abspath(output_filepath)

        app_type = (
            "Word"
            if filename.endswith(("doc", "docx"))
            else "Excel" if filename.endswith(("xls", "xlsx")) else "PowerPoint"
        )
        try:
            logger.info(
                f"Converting {input_filepath} to {output_filepath} as {app_type}."
            )
            office_to_pdf(input_filepath, output_filepath, app_type)
        except Exception as e:
            logger.error(e)
            logger.error(f"Error during conversion: {str(e)}")
            return jsonify({"error": str(e)}), 500

        logger.info(f"File converted successfully. Sending file: {output_filepath}")
        return send_file(output_filepath, as_attachment=True)

    logger.error("Invalid file format.")
    return jsonify({"error": "Invalid file format"}), 400


# 检查文件扩展名是否合法
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def main():
    app.run(port=int(os.environ.get("PORT", 80)))


if __name__ == "__main__":
    main()
