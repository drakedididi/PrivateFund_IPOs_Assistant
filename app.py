from __future__ import annotations

import io
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

from module_tools.invoice.invoice_organization_local import process_zip_file


app = Flask(__name__)
CORS(app)


def developing():
    return "功能开发中", 200


@app.route("/api/invoice", methods=["POST"])
def api_invoice():
    uploaded_file = request.files.get("file")
    if uploaded_file is None or not uploaded_file.filename:
        return jsonify({"error": "请上传文件"}), 400

    filename = secure_filename(uploaded_file.filename) or "invoice_input.zip"
    if not filename.lower().endswith(".zip"):
        return jsonify({"error": "仅支持 ZIP 文件"}), 400

    try:
        with tempfile.TemporaryDirectory(prefix="invoice_api_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_zip = tmp_path / filename
            output_zip = tmp_path / f"{Path(filename).stem}_classified.zip"

            uploaded_file.save(str(input_zip))
            process_zip_file(input_zip, output_zip)
            result_bytes = output_zip.read_bytes()

        return send_file(
            io.BytesIO(result_bytes),
            mimetype="application/zip",
            as_attachment=True,
            download_name=output_zip.name,
        )
    except Exception as exc:
        return jsonify({"error": f"处理失败: {exc}"}), 500


@app.route("/api/video", methods=["POST"])
def api_video():
    return developing()


@app.route("/api/redemption", methods=["POST"])
def api_redemption():
    return developing()


@app.route("/api/excel", methods=["POST"])
def api_excel():
    return developing()


@app.route("/api/fund", methods=["POST"])
def api_fund():
    return developing()


@app.route("/api/pcf/crawl", methods=["POST"])
def api_pcf_crawl():
    return developing()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
