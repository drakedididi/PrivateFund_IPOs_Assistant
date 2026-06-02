from __future__ import annotations

import io
import hmac
import os
import shutil
import tempfile
import time
import zipfile
from pathlib import Path
from urllib.parse import quote

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

from module_tools.big_redemption.redemption_word import process_excel_files as process_redemption_files
from module_tools.extra_revenue.extra_revenue import analyze_extra_revenue_excel
from module_tools.invoice.invoice_organization_local import process_zip_file
from module_tools.inquiry_video.video import process_excel_files as process_inquiry_video_files
from module_tools.related_deal.convert_openyxl import process_excel_files as process_related_decision_files
from module_tools.related_deal.multi_fund import process_excel_files as process_related_notice_files
from module_tools.valuation_table.valuation_table import process_excel_files as process_valuation_table_files


RENDER_SERVICE_URL = "https://privatefund-ipos-assistant-km21.onrender.com"
APP_VERSION = "valuation_table"
EXPOSED_HEADERS = [
    "Content-Disposition",
    "X-App-Version",
    "X-Extra-Revenue-Frequency",
    "X-Max-Recovery-Period",
    "X-Recovery-Periods",
]

app = Flask(__name__)
CORS(app, expose_headers=EXPOSED_HEADERS)


def debug_log(message: str) -> None:
    print(f"[valuation_table] {message}", flush=True)


@app.after_request
def add_app_headers(response):
    response.headers["X-App-Version"] = APP_VERSION
    exposed_headers = [
        header.strip()
        for header in response.headers.get("Access-Control-Expose-Headers", "").split(",")
        if header.strip()
    ]
    for header in EXPOSED_HEADERS:
        if header not in exposed_headers:
            exposed_headers.append(header)
    response.headers["Access-Control-Expose-Headers"] = ", ".join(exposed_headers)
    return response


@app.route("/", methods=["GET"])
@app.route("/api/health", methods=["GET"])
def health_check():
    return jsonify({"status": "ok", "service_url": RENDER_SERVICE_URL, "version": APP_VERSION}), 200


def get_render_port() -> int:
    port = os.environ.get("PORT")
    if not port:
        raise RuntimeError(
            "PORT environment variable is required. "
            "On Render, use: gunicorn app:app --bind 0.0.0.0:$PORT"
        )
    return int(port)


def developing():
    return "功能开发中", 200


def get_request_secret_token() -> str:
    auth_header = request.headers.get("Authorization", "")
    if auth_header.lower().startswith("bearer "):
        return auth_header[7:].strip()
    return request.headers.get("X-App-Secret-Token", "").strip()


def require_secret_token():
    expected_token = os.environ.get("APP_SECRET_TOKEN", "").strip()
    if not expected_token:
        return jsonify({"error": "服务端未配置 APP_SECRET_TOKEN"}), 503

    supplied_token = get_request_secret_token()
    if not supplied_token or not hmac.compare_digest(supplied_token, expected_token):
        return jsonify({"error": "APP_SECRET_TOKEN 缺失或无效"}), 401

    return None


def safe_upload_filename(original_name: str, fallback_stem: str) -> tuple[str, str]:
    original_basename = Path(original_name.replace("\\", "/")).name
    original_path = Path(original_basename)
    suffix = original_path.suffix.lower()
    safe_stem = "".join(
        ch for ch in original_path.stem
        if ch.isalnum() or ch in (" ", "-", "_")
    ).strip()
    safe_stem = safe_stem or secure_filename(original_path.stem) or fallback_stem
    return f"{safe_stem}{suffix}", suffix


def decode_zip_member_name(name: str) -> str:
    try:
        repaired = name.encode("cp437").decode("gbk")
    except Exception:
        return name
    old_cjk = sum(1 for ch in name if "\u4e00" <= ch <= "\u9fff")
    new_cjk = sum(1 for ch in repaired if "\u4e00" <= ch <= "\u9fff")
    return repaired if new_cjk > old_cjk else name


def safe_target_path(base_dir: Path, member_name: str) -> Path:
    member_path = Path(member_name)
    target = (base_dir / member_path).resolve()
    base = base_dir.resolve()
    if base != target and base not in target.parents:
        raise ValueError(f"非法路径: {member_name}")
    return target


def extract_uploaded_zip(input_zip: Path, extract_dir: Path) -> int:
    count = 0
    with zipfile.ZipFile(input_zip) as zf:
        for info in zf.infolist():
            fixed_name = decode_zip_member_name(info.filename)
            if not fixed_name or fixed_name.endswith("/"):
                continue
            target_path = safe_target_path(extract_dir, fixed_name)
            target_path.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src, open(target_path, "wb") as dst:
                shutil.copyfileobj(src, dst)
            count += 1
    return count


def prepare_uploaded_excel_input(uploaded_file, input_dir: Path) -> str:
    if uploaded_file is None or not uploaded_file.filename:
        raise ValueError("请上传文件")

    filename, suffix = safe_upload_filename(uploaded_file.filename, "input")
    if suffix not in {".zip", ".xlsx", ".xls"}:
        raise ValueError("仅支持 ZIP、XLSX 或 XLS 文件")

    if suffix == ".zip":
        input_dir.mkdir(parents=True, exist_ok=True)
        saved_path = input_dir / filename
        uploaded_file.save(str(saved_path))
        extract_dir = input_dir / "_unzipped"
        extract_dir.mkdir(parents=True, exist_ok=True)
        extract_uploaded_zip(saved_path, extract_dir)
        saved_path.unlink(missing_ok=True)
        return Path(filename).stem

    file_dir = input_dir / Path(filename).stem
    file_dir.mkdir(parents=True, exist_ok=True)
    saved_path = file_dir / filename
    uploaded_file.save(str(saved_path))
    return Path(filename).stem


def prepare_uploaded_pdf_input(uploaded_file, input_dir: Path) -> str:
    if uploaded_file is None or not uploaded_file.filename:
        raise ValueError("请上传文件")

    filename, suffix = safe_upload_filename(uploaded_file.filename, "valuation_input")
    debug_log(f"upload received: filename={filename}, suffix={suffix}")
    if suffix not in {".zip", ".pdf"}:
        raise ValueError("仅支持 ZIP 或 PDF 文件")

    if suffix == ".zip":
        input_dir.mkdir(parents=True, exist_ok=True)
        saved_path = input_dir / filename
        uploaded_file.save(str(saved_path))
        debug_log(f"zip saved: path={saved_path}, size={saved_path.stat().st_size}")
        extract_dir = input_dir / "_unzipped"
        extract_dir.mkdir(parents=True, exist_ok=True)
        extracted_count = extract_uploaded_zip(saved_path, extract_dir)
        pdf_count = sum(1 for p in extract_dir.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
        debug_log(f"zip extracted: files={extracted_count}, pdf_count={pdf_count}, dir={extract_dir}")
        saved_path.unlink(missing_ok=True)
        return Path(filename).stem

    file_dir = input_dir / Path(filename).stem
    file_dir.mkdir(parents=True, exist_ok=True)
    saved_path = file_dir / filename
    uploaded_file.save(str(saved_path))
    debug_log(f"pdf saved: path={saved_path}, size={saved_path.stat().st_size}")
    return Path(filename).stem


def zip_output_dir(output_dir: Path, output_zip: Path) -> None:
    files = [p for p in output_dir.rglob("*") if p.is_file()]
    if not files:
        raise RuntimeError("未生成任何结果文件")

    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in files:
            zf.write(file_path, file_path.relative_to(output_dir))


def run_document_tool(processor, download_suffix: str):
    auth_error = require_secret_token()
    if auth_error:
        return auth_error

    uploaded_file = request.files.get("file")
    try:
        with tempfile.TemporaryDirectory(prefix="private_tool_api_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_dir = tmp_path / "input"
            output_dir = tmp_path / "output"
            upload_stem = prepare_uploaded_excel_input(uploaded_file, input_dir)

            processor(input_dir, output_dir)

            output_zip = tmp_path / f"{upload_stem}{download_suffix}"
            zip_output_dir(output_dir, output_zip)
            result_bytes = output_zip.read_bytes()

        return send_file(
            io.BytesIO(result_bytes),
            mimetype="application/zip",
            as_attachment=True,
            download_name=output_zip.name,
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"处理失败: {exc}"}), 500


def run_pdf_document_tool(processor, download_suffix: str):
    auth_error = require_secret_token()
    if auth_error:
        return auth_error

    uploaded_file = request.files.get("file")
    started_at = time.perf_counter()
    debug_log("request started")
    try:
        with tempfile.TemporaryDirectory(prefix="private_tool_api_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_dir = tmp_path / "input"
            output_dir = tmp_path / "output"
            upload_stem = prepare_uploaded_pdf_input(uploaded_file, input_dir)

            process_started_at = time.perf_counter()
            processor(input_dir, output_dir)
            debug_log(f"processor finished: seconds={time.perf_counter() - process_started_at:.2f}")

            output_zip = tmp_path / f"{upload_stem}{download_suffix}"
            zip_output_dir(output_dir, output_zip)
            debug_log(f"output zipped: path={output_zip}, size={output_zip.stat().st_size}")
            result_bytes = output_zip.read_bytes()

        debug_log(f"request finished: seconds={time.perf_counter() - started_at:.2f}")
        return send_file(
            io.BytesIO(result_bytes),
            mimetype="application/zip",
            as_attachment=True,
            download_name=output_zip.name,
        )
    except ValueError as exc:
        debug_log(f"request value error: {exc}")
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        debug_log(f"request failed: {exc}")
        return jsonify({"error": f"处理失败: {exc}"}), 500


def run_extra_revenue_tool():
    uploaded_file = request.files.get("file")
    if uploaded_file is None or not uploaded_file.filename:
        return jsonify({"error": "请上传文件"}), 400

    filename, suffix = safe_upload_filename(uploaded_file.filename, "extra_revenue_input")
    if suffix not in {".xlsx", ".xls"}:
        return jsonify({"error": "仅支持 XLSX 或 XLS 文件"}), 400

    try:
        with tempfile.TemporaryDirectory(prefix="extra_revenue_api_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            input_path = tmp_path / filename
            output_dir = tmp_path / "output"
            uploaded_file.save(str(input_path))

            result = analyze_extra_revenue_excel(input_path, output_dir)
            output_zip = tmp_path / f"{result['download_stem']}.zip"
            zip_output_dir(output_dir, output_zip)
            result_bytes = output_zip.read_bytes()

        response = send_file(
            io.BytesIO(result_bytes),
            mimetype="application/zip",
            as_attachment=True,
            download_name=output_zip.name,
        )
        response.headers["X-Extra-Revenue-Frequency"] = quote(str(result["frequency"]), safe="")
        response.headers["X-Max-Recovery-Period"] = quote(
            str(result["max_recovery_period"]),
            safe="",
        )
        response.headers["X-Recovery-Periods"] = quote(
            str(result["recovery_periods_text"]),
            safe="",
        )
        return response
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"处理失败: {exc}"}), 500


@app.route("/api/invoice", methods=["POST"])
def api_invoice():
    uploaded_file = request.files.get("file")
    if uploaded_file is None or not uploaded_file.filename:
        return jsonify({"error": "请上传文件"}), 400

    filename, suffix = safe_upload_filename(uploaded_file.filename, "invoice_input")
    if suffix != ".zip":
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
    return run_document_tool(process_inquiry_video_files, "_inquiry_video_docs.zip")


@app.route("/api/extra-revenue", methods=["POST"])
def api_extra_revenue():
    return run_extra_revenue_tool()


@app.route("/api/redemption", methods=["POST"])
def api_redemption():
    return run_document_tool(process_redemption_files, "_redemption_docs.zip")


@app.route("/api/valuation-table", methods=["POST"])
def api_valuation_table():
    return run_pdf_document_tool(process_valuation_table_files, "_valuation_table.zip")


@app.route("/api/excel", methods=["POST"])
def api_excel():
    return run_document_tool(process_related_decision_files, "_related_decision_docs.zip")


@app.route("/api/fund", methods=["POST"])
def api_fund():
    return run_document_tool(process_related_notice_files, "_related_notice_docs.zip")


@app.route("/api/pcf/crawl", methods=["POST"])
def api_pcf_crawl():
    return developing()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=get_render_port())
