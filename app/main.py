"""
Flask application — file upload, processing, PDF download.
"""

import os
import uuid
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template

from .parser import parse_excel, ParseResult
from .renderer import render_to_pdf_merged

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

UPLOAD_DIR = Path(__file__).parent.parent / "uploads"
OUTPUT_DIR = Path(__file__).parent.parent / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}


def _allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload_and_process():
    """
    Upload one or more Excel files, parse them, generate PDF.
    Returns JSON with processing report and download URL.
    """
    if "files" not in request.files:
        return jsonify({"error": "Файлы не найдены в запросе"}), 400

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "Файлы не выбраны"}), 400

    session_id = uuid.uuid4().hex[:12]
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(exist_ok=True)

    parse_results: list[ParseResult] = []
    report = {"files": [], "processed_sheets": [], "skipped_sheets": [], "errors": []}

    for f in files:
        if not f.filename:
            continue
        if not _allowed_file(f.filename):
            report["errors"].append(
                f"Файл «{f.filename}» — неподдерживаемый формат. Требуется .xlsx"
            )
            continue

        safe_name = f"{uuid.uuid4().hex[:8]}_{f.filename}"
        fpath = session_dir / safe_name
        f.save(str(fpath))

        pr = parse_excel(fpath)
        parse_results.append(pr)

        report["files"].append(f.filename)
        for s in pr.sheets:
            report["processed_sheets"].append(
                {"file": f.filename, "sheet": s.name, "title": s.title}
            )
        for name, reason in pr.skipped:
            report["skipped_sheets"].append(
                {"file": f.filename, "sheet": name, "reason": reason}
            )
        report["errors"].extend(pr.errors)

    total_sheets = sum(len(pr.sheets) for pr in parse_results)
    if total_sheets == 0:
        return jsonify({
            "error": "Не найдено листов с расписанием для обработки",
            "report": report,
        }), 400

    # Generate PDF
    try:
        pdf_bytes = render_to_pdf_merged(parse_results)
    except Exception as e:
        return jsonify({"error": f"Ошибка генерации PDF: {e}", "report": report}), 500

    pdf_path = OUTPUT_DIR / f"{session_id}.pdf"
    pdf_path.write_bytes(pdf_bytes)

    report["pdf_url"] = f"/api/download/{session_id}"
    report["total_pages"] = total_sheets

    return jsonify({"success": True, "report": report})


@app.route("/api/download/<session_id>")
def download(session_id: str):
    """Download generated PDF."""
    pdf_path = OUTPUT_DIR / f"{session_id}.pdf"
    if not pdf_path.exists():
        return jsonify({"error": "Файл не найден"}), 404
    return send_file(
        str(pdf_path),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="Расписание.pdf",
    )
