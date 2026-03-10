"""
Flask web application.

Provides file upload, processing, and PDF download endpoints.
"""

import os
import uuid
import zipfile
import io
import shutil

from flask import Flask, request, jsonify, send_file, send_from_directory

from . import config
from .parser import parse_workbook
from .pdf_export import generate_combined_pdf, generate_separate_pdfs


def create_app():
    app = Flask(__name__, static_folder="static")
    app.config["MAX_CONTENT_LENGTH"] = config.MAX_FILE_SIZE_MB * 1024 * 1024

    # Ensure upload/output directories exist
    upload_dir = os.path.join(app.root_path, "..", config.UPLOAD_FOLDER)
    output_dir = os.path.join(app.root_path, "..", config.OUTPUT_FOLDER)
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    @app.route("/")
    def index():
        return send_from_directory(app.static_folder, "index.html")

    @app.route("/api/process", methods=["POST"])
    def process_files():
        """
        Accept uploaded Excel files, parse them, generate PDF.

        Returns JSON with processing results and download URL.
        """
        if "files" not in request.files:
            return jsonify({"error": "Файлы не загружены"}), 400

        files = request.files.getlist("files")
        if not files or all(f.filename == "" for f in files):
            return jsonify({"error": "Не выбраны файлы"}), 400

        # Generate unique session ID
        session_id = str(uuid.uuid4())[:8]
        session_dir = os.path.join(upload_dir, session_id)
        os.makedirs(session_dir, exist_ok=True)

        all_sheets = []
        file_results = []

        for file in files:
            if not file.filename:
                continue

            # Validate extension
            ext = os.path.splitext(file.filename)[1].lower()
            if ext not in config.ALLOWED_EXTENSIONS:
                file_results.append({
                    "filename": file.filename,
                    "status": "error",
                    "message": f"Неподдерживаемый формат: {ext}. Допустимые: {', '.join(config.ALLOWED_EXTENSIONS)}",
                    "sheets": [],
                })
                continue

            # Save file
            safe_filename = f"{uuid.uuid4().hex[:8]}_{file.filename}"
            filepath = os.path.join(session_dir, safe_filename)
            file.save(filepath)

            # Parse
            try:
                sheets = parse_workbook(filepath, file.filename)
            except Exception as e:
                file_results.append({
                    "filename": file.filename,
                    "status": "error",
                    "message": f"Ошибка при чтении файла: {str(e)}",
                    "sheets": [],
                })
                continue

            sheet_info = []
            for s in sheets:
                info = {
                    "name": s.name,
                    "skipped": s.skipped,
                    "skip_reason": s.skip_reason,
                    "rows": s.num_rows,
                    "cols": s.num_cols,
                    "title": s.title,
                }
                sheet_info.append(info)

            file_results.append({
                "filename": file.filename,
                "status": "ok",
                "message": f"Обработано листов: {len([s for s in sheets if not s.skipped])} из {len(sheets)}",
                "sheets": sheet_info,
            })
            all_sheets.extend(sheets)

        printable = [s for s in all_sheets if not s.skipped]

        if not printable:
            # Clean up
            shutil.rmtree(session_dir, ignore_errors=True)
            return jsonify({
                "error": "Не найдено ни одного листа с расписанием для печати",
                "files": file_results,
            }), 400

        # Generate combined PDF
        try:
            pdf_bytes = generate_combined_pdf(all_sheets)
        except Exception as e:
            shutil.rmtree(session_dir, ignore_errors=True)
            return jsonify({
                "error": f"Ошибка при генерации PDF: {str(e)}",
                "files": file_results,
            }), 500

        # Save PDF
        output_session_dir = os.path.join(output_dir, session_id)
        os.makedirs(output_session_dir, exist_ok=True)
        pdf_path = os.path.join(output_session_dir, config.COMBINED_PDF_FILENAME)
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        # Also generate separate PDFs as a zip
        try:
            separate = generate_separate_pdfs(all_sheets)
            if len(separate) > 1:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, fbytes in separate:
                        zf.writestr(fname, fbytes)
                zip_buffer.seek(0)
                zip_path = os.path.join(output_session_dir, "расписание_по_классам.zip")
                with open(zip_path, "wb") as f:
                    f.write(zip_buffer.read())
        except Exception:
            pass  # Separate PDFs are optional — don't fail if they error

        # Clean up uploads
        shutil.rmtree(session_dir, ignore_errors=True)

        return jsonify({
            "status": "ok",
            "session_id": session_id,
            "total_pages": len(printable),
            "files": file_results,
            "downloads": {
                "combined_pdf": f"/api/download/{session_id}/combined",
                "separate_zip": f"/api/download/{session_id}/separate",
            },
        })

    @app.route("/api/download/<session_id>/combined")
    def download_combined(session_id):
        """Download the combined PDF."""
        output_session_dir = os.path.join(output_dir, session_id)
        pdf_path = os.path.join(output_session_dir, config.COMBINED_PDF_FILENAME)

        if not os.path.exists(pdf_path):
            return jsonify({"error": "Файл не найден. Возможно, сессия истекла."}), 404

        return send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=config.COMBINED_PDF_FILENAME,
        )

    @app.route("/api/download/<session_id>/separate")
    def download_separate(session_id):
        """Download separate PDFs as a ZIP archive."""
        output_session_dir = os.path.join(output_dir, session_id)
        zip_path = os.path.join(output_session_dir, "расписание_по_классам.zip")

        if not os.path.exists(zip_path):
            return jsonify({"error": "Файл не найден"}), 404

        return send_file(
            zip_path,
            mimetype="application/zip",
            as_attachment=True,
            download_name="расписание_по_классам.zip",
        )

    @app.route("/api/preview/<session_id>")
    def preview(session_id):
        """Return the combined PDF for inline preview."""
        output_session_dir = os.path.join(output_dir, session_id)
        pdf_path = os.path.join(output_session_dir, config.COMBINED_PDF_FILENAME)

        if not os.path.exists(pdf_path):
            return jsonify({"error": "Файл не найден"}), 404

        return send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=False,
        )

    return app
