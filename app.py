"""Invoice Parser — Flask API and static frontend."""

from __future__ import annotations

import os
from datetime import datetime

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, request, send_from_directory

from exporter.excel_export import build_summary_workbook_bytes, build_workbook_bytes
from parser.pdf_parser import parse_pdf_bytes

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB uploads


@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/api/parse", methods=["POST"])
def api_parse():
    """Accept multipart PDF uploads; return parsed structures per file."""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    results: list[dict] = []
    for f in files:
        if not f.filename:
            continue
        if not f.filename.lower().endswith(".pdf"):
            results.append(
                {
                    "ok": False,
                    "filename": f.filename,
                    "invoice": {},
                    "shipments": [],
                    "adjustments": [],
                    "warnings": [f"{f.filename}: Only PDF files are supported."],
                    "detected_type": "",
                }
            )
            continue
        data = f.read()
        parsed = parse_pdf_bytes(data, f.filename)
        results.append(parsed)

    return jsonify({"results": results})


@app.route("/api/export", methods=["POST"])
def api_export():
    """Build Excel from JSON payload of invoices / shipments / adjustments."""
    payload = request.get_json(force=True, silent=True) or {}
    invoices = payload.get("invoices") or []
    shipments = payload.get("shipments") or []
    adjustments = payload.get("adjustments") or []

    if not invoices:
        return jsonify({"error": "No invoice rows to export"}), 400

    try:
        xlsx = build_workbook_bytes(invoices, shipments, adjustments)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    custom_name = payload.get("filename", "").strip()
    if custom_name:
        name = custom_name if custom_name.endswith(".xlsx") else custom_name + ".xlsx"
    else:
        name = f"invoice_detail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return Response(
        xlsx,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{name}"',
            "Content-Length": str(len(xlsx)),
        },
    )


@app.route("/api/export/summary", methods=["POST"])
def api_export_summary():
    """
    Build the Summary Excel (invoice-level, 11 columns matching the UI summary table).
    Accepts optional 'filename' in the JSON body to name the download.
    """
    payload = request.get_json(force=True, silent=True) or {}
    invoices = payload.get("invoices") or []
    custom_name = payload.get("filename", "").strip()

    if not invoices:
        return jsonify({"error": "No invoice rows to export"}), 400

    try:
        xlsx = build_summary_workbook_bytes(invoices)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if custom_name:
        name = custom_name if custom_name.endswith(".xlsx") else custom_name + ".xlsx"
    else:
        name = f"invoice_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return Response(
        xlsx,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{name}"',
            "Content-Length": str(len(xlsx)),
        },
    )


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})



if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))