"""Invoice Parser — Flask API and static frontend."""

from __future__ import annotations

import os
from datetime import datetime

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, request, send_from_directory

from exporter.excel_export import build_workbook_bytes
from parser.ups_parser import UPSParseError, parse_invoice

load_dotenv()

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/api/parse", methods=["POST"])
def api_parse():
    """Accept multipart PDF uploads; return parsed invoice dicts per file."""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    results: list[dict] = []
    for f in files:
        if not f.filename:
            continue
        if not f.filename.lower().endswith(".pdf"):
            results.append({
                "ok": False,
                "filename": f.filename,
                "invoice": {},
                "warnings": [f"{f.filename}: Only PDF files are supported."],
                "detected_type": "",
            })
            continue
        data = f.read()
        try:
            inv = parse_invoice(data, filename=f.filename)
            results.append({
                "ok": True,
                "filename": f.filename,
                "invoice": inv,
                "warnings": [],
                "detected_type": inv.get("invoice_type_label", ""),
            })
        except UPSParseError as e:
            results.append({
                "ok": False,
                "filename": f.filename,
                "invoice": {},
                "warnings": [str(e)],
                "detected_type": "",
            })
        except Exception as e:
            results.append({
                "ok": False,
                "filename": f.filename,
                "invoice": {},
                "warnings": [f"{f.filename}: Unexpected error — {e}"],
                "detected_type": "",
            })

    return jsonify({"results": results})


@app.route("/api/export/summary", methods=["POST"])
def api_export_summary():
    """Build the two-table Excel workbook from a JSON array of invoice dicts."""
    payload     = request.get_json(force=True, silent=True) or {}
    invoices    = payload.get("invoices") or []
    custom_name = payload.get("filename", "").strip()

    if not invoices:
        return jsonify({"error": "No invoice rows to export"}), 400

    try:
        xlsx = build_workbook_bytes(invoices)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if custom_name:
        name = custom_name if custom_name.endswith(".xlsx") else custom_name + ".xlsx"
    else:
        name = f"invoice_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return Response(
        xlsx,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
