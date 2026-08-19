"""
ROHI Excel Sync Server
----------------------
Optional small HTTP endpoint for the ROHI Attendance Android app.

Run on the ROHI server:
    pip install flask
    python excel_sync_server.py

Set the Android app's four "upload endpoint" fields to:
    http://SERVER_IP:8080/upload

Files are stored in:
    ./excel_sync/attendance
    ./excel_sync/timesheet
    ./excel_sync/leave
    ./excel_sync/staff

If these folders are inside a Google Drive/OneDrive/SharePoint sync folder,
the exported workbooks will then be copied to that cloud folder automatically.

For production, place this endpoint behind HTTPS and authentication.
"""
import os
from flask import Flask, request, jsonify

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
ROOT = os.path.join(BASE_DIR, "excel_sync")
ALLOWED = {"attendance", "timesheet", "leave", "staff"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


def safe_name(value):
    value = os.path.basename(str(value or "report.xlsx"))
    return "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in value)


@app.post("/upload")
def upload():
    report_type = str(request.form.get("report_type") or "").strip().lower()
    if report_type not in ALLOWED:
        return jsonify(ok=False, error="Invalid report_type"), 400

    uploaded = request.files.get("file")
    if uploaded is None:
        return jsonify(ok=False, error="Missing file"), 400

    folder = os.path.join(ROOT, report_type)
    os.makedirs(folder, exist_ok=True)
    filename = safe_name(request.form.get("filename") or uploaded.filename)
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    destination = os.path.join(folder, filename)
    uploaded.save(destination)
    return jsonify(ok=True, report_type=report_type, filename=filename)


@app.get("/health")
def health():
    return jsonify(ok=True, service="ROHI Excel Sync Server")


if __name__ == "__main__":
    os.makedirs(ROOT, exist_ok=True)
    app.run(host="0.0.0.0", port=8080, debug=False)
