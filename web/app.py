"""
Flask backend for the Letter Exporter.
Serves a single-page UI and exposes an API endpoint to generate Word documents.
"""
import io
import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, request, jsonify, send_file, send_from_directory

# Load .env from project root
PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

# Add tools/ to path so we can import the extraction module
sys.path.insert(0, str(PROJECT_ROOT / "tools"))
import letter_extract as le

app = Flask(__name__, static_folder="static")

# Matches engagement URLs - document ID is optional
CW_URL_PATTERN = re.compile(
    r"https?://([^/]+)/([^/]+)/e/eng/([^/]+)"
)
CW_DOC_PATTERN = re.compile(
    r"#/(?:efinancials|letter)/([^/?\s]+)"
)


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()

    if not url:
        return jsonify({"error": "URL is required."}), 400

    match = CW_URL_PATTERN.search(url)
    if not match:
        return jsonify({
            "error": "Invalid Caseware URL. Expected format: "
                     "https://<host>/<tenant>/e/eng/<engagementId>/..."
        }), 400

    host_name = match.group(1)
    tenant = match.group(2)
    engagement_id = match.group(3)
    template_name = (data.get("templateName") or "Letter").strip()

    # Document ID is required for letters
    doc_match = CW_DOC_PATTERN.search(url)
    if not doc_match:
        return jsonify({
            "error": "URL must include a document fragment "
                     "(e.g. #/letter/<documentId>)"
        }), 400
    document_id = doc_match.group(1)

    host = f"https://{host_name}"

    try:
        docx_bytes = le.generate_report_bytes(
            engagement_id=engagement_id,
            document_id=document_id,
            host=host,
            tenant=tenant,
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 502
    except Exception as e:
        return jsonify({"error": f"Unexpected error: {e}"}), 500

    safe_name = re.sub(r'[^\w\s-]', '', template_name).strip().replace(' ', '_')
    filename = f"{safe_name}_letter.docx" if safe_name else f"letter_{engagement_id[:8]}.docx"

    return send_file(
        io.BytesIO(docx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "1") == "1", port=5001)
