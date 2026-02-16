"""
ATSReady — Flask API
====================
Serves two endpoints:
  POST /convert  → Returns the ATS-optimised .docx file
  POST /score    → Returns JSON with ATS score and details
  GET  /health   → Health check

Both accept multipart form data with a 'resume' PDF file.
Secured with a simple API key header: X-API-Key

Run locally for testing:
    pip install flask pdfplumber python-docx
    python -m flask --app app run

Then open your site with ?test=true:
    index.html?test=true

The test mode will POST directly to http://127.0.0.1:5000
"""

import os
import tempfile
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # Required for test mode (browser → localhost)
from ats_converter import convert_resume, extract_text_from_pdf, detect_sections, score_ats_readability

app = Flask(__name__)

# Enable CORS so the frontend (served from file:// or Netlify) can
# call the local API during testing. In production you'd restrict this.
CORS(app)

# Simple API key for security (change this to a random string)
API_KEY = os.environ.get("ATS_API_KEY", "your-secret-key-change-this")


def check_api_key():
    """Verify the API key from request headers."""
    if request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401
    return None


def get_uploaded_file():
    """Validate and return the uploaded resume file."""
    if "resume" not in request.files:
        return None, (jsonify({"error": "No file uploaded"}), 400)

    file = request.files["resume"]
    if not file.filename.lower().endswith(".pdf"):
        return None, (jsonify({"error": "Only PDF files accepted"}), 400)

    return file, None


@app.route("/convert", methods=["POST"])
def convert():
    """Convert a PDF resume to an ATS-optimised .docx file."""
    auth_error = check_api_key()
    if auth_error:
        return auth_error

    file, file_error = get_uploaded_file()
    if file_error:
        return file_error

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.pdf")
        output_path = os.path.join(tmpdir, "ats_optimised.docx")

        file.save(input_path)
        result = convert_resume(input_path, output_path)

        if result["success"]:
            return send_file(
                output_path,
                as_attachment=True,
                download_name="ats_optimised_resume.docx",
                mimetype=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml.document"
                ),
            )
        else:
            return jsonify(result), 422


@app.route("/score", methods=["POST"])
def score():
    """
    Score a PDF resume for ATS readability.
    Returns JSON with score, issues, passed checks, word count,
    and sections detected — without generating the .docx.
    """
    auth_error = check_api_key()
    if auth_error:
        return auth_error

    file, file_error = get_uploaded_file()
    if file_error:
        return file_error

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.pdf")
        file.save(input_path)

        # Extract and analyse
        raw_text = extract_text_from_pdf(input_path)

        if not raw_text.strip():
            return jsonify({
                "success": False,
                "score": 0,
                "error": (
                    "Could not extract any text from this PDF. "
                    "It may be an image-only PDF."
                ),
            }), 422

        sections = detect_sections(raw_text)
        score_data = score_ats_readability(raw_text, sections)

        return jsonify({
            "success": True,
            "score": score_data["score"],
            "issues": score_data["issues"],
            "passed": score_data["passed"],
            "word_count": score_data["word_count"],
            "sections_detected": len(
                [s for s in sections if s["heading"] != "HEADER"]
            ),
        })


@app.route("/health", methods=["GET"])
def health():
    """Simple health check."""
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(debug=True)
