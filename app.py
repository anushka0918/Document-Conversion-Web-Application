"""
Flask web app for PDF to DOCX conversion
Uses PDFToDocxConverter from pipeline.py with security features
"""

from flask import Flask, render_template, request, send_file, jsonify
import os
from pathlib import Path
from werkzeug.utils import secure_filename
import logging

from generalised_converter import PDFToDocxConverter

# ------------------------------------------------------------------
# Flask setup
# ------------------------------------------------------------------
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {"pdf"}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

 
converter = PDFToDocxConverter(
    preserve_images=True,
    preserve_tables=True,
    preserve_fonts=True,
    multi_processing=False,  # Flask handles single requests
    max_workers=1,
)


# ------------------------------------------------------------------
# Helper functions
# ------------------------------------------------------------------
def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed"""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------
@app.route("/")
def index():
     
    try:
        return render_template("index.html")
    except:
        # Fallback if no template exists
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <title>PDF to DOCX Converter</title>
            <style>
                body { font-family: Arial, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; }
                h1 { color: #333; }
                form { background: #f4f4f4; padding: 20px; border-radius: 8px; }
                input[type="file"] { margin: 10px 0; }
                button { background: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
                button:hover { background: #0056b3; }
            </style>
        </head>
        <body>
            <h1>🚀 PDF to DOCX Converter</h1>
            <form action="/convert" method="post" enctype="multipart/form-data">
                <p>Select a PDF file to convert:</p>
                <input type="file" name="pdf_file" accept=".pdf" required>
                <br><br>
                <button type="submit">Convert to DOCX</button>
            </form>
            <p><small>Maximum file size: 50 MB</small></p>
        </body>
        </html>
        """


@app.route("/convert", methods=["POST"])
def convert():
    """Handle PDF upload and conversion"""

    # Validate request
    if "pdf_file" not in request.files:
        logger.warning("No file part in request")
        return jsonify({"success": False, "message": "No file uploaded"}), 400

    file = request.files["pdf_file"]

    if file.filename == "":
        logger.warning("Empty filename")
        return jsonify({"success": False, "message": "No file selected"}), 400

    # Validate file type
    if not allowed_file(file.filename):
        logger.warning(f"Invalid file type: {file.filename}")
        return (
            jsonify(
                {
                    "success": False,
                    "message": "Invalid file type. Only PDF files are allowed.",
                }
            ),
            400,
        )

    try:
        # Secure filename (prevents directory traversal attacks)
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)

        # Save uploaded file
        file.save(pdf_path)
        logger.info(f"File uploaded: {filename}")

        # Generate output path
        output_name = Path(filename).with_suffix(".docx").name
        output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_name)

        # Convert PDF to DOCX
        logger.info(f"Converting: {filename}")
        success, result = converter.convert_single(
            pdf_path, output_path, verbose=False
        )

        if not success:
            logger.error(f"Conversion failed: {result}")
            return jsonify({"success": False, "message": result}), 500

        logger.info(f"Conversion successful: {output_name}")

        # Send converted file
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        logger.error(f"Server error: {e}", exc_info=True)
        return (
            jsonify({"success": False, "message": f"Server error: {str(e)}"}),
            500,
        )


@app.route("/health")
def health():
    """Health check endpoint"""
    return jsonify({"status": "healthy", "service": "PDF to DOCX Converter"}), 200


# ------------------------------------------------------------------
# Error handlers
# ------------------------------------------------------------------
@app.errorhandler(413)
def file_too_large(e):
    """Handle file size limit exceeded"""
    logger.warning("File size limit exceeded")
    return (
        jsonify(
            {
                "success": False,
                "message": "File too large. Maximum size is 50 MB.",
            }
        ),
        413,
    )


@app.errorhandler(404)
def not_found(e):
    """Handle 404 errors"""
    return jsonify({"error": "Endpoint not found"}), 404


@app.errorhandler(500)
def internal_error(e):
    """Handle 500 errors"""
    logger.error(f"Internal server error: {e}")
    return jsonify({"error": "Internal server error"}), 500


# ------------------------------------------------------------------
# Main
# ------------------------------------------------------------------
if __name__ == "__main__":
    logger.info("=" * 70)
    logger.info("🚀 PDF to DOCX Converter - Flask Web App")
    logger.info("=" * 70)
    logger.info(f"📁 Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    logger.info(f"📁 Output folder: {os.path.abspath(OUTPUT_FOLDER)}")
    logger.info(f"📏 Max file size: {MAX_CONTENT_LENGTH / (1024*1024):.0f} MB")
    logger.info("=" * 70)

     
    app.run(debug=True, host="0.0.0.0", port=5000)


 