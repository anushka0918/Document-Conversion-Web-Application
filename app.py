"""
🚀 MEMORY-OPTIMIZED FLASK APP FOR PDF TO DOCX
✅ Optimized for Render/Heroku FREE tier (512MB RAM)
✅ File size limits, timeouts, memory monitoring
"""

from flask import Flask, render_template, request, send_file, jsonify, after_this_request
import os
import logging
import gc
import psutil
from pathlib import Path
from werkzeug.utils import secure_filename

# Import optimized converter
from generalised_converter import MemoryOptimizedConverter

# Flask setup
app = Flask(__name__)

# CRITICAL: Smaller file size limit for free tier
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {"pdf"}
MAX_CONTENT_LENGTH = 10 * 1024 * 1024  # 10 MB (reduced for free tier)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

# Logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize memory-optimized converter
converter = MemoryOptimizedConverter(
    max_memory_mb=400,  # Leave 100MB buffer for OS/Flask
    enable_images=True,
    fallback_mode=True  # Auto-switch to text-only if needed
)


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed"""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_memory_usage():
    """Get current memory usage"""
    try:
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / (1024 * 1024)
    except:
        return 0


def cleanup_files(*file_paths):
    """Clean up temporary files"""
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"🗑️ Cleaned up: {file_path}")
        except Exception as e:
            logger.warning(f"⚠️ Could not delete {file_path}: {e}")


@app.route("/")
def index():
    """Landing page"""
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>PDF to DOCX Converter</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body {
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
                display: flex;
                align-items: center;
                justify-content: center;
                padding: 20px;
            }
            .container {
                background: white;
                border-radius: 20px;
                padding: 40px;
                max-width: 500px;
                width: 100%;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            }
            h1 {
                color: #333;
                margin-bottom: 10px;
                font-size: 28px;
            }
            .subtitle {
                color: #666;
                margin-bottom: 30px;
                font-size: 14px;
            }
            .upload-area {
                border: 2px dashed #667eea;
                border-radius: 10px;
                padding: 40px 20px;
                text-align: center;
                background: #f8f9ff;
                cursor: pointer;
                transition: all 0.3s;
            }
            .upload-area:hover {
                background: #f0f2ff;
                border-color: #764ba2;
            }
            .upload-area.dragover {
                background: #e8ebff;
                border-color: #764ba2;
            }
            input[type="file"] { display: none; }
            .upload-icon {
                font-size: 48px;
                margin-bottom: 10px;
            }
            .upload-text {
                color: #667eea;
                font-weight: 600;
                margin-bottom: 5px;
            }
            .upload-hint {
                color: #999;
                font-size: 12px;
            }
            .btn {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 10px;
                font-size: 16px;
                font-weight: 600;
                cursor: pointer;
                width: 100%;
                margin-top: 20px;
                transition: transform 0.2s;
            }
            .btn:hover {
                transform: translateY(-2px);
            }
            .btn:disabled {
                opacity: 0.6;
                cursor: not-allowed;
            }
            .status {
                margin-top: 20px;
                padding: 15px;
                border-radius: 10px;
                display: none;
            }
            .status.success {
                background: #d4edda;
                color: #155724;
                display: block;
            }
            .status.error {
                background: #f8d7da;
                color: #721c24;
                display: block;
            }
            .status.loading {
                background: #d1ecf1;
                color: #0c5460;
                display: block;
            }
            .spinner {
                border: 3px solid #f3f3f3;
                border-top: 3px solid #667eea;
                border-radius: 50%;
                width: 30px;
                height: 30px;
                animation: spin 1s linear infinite;
                margin: 0 auto 10px;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            .features {
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 15px;
                margin-top: 30px;
                padding-top: 30px;
                border-top: 1px solid #eee;
            }
            .feature {
                display: flex;
                align-items: center;
                gap: 10px;
                font-size: 13px;
                color: #666;
            }
            .feature-icon {
                color: #667eea;
                font-size: 18px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🚀 PDF to DOCX</h1>
            <p class="subtitle">Convert your PDF files to Word documents instantly</p>
            
            <form id="uploadForm" enctype="multipart/form-data">
                <div class="upload-area" id="uploadArea">
                    <div class="upload-icon">📄</div>
                    <div class="upload-text">Click to upload or drag & drop</div>
                    <div class="upload-hint">PDF files only • Max 10MB</div>
                    <input type="file" id="pdfFile" name="pdf_file" accept=".pdf" required>
                </div>
                
                <button type="submit" class="btn" id="convertBtn">
                    Convert to DOCX
                </button>
            </form>
            
            <div class="status" id="status"></div>
            
            <div class="features">
                <div class="feature">
                    <span class="feature-icon">✅</span>
                    <span>Preserves formatting</span>
                </div>
                <div class="feature">
                    <span class="feature-icon">⚡</span>
                    <span>Fast conversion</span>
                </div>
                <div class="feature">
                    <span class="feature-icon">🔒</span>
                    <span>Secure & private</span>
                </div>
                <div class="feature">
                    <span class="feature-icon">🎯</span>
                    <span>Auto cleanup</span>
                </div>
            </div>
        </div>
        
        <script>
            const uploadArea = document.getElementById('uploadArea');
            const fileInput = document.getElementById('pdfFile');
            const form = document.getElementById('uploadForm');
            const status = document.getElementById('status');
            const convertBtn = document.getElementById('convertBtn');
            
            // Click to upload
            uploadArea.addEventListener('click', () => fileInput.click());
            
            // Drag and drop
            uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                uploadArea.classList.add('dragover');
            });
            
            uploadArea.addEventListener('dragleave', () => {
                uploadArea.classList.remove('dragover');
            });
            
            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('dragover');
                
                const files = e.dataTransfer.files;
                if (files.length > 0 && files[0].type === 'application/pdf') {
                    fileInput.files = files;
                    updateUploadText(files[0].name);
                }
            });
            
            // File selected
            fileInput.addEventListener('change', (e) => {
                if (e.target.files.length > 0) {
                    updateUploadText(e.target.files[0].name);
                }
            });
            
            function updateUploadText(filename) {
                document.querySelector('.upload-text').textContent = filename;
                document.querySelector('.upload-hint').textContent = 'Ready to convert';
            }
            
            // Form submission
            form.addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const formData = new FormData(form);
                
                // Show loading
                status.className = 'status loading';
                status.innerHTML = '<div class="spinner"></div><div>Converting... This may take 30-60 seconds</div>';
                convertBtn.disabled = true;
                
                try {
                    const response = await fetch('/convert', {
                        method: 'POST',
                        body: formData
                    });
                    
                    if (response.ok) {
                        // Download file
                        const blob = await response.blob();
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'converted.docx';
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        document.body.removeChild(a);
                        
                        // Show success
                        status.className = 'status success';
                        status.textContent = '✅ Success! Your file has been downloaded.';
                    } else {
                        const error = await response.json();
                        throw new Error(error.message || 'Conversion failed');
                    }
                } catch (error) {
                    status.className = 'status error';
                    status.textContent = '❌ ' + error.message;
                } finally {
                    convertBtn.disabled = false;
                }
            });
        </script>
    </body>
    </html>
    """


@app.route("/convert", methods=["POST"])
def convert():
    """Handle PDF upload and conversion"""
    
    # Log initial memory
    initial_memory = get_memory_usage()
    logger.info(f"📊 Initial memory: {initial_memory:.1f}MB")
    
    pdf_path = None
    output_path = None
    
    try:
        # Validate request
        if "pdf_file" not in request.files:
            return jsonify({"success": False, "message": "No file uploaded"}), 400
        
        file = request.files["pdf_file"]
        
        if file.filename == "":
            return jsonify({"success": False, "message": "No file selected"}), 400
        
        if not allowed_file(file.filename):
            return jsonify({
                "success": False,
                "message": "Invalid file type. Only PDF files allowed."
            }), 400
        
        # Secure filename
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        
        # Save uploaded file
        file.save(pdf_path)
        logger.info(f"📄 File uploaded: {filename} ({os.path.getsize(pdf_path) / (1024*1024):.2f}MB)")
        
        # Generate output path
        output_name = Path(filename).with_suffix(".docx").name
        output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_name)
        
        # Convert with memory monitoring
        logger.info(f"🔄 Starting conversion: {filename}")
        success, result = converter.convert_with_memory_monitoring(
            pdf_path,
            output_path,
            verbose=True
        )
        
        if not success:
            logger.error(f"❌ Conversion failed: {result}")
            return jsonify({"success": False, "message": result}), 500
        
        logger.info(f"✅ Conversion successful: {output_name}")
        
        # Send file and cleanup after
        @after_this_request
        def cleanup(response):
            try:
                cleanup_files(pdf_path, output_path)
                gc.collect()  # Force garbage collection
            except Exception as e:
                logger.error(f"Cleanup error: {e}")
            return response
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    except Exception as e:
        logger.error(f"❌ Server error: {e}", exc_info=True)
        
        # Cleanup on error
        if pdf_path:
            cleanup_files(pdf_path)
        if output_path:
            cleanup_files(output_path)
        
        gc.collect()
        
        return jsonify({
            "success": False,
            "message": f"Conversion error: {str(e)}"
        }), 500


@app.route("/health")
def health():
    """Health check with memory info"""
    memory_mb = get_memory_usage()
    return jsonify({
        "status": "healthy",
        "service": "PDF to DOCX Converter",
        "memory_mb": round(memory_mb, 2)
    }), 200


@app.errorhandler(413)
def file_too_large(e):
    """Handle file size limit exceeded"""
    return jsonify({
        "success": False,
        "message": "File too large. Maximum size is 10 MB for free hosting."
    }), 413


@app.errorhandler(500)
def internal_error(e):
    """Handle 500 errors"""
    logger.error(f"Internal error: {e}")
    gc.collect()  # Cleanup on error
    return jsonify({"error": "Internal server error"}), 500


if __name__ == "__main__":
    logger.info("=" * 70)
    logger.info("🚀 Memory-Optimized PDF to DOCX Converter")
    logger.info("=" * 70)
    logger.info(f"📂 Upload folder: {os.path.abspath(UPLOAD_FOLDER)}")
    logger.info(f"📂 Output folder: {os.path.abspath(OUTPUT_FOLDER)}")
    logger.info(f"📏 Max file size: {MAX_CONTENT_LENGTH / (1024*1024):.0f} MB")
    logger.info(f"💾 Max memory: 400 MB")
    logger.info("=" * 70)
    
    # For local development
    app.run(debug=False, host="0.0.0.0", port=5000)



 