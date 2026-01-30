"""
Flask web application for invoice data extraction.
Provides a web interface to upload invoices and download extracted data as Excel.
"""

import os
import uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, url_for
from werkzeug.utils import secure_filename

from extractor import InvoiceExtractor
from excel_generator import ExcelGenerator

app = Flask(__name__)

# Configuration
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'png', 'jpg', 'jpeg', 'bmp', 'tiff', 'gif', 'docx', 'doc'}

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Initialize extractors
extractor = InvoiceExtractor()
excel_gen = ExcelGenerator()


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    """Render main page."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads and process invoices."""
    if 'files' not in request.files:
        return jsonify({'error': 'No files provided'}), 400

    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No files selected'}), 400

    # Create unique session folder
    session_id = str(uuid.uuid4())[:8]
    session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    os.makedirs(session_folder, exist_ok=True)

    uploaded_files = []
    for file in files:
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(session_folder, filename)
            file.save(filepath)
            uploaded_files.append(filepath)

    if not uploaded_files:
        return jsonify({'error': 'No valid files uploaded. Allowed: PDF, PNG, JPG, JPEG, BMP, TIFF, GIF'}), 400

    # Process invoices
    try:
        results = extractor.process_multiple_files(uploaded_files)

        # Generate Excel
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f'facturas_{session_id}_{timestamp}.xlsx'
        excel_path = excel_gen.generate_from_invoices(
            results,
            app.config['OUTPUT_FOLDER'],
            excel_filename
        )

        # Prepare response data
        response_data = {
            'success': True,
            'session_id': session_id,
            'files_processed': len(uploaded_files),
            'results': [],
            'excel_filename': excel_filename,
            'download_url': url_for('download_excel', filename=excel_filename)
        }

        # Add preview of extracted data (include error messages)
        for r in results:
            preview = {k: v for k, v in r.items() if k != 'raw_text'}
            response_data['results'].append(preview)
            # Print errors for debugging
            if r.get('status') == 'error':
                print(f"ERROR processing {r.get('file_name')}: {r.get('error')}")

        return jsonify(response_data)

    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500


@app.route('/download/<filename>')
def download_excel(filename: str):
    """Download generated Excel file."""
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
    if os.path.exists(filepath):
        return send_file(
            filepath,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    return jsonify({'error': 'File not found'}), 404


@app.route('/health')
def health():
    """Health check endpoint."""
    return jsonify({'status': 'ok', 'timestamp': datetime.now().isoformat()})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('RAILWAY_ENVIRONMENT') is None  # Solo debug en local
    print(f"\n✅ Servidor listo en: http://127.0.0.1:{port}\n")
    app.run(host='0.0.0.0', port=port, debug=debug)
