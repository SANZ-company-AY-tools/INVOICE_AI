"""
Flask web application for invoice data extraction.
Provides a web interface to upload invoices and download extracted data as Excel.
With Microsoft 365 SSO authentication.
"""

import os
import uuid
import msal
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, jsonify, send_file, url_for, redirect, session
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix

from extractor import InvoiceExtractor
from excel_generator import ExcelGenerator, SAPCSVGenerator

app = Flask(__name__)

# Fix for running behind a proxy (Railway, Heroku, etc.) - ensures HTTPS URLs
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.config['PREFERRED_URL_SCHEME'] = 'https'

# Secret key for session
app.secret_key = os.environ.get('FLASK_SECRET_KEY', os.urandom(24))

# Configuration
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'png', 'jpg', 'jpeg', 'bmp', 'tiff', 'gif', 'docx', 'doc'}
app.config['SESSION_TYPE'] = 'filesystem'

# Microsoft 365 / Azure AD Configuration
AZURE_CLIENT_ID = os.environ.get('AZURE_CLIENT_ID', 'dd46c7c1-75d6-4e18-88ea-fcd4a92f6ccd')
AZURE_CLIENT_SECRET = os.environ.get('AZURE_CLIENT_SECRET', '')
AZURE_TENANT_ID = os.environ.get('AZURE_TENANT_ID', 'bdab45bf-1644-4fe5-ae0f-66bde297f9f0')
AZURE_AUTHORITY = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
AZURE_SCOPE = ["User.Read"]
SHAREPOINT_SCOPE = ["User.Read", "Files.Read.All", "Sites.Read.All"]

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Initialize extractors
extractor = InvoiceExtractor()
excel_gen = ExcelGenerator()
csv_gen = SAPCSVGenerator()


def get_msal_app():
    """Create MSAL confidential client application."""
    return msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=AZURE_AUTHORITY,
        client_credential=AZURE_CLIENT_SECRET,
    )


def login_required(f):
    """Decorator to require login for routes."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/login')
def login():
    """Redirect to Microsoft login."""
    msal_app = get_msal_app()

    # Get the redirect URI based on the request
    redirect_uri = url_for('auth_callback', _external=True)

    auth_url = msal_app.get_authorization_request_url(
        AZURE_SCOPE,
        redirect_uri=redirect_uri,
        prompt="select_account"
    )
    return redirect(auth_url)


@app.route('/auth/callback')
def auth_callback():
    """Handle Microsoft OAuth callback."""
    if 'error' in request.args:
        return f"Error: {request.args.get('error_description', 'Unknown error')}", 400

    if 'code' not in request.args:
        return redirect(url_for('login'))

    msal_app = get_msal_app()
    redirect_uri = url_for('auth_callback', _external=True)

    result = msal_app.acquire_token_by_authorization_code(
        request.args['code'],
        scopes=AZURE_SCOPE,
        redirect_uri=redirect_uri
    )

    if 'error' in result:
        return f"Error: {result.get('error_description', 'Could not acquire token')}", 400

    # Store user info in session
    session['user'] = result.get('id_token_claims')
    session['access_token'] = result.get('access_token')

    return redirect(url_for('index'))


@app.route('/logout')
def logout():
    """Clear session and logout."""
    session.clear()
    # Redirect to Microsoft logout
    logout_url = f"{AZURE_AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('login', _external=True)}"
    return redirect(logout_url)


@app.route('/')
@login_required
def index():
    """Render main page."""
    user = session.get('user', {})
    return render_template('index.html', user=user, client_id=AZURE_CLIENT_ID)


@app.route('/sharepoint/login')
@login_required
def sharepoint_login():
    """Redirect to Microsoft login with SharePoint/Files permissions."""
    msal_app = get_msal_app()
    redirect_uri = url_for('sharepoint_callback', _external=True)

    auth_url = msal_app.get_authorization_request_url(
        SHAREPOINT_SCOPE,
        redirect_uri=redirect_uri,
        prompt="consent"
    )
    return redirect(auth_url)


@app.route('/sharepoint/callback')
@login_required
def sharepoint_callback():
    """Handle SharePoint OAuth callback - store Graph token in session."""
    if 'error' in request.args:
        return f"Error: {request.args.get('error_description', 'Unknown error')}", 400

    if 'code' not in request.args:
        return redirect(url_for('sharepoint_login'))

    msal_app = get_msal_app()
    redirect_uri = url_for('sharepoint_callback', _external=True)

    result = msal_app.acquire_token_by_authorization_code(
        request.args['code'],
        scopes=SHAREPOINT_SCOPE,
        redirect_uri=redirect_uri
    )

    if 'error' in result:
        return f"Error SharePoint: {result.get('error_description', 'Could not acquire token')}", 400

    # Store Graph access token in session
    session['graph_token'] = result.get('access_token')
    return redirect(url_for('index'))


@app.route('/sharepoint/browse')
@login_required
def sharepoint_browse():
    """Browse SharePoint/OneDrive files. Supports folder navigation."""
    import requests as req

    token = session.get('graph_token')
    if not token:
        return jsonify({'error': 'no_token', 'login_url': url_for('sharepoint_login')}), 401

    # Get folder_id param for navigation (empty = root)
    folder_id = request.args.get('folder_id', '')
    source = request.args.get('source', 'onedrive')  # onedrive or site

    try:
        headers = {'Authorization': f'Bearer {token}'}

        if source == 'onedrive':
            if folder_id:
                url = f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children?$top=200'
            else:
                url = 'https://graph.microsoft.com/v1.0/me/drive/root/children?$top=200'
        else:
            # SharePoint site drive
            site_id = request.args.get('site_id', '')
            drive_id = request.args.get('drive_id', '')
            if folder_id:
                url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children?$top=200'
            else:
                url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children?$top=200'

        resp = req.get(url, headers=headers)

        if resp.status_code == 401:
            session.pop('graph_token', None)
            return jsonify({'error': 'no_token', 'login_url': url_for('sharepoint_login')}), 401

        if not resp.ok:
            return jsonify({'error': f'Graph API error: {resp.status_code}'}), 500

        data = resp.json()
        items = []
        allowed_exts = {'.pdf', '.png', '.jpg', '.jpeg', '.docx', '.tiff', '.bmp', '.gif'}

        for item in data.get('value', []):
            is_folder = 'folder' in item
            name = item.get('name', '')
            ext = os.path.splitext(name)[1].lower()

            if is_folder or ext in allowed_exts:
                items.append({
                    'id': item['id'],
                    'name': name,
                    'is_folder': is_folder,
                    'size': item.get('size', 0),
                    'modified': item.get('lastModifiedDateTime', ''),
                    'download_url': item.get('@microsoft.graph.downloadUrl', ''),
                })

        return jsonify({'items': items})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/sharepoint/sites')
@login_required
def sharepoint_sites():
    """List SharePoint sites the user has access to."""
    import requests as req

    token = session.get('graph_token')
    if not token:
        return jsonify({'error': 'no_token', 'login_url': url_for('sharepoint_login')}), 401

    try:
        headers = {'Authorization': f'Bearer {token}'}

        # Search for sites the user follows or has access to
        resp = req.get('https://graph.microsoft.com/v1.0/sites?search=*', headers=headers)

        if resp.status_code == 401:
            session.pop('graph_token', None)
            return jsonify({'error': 'no_token', 'login_url': url_for('sharepoint_login')}), 401

        sites = []
        if resp.ok:
            for site in resp.json().get('value', []):
                # Get drives for each site
                drives_resp = req.get(f'https://graph.microsoft.com/v1.0/sites/{site["id"]}/drives', headers=headers)
                drives = []
                if drives_resp.ok:
                    for d in drives_resp.json().get('value', []):
                        drives.append({'id': d['id'], 'name': d.get('name', 'Documents')})

                sites.append({
                    'id': site['id'],
                    'name': site.get('displayName', site.get('name', '')),
                    'url': site.get('webUrl', ''),
                    'drives': drives
                })

        return jsonify({'sites': sites})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/sharepoint/download', methods=['POST'])
@login_required
def sharepoint_download():
    """Download files from SharePoint/OneDrive and add them to upload queue."""
    import requests as req

    token = session.get('graph_token')
    if not token:
        return jsonify({'error': 'no_token'}), 401

    data = request.get_json()
    file_ids = data.get('file_ids', [])

    if not file_ids:
        return jsonify({'error': 'No files selected'}), 400

    headers = {'Authorization': f'Bearer {token}'}
    session_id = str(uuid.uuid4())[:8]
    session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    os.makedirs(session_folder, exist_ok=True)

    downloaded = []
    for file_info in file_ids:
        try:
            file_id = file_info['id']
            filename = secure_filename(file_info['name'])
            drive_id = file_info.get('drive_id', '')

            # Get download URL from Graph API
            if drive_id:
                url = f'https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content'
            else:
                url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content'

            resp = req.get(url, headers=headers, allow_redirects=True)
            if resp.ok:
                filepath = os.path.join(session_folder, filename)
                with open(filepath, 'wb') as f:
                    f.write(resp.content)
                downloaded.append(filepath)
        except Exception as e:
            print(f"Error downloading {file_info.get('name')}: {e}")

    if not downloaded:
        return jsonify({'error': 'No files could be downloaded'}), 500

    # Process invoices
    try:
        results = extractor.process_multiple_files(downloaded)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f'facturas_{session_id}_{timestamp}.xlsx'
        excel_path = excel_gen.generate_from_invoices(
            results, app.config['OUTPUT_FOLDER'], excel_filename
        )

        csv_filename = f'facturas_sap_{session_id}_{timestamp}.csv'
        csv_path = csv_gen.generate_csv(
            results, app.config['OUTPUT_FOLDER'], csv_filename
        )

        response_data = {
            'success': True,
            'session_id': session_id,
            'files_processed': len(downloaded),
            'results': [],
            'excel_filename': excel_filename,
            'download_url': url_for('download_excel', filename=excel_filename),
            'csv_filename': csv_filename,
            'csv_download_url': url_for('download_csv', filename=csv_filename)
        }

        for r in results:
            preview = {k: v for k, v in r.items() if k != 'raw_text'}
            response_data['results'].append(preview)

        return jsonify(response_data)

    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500


@app.route('/upload', methods=['POST'])
@login_required
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
        return jsonify({'error': 'No valid files uploaded. Allowed: PDF, PNG, JPG, JPEG, BMP, TIFF, GIF, DOCX'}), 400

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

        # Generate SAP CSV
        csv_filename = f'facturas_sap_{session_id}_{timestamp}.csv'
        csv_path = csv_gen.generate_csv(
            results,
            app.config['OUTPUT_FOLDER'],
            csv_filename
        )

        # Prepare response data
        response_data = {
            'success': True,
            'session_id': session_id,
            'files_processed': len(uploaded_files),
            'results': [],
            'excel_filename': excel_filename,
            'download_url': url_for('download_excel', filename=excel_filename),
            'csv_filename': csv_filename,
            'csv_download_url': url_for('download_csv', filename=csv_filename)
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


@app.route('/download/csv/<filename>')
@login_required
def download_csv(filename: str):
    """Download generated SAP CSV file."""
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
    if os.path.exists(filepath):
        return send_file(
            filepath,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
    return jsonify({'error': 'File not found'}), 404


@app.route('/download/<filename>')
@login_required
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
