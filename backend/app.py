"""
Flask Backend for Media Reconciliation System
Modern web application with Google Sheets integration
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import sys
from werkzeug.utils import secure_filename
import pandas as pd
from datetime import datetime
import tempfile
import json

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import existing functionality
from func.matching import match_media_watch_with_tc
from func.format import standardize_dataframe, normalize_theme_name
from func.create_pdf import create_summary_pdf_tables_only

# Import Google Sheets and PDF conversion modules
from app.google_sheets import GoogleSheetsService
from app.pdf_converter import TCPDFConverter

app = Flask(__name__)
CORS(app)  # Enable CORS for React frontend

# Configuration
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'reconciliation_uploads')
ALLOWED_EXTENSIONS = {'pdf', 'xlsx', 'xls'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Global storage for session data (in production, use Redis or database)
sessions = {}


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_session(session_id):
    """Get or create session data"""
    if session_id not in sessions:
        sessions[session_id] = {
            'lmrb_data': None,
            'tc_data': None,
            'theme_mappings': [],
            'matched_data': None,
            'unmatched_lmrb': None,
            'unmatched_tc': None,
            'google_sheet_id': None,
            'date_range': None
        }
    return sessions[session_id]


@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'message': 'Reconciliation API is running'})


@app.route('/api/google-sheets/connect', methods=['POST'])
def connect_google_sheet():
    """
    Connect to Google Sheet for LMRB data
    Expected: { "sheet_id": "your-google-sheet-id", "session_id": "unique-session" }
    """
    try:
        data = request.get_json()
        sheet_id = data.get('sheet_id')
        session_id = data.get('session_id', 'default')

        if not sheet_id:
            return jsonify({'error': 'sheet_id is required'}), 400

        # Initialize Google Sheets service
        gs_service = GoogleSheetsService()

        # Get sheet metadata
        metadata = gs_service.get_sheet_metadata(sheet_id)

        # Store sheet ID in session
        session = get_session(session_id)
        session['google_sheet_id'] = sheet_id

        return jsonify({
            'success': True,
            'message': 'Connected to Google Sheet',
            'metadata': metadata
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/google-sheets/lmrb-data', methods=['POST'])
def get_lmrb_data():
    """
    Fetch LMRB data from Google Sheet with optional date range
    Expected: {
        "session_id": "unique-session",
        "start_date": "2024-01-01" (optional),
        "end_date": "2024-12-31" (optional),
        "channel": "Channel Name" (optional)
    }
    """
    try:
        data = request.get_json()
        session_id = data.get('session_id', 'default')
        start_date = data.get('start_date')
        end_date = data.get('end_date')
        channel = data.get('channel')

        session = get_session(session_id)
        sheet_id = session.get('google_sheet_id')

        if not sheet_id:
            return jsonify({'error': 'No Google Sheet connected. Call /connect first'}), 400

        # Fetch data from Google Sheets
        gs_service = GoogleSheetsService()
        lmrb_df = gs_service.get_lmrb_data(
            sheet_id,
            start_date=start_date,
            end_date=end_date
        )

        # Standardize column names
        lmrb_df = standardize_dataframe(lmrb_df)

        # Filter by channel if provided
        if channel and channel != 'All':
            lmrb_df = lmrb_df[lmrb_df['Channel'].str.upper() == channel.upper()]

        # Store in session
        session['lmrb_data'] = lmrb_df
        session['date_range'] = {'start': start_date, 'end': end_date}

        # Get unique channels and products for filtering
        channels = sorted(lmrb_df['Channel'].unique().tolist())
        products = sorted(lmrb_df['Product'].unique().tolist()) if 'Product' in lmrb_df.columns else []

        return jsonify({
            'success': True,
            'record_count': len(lmrb_df),
            'channels': channels,
            'products': products,
            'date_range': {
                'min': lmrb_df['Date'].min().strftime('%Y-%m-%d') if 'Date' in lmrb_df.columns else None,
                'max': lmrb_df['Date'].max().strftime('%Y-%m-%d') if 'Date' in lmrb_df.columns else None
            },
            'sample_data': lmrb_df.head(10).to_dict('records')
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload-tc', methods=['POST'])
def upload_tc_pdf():
    """
    Upload and convert TC PDF to Excel data
    Expected: multipart/form-data with 'file' and 'session_id'
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']
        session_id = request.form.get('session_id', 'default')

        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Only PDF files are allowed'}), 400

        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{filename}")
        file.save(filepath)

        # Convert PDF to Excel data
        converter = TCPDFConverter()
        tc_df = converter.convert_pdf_to_dataframe(filepath)

        # Standardize column names
        tc_df = standardize_dataframe(tc_df)

        # Store in session
        session = get_session(session_id)
        session['tc_data'] = tc_df

        # Get unique themes and channels
        themes = sorted(tc_df['Theme'].unique().tolist()) if 'Theme' in tc_df.columns else []

        return jsonify({
            'success': True,
            'message': f'TC PDF processed successfully',
            'record_count': len(tc_df),
            'themes': themes,
            'sample_data': tc_df.head(10).to_dict('records')
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

    finally:
        # Clean up uploaded file
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except:
                pass


@app.route('/api/theme-mappings', methods=['GET', 'POST', 'DELETE'])
def manage_theme_mappings():
    """
    Manage theme mappings
    GET: Retrieve all mappings
    POST: Add new mapping
    DELETE: Remove mapping by index
    """
    session_id = request.args.get('session_id', 'default')
    session = get_session(session_id)

    if request.method == 'GET':
        return jsonify({
            'success': True,
            'mappings': session['theme_mappings']
        })

    elif request.method == 'POST':
        try:
            mapping = request.get_json()

            # Validate required fields
            required_fields = ['tc_theme', 'tc_duration', 'lmrb_theme', 'lmrb_duration', 'mapping_type']
            for field in required_fields:
                if field not in mapping:
                    return jsonify({'error': f'Missing required field: {field}'}), 400

            # Add mapping
            session['theme_mappings'].append(mapping)

            return jsonify({
                'success': True,
                'message': 'Mapping added successfully',
                'total_mappings': len(session['theme_mappings'])
            })

        except Exception as e:
            return jsonify({'error': str(e)}), 500

    elif request.method == 'DELETE':
        try:
            data = request.get_json()
            index = data.get('index')

            if index is None or index < 0 or index >= len(session['theme_mappings']):
                return jsonify({'error': 'Invalid mapping index'}), 400

            # Remove mapping
            del session['theme_mappings'][index]

            return jsonify({
                'success': True,
                'message': 'Mapping deleted successfully',
                'total_mappings': len(session['theme_mappings'])
            })

        except Exception as e:
            return jsonify({'error': str(e)}), 500


@app.route('/api/reconcile', methods=['POST'])
def run_reconciliation():
    """
    Run reconciliation matching
    Expected: {
        "session_id": "unique-session",
        "time_tolerance": 30,
        "ignore_date": false,
        "products": ["Product1", "Product2"] (optional)
    }
    """
    try:
        data = request.get_json()
        session_id = data.get('session_id', 'default')
        time_tolerance = data.get('time_tolerance', 30)
        ignore_date = data.get('ignore_date', False)
        selected_products = data.get('products', [])

        session = get_session(session_id)

        # Validate data availability
        if session['lmrb_data'] is None:
            return jsonify({'error': 'No LMRB data loaded. Fetch from Google Sheets first'}), 400

        if session['tc_data'] is None:
            return jsonify({'error': 'No TC data uploaded'}), 400

        if not session['theme_mappings']:
            return jsonify({'error': 'No theme mappings defined'}), 400

        # Filter LMRB data by selected products
        lmrb_data = session['lmrb_data'].copy()
        if selected_products:
            lmrb_data = lmrb_data[lmrb_data['Product'].isin(selected_products)]

        # Run matching algorithm
        matched_df, matched_indices, program_mismatched_df, matched_tc_keys = match_media_watch_with_tc(
            media_watch_df=lmrb_data,
            tc_df=session['tc_data'],
            theme_mapping=session['theme_mappings'],
            ignore_date=ignore_date,
            time_tolerance=time_tolerance
        )

        # Calculate unmatched records
        unmatched_lmrb = lmrb_data[~lmrb_data.index.isin(matched_indices)]
        unmatched_tc = session['tc_data'][~session['tc_data'].index.isin([k[0] for k in matched_tc_keys])]

        # Store results in session
        session['matched_data'] = matched_df
        session['unmatched_lmrb'] = unmatched_lmrb
        session['unmatched_tc'] = unmatched_tc

        # Calculate statistics
        total_lmrb = len(lmrb_data)
        total_matched = len(matched_df)
        match_rate = (total_matched / total_lmrb * 100) if total_lmrb > 0 else 0

        return jsonify({
            'success': True,
            'statistics': {
                'total_lmrb_records': total_lmrb,
                'total_tc_records': len(session['tc_data']),
                'matched_records': total_matched,
                'unmatched_lmrb': len(unmatched_lmrb),
                'unmatched_tc': len(unmatched_tc),
                'match_rate': round(match_rate, 2),
                'program_mismatched': len(program_mismatched_df)
            },
            'matched_data': matched_df.head(50).to_dict('records'),  # Send first 50 for preview
            'unmatched_lmrb_sample': unmatched_lmrb.head(20).to_dict('records'),
            'unmatched_tc_sample': unmatched_tc.head(20).to_dict('records')
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/results/matched', methods=['GET'])
def get_matched_results():
    """Get all matched results with pagination"""
    try:
        session_id = request.args.get('session_id', 'default')
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 50))

        session = get_session(session_id)

        if session['matched_data'] is None:
            return jsonify({'error': 'No reconciliation results available'}), 400

        matched_df = session['matched_data']

        # Pagination
        start_idx = (page - 1) * per_page
        end_idx = start_idx + per_page

        total_records = len(matched_df)
        total_pages = (total_records + per_page - 1) // per_page

        return jsonify({
            'success': True,
            'data': matched_df.iloc[start_idx:end_idx].to_dict('records'),
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total_records': total_records,
                'total_pages': total_pages
            }
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/results/unmatched', methods=['GET'])
def get_unmatched_results():
    """Get unmatched LMRB and TC records"""
    try:
        session_id = request.args.get('session_id', 'default')
        data_type = request.args.get('type', 'lmrb')  # 'lmrb' or 'tc'

        session = get_session(session_id)

        if data_type == 'lmrb':
            if session['unmatched_lmrb'] is None:
                return jsonify({'error': 'No unmatched LMRB data available'}), 400
            data = session['unmatched_lmrb']
        else:
            if session['unmatched_tc'] is None:
                return jsonify({'error': 'No unmatched TC data available'}), 400
            data = session['unmatched_tc']

        return jsonify({
            'success': True,
            'data': data.to_dict('records'),
            'count': len(data)
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/export/excel', methods=['POST'])
def export_to_excel():
    """Export reconciliation results to Excel"""
    try:
        data = request.get_json()
        session_id = data.get('session_id', 'default')

        session = get_session(session_id)

        if session['matched_data'] is None:
            return jsonify({'error': 'No reconciliation results to export'}), 400

        # Create Excel file
        output_filename = f"reconciliation_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Matched data
            session['matched_data'].to_excel(writer, sheet_name='Matched Records', index=False)

            # Unmatched LMRB
            if session['unmatched_lmrb'] is not None and len(session['unmatched_lmrb']) > 0:
                session['unmatched_lmrb'].to_excel(writer, sheet_name='Unmatched LMRB', index=False)

            # Unmatched TC
            if session['unmatched_tc'] is not None and len(session['unmatched_tc']) > 0:
                session['unmatched_tc'].to_excel(writer, sheet_name='Unmatched TC', index=False)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/export/pdf', methods=['POST'])
def export_to_pdf():
    """Export summary report to PDF"""
    try:
        data = request.get_json()
        session_id = data.get('session_id', 'default')
        channel = data.get('channel', 'All Channels')

        session = get_session(session_id)

        if session['matched_data'] is None:
            return jsonify({'error': 'No reconciliation results to export'}), 400

        # Create PDF
        output_filename = f"reconciliation_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        # Prepare data for PDF
        total_lmrb = len(session['lmrb_data']) if session['lmrb_data'] is not None else 0
        total_matched = len(session['matched_data'])
        total_unmatched_lmrb = len(session['unmatched_lmrb']) if session['unmatched_lmrb'] is not None else 0
        total_unmatched_tc = len(session['unmatched_tc']) if session['unmatched_tc'] is not None else 0

        summary_stats = {
            'channel': channel,
            'total_lmrb': total_lmrb,
            'total_tc': len(session['tc_data']) if session['tc_data'] is not None else 0,
            'matched': total_matched,
            'unmatched_lmrb': total_unmatched_lmrb,
            'unmatched_tc': total_unmatched_tc,
            'match_rate': (total_matched / total_lmrb * 100) if total_lmrb > 0 else 0
        }

        create_summary_pdf_tables_only(
            output_path,
            channel,
            session['matched_data'],
            summary_stats
        )

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/pdf'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/manual-match', methods=['POST'])
def manual_match():
    """Manually match unmatched LMRB and TC records"""
    try:
        data = request.get_json()
        session_id = data.get('session_id', 'default')
        lmrb_index = data.get('lmrb_index')
        tc_index = data.get('tc_index')

        session = get_session(session_id)

        if session['unmatched_lmrb'] is None or session['unmatched_tc'] is None:
            return jsonify({'error': 'No unmatched records available'}), 400

        # Get records
        lmrb_record = session['unmatched_lmrb'].iloc[lmrb_index]
        tc_record = session['unmatched_tc'].iloc[tc_index]

        # Create manual match (combine both records)
        manual_match_record = pd.concat([lmrb_record, tc_record])

        # Add to matched data
        session['matched_data'] = pd.concat([
            session['matched_data'],
            manual_match_record.to_frame().T
        ], ignore_index=True)

        # Remove from unmatched
        session['unmatched_lmrb'] = session['unmatched_lmrb'].drop(session['unmatched_lmrb'].index[lmrb_index])
        session['unmatched_tc'] = session['unmatched_tc'].drop(session['unmatched_tc'].index[tc_index])

        return jsonify({
            'success': True,
            'message': 'Records matched manually',
            'remaining_unmatched_lmrb': len(session['unmatched_lmrb']),
            'remaining_unmatched_tc': len(session['unmatched_tc'])
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("=" * 60)
    print("Media Reconciliation API Server")
    print("=" * 60)
    print(f"Server starting on http://localhost:5000")
    print(f"API Documentation: http://localhost:5000/api/health")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)
