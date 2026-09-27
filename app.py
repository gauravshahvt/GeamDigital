import os
import uuid
import json
import re
from flask import Flask, request, jsonify, render_template, send_file, send_from_directory
from werkzeug.utils import secure_filename
from io import BytesIO, StringIO
import pymupdf

from pdf_engine.extractor import process_pdf, render_page_to_base64
from pdf_engine.excel_builder import create_excel_workbook, THEMES
from pdf_engine.ocr_engine import is_tesseract_available
from pdf_engine.voter_extractor import is_voter_list_pdf, extract_voters, process_multiple_wards

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max upload for batch voter lists
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
EXPORT_FOLDER = os.path.join(os.path.dirname(__file__), 'exports')
SAMPLES_FOLDER = os.path.join(os.path.dirname(__file__), 'static', 'samples')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(SAMPLES_FOLDER, exist_ok=True)

# In-memory session store for processed results
SESSION_CACHE = {}

def process_batch_wards(pdf_paths: list) -> dict:
    """Processes multiple or single voter list PDFs into a structured multi-sheet dataset."""
    res = process_multiple_wards(pdf_paths)
    
    # Generate page previews for each ward
    previews = {}
    for idx, path in enumerate(pdf_paths):
        try:
            doc = pymupdf.open(path)
            if len(doc) > 0:
                previews[f"ward_{idx+1}_p1"] = render_page_to_base64(doc[0])
            if len(doc) > 2:
                previews[f"ward_{idx+1}_p3"] = render_page_to_base64(doc[2])
            doc.close()
        except Exception:
            pass

    metadata = {
        'कंपनी': 'Geam Digital',
        'कुल वार्ड्स संख्या': f"{res['total_wards']} वार्ड",
        'कुल वैध सक्रिय मतदाता': f"{res['total_active_voters']} मतदाता",
        'कुल विलोपित (Deleted)': f"{res['total_deleted_voters']} मतदाता",
        'शीट 1': 'वार्ड_सारांश (Summary of All Wards)',
        'शीट 2': 'समस्त_मतदाता_सूची (11 Columns, Sequence-wise)',
    }

    return {
        'is_multi_ward': True,
        'total_wards': res['total_wards'],
        'total_active_voters': res['total_active_voters'],
        'total_deleted_voters': res['total_deleted_voters'],
        'total_scanned_voters': res['total_scanned_voters'],
        'summary_grid': res['summary_grid'],
        'master_grid': res['master_grid'],
        'wards': res['wards'],
        'tables': res['all_tables'],
        'table_names': res['all_names'],
        'tables_count': len(res['all_tables']),
        'metadata': metadata,
        'invoice_metadata': metadata,
        'previews': previews,
        'is_ocr_available': is_tesseract_available()
    }

@app.route('/')
def index():
    return render_template('index.html', themes=THEMES, ocr_ready=is_tesseract_available())

@app.route('/api/status')
def api_status():
    return jsonify({
        'status': 'online',
        'app_name': 'Geam Digital Multi-Ward Voter List to Excel Converter',
        'ocr_available': is_tesseract_available(),
        'available_themes': list(THEMES.keys())
    })

@app.route('/api/load_sample', methods=['POST'])
def load_sample():
    try:
        data = request.get_json() or {}
        sample_type = data.get('sample_type', 'multi_ward')

        if sample_type == 'single':
            sample_paths = [os.path.join(SAMPLES_FOLDER, 'sample_voter_list_rajasthan.pdf')]
            display_name = 'Geam_Digital_Ward_001.xlsx'
        else:
            # Multi-ward batch sample (Ward 1, Ward 2, Ward 7)
            sample_files = [
                'sample_voter_list_rajasthan.pdf',
                'sample_voter_list_ward_002.pdf',
                'sample_voter_list_ward_007.pdf'
            ]
            sample_paths = [os.path.join(SAMPLES_FOLDER, f) for f in sample_files if os.path.exists(os.path.join(SAMPLES_FOLDER, f))]
            display_name = f'Geam_Digital_Panchayat_3_Wards.xlsx'

        if not sample_paths:
            return jsonify({'error': 'Sample files not found'}), 404

        result = process_batch_wards(sample_paths)

        session_id = str(uuid.uuid4())
        SESSION_CACHE[session_id] = {
            'file_name': display_name,
            'result': result
        }

        return jsonify({
            'session_id': session_id,
            'file_name': display_name,
            'data': result
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def upload_pdf():
    try:
        # Collect all uploaded files (single or multiple)
        uploaded_files = []
        if 'pdf_files' in request.files:
            uploaded_files = request.files.getlist('pdf_files')
        elif 'pdf_file' in request.files:
            uploaded_files = request.files.getlist('pdf_file')

        uploaded_files = [f for f in uploaded_files if f and f.filename and f.filename.lower().endswith('.pdf')]
        if not uploaded_files:
            return jsonify({'error': 'कृपया कम से कम एक वैध .pdf मतदाता सूची फ़ाइल अपलोड करें'}), 400

        session_id = str(uuid.uuid4())
        saved_paths = []
        for f in uploaded_files:
            sec_name = secure_filename(f.filename)
            save_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{sec_name}")
            f.save(save_path)
            saved_paths.append(save_path)

        # Process multiple or single ward PDFs
        result = process_batch_wards(saved_paths)

        if len(saved_paths) == 1:
            base_name = os.path.splitext(uploaded_files[0].filename)[0]
            display_name = f"Geam_Digital_{base_name}.xlsx"
        else:
            display_name = f"Geam_Digital_Panchayat_{len(saved_paths)}_Wards.xlsx"

        SESSION_CACHE[session_id] = {
            'file_name': display_name,
            'saved_paths': saved_paths,
            'result': result
        }

        return jsonify({
            'session_id': session_id,
            'file_name': display_name,
            'data': result
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/export_excel', methods=['POST'])
def export_excel():
    try:
        payload = request.get_json() or {}
        tables = payload.get('tables', [])
        table_names = payload.get('table_names', ['वार्ड_सारांश', 'समस्त_मतदाता_सूची'])
        theme = payload.get('theme', 'geam_digital')
        layout_mode = payload.get('layout_mode', 'multi_sheet')
        metadata = payload.get('metadata', None)
        file_name = payload.get('file_name', 'Geam_Digital_Panchayat_Voter_List.xlsx')

        if not file_name.endswith('.xlsx'):
            file_name = os.path.splitext(file_name)[0] + '.xlsx'

        # Generate excel bytes
        excel_bytes = create_excel_workbook(
            tables=tables,
            table_names=table_names,
            theme=theme,
            mode=layout_mode,
            metadata=metadata
        )

        return send_file(
            BytesIO(excel_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/export_csv', methods=['POST'])
def export_csv():
    try:
        payload = request.get_json() or {}
        table = payload.get('table', [])
        file_name = payload.get('file_name', 'extracted_table.csv')

        if not file_name.endswith('.csv'):
            file_name = os.path.splitext(file_name)[0] + '.csv'

        import csv
        output = BytesIO()
        text_stream = StringIO()
        writer = csv.writer(text_stream)
        for row in table:
            writer.writerow(row)
        
        output.write(text_stream.getvalue().encode('utf-8-sig'))
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=file_name,
            mimetype='text/csv'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Starting PDF to Excel Studio on http://127.0.0.1:{port}")
    app.run(host='127.0.0.1', port=port, debug=False)
