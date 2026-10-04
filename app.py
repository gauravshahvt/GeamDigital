import os
import uuid
import json
import re
import base64
import threading
from typing import Optional, Any
from flask import Flask, request, jsonify, render_template, send_file, send_from_directory
from werkzeug.utils import secure_filename
from io import BytesIO, StringIO
import pymupdf

from pdf_engine.extractor import process_pdf, render_page_to_base64
from pdf_engine.excel_builder import create_excel_workbook, THEMES
from pdf_engine.ocr_engine import is_tesseract_available
from pdf_engine.voter_extractor import (
    is_voter_list_pdf,
    extract_voters,
    process_multiple_wards,
    parse_panchayat_and_wards_from_filenames
)
from pdf_engine.history_manager import (
    log_activity,
    get_recent_history,
    clear_all_history,
    delete_history_item
)
from pdf_engine.parchi_generator import (
    parse_excel_for_parchi,
    generate_parchi_html,
    generate_parchi_pdf,
    predict_gender,
    clean_val_str
)
from pdf_engine.color_parchi_generator import (
    generate_color_parchi_html,
    generate_color_parchi_pdf,
    SYMBOL_PRESETS
)
from pdf_engine.alphabetical_voter_list import (
    hindi_to_english_sort_key,
    sort_voters_alphabetically,
    generate_alphabetical_html,
    generate_alphabetical_pdf,
    generate_alphabetical_excel
)
from pdf_engine.family_voter_list import (
    group_and_sort_families,
    generate_family_html,
    generate_family_pdf,
    generate_family_excel
)
from pdf_engine.age_voter_list import (
    filter_and_sort_by_age,
    generate_age_html,
    generate_age_pdf,
    generate_age_excel
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max upload for batch voter lists
app.config['TEMPLATES_AUTO_RELOAD'] = True
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
EXPORT_FOLDER = os.path.join(os.path.dirname(__file__), 'exports')
SAMPLES_FOLDER = os.path.join(os.path.dirname(__file__), 'static', 'samples')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(SAMPLES_FOLDER, exist_ok=True)

# In-memory session stores
SESSION_CACHE = {}
PARCHI_CACHE = {}
COLOR_PARCHI_CACHE = {}
ALPHA_CACHE = {}
FAMILY_CACHE = {}
AGE_CACHE = {}
UPLOAD_TASKS = {}

def process_batch_wards(pdf_paths: list, progress_callback: Optional[Any] = None) -> dict:
    """Processes multiple or single voter list PDFs into a structured multi-sheet dataset."""
    res = process_multiple_wards(pdf_paths, progress_callback=progress_callback)
    
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
        'शीट 2': 'समस्त_मतदाता_सूची (13 Columns, Sequence-wise)',
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
        'is_ocr_available': is_tesseract_available(),
        'panchayat_name': res.get('panchayat_name', ''),
        'start_ward': res.get('start_ward', 1),
        'last_ward': res.get('last_ward', 1),
        'suggested_filename': res.get('suggested_filename', ''),
        'file_name': res.get('suggested_filename', '')
    }

def _worker_process_upload(task_id: str, saved_paths: list, display_name: str, uploaded_filenames: list, session_id: str):
    """Background worker for live streaming voter extraction with real-time progress & live rows."""
    try:
        def on_progress(info):
            if task_id not in UPLOAD_TASKS:
                return
            t = UPLOAD_TASKS[task_id]
            w_idx = info.get('ward_idx', 1)
            t_wards = max(1, info.get('total_wards', 1))
            page = info.get('page', 1)
            t_pages = max(1, info.get('total_pages', 1))

            base_pct = ((w_idx - 1) / t_wards) * 100
            page_pct = (page / t_pages) * (100 / t_wards)
            pct = int(min(98, round(base_pct + page_pct)))
            t['percent'] = max(t.get('percent', 0), pct)
            t['current_ward'] = w_idx
            t['total_wards'] = t_wards
            t['current_file'] = info.get('filename', '')
            t['current_page'] = page
            t['total_pages'] = t_pages
            t['total_voters'] = info.get('total_voters', t.get('total_voters', 0))

            if info.get('event') == 'voter' and 'voter' in info:
                v = info['voter']
                live_row = {
                    'वार्ड नं.': v.get('वार्ड नं.', ''),
                    'जि. प.': v.get('जि. प.', ''),
                    'पं. स.': v.get('पं. स.', ''),
                    'क्रम संख्या': v.get('क्रम संख्या', ''),
                    'नाम': v.get('नाम', ''),
                    'पिता/पति का नाम': v.get('पिता/पति का नाम', ''),
                    'आयु': v.get('आयु', ''),
                    'लिंग': v.get('लिंग', ''),
                    'हाउस नंबर': v.get('हाउस नंबर', ''),
                    'वोटर ID': v.get('वोटर ID', ''),
                    'Status': v.get('Status', 'Active')
                }
                t['recent_voters'].append(live_row)
                if len(t['recent_voters']) > 35:
                    t['recent_voters'].pop(0)
                t['latest_voter'] = live_row
                t['message'] = f"वार्ड {w_idx}/{t_wards} ({info.get('filename', '')}): पेज {page}/{t_pages} स्कैन जारी • {t['total_voters']} मतदाता प्राप्त"

        result = process_batch_wards(saved_paths, progress_callback=on_progress)

        final_filename = result.get('suggested_filename') or display_name
        display_name = final_filename

        SESSION_CACHE[session_id] = {
            'file_name': display_name,
            'saved_paths': saved_paths,
            'result': result
        }

        # Log activity into history (only file names stored, no PDF content)
        log_activity(
            feature_id='feature1',
            feature_name='फ़ीचर 1: वोटर लिस्ट PDF से Excel',
            file_names=uploaded_filenames,
            action='PDF अपलोड एवं 13-कॉलम एक्सेल निर्माण',
            details=f"{len(saved_paths)} वार्ड्स • {result.get('total_active_voters', 0)} सक्रिय मतदाता ({result.get('total_deleted_voters', 0)} विलोपित) • {display_name}"
        )

        UPLOAD_TASKS[task_id]['status'] = 'completed'
        UPLOAD_TASKS[task_id]['percent'] = 100
        UPLOAD_TASKS[task_id]['data'] = result
        UPLOAD_TASKS[task_id]['session_id'] = session_id
        UPLOAD_TASKS[task_id]['file_name'] = display_name
        UPLOAD_TASKS[task_id]['suggested_filename'] = display_name
        UPLOAD_TASKS[task_id]['message'] = f'सम्पूर्ण पंचायत 13-कॉलम एक्सेल ({display_name}) सफलतापूर्वक तैयार!'
    except Exception as e:
        import traceback
        traceback.print_exc()
        if task_id in UPLOAD_TASKS:
            UPLOAD_TASKS[task_id]['status'] = 'error'
            UPLOAD_TASKS[task_id]['error'] = str(e)
            UPLOAD_TASKS[task_id]['message'] = f'त्रुटि: {str(e)}'

@app.route('/')
def index():
    return render_template('index.html', themes=THEMES, ocr_ready=is_tesseract_available())

@app.route('/api/status')
def api_status():
    return jsonify({
        'status': 'online',
        'app_name': 'Geam Digital Multi-Ward Voter List to Excel Converter',
        'ocr_available': is_tesseract_available(),
        'available_themes': list(THEMES.keys()),
        'history_count': len(get_recent_history(200)),
        'recent_history': get_recent_history(5)
    })

# ==========================================
# HISTORY API (CROSS-FEATURE ACTIVITY LOG)
# ==========================================

@app.route('/api/history', methods=['GET'])
def api_get_history():
    feature_id = request.args.get('feature_id', None)
    records = get_recent_history(limit=100, feature_id=feature_id)
    return jsonify({
        'status': 'success',
        'history': records,
        'total': len(records)
    })

@app.route('/api/history/log', methods=['POST'])
def api_log_history():
    try:
        data = request.get_json() or {}
        feature_id = data.get('feature_id', 'feature1')
        feature_name = data.get('feature_name', 'फ़ीचर 1')
        file_names = data.get('file_names', data.get('file_name', 'अज्ञात फ़ाइल'))
        action = data.get('action', 'संपादन / एडिट')
        details = data.get('details', '')
        rec = log_activity(feature_id, feature_name, file_names, action, details)
        return jsonify({'status': 'success', 'record': rec})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/history/clear', methods=['POST'])
def api_clear_history():
    clear_all_history()
    return jsonify({'status': 'success', 'message': 'सभी इतिहास सफलतापूर्वक हटा दिया गया'})

@app.route('/api/history/delete', methods=['POST'])
def api_delete_history():
    data = request.get_json() or {}
    record_id = data.get('id', '')
    if not record_id:
        return jsonify({'error': 'id missing'}), 400
    delete_history_item(record_id)
    return jsonify({'status': 'success'})

# ==========================================
# FEATURE 1: LIVE STREAMING UPLOAD ENDPOINTS
# ==========================================

@app.route('/api/upload_start', methods=['POST'])
def upload_start():
    """Starts asynchronous voter extraction task with live preview support."""
    try:
        uploaded_files = []
        if 'pdf_files' in request.files:
            uploaded_files = request.files.getlist('pdf_files')
        elif 'pdf_file' in request.files:
            uploaded_files = request.files.getlist('pdf_file')

        uploaded_files = [f for f in uploaded_files if f and f.filename and f.filename.lower().endswith('.pdf')]
        if not uploaded_files:
            return jsonify({'error': 'कृपया कम से कम एक वैध .pdf मतदाता सूची फ़ाइल अपलोड करें'}), 400

        task_id = str(uuid.uuid4())
        session_id = str(uuid.uuid4())
        saved_paths = []
        uploaded_filenames = []
        for f in uploaded_files:
            sec_name = secure_filename(f.filename)
            uploaded_filenames.append(f.filename)
            save_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{sec_name}")
            f.save(save_path)
            saved_paths.append(save_path)

        init_panchayat, init_start, init_last = parse_panchayat_and_wards_from_filenames(uploaded_filenames)
        display_name = f"{init_panchayat} Ward {init_start} to {init_last}.xlsx"

        UPLOAD_TASKS[task_id] = {
            'task_id': task_id,
            'session_id': session_id,
            'status': 'processing',
            'percent': 0,
            'current_ward': 1,
            'total_wards': len(saved_paths),
            'current_file': uploaded_filenames[0] if uploaded_filenames else '',
            'current_page': 1,
            'total_pages': 1,
            'total_voters': 0,
            'recent_voters': [],
            'latest_voter': None,
            'message': 'प्रारंभिक विश्लेषण शुरू हो रहा है...',
            'file_name': display_name,
            'result': None,
            'error': None
        }

        t = threading.Thread(
            target=_worker_process_upload,
            args=(task_id, saved_paths, display_name, uploaded_filenames, session_id),
            daemon=True
        )
        t.start()

        return jsonify({
            'task_id': task_id,
            'session_id': session_id,
            'file_name': display_name,
            'total_files': len(saved_paths)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload_progress/<task_id>', methods=['GET'])
def upload_progress(task_id):
    """Returns real-time progress state and newly extracted live voter rows."""
    task = UPLOAD_TASKS.get(task_id)
    if not task:
        return jsonify({'error': 'टास्क उपलब्ध नहीं है'}), 404
    return jsonify(task)

@app.route('/api/load_sample', methods=['POST'])
def load_sample():
    try:
        data = request.get_json() or {}
        sample_type = data.get('sample_type', 'multi_ward')

        if sample_type == 'single':
            sample_paths = [os.path.join(SAMPLES_FOLDER, 'sample_voter_list_rajasthan.pdf')]
            display_name = 'Bagmali Ward 1 to 1.xlsx'
        else:
            # Multi-ward batch sample (Ward 1, Ward 2, Ward 7)
            sample_files = [
                'sample_voter_list_rajasthan.pdf',
                'sample_voter_list_ward_002.pdf',
                'sample_voter_list_ward_007.pdf'
            ]
            sample_paths = [os.path.join(SAMPLES_FOLDER, f) for f in sample_files if os.path.exists(os.path.join(SAMPLES_FOLDER, f))]
            display_name = 'Bagmali Ward 1 to 7.xlsx'

        if not sample_paths:
            return jsonify({'error': 'Sample files not found'}), 404

        result = process_batch_wards(sample_paths)
        final_filename = result.get('suggested_filename') or display_name
        display_name = final_filename

        session_id = str(uuid.uuid4())
        SESSION_CACHE[session_id] = {
            'file_name': display_name,
            'result': result
        }

        log_activity(
            feature_id='feature1',
            feature_name='फ़ीचर 1: वोटर लिस्ट PDF से Excel',
            file_names=[display_name],
            action='डेमो 3-वार्ड पंचायत लोड',
            details=f"डेमो • {result.get('total_active_voters', 0)} मतदाता • {display_name}"
        )

        return jsonify({
            'session_id': session_id,
            'file_name': display_name,
            'suggested_filename': display_name,
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
        uploaded_filenames = []
        for f in uploaded_files:
            sec_name = secure_filename(f.filename)
            uploaded_filenames.append(f.filename)
            save_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{sec_name}")
            f.save(save_path)
            saved_paths.append(save_path)

        # Process multiple or single ward PDFs
        result = process_batch_wards(saved_paths)

        final_filename = result.get('suggested_filename')
        if not final_filename:
            p_name, s_ward, l_ward = parse_panchayat_and_wards_from_filenames(uploaded_filenames)
            final_filename = f"{p_name} Ward {s_ward} to {l_ward}.xlsx"
        display_name = final_filename

        SESSION_CACHE[session_id] = {
            'file_name': display_name,
            'saved_paths': saved_paths,
            'result': result
        }

        log_activity(
            feature_id='feature1',
            feature_name='फ़ीचर 1: वोटर लिस्ट PDF से Excel',
            file_names=uploaded_filenames,
            action='PDF अपलोड एवं 13-कॉलम एक्सेल निर्माण',
            details=f"{len(saved_paths)} वार्ड्स • {result.get('total_active_voters', 0)} सक्रिय मतदाता • {display_name}"
        )

        return jsonify({
            'session_id': session_id,
            'file_name': display_name,
            'suggested_filename': display_name,
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

        log_activity(
            feature_id='feature1',
            feature_name='फ़ीचर 1: वोटर लिस्ट PDF से Excel',
            file_names=[file_name],
            action='एक्सेल शीट डाउनलोड (.xlsx)',
            details=f"{len(tables)} शीट्स • थीम: {theme}"
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

# ==========================================
# FEATURE 2: PLAIN MATDATA PARCHI ENDPOINTS
# ==========================================

@app.route('/api/parchi/upload_excel', methods=['POST'])
def api_parchi_upload_excel():
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'कृपया एक वैध एक्सेल फ़ाइल (.xlsx / .xls) चुनें'}), 400
        
        f = request.files['excel_file']
        if not f or not f.filename:
            return jsonify({'error': 'फ़ाइल उपलब्ध नहीं है'}), 400
            
        filename = secure_filename(f.filename)
        session_id = str(uuid.uuid4())
        save_path = os.path.join(UPLOAD_FOLDER, f"parchi_{session_id}_{filename}")
        f.save(save_path)
        
        parsed = parse_excel_for_parchi(save_path)
        
        PARCHI_CACHE[session_id] = {
            'file_path': save_path,
            'filename': filename,
            'parsed': parsed
        }

        log_activity(
            feature_id='feature2',
            feature_name='फ़ीचर 2: Plain मतदाता पर्ची',
            file_names=[filename],
            action='पर्ची हेतु एक्सेल अपलोड',
            details=f"{parsed['total_voters']} मतदाता • {len(parsed['sheet_names'])} शीट्स"
        )
        
        return jsonify({
            'session_id': session_id,
            'filename': filename,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': parsed['voters'][:12]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/parchi/select_sheet', methods=['POST'])
def api_parchi_select_sheet():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        sheet_name = data.get('sheet_name')
        
        if not session_id or session_id not in PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः अपलोड करें।'}), 404
            
        file_path = PARCHI_CACHE[session_id]['file_path']
        parsed = parse_excel_for_parchi(file_path, sheet_name=sheet_name)
        PARCHI_CACHE[session_id]['parsed'] = parsed
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': parsed['voters'][:12]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/parchi/from_session', methods=['POST'])
def api_parchi_from_session():
    try:
        data = request.get_json() or {}
        feature1_session_id = data.get('session_id')
        
        # If no session_id provided, take the latest one from SESSION_CACHE
        if not feature1_session_id and SESSION_CACHE:
            feature1_session_id = list(SESSION_CACHE.keys())[-1]
            
        if not feature1_session_id or feature1_session_id not in SESSION_CACHE:
            return jsonify({'error': 'फ़ीचर 1 से कोई मतदाता डेटा उपलब्ध नहीं है। कृपया पहले पीडीएफ प्रोसेस करें या एक्सेल अपलोड करें।'}), 404
            
        res = SESSION_CACHE[feature1_session_id]['result']
        voters = []
        
        for w in res.get('wards', []):
            grid = w.get('grid', [])
            if len(grid) > 1:
                headers = grid[0]
                col_idx = {h: idx for idx, h in enumerate(headers)}
                for row in grid[1:]:
                    name = row[col_idx.get('नाम', 2)] if 'नाम' in col_idx else ''
                    if not name:
                        continue
                    serial = row[col_idx.get('क्रम संख्या', 1)] if 'क्रम संख्या' in col_idx else len(voters) + 1
                    part_no = (
                        row[col_idx['वार्ड नं.']] if 'वार्ड नं.' in col_idx else
                        row[col_idx['वार्ड संख्या']] if 'वार्ड संख्या' in col_idx else
                        row[col_idx.get('भाग संख्या', 0)] if 'भाग संख्या' in col_idx else '1'
                    )
                    jp = (
                        row[col_idx['जि. प.']] if 'जि. प.' in col_idx else
                        w.get('jila_parishad', '')
                    )
                    ps = (
                        row[col_idx['पं. स.']] if 'पं. स.' in col_idx else
                        w.get('panchayat_samiti', '')
                    )
                    rel = row[col_idx.get('पिता/पति का नाम', 3)] if 'पिता/पति का नाम' in col_idx else ''
                    age = row[col_idx.get('आयु', 4)] if 'आयु' in col_idx else ''
                    epic = row[col_idx.get('वोटर ID', 6)] if 'वोटर ID' in col_idx else ''
                    house = row[col_idx.get('हाउस नंबर', 7)] if 'हाउस नंबर' in col_idx else '00'
                    booth = row[col_idx.get('बूथ का पता', 9)] if 'बूथ का पता' in col_idx else ''
                    status = row[col_idx.get('Status', 10)] if 'Status' in col_idx else 'Active'
                    
                    voters.append({
                        'serial': serial,
                        'part_no': part_no,
                        'jila_parishad': jp,
                        'panchayat_samiti': ps,
                        'name': name,
                        'relative_name': rel,
                        'age': age,
                        'gender': predict_gender(name, rel),
                        'epic': epic,
                        'house': house if house else '00',
                        'booth_address': booth,
                        'status': status
                    })
                    
        parchi_session_id = str(uuid.uuid4())
        default_booth = voters[0].get('booth_address', '') if voters else ""
        
        # Detect clean Hindi Gram Panchayat name
        src_name = SESSION_CACHE.get(feature1_session_id, {}).get('file_name', '')
        p_name_cand = res.get('panchayat_name', '')
        
        from pdf_engine.parchi_generator import ENGLISH_TO_HINDI_PANCHAYAT
        hindi_panchayat = ""
        if p_name_cand:
            p_upper = p_name_cand.strip().upper()
            if p_upper in ENGLISH_TO_HINDI_PANCHAYAT:
                hindi_panchayat = ENGLISH_TO_HINDI_PANCHAYAT[p_upper]
            elif any(c in p_name_cand for c in 'अआइईउऊएऐओऔकखगघचछजझटठडढतथदधनपफबभमयरलवशषसह'):
                hindi_panchayat = p_name_cand.strip()

        if not hindi_panchayat and default_booth:
            m_b = re.search(r'(?:विद्यालय|वि\.|स्कूल|भवन|केंन्द|केंद्र)\s*([A-Za-z\u0900-\u097F]+)', default_booth)
            if m_b:
                c = m_b.group(1).strip()
                if c not in ['नवीन', 'पुराना', 'कमरा', 'क', 'कक्ष', 'प्राइमरी', 'उच्च', 'माध्यमिक', 'संख्या', 'न', 'नं']:
                    hindi_panchayat = c

        if not hindi_panchayat and src_name:
            for eng, hin in ENGLISH_TO_HINDI_PANCHAYAT.items():
                if eng.lower() in src_name.lower():
                    hindi_panchayat = hin
                    break

        default_panchayat = f"ग्राम पंचायत {hindi_panchayat}" if hindi_panchayat else "ग्राम पंचायत"
        default_jp = next((str(v['jila_parishad']).strip() for v in voters if v.get('jila_parishad')), "")
        default_ps = next((str(v['panchayat_samiti']).strip() for v in voters if v.get('panchayat_samiti')), "")
        
        parsed = {
            'sheet_names': ['फ़ीचर 1 मतदाता सूची'],
            'active_sheet': 'फ़ीचर 1 मतदाता सूची',
            'voters': voters,
            'total_voters': len(voters),
            'active_count': len([v for v in voters if v['status'] == 'Active']),
            'deleted_count': len([v for v in voters if v['status'] == 'Deleted']),
            'default_panchayat_name': default_panchayat,
            'default_booth_address': default_booth,
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps
        }
        
        PARCHI_CACHE[parchi_session_id] = {
            'filename': 'फ़ीचर_1_मतदाता_सूची.xlsx',
            'parsed': parsed
        }

        src_name = SESSION_CACHE.get(feature1_session_id, {}).get('file_name', 'वोटर_लिस्ट.xlsx')
        log_activity(
            feature_id='feature2',
            feature_name='फ़ीचर 2: Plain मतदाता पर्ची',
            file_names=[src_name],
            action='फ़ीचर 1 से पर्ची डेटा लोड',
            details=f"{parsed['total_voters']} मतदाता"
        )
        
        return jsonify({
            'session_id': parchi_session_id,
            'filename': 'फ़ीचर 1 मतदाता सूची',
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': default_panchayat,
            'default_booth_address': default_booth,
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps,
            'voters_sample': voters[:12]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/parchi/preview_html', methods=['POST'])
def api_parchi_preview_html():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        voting_time = data.get('voting_time', 'प्रातः 7 से सायं 6 बजे तक')
        filter_active_only = data.get('filter_active_only', True)
        layout_type = str(data.get('layout_type', '12'))
        show_page_number = data.get('show_page_number', True)
        if isinstance(show_page_number, str):
            show_page_number = (show_page_number.lower() in ['true', '1', 'yes'])
        range_from = data.get('range_from')
        range_to = data.get('range_to')
        page = int(data.get('page', 1))
        
        if not session_id or session_id not in PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया पुनः फ़ाइल चुनें।'}), 404
            
        voters = PARCHI_CACHE[session_id]['parsed']['voters']
        panchayat_name = data.get('panchayat_name')
        if not panchayat_name or not str(panchayat_name).strip():
            panchayat_name = PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_name') or 'ग्राम पंचायत'
        
        # Override booth address if user customized it
        custom_booth = data.get('booth_address')
        if custom_booth:
            for v in voters:
                if not v.get('booth_address'):
                    v['booth_address'] = custom_booth
                    
        jp = data.get('jila_parishad') or PARCHI_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')

        html, total_pages, total_voters = generate_parchi_html(
            voters=voters,
            panchayat_name=panchayat_name,
            voting_time=voting_time,
            filter_active_only=filter_active_only,
            layout_type=layout_type,
            show_page_number=show_page_number,
            range_from=int(range_from) if range_from else None,
            range_to=int(range_to) if range_to else None,
            for_preview=True,
            preview_page=page,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        return jsonify({
            'html': html,
            'total_pages': total_pages,
            'current_page': page,
            'total_voters': total_voters
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/parchi/export_pdf', methods=['POST'])
def api_parchi_export_pdf():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        if not session_id or session_id not in PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया पुनः फ़ाइल चुनें।'}), 404

        panchayat_name = data.get('panchayat_name')
        if not panchayat_name or not str(panchayat_name).strip():
            panchayat_name = PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_name') or 'ग्राम पंचायत'
        voting_time = data.get('voting_time', 'प्रातः 7 से सायं 6 बजे तक')
        filter_active_only = data.get('filter_active_only', True)
        layout_type = str(data.get('layout_type', '12'))
        show_page_number = data.get('show_page_number', True)
        if isinstance(show_page_number, str):
            show_page_number = (show_page_number.lower() in ['true', '1', 'yes'])
        range_from = data.get('range_from')
        range_to = data.get('range_to')
        
        if not session_id or session_id not in PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया पुनः फ़ाइल चुनें।'}), 404
            
        voters = PARCHI_CACHE[session_id]['parsed']['voters']
        
        # Custom booth if provided
        custom_booth = data.get('booth_address')
        if custom_booth:
            for v in voters:
                if not v.get('booth_address'):
                    v['booth_address'] = custom_booth
                    
        jp = data.get('jila_parishad') or PARCHI_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')

        pdf_bytes = generate_parchi_pdf(
            voters=voters,
            panchayat_name=panchayat_name,
            voting_time=voting_time,
            filter_active_only=filter_active_only,
            layout_type=layout_type,
            show_page_number=show_page_number,
            range_from=int(range_from) if range_from else None,
            range_to=int(range_to) if range_to else None,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        clean_title = re.sub(r'[^\w\s-]', '', panchayat_name).strip() or "Voter_Parchi"
        file_name = f"Geam_Digital_Matdata_Parchi_{clean_title}_{layout_type}_slips.pdf"

        log_activity(
            feature_id='feature2',
            feature_name='फ़ीचर 2: Plain मतदाता पर्ची',
            file_names=[file_name],
            action='12-पर्ची Plain PDF जनरेट',
            details=f"{len(voters)} मतदाता • {layout_type}-पर्ची लेआउट"
        )
        
        return send_file(
            BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/parchi/print', methods=['GET', 'POST'])
def parchi_print_standalone():
    try:
        session_id = request.args.get('session_id') or (request.form.get('session_id') if request.method == 'POST' else None)
        if not session_id or session_id not in PARCHI_CACHE:
            return "<h3>सत्र समाप्त हो गया है। कृपया मुख्य पोर्टल से पुनः प्रिंट बटन दबाएं।</h3>", 404
            
        voters = PARCHI_CACHE[session_id]['parsed']['voters']
        panchayat_name = request.args.get('panchayat_name')
        if not panchayat_name or not str(panchayat_name).strip():
            panchayat_name = PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_name') or 'ग्राम पंचायत'
        voting_time = request.args.get('voting_time') or 'प्रातः 7 से सायं 6 बजे तक'
        filter_active_only = (request.args.get('filter_active_only', 'true').lower() == 'true')
        layout_type = request.args.get('layout_type', '12')
        show_page_str = request.args.get('show_page_number', 'true')
        show_page_number = (show_page_str.lower() in ['true', '1', 'yes'])
        range_from = request.args.get('range_from')
        range_to = request.args.get('range_to')
        
        custom_booth = request.args.get('booth_address')
        if custom_booth:
            for v in voters:
                if not v.get('booth_address'):
                    v['booth_address'] = custom_booth

        jp = request.args.get('jila_parishad') or PARCHI_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = request.args.get('panchayat_samiti') or PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')

        html, total_pages, total_voters = generate_parchi_html(
            voters=voters,
            panchayat_name=panchayat_name,
            voting_time=voting_time,
            filter_active_only=filter_active_only,
            layout_type=layout_type,
            show_page_number=show_page_number,
            range_from=int(range_from) if range_from else None,
            range_to=int(range_to) if range_to else None,
            for_preview=False,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        auto_print_script = """
        <script>
            window.addEventListener('DOMContentLoaded', () => {
                setTimeout(() => {
                    window.print();
                }, 500);
            });
        </script>
        """
        html_with_print = html.replace('</body>', f'{auto_print_script}</body>')
        return html_with_print
    except Exception as e:
        return f"<h3>त्रुटि: {str(e)}</h3>", 500

# ==============================================================
# FEATURE 3: COLOR VOTER SLIP (CANDIDATE POSTER) ENDPOINTS
# ==============================================================

@app.route('/api/color_parchi/upload_excel', methods=['POST'])
def api_color_parchi_upload_excel():
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'कृपया एक वैध एक्सेल फ़ाइल (.xlsx / .xls) चुनें'}), 400
        
        f = request.files['excel_file']
        if not f or not f.filename:
            return jsonify({'error': 'फ़ाइल उपलब्ध नहीं है'}), 400
            
        filename = secure_filename(f.filename)
        session_id = str(uuid.uuid4())
        save_path = os.path.join(UPLOAD_FOLDER, f"color_parchi_{session_id}_{filename}")
        f.save(save_path)
        
        parsed = parse_excel_for_parchi(save_path)
        
        COLOR_PARCHI_CACHE[session_id] = {
            'file_path': save_path,
            'filename': filename,
            'parsed': parsed
        }

        log_activity(
            feature_id='feature3',
            feature_name='फ़ीचर 3: Color Voter Slip',
            file_names=[filename],
            action='कलर पर्ची हेतु एक्सेल अपलोड',
            details=f"{parsed['total_voters']} मतदाता"
        )
        
        return jsonify({
            'session_id': session_id,
            'filename': filename,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': parsed['voters'][:8]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/color_parchi/select_sheet', methods=['POST'])
def api_color_parchi_select_sheet():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        sheet_name = data.get('sheet_name')
        
        if not session_id or session_id not in COLOR_PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः अपलोड करें।'}), 404
            
        file_path = COLOR_PARCHI_CACHE[session_id]['file_path']
        parsed = parse_excel_for_parchi(file_path, sheet_name=sheet_name)
        COLOR_PARCHI_CACHE[session_id]['parsed'] = parsed
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': parsed['voters'][:8]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/color_parchi/from_session', methods=['POST'])
def api_color_parchi_from_session():
    try:
        data = request.get_json() or {}
        feature1_session_id = data.get('session_id')
        
        if not feature1_session_id and SESSION_CACHE:
            feature1_session_id = list(SESSION_CACHE.keys())[-1]
            
        if not feature1_session_id or feature1_session_id not in SESSION_CACHE:
            return jsonify({'error': 'फ़ीचर 1 से कोई मतदाता डेटा उपलब्ध नहीं है। कृपया पहले पीडीएफ प्रोसेस करें या एक्सेल अपलोड करें।'}), 404
            
        res = SESSION_CACHE[feature1_session_id]['result']
        voters = []
        
        for w in res.get('wards', []):
            grid = w.get('grid', [])
            if len(grid) > 1:
                headers = grid[0]
                col_idx = {h: idx for idx, h in enumerate(headers)}
                for row in grid[1:]:
                    name = row[col_idx.get('नाम', 2)] if 'नाम' in col_idx else ''
                    if not name:
                        continue
                    serial = row[col_idx.get('क्रम संख्या', 1)] if 'क्रम संख्या' in col_idx else len(voters) + 1
                    part_no = (
                        row[col_idx['वार्ड नं.']] if 'वार्ड नं.' in col_idx else
                        row[col_idx['वार्ड संख्या']] if 'वार्ड संख्या' in col_idx else
                        row[col_idx.get('भाग संख्या', 0)] if 'भाग संख्या' in col_idx else '1'
                    )
                    jp = (
                        row[col_idx['जि. प.']] if 'जि. प.' in col_idx else
                        w.get('jila_parishad', '')
                    )
                    ps = (
                        row[col_idx['पं. स.']] if 'पं. स.' in col_idx else
                        w.get('panchayat_samiti', '')
                    )
                    rel = row[col_idx.get('पिता/पति का नाम', 3)] if 'पिता/पति का नाम' in col_idx else ''
                    age = row[col_idx.get('आयु', 4)] if 'आयु' in col_idx else ''
                    epic = row[col_idx.get('वोटर ID', 6)] if 'वोटर ID' in col_idx else ''
                    house = row[col_idx.get('हाउस नंबर', 7)] if 'हाउस नंबर' in col_idx else '1'
                    booth = row[col_idx.get('बूथ का पता', 9)] if 'बूथ का पता' in col_idx else ''
                    status = row[col_idx.get('Status', 10)] if 'Status' in col_idx else 'Active'
                    
                    voters.append({
                        'serial': serial,
                        'part_no': part_no,
                        'ward': part_no,
                        'jila_parishad': jp,
                        'panchayat_samiti': ps,
                        'name': name,
                        'relative_name': rel,
                        'age': age,
                        'gender': predict_gender(name, rel),
                        'epic': epic,
                        'house': house if house else '1',
                        'booth_address': booth,
                        'status': status
                    })
                    
        color_session_id = str(uuid.uuid4())
        default_panchayat = f"वार्ड नं. {voters[0]['part_no']}" if voters else "वार्ड नं. 3"
        default_booth = voters[0].get('booth_address', '') if voters else ""
        default_jp = next((str(v['jila_parishad']).strip() for v in voters if v.get('jila_parishad')), "")
        default_ps = next((str(v['panchayat_samiti']).strip() for v in voters if v.get('panchayat_samiti')), "")
        
        parsed = {
            'sheet_names': ['फ़ीचर 1 मतदाता सूची'],
            'active_sheet': 'फ़ीचर 1 मतदाता सूची',
            'voters': voters,
            'total_voters': len(voters),
            'active_count': len([v for v in voters if v['status'] == 'Active']),
            'deleted_count': len([v for v in voters if v['status'] == 'Deleted']),
            'default_panchayat_name': default_panchayat,
            'default_booth_address': default_booth,
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps
        }
        
        COLOR_PARCHI_CACHE[color_session_id] = {
            'filename': 'फ़ीचर_1_मतदाता_सूची.xlsx',
            'parsed': parsed
        }

        src_name = SESSION_CACHE.get(feature1_session_id, {}).get('file_name', 'वोटर_लिस्ट.xlsx')
        log_activity(
            feature_id='feature3',
            feature_name='फ़ीचर 3: Color Voter Slip',
            file_names=[src_name],
            action='फ़ीचर 1 से कलर पर्ची लोड',
            details=f"{len(voters)} मतदाता"
        )
        
        return jsonify({
            'session_id': color_session_id,
            'filename': 'फ़ीचर 1 मतदाता सूची',
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': default_panchayat,
            'default_booth_address': default_booth,
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps,
            'voters_sample': voters[:8]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/color_parchi/defaults', methods=['GET'])
def api_color_parchi_defaults():
    from pdf_engine.color_parchi_generator import DEFAULT_HAND_DATA, DEFAULT_LOTUS_DATA, DEFAULT_CYCLE_DATA, DEFAULT_CANDIDATE_DATA
    return jsonify({
        'hand': DEFAULT_HAND_DATA,
        'lotus': DEFAULT_LOTUS_DATA,
        'cycle': DEFAULT_CYCLE_DATA,
        'candidate_photo': DEFAULT_CANDIDATE_DATA
    })

@app.route('/api/color_parchi/upload_asset', methods=['POST'])
def api_color_parchi_upload_asset():
    """Converts uploaded image (poster, photo, or symbol) into a base64 Data URL."""
    try:
        if 'image_file' not in request.files:
            return jsonify({'error': 'कृपया एक वैध इमेज फ़ाइल चुनें'}), 400
        f = request.files['image_file']
        if not f or not f.filename:
            return jsonify({'error': 'फ़ाइल उपलब्ध नहीं है'}), 400
            
        mimetype = f.mimetype or 'image/jpeg'
        img_bytes = f.read()
        encoded = base64.b64encode(img_bytes).decode('ascii')
        data_url = f"data:{mimetype};base64,{encoded}"

        log_activity(
            feature_id='feature3',
            feature_name='फ़ीचर 3: Color Voter Slip',
            file_names=[f.filename],
            action='प्रत्याशी पोस्टर / लोगो इमेज अपलोड व एडिट',
            details=f"फ़ाइल: {f.filename}"
        )
        
        return jsonify({
            'success': True,
            'filename': secure_filename(f.filename),
            'data_url': data_url
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/color_parchi/preview_html', methods=['POST'])
def api_color_parchi_preview_html():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        candidate_info = data.get('candidate_info', {})
        layout_type = str(data.get('layout_type', '8'))
        show_page_number = data.get('show_page_number', True)
        if isinstance(show_page_number, str):
            show_page_number = (show_page_number.lower() in ['true', '1', 'yes'])
        filter_active_only = data.get('filter_active_only', True)
        page = int(data.get('page', 1))
        
        if not session_id or session_id not in COLOR_PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया पुनः फ़ाइल चुनें।'}), 404
            
        voters = COLOR_PARCHI_CACHE[session_id]['parsed']['voters']
        candidate_info['jila_parishad'] = candidate_info.get('jila_parishad') or data.get('jila_parishad') or COLOR_PARCHI_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        candidate_info['panchayat_samiti'] = candidate_info.get('panchayat_samiti') or data.get('panchayat_samiti') or COLOR_PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        COLOR_PARCHI_CACHE[session_id]['custom_candidate_info'] = candidate_info
        
        html, total_pages, total_voters = generate_color_parchi_html(
            voters=voters,
            candidate_info=candidate_info,
            layout_type=layout_type,
            show_page_number=show_page_number,
            filter_active_only=filter_active_only,
            for_preview=True,
            preview_page=page
        )
        
        return jsonify({
            'html': html,
            'total_pages': total_pages,
            'current_page': page,
            'total_voters': total_voters
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/color_parchi/export_pdf', methods=['POST'])
def api_color_parchi_export_pdf():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        candidate_info = data.get('candidate_info', {})
        layout_type = str(data.get('layout_type', '8'))
        show_page_number = data.get('show_page_number', True)
        if isinstance(show_page_number, str):
            show_page_number = (show_page_number.lower() in ['true', '1', 'yes'])
        filter_active_only = data.get('filter_active_only', True)
        
        if not session_id or session_id not in COLOR_PARCHI_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया पुनः फ़ाइल चुनें।'}), 404
            
        voters = COLOR_PARCHI_CACHE[session_id]['parsed']['voters']
        candidate_info['jila_parishad'] = candidate_info.get('jila_parishad') or data.get('jila_parishad') or COLOR_PARCHI_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        candidate_info['panchayat_samiti'] = candidate_info.get('panchayat_samiti') or data.get('panchayat_samiti') or COLOR_PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        COLOR_PARCHI_CACHE[session_id]['custom_candidate_info'] = candidate_info
        
        pdf_bytes = generate_color_parchi_pdf(
            voters=voters,
            candidate_info=candidate_info,
            layout_type=layout_type,
            show_page_number=show_page_number,
            filter_active_only=filter_active_only
        )
        
        cand_name = re.sub(r'[^\w\s-]', '', candidate_info.get('candidate_name', 'Pratyashi')).strip() or "Pratyashi"
        file_name = f"Geam_Digital_Color_Parchi_{cand_name}_A4_{layout_type}.pdf"

        log_activity(
            feature_id='feature3',
            feature_name='फ़ीचर 3: Color Voter Slip',
            file_names=[file_name],
            action='कलर पर्ची (पोस्टर सहित) PDF जनरेट',
            details=f"{len(voters)} पर्चियां • लेआउट: {layout_type} Slips"
        )
        
        return send_file(
            BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/color_parchi/print', methods=['GET', 'POST'])
def color_parchi_print_standalone():
    try:
        session_id = request.args.get('session_id') or (request.form.get('session_id') if request.method == 'POST' else None)
        layout_type = request.args.get('layout_type', '8')
        show_page_str = request.args.get('show_page_number', 'true')
        show_page_number = (show_page_str.lower() in ['true', '1', 'yes'])
        filter_active_only = (request.args.get('filter_active_only', 'true').lower() == 'true')
        
        if not session_id or session_id not in COLOR_PARCHI_CACHE:
            return "<h3>सत्र समाप्त हो गया है। कृपया मुख्य पोर्टल से पुनः प्रिंट बटन दबाएं।</h3>", 404
            
        voters = COLOR_PARCHI_CACHE[session_id]['parsed']['voters']
        
        candidate_info = {
            'candidate_name': request.args.get('candidate_name', 'संजय कुमार मेवाड़ा'),
            'header_prefix': request.args.get('header_prefix', 'पार्षद पद हेतु वार्ड नं'),
            'ward_badge': request.args.get('ward_badge', '03'),
            'header_suffix': request.args.get('header_suffix', 'से लोकप्रिय प्रत्याशी'),
            'button_no': request.args.get('button_no', '2'),
            'slogan': request.args.get('slogan', 'को हाथ के निशान पर बटन दबा कर विजयी बनावें।'),
            'voting_time': request.args.get('voting_time', 'प्रातः 7 से सायं 6 बजे तक'),
            'time_color': request.args.get('time_color', '#e11d48'),
            'poster_bg_color': request.args.get('poster_bg_color', '#fff033'),
            'candidate_name_color': request.args.get('candidate_name_color', '#b91c1c'),
            'btn_box_color': request.args.get('btn_box_color', '#0284c7'),
            'symbol_preset': request.args.get('symbol_preset', 'hand'),
            'jila_parishad': request.args.get('jila_parishad') or COLOR_PARCHI_CACHE[session_id]['parsed'].get('default_jila_parishad', ''),
            'panchayat_samiti': request.args.get('panchayat_samiti') or COLOR_PARCHI_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        }
        
        if 'custom_candidate_info' in COLOR_PARCHI_CACHE[session_id]:
            cached_info = COLOR_PARCHI_CACHE[session_id]['custom_candidate_info']
            for k, val in cached_info.items():
                if val:
                    candidate_info[k] = val
        if request.args.get('jila_parishad'):
            candidate_info['jila_parishad'] = request.args.get('jila_parishad')
        if request.args.get('panchayat_samiti'):
            candidate_info['panchayat_samiti'] = request.args.get('panchayat_samiti')

        html, total_pages, total_voters = generate_color_parchi_html(
            voters=voters,
            candidate_info=candidate_info,
            layout_type=layout_type,
            show_page_number=show_page_number,
            filter_active_only=filter_active_only,
            for_preview=False
        )
        
        auto_print_script = """
        <script>
            window.addEventListener('DOMContentLoaded', () => {
                setTimeout(() => {
                    window.print();
                }, 500);
            });
        </script>
        """
        html_with_print = html.replace('</body>', f'{auto_print_script}</body>')
        return html_with_print
    except Exception as e:
        return f"<h3>त्रुटि: {str(e)}</h3>", 500

# ==============================================================
# FEATURE 4: ALPHABETICAL (ABCD) VOTER LIST CONTROLLERS
# ==============================================================

@app.route('/api/alphabetical/from_session', methods=['POST'])
def api_alpha_from_session():
    try:
        data = request.get_json() or {}
        feature1_session_id = data.get('session_id')
        
        if not feature1_session_id and SESSION_CACHE:
            feature1_session_id = list(SESSION_CACHE.keys())[-1]
            
        if not feature1_session_id or feature1_session_id not in SESSION_CACHE:
            return jsonify({'error': 'फ़ीचर 1 से कोई सत्र डेटा उपलब्ध नहीं है। कृपया पहले फ़ाइल प्रोसेस करें या एक्सेल अपलोड करें।'}), 404
            
        res = SESSION_CACHE[feature1_session_id]['result']
        voters = []
        
        for w in res.get('wards', []):
            grid = w.get('grid', [])
            if len(grid) > 1:
                headers = grid[0]
                col_idx = {h: idx for idx, h in enumerate(headers)}
                for row in grid[1:]:
                    name = row[col_idx.get('नाम', 2)] if 'नाम' in col_idx else ''
                    if not name:
                        continue
                    serial = row[col_idx.get('क्रम संख्या', 1)] if 'क्रम संख्या' in col_idx else len(voters) + 1
                    part_no = (
                        row[col_idx['वार्ड नं.']] if 'वार्ड नं.' in col_idx else
                        row[col_idx['वार्ड संख्या']] if 'वार्ड संख्या' in col_idx else
                        row[col_idx.get('भाग संख्या', 0)] if 'भाग संख्या' in col_idx else '1'
                    )
                    jp = (
                        row[col_idx['जि. प.']] if 'जि. प.' in col_idx else
                        w.get('jila_parishad', '')
                    )
                    ps = (
                        row[col_idx['पं. स.']] if 'पं. स.' in col_idx else
                        w.get('panchayat_samiti', '')
                    )
                    rel = row[col_idx.get('पिता/पति का नाम', 3)] if 'पिता/पति का नाम' in col_idx else ''
                    age = row[col_idx.get('आयु', 4)] if 'आयु' in col_idx else ''
                    epic = row[col_idx.get('वोटर ID', 6)] if 'वोटर ID' in col_idx else ''
                    house = row[col_idx.get('हाउस नंबर', 7)] if 'हाउस नंबर' in col_idx else ''
                    booth = row[col_idx.get('बूथ का पता', 9)] if 'बूथ का पता' in col_idx else ''
                    status = row[col_idx.get('Status', 10)] if 'Status' in col_idx else 'Active'
                    
                    voters.append({
                        'serial': serial,
                        'part_no': part_no,
                        'ward': part_no,
                        'jila_parishad': jp,
                        'panchayat_samiti': ps,
                        'name': name,
                        'relative_name': rel,
                        'age': age,
                        'gender': predict_gender(name, rel),
                        'epic': epic,
                        'house': house,
                        'booth_address': booth,
                        'status': status
                    })
                    
        alpha_session_id = str(uuid.uuid4())
        default_ward = str(voters[0].get('part_no', '8')) if voters else "8"
        default_booth = voters[0].get('booth_address', '') if voters else "28 - राजकीय बालिका उच्च माध्यमिक विद्यालय बापूनगर भीलवाडा कमरा न. 7"
        default_jp = next((str(v['jila_parishad']).strip() for v in voters if v.get('jila_parishad')), "")
        default_ps = next((str(v['panchayat_samiti']).strip() for v in voters if v.get('panchayat_samiti')), "")

        p_name = res.get('panchayat_name', '')
        if p_name and not str(p_name).strip().startswith('ग्राम पंचायत'):
            default_panchayat = f"ग्राम पंचायत {str(p_name).strip()}"
        elif p_name:
            default_panchayat = str(p_name).strip()
        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        default_ward = raw_wards[0] if raw_wards else "1"
        
        parsed = {
            'sheet_names': ['फ़ीचर 1 मतदाता सूची'],
            'active_sheet': 'फ़ीचर 1 मतदाता सूची',
            'voters': voters,
            'wards': raw_wards,
            'total_voters': len(voters),
            'active_count': len([v for v in voters if v['status'] == 'Active']),
            'deleted_count': len([v for v in voters if v['status'] == 'Deleted']),
            'default_panchayat_name': default_panchayat,
            'default_part_no': default_ward,
            'default_booth_address': default_booth,
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps
        }
        
        ALPHA_CACHE[alpha_session_id] = {
            'filename': 'फ़ीचर_1_मतदाता_सूची.xlsx',
            'parsed': parsed
        }

        src_name = SESSION_CACHE.get(feature1_session_id, {}).get('file_name', 'वोटर_लिस्ट.xlsx')
        log_activity(
            feature_id='feature4',
            feature_name='फ़ीचर 4: अल्फाबेटिक वोटर लिस्ट',
            file_names=[src_name],
            action='फ़ीचर 1 से अल्फाबेटिक लिस्ट लोड',
            details=f"{len(voters)} मतदाता"
        )
        
        return jsonify({
            'session_id': alpha_session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': parsed['wards'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_part_no': parsed['default_part_no'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps,
            'voters_sample': voters[:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/alphabetical/upload_excel', methods=['POST'])
def api_alpha_upload_excel():
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'कृपया एक वैध एक्सेल फ़ाइल (.xlsx/.xls) चुनें'}), 400
        f = request.files['excel_file']
        if not f or not f.filename:
            return jsonify({'error': 'फ़ाइल उपलब्ध नहीं है'}), 400
            
        file_path = os.path.join(UPLOAD_FOLDER, f"alpha_{uuid.uuid4().hex[:8]}_{secure_filename(f.filename)}")
        f.save(file_path)
        
        parsed = parse_excel_for_parchi(file_path)
        voters = parsed['voters']
        
        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        parsed['wards'] = raw_wards
        
        alpha_session_id = str(uuid.uuid4())
        default_ward = raw_wards[0] if raw_wards else "1"
        
        ALPHA_CACHE[alpha_session_id] = {
            'file_path': file_path,
            'filename': f.filename,
            'parsed': parsed
        }

        log_activity(
            feature_id='feature4',
            feature_name='फ़ीचर 4: अल्फाबेटिक वोटर लिस्ट',
            file_names=[f.filename],
            action='अल्फाबेटिक लिस्ट हेतु एक्सेल अपलोड',
            details=f"{parsed['total_voters']} मतदाता • {len(raw_wards)} वार्ड्स"
        )
        
        return jsonify({
            'session_id': alpha_session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': parsed['wards'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed.get('default_panchayat_name') or "ग्राम पंचायत",
            'default_part_no': parsed.get('default_part_no') or default_ward,
            'default_booth_address': parsed.get('default_booth_address') or "28 - राजकीय बालिका उच्च माध्यमिक विद्यालय बापूनगर भीलवाडा कमरा न. 7",
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': voters[:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/alphabetical/select_sheet', methods=['POST'])
def api_alpha_select_sheet():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        sheet_name = data.get('sheet_name')
        
        if not session_id or session_id not in ALPHA_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        file_path = ALPHA_CACHE[session_id].get('file_path')
        if not file_path:
            return jsonify({'error': 'सत्र फ़ाइल उपलब्ध नहीं है'}), 400
            
        parsed = parse_excel_for_parchi(file_path, sheet_name=sheet_name)
        voters = parsed['voters']
        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        parsed['wards'] = raw_wards
        ALPHA_CACHE[session_id]['parsed'] = parsed
        
        default_ward = raw_wards[0] if raw_wards else "1"
        return jsonify({
            'session_id': session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': parsed['wards'],
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed.get('default_panchayat_name') or "ग्राम पंचायत",
            'default_part_no': default_ward,
            'default_booth_address': parsed.get('default_booth_address') or "28 - राजकीय बालिका उच्च माध्यमिक विद्यालय बापूनगर भीलवाडा कमरा न. 7",
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': voters[:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/alphabetical/preview_html', methods=['POST'])
def api_alpha_preview_html():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'वार्ड नं.- 8')
        part_no = data.get('part_no', '1')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', 'अल्फाबेटिक ABCD से वोटर लिस्ट')
        selected_ward = data.get('selected_ward', 'all')
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
        page = int(data.get('page', 1))
        
        if not session_id or session_id not in ALPHA_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = ALPHA_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or ALPHA_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or ALPHA_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        ALPHA_CACHE[session_id]['config'] = {
            'ward_title': ward_title,
            'part_no': part_no,
            'booth_address': booth_address,
            'list_title': list_title,
            'selected_ward': selected_ward,
            'filter_active_only': filter_active_only,
            'jila_parishad': jp,
            'panchayat_samiti': ps
        }
        
        html, total_pages, total_voters, page_meta = generate_alphabetical_html(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            for_preview=True,
            preview_page=page,
            rows_per_page=31,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        return jsonify({
            'html': html,
            'total_pages': total_pages,
            'total_voters': total_voters,
            'current_page': page,
            'current_ward': page_meta.get('current_ward', '1'),
            'page_in_ward': page_meta.get('page_in_ward', 1),
            'total_in_ward': page_meta.get('total_in_ward', 1),
            'ward_voters': page_meta.get('ward_voters', total_voters),
            'total_wards': page_meta.get('total_wards', 1)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/alphabetical/export_pdf', methods=['POST'])
def api_alpha_export_pdf():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'वार्ड नं.- 8')
        part_no = data.get('part_no', '1')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', 'अल्फाबेटिक ABCD से वोटर लिस्ट')
        selected_ward = data.get('selected_ward', 'all')
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
            
        if not session_id or session_id not in ALPHA_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = ALPHA_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or ALPHA_CACHE[session_id].get('config', {}).get('jila_parishad') or ALPHA_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or ALPHA_CACHE[session_id].get('config', {}).get('panchayat_samiti') or ALPHA_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        pdf_bytes = generate_alphabetical_pdf(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        safe_title = re.sub(r'[^\w\s-]', '', ward_title).strip() or "Alphabetical_List"
        file_name = f"Geam_Digital_{safe_title}_ABCD_Voter_List.pdf"

        log_activity(
            feature_id='feature4',
            feature_name='फ़ीचर 4: अल्फाबेटिक वोटर लिस्ट',
            file_names=[file_name],
            action='A-Z अल्फाबेटिक PDF जनरेट',
            details=f"{len(voters)} मतदाता • {ward_title}"
        )
        
        return send_file(
            BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/alphabetical/export_excel', methods=['POST'])
def api_alpha_export_excel():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'वार्ड नं.- 8')
        part_no = data.get('part_no', '1')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', 'अल्फाबेटिक ABCD से वोटर लिस्ट')
        selected_ward = data.get('selected_ward', 'all')
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
            
        if not session_id or session_id not in ALPHA_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = ALPHA_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or ALPHA_CACHE[session_id].get('config', {}).get('jila_parishad') or ALPHA_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or ALPHA_CACHE[session_id].get('config', {}).get('panchayat_samiti') or ALPHA_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        excel_bytes = generate_alphabetical_excel(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        safe_title = re.sub(r'[^\w\s-]', '', ward_title).strip() or "Alphabetical_List"
        file_name = f"Geam_Digital_{safe_title}_ABCD_Voter_List.xlsx"

        log_activity(
            feature_id='feature4',
            feature_name='फ़ीचर 4: अल्फाबेटिक वोटर लिस्ट',
            file_names=[file_name],
            action='A-Z अल्फाबेटिक एक्सेल जनरेट',
            details=f"{len(voters)} मतदाता • {ward_title}"
        )
        
        return send_file(
            BytesIO(excel_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/alphabetical/print', methods=['GET', 'POST'])
def alpha_print_standalone():
    try:
        session_id = request.args.get('session_id') or (request.form.get('session_id') if request.method == 'POST' else None)
        ward_title = request.args.get('ward_title', 'वार्ड नं.- 8')
        part_no = request.args.get('part_no', '1')
        booth_address = request.args.get('booth_address', '')
        list_title = request.args.get('list_title', 'अल्फाबेटिक ABCD से वोटर लिस्ट')
        selected_ward = request.args.get('selected_ward', 'all')
        filter_active_only = (request.args.get('filter_active_only', 'true').lower() == 'true')
        jp = request.args.get('jila_parishad') or (request.form.get('jila_parishad') if request.method == 'POST' else None)
        ps = request.args.get('panchayat_samiti') or (request.form.get('panchayat_samiti') if request.method == 'POST' else None)
        
        if not session_id or session_id not in ALPHA_CACHE:
            return "<h3>सत्र समाप्त हो गया है। कृपया मुख्य पोर्टल से पुनः प्रिंट बटन दबाएं।</h3>", 404
            
        voters = ALPHA_CACHE[session_id]['parsed']['voters']
        
        if 'config' in ALPHA_CACHE[session_id]:
            c = ALPHA_CACHE[session_id]['config']
            ward_title = c.get('ward_title', ward_title)
            part_no = c.get('part_no', part_no)
            booth_address = c.get('booth_address', booth_address)
            list_title = c.get('list_title', list_title)
            selected_ward = c.get('selected_ward', selected_ward)
            filter_active_only = c.get('filter_active_only', filter_active_only)
            jp = c.get('jila_parishad', jp)
            ps = c.get('panchayat_samiti', ps)

        if not jp:
            jp = ALPHA_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        if not ps:
            ps = ALPHA_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')

        html, total_pages, total_voters, _ = generate_alphabetical_html(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            for_preview=False,
            rows_per_page=31,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        auto_print_script = """
        <script>
            window.addEventListener('DOMContentLoaded', () => {
                setTimeout(() => {
                    window.print();
                }, 500);
            });
        </script>
        """
        html_with_print = html.replace('</body>', f'{auto_print_script}</body>')
        return html_with_print
    except Exception as e:
        return f"<h3>त्रुटि: {str(e)}</h3>", 500

# ==============================================================
# FEATURE 5: बड़े-घर / बड़े परिवारों की लिस्ट ENDPOINTS
# ==============================================================

@app.route('/api/family/from_session', methods=['POST'])
def api_family_from_session():
    try:
        data = request.get_json() or {}
        feature1_session_id = data.get('session_id')
        
        if not feature1_session_id and SESSION_CACHE:
            feature1_session_id = list(SESSION_CACHE.keys())[-1]
            
        if not feature1_session_id or feature1_session_id not in SESSION_CACHE:
            return jsonify({'error': 'फ़ीचर 1 से कोई मतदाता डेटा उपलब्ध नहीं है। कृपया पहले पीडीएफ प्रोसेस करें या एक्सेल अपलोड करें।'}), 404
            
        res = SESSION_CACHE[feature1_session_id]['result']
        voters = []
        
        for w in res.get('wards', []):
            grid = w.get('grid', [])
            if len(grid) > 1:
                headers = grid[0]
                col_idx = {}
                for idx, h in enumerate(headers):
                    hl = str(h).strip().lower()
                    if 'क्रम' in hl or 'serial' in hl: col_idx['serial'] = idx
                    elif 'जि' in hl and 'प' in hl: col_idx['jila_parishad'] = idx
                    elif 'पं' in hl and 'स' in hl: col_idx['panchayat_samiti'] = idx
                    elif 'भाग' in hl or 'वार्ड' in hl or 'ward' in hl: col_idx['part_no'] = idx
                    elif 'नाम' in hl and 'पिता' not in hl and 'पति' not in hl and 'बूथ' not in hl: col_idx['name'] = idx
                    elif 'पिता' in hl or 'पति' in hl or 'संबंधी' in hl: col_idx['relative_name'] = idx
                    elif 'आयु' in hl or 'उम्र' in hl or 'age' in hl: col_idx['age'] = idx
                    elif 'लिंग' in hl or 'gender' in hl or 'sex' in hl: col_idx['gender'] = idx
                    elif 'पहचान' in hl or 'epic' in hl or 'वोटर' in hl: col_idx['epic'] = idx
                    elif 'मकान' in hl or 'हाउस' in hl or 'house' in hl: col_idx['house'] = idx
                    elif 'बूथ' in hl or 'केंद्र' in hl or 'address' in hl: col_idx['booth_address'] = idx
                    elif 'status' in hl or 'स्थिति' in hl: col_idx['status'] = idx

                for row in grid[1:]:
                    if not any(row): continue
                    def gv(k):
                        idx = col_idx.get(k)
                        return clean_val_str(row[idx]) if idx is not None and idx < len(row) else ''
                    
                    voters.append({
                        'serial': gv('serial') or str(len(voters) + 1),
                        'jila_parishad': gv('jila_parishad') or str(w.get('jila_parishad', res.get('jila_parishad', ''))),
                        'panchayat_samiti': gv('panchayat_samiti') or str(w.get('panchayat_samiti', res.get('panchayat_samiti', ''))),
                        'part_no': gv('part_no') or str(w.get('ward_no', '1')),
                        'name': gv('name'),
                        'relative_name': gv('relative_name'),
                        'age': gv('age'),
                        'gender': gv('gender'),
                        'epic': gv('epic'),
                        'house': gv('house'),
                        'booth_address': gv('booth_address') or w.get('polling_station', ''),
                        'status': gv('status') or 'Active'
                    })

        if not voters:
            return jsonify({'error': 'वर्तमान सत्र में कोई मतदाता रिकॉर्ड नहीं मिले।'}), 400

        session_id = str(uuid.uuid4())
        p_name = res.get('panchayat_name', '')
        if p_name and not str(p_name).strip().startswith('ग्राम पंचायत'):
            default_panchayat = f"ग्राम पंचायत {str(p_name).strip()}"
        elif p_name:
            default_panchayat = str(p_name).strip()
        else:
            default_panchayat = "ग्राम पंचायत"
        default_part = str(voters[0].get('part_no', '1')) if voters else '1'
        default_booth = res.get('polling_station', '')
        if not default_booth and voters:
            default_booth = voters[0].get('booth_address', '')
            
        default_jp = res.get('jila_parishad', '')
        default_ps = res.get('panchayat_samiti', '')
        if not default_jp and voters:
            default_jp = voters[0].get('jila_parishad', '')
        if not default_ps and voters:
            default_ps = voters[0].get('panchayat_samiti', '')

        FAMILY_CACHE[session_id] = {
            'file_name': 'Session_Voters.xlsx',
            'parsed': {
                'sheet_names': ['वर्तमान_सत्र_डेटा'],
                'active_sheet': 'वर्तमान_सत्र_डेटा',
                'voters': voters,
                'total_voters': len(voters),
                'default_panchayat_name': default_panchayat,
                'default_part_no': default_part,
                'default_booth_address': default_booth,
                'default_jila_parishad': default_jp,
                'default_panchayat_samiti': default_ps
            }
        }

        src_name = SESSION_CACHE.get(feature1_session_id, {}).get('file_name', 'वोटर_लिस्ट.xlsx')
        log_activity(
            feature_id='feature5',
            feature_name='फ़ीचर 5: बड़े परिवारों की लिस्ट',
            file_names=[src_name],
            action='फ़ीचर 1 से परिवार लिस्ट लोड',
            details=f"{len(voters)} मतदाता"
        )

        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': ['वर्तमान_सत्र_डेटा'],
            'active_sheet': 'वर्तमान_सत्र_डेटा',
            'wards': raw_wards,
            'total_voters': len(voters),
            'default_panchayat_name': default_panchayat,
            'default_part_no': default_part,
            'default_booth_address': default_booth,
            'default_jila_parishad': default_jp,
            'default_panchayat_samiti': default_ps,
            'voters_sample': voters[:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/family/upload_excel', methods=['POST'])
def api_family_upload_excel():
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'कृपया एक वैध एक्सेल फ़ाइल (.xlsx / .xls) चुनें'}), 400
            
        file = request.files['excel_file']
        if not file.filename:
            return jsonify({'error': 'कोई फ़ाइल नहीं चुनी गई'}), 400
            
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ['.xlsx', '.xls']:
            return jsonify({'error': 'केवल .xlsx या .xls प्रारूप ही समर्थित हैं'}), 400
            
        session_id = str(uuid.uuid4())
        save_name = f"family_{session_id}_{secure_filename(file.filename)}"
        save_path = os.path.join(UPLOAD_FOLDER, save_name)
        file.save(save_path)
        
        parsed = parse_excel_for_parchi(save_path)
        FAMILY_CACHE[session_id] = {
            'file_path': save_path,
            'file_name': file.filename,
            'parsed': parsed
        }

        log_activity(
            feature_id='feature5',
            feature_name='फ़ीचर 5: बड़े परिवारों की लिस्ट',
            file_names=[file.filename],
            action='परिवार लिस्ट हेतु एक्सेल अपलोड',
            details=f"{parsed['total_voters']} मतदाता"
        )

        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in parsed['voters'] if v.get('part_no'))), key=extract_w_key)
        parsed['wards'] = raw_wards
        
        return jsonify({
            'session_id': session_id,
            'file_name': file.filename,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': raw_wards,
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed.get('default_panchayat_name') or "ग्राम पंचायत",
            'default_part_no': parsed.get('default_part_no') or (raw_wards[0] if raw_wards else "1"),
            'default_booth_address': parsed.get('default_booth_address') or "",
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': parsed['voters'][:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/family/select_sheet', methods=['POST'])
def api_family_select_sheet():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        sheet_name = data.get('sheet_name')
        
        if not session_id or session_id not in FAMILY_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः अपलोड करें।'}), 404
            
        file_path = FAMILY_CACHE[session_id]['file_path']
        parsed = parse_excel_for_parchi(file_path, sheet_name=sheet_name)
        FAMILY_CACHE[session_id]['parsed'] = parsed

        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in parsed['voters'] if v.get('part_no'))), key=extract_w_key)
        parsed['wards'] = raw_wards
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': raw_wards,
            'total_voters': parsed['total_voters'],
            'active_count': parsed['active_count'],
            'deleted_count': parsed['deleted_count'],
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': parsed['voters'][:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/family/preview_html', methods=['POST'])
def api_family_preview_html():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'वार्ड न.- 3')
        part_no = data.get('part_no', '1')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', 'बड़े-घर की लिस्ट')
        selected_ward = data.get('selected_ward', 'all')
        min_family_size = int(data.get('min_family_size', 2))
        ignore_zero = data.get('ignore_zero_houses', True)
        if isinstance(ignore_zero, str):
            ignore_zero = (ignore_zero.lower() in ['true', '1', 'yes'])
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
        page = int(data.get('page', 1))
        
        if not session_id or session_id not in FAMILY_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = FAMILY_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or FAMILY_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or FAMILY_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        FAMILY_CACHE[session_id]['config'] = {
            'ward_title': ward_title,
            'part_no': part_no,
            'booth_address': booth_address,
            'list_title': list_title,
            'selected_ward': selected_ward,
            'min_family_size': min_family_size,
            'ignore_zero_houses': ignore_zero,
            'filter_active_only': filter_active_only,
            'jila_parishad': jp,
            'panchayat_samiti': ps
        }
        
        html, total_pages, total_voters, total_families, page_meta = generate_family_html(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            min_family_size=min_family_size,
            ignore_zero_houses=ignore_zero,
            selected_ward=selected_ward,
            filter_active_only=filter_active_only,
            for_preview=True,
            preview_page=page,
            max_lines_per_page=27,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        return jsonify({
            'html': html,
            'total_pages': total_pages,
            'total_voters': total_voters,
            'total_families': total_families,
            'current_page': page,
            'current_ward': page_meta.get('current_ward', '1'),
            'page_in_ward': page_meta.get('page_in_ward', 1),
            'total_in_ward': page_meta.get('total_in_ward', 1),
            'ward_voters': page_meta.get('ward_voters', total_voters),
            'ward_families': page_meta.get('ward_families', total_families),
            'total_wards': page_meta.get('total_wards', 1)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/family/export_pdf', methods=['POST'])
def api_family_export_pdf():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'वार्ड न.- 3')
        part_no = data.get('part_no', '1')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', 'बड़े-घर की लिस्ट')
        selected_ward = data.get('selected_ward', 'all')
        min_family_size = int(data.get('min_family_size', 2))
        ignore_zero = data.get('ignore_zero_houses', True)
        if isinstance(ignore_zero, str):
            ignore_zero = (ignore_zero.lower() in ['true', '1', 'yes'])
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
            
        if not session_id or session_id not in FAMILY_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = FAMILY_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or FAMILY_CACHE[session_id].get('config', {}).get('jila_parishad') or FAMILY_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or FAMILY_CACHE[session_id].get('config', {}).get('panchayat_samiti') or FAMILY_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        pdf_bytes = generate_family_pdf(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            min_family_size=min_family_size,
            ignore_zero_houses=ignore_zero,
            selected_ward=selected_ward,
            filter_active_only=filter_active_only,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        safe_title = re.sub(r'[^\w\s-]', '', ward_title).strip() or "Family_List"
        file_name = f"Geam_Digital_{safe_title}_Bade_Ghar_List.pdf"

        log_activity(
            feature_id='feature5',
            feature_name='फ़ीचर 5: बड़े परिवारों की लिस्ट',
            file_names=[file_name],
            action='बड़े परिवार PDF जनरेट',
            details=f"{len(voters)} मतदाता • {ward_title}"
        )
        
        return send_file(
            BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/family/export_excel', methods=['POST'])
def api_family_export_excel():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'वार्ड न.- 3')
        part_no = data.get('part_no', '1')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', 'बड़े-घर की लिस्ट')
        selected_ward = data.get('selected_ward', 'all')
        min_family_size = int(data.get('min_family_size', 2))
        ignore_zero = data.get('ignore_zero_houses', True)
        if isinstance(ignore_zero, str):
            ignore_zero = (ignore_zero.lower() in ['true', '1', 'yes'])
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
            
        if not session_id or session_id not in FAMILY_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = FAMILY_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or FAMILY_CACHE[session_id].get('config', {}).get('jila_parishad') or FAMILY_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or FAMILY_CACHE[session_id].get('config', {}).get('panchayat_samiti') or FAMILY_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        excel_bytes = generate_family_excel(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            min_family_size=min_family_size,
            ignore_zero_houses=ignore_zero,
            selected_ward=selected_ward,
            filter_active_only=filter_active_only,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        safe_title = re.sub(r'[^\w\s-]', '', ward_title).strip() or "Family_List"
        file_name = f"Geam_Digital_{safe_title}_Bade_Ghar_List.xlsx"

        log_activity(
            feature_id='feature5',
            feature_name='फ़ीचर 5: बड़े परिवारों की लिस्ट',
            file_names=[file_name],
            action='बड़े परिवार एक्सेल जनरेट',
            details=f"{len(voters)} मतदाता • {ward_title}"
        )
        
        return send_file(
            BytesIO(excel_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/family/print', methods=['GET', 'POST'])
def family_print_standalone():
    try:
        session_id = request.args.get('session_id') or (request.form.get('session_id') if request.method == 'POST' else None)
        ward_title = request.args.get('ward_title', 'वार्ड न.- 3')
        part_no = request.args.get('part_no', '1')
        booth_address = request.args.get('booth_address', '')
        list_title = request.args.get('list_title', 'बड़े-घर की लिस्ट')
        selected_ward = request.args.get('selected_ward', 'all')
        min_family_size = int(request.args.get('min_family_size', 2))
        ignore_zero = (request.args.get('ignore_zero_houses', 'true').lower() == 'true')
        filter_active_only = (request.args.get('filter_active_only', 'true').lower() == 'true')
        jp = request.args.get('jila_parishad') or (request.form.get('jila_parishad') if request.method == 'POST' else None)
        ps = request.args.get('panchayat_samiti') or (request.form.get('panchayat_samiti') if request.method == 'POST' else None)
        
        if not session_id or session_id not in FAMILY_CACHE:
            return "<h3>सत्र समाप्त हो गया है। कृपया मुख्य पोर्टल से पुनः प्रिंट बटन दबाएं।</h3>", 404
            
        voters = FAMILY_CACHE[session_id]['parsed']['voters']
        
        if 'config' in FAMILY_CACHE[session_id]:
            c = FAMILY_CACHE[session_id]['config']
            ward_title = c.get('ward_title', ward_title)
            part_no = c.get('part_no', part_no)
            booth_address = c.get('booth_address', booth_address)
            list_title = c.get('list_title', list_title)
            selected_ward = c.get('selected_ward', selected_ward)
            min_family_size = c.get('min_family_size', min_family_size)
            ignore_zero = c.get('ignore_zero_houses', ignore_zero)
            filter_active_only = c.get('filter_active_only', filter_active_only)
            jp = c.get('jila_parishad', jp)
            ps = c.get('panchayat_samiti', ps)

        if not jp:
            jp = FAMILY_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        if not ps:
            ps = FAMILY_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')

        html, total_pages, total_voters, total_families, _ = generate_family_html(
            voters=voters,
            ward_title=ward_title,
            part_no=part_no,
            booth_address=booth_address,
            list_title=list_title,
            min_family_size=min_family_size,
            ignore_zero_houses=ignore_zero,
            selected_ward=selected_ward,
            filter_active_only=filter_active_only,
            for_preview=False,
            max_lines_per_page=27,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        auto_print_script = """
        <script>
            window.addEventListener('DOMContentLoaded', () => {
                setTimeout(() => {
                    window.print();
                }, 500);
            });
        </script>
        """
        html_with_print = html.replace('</body>', f'{auto_print_script}</body>')
        return html_with_print
    except Exception as e:
        return f"<h3>त्रुटि: {str(e)}</h3>", 500

# ==============================================================
# FEATURE 6: आयु अनुसार (युवा 18-26 एवं बुजुर्ग 120-70) वोटर लिस्ट ENDPOINTS
# ==============================================================

@app.route('/api/age_list/from_session', methods=['POST'])
def api_age_from_session():
    try:
        data = request.get_json() or {}
        feature1_session_id = data.get('session_id')
        
        if not feature1_session_id and SESSION_CACHE:
            feature1_session_id = list(SESSION_CACHE.keys())[-1]
            
        if not feature1_session_id or feature1_session_id not in SESSION_CACHE:
            return jsonify({'error': 'फ़ीचर 1 से कोई मतदाता डेटा उपलब्ध नहीं है। कृपया पहले पीडीएफ प्रोसेस करें या एक्सेल अपलोड करें।'}), 404
            
        res = SESSION_CACHE[feature1_session_id]['result']
        voters = []
        
        for w in res.get('wards', []):
            grid = w.get('grid', [])
            if len(grid) > 1:
                headers = grid[0]
                col_idx = {}
                for idx, h in enumerate(headers):
                    hl = str(h).strip().lower()
                    if 'क्रम' in hl or 'serial' in hl: col_idx['serial'] = idx
                    elif 'जि' in hl and 'प' in hl: col_idx['jila_parishad'] = idx
                    elif 'पं' in hl and 'स' in hl: col_idx['panchayat_samiti'] = idx
                    elif 'भाग' in hl or 'वार्ड' in hl or 'ward' in hl: col_idx['part_no'] = idx
                    elif 'नाम' in hl and 'पिता' not in hl and 'पति' not in hl and 'बूथ' not in hl: col_idx['name'] = idx
                    elif 'पिता' in hl or 'पति' in hl or 'संबंधी' in hl: col_idx['relative_name'] = idx
                    elif 'आयु' in hl or 'उम्र' in hl or 'age' in hl: col_idx['age'] = idx
                    elif 'लिंग' in hl or 'gender' in hl or 'sex' in hl: col_idx['gender'] = idx
                    elif 'पहचान' in hl or 'epic' in hl or 'वोटर' in hl: col_idx['epic'] = idx
                    elif 'मकान' in hl or 'हाउस' in hl or 'house' in hl: col_idx['house'] = idx
                    elif 'बूथ' in hl or 'केंद्र' in hl or 'address' in hl: col_idx['booth_address'] = idx
                    elif 'status' in hl or 'स्थिति' in hl: col_idx['status'] = idx

                for row in grid[1:]:
                    if not any(row): continue
                    def gv(k):
                        idx = col_idx.get(k)
                        return clean_val_str(row[idx]) if idx is not None and idx < len(row) else ''
                    
                    voters.append({
                        'serial': gv('serial') or str(len(voters) + 1),
                        'jila_parishad': gv('jila_parishad') or str(w.get('jila_parishad', res.get('jila_parishad', ''))),
                        'panchayat_samiti': gv('panchayat_samiti') or str(w.get('panchayat_samiti', res.get('panchayat_samiti', ''))),
                        'part_no': gv('part_no') or str(w.get('ward_no', '1')),
                        'name': gv('name'),
                        'relative_name': gv('relative_name'),
                        'age': gv('age'),
                        'gender': gv('gender'),
                        'epic': gv('epic'),
                        'house': gv('house'),
                        'booth_address': gv('booth_address') or w.get('polling_station', ''),
                        'status': gv('status') or 'Active'
                    })

        if not voters:
            return jsonify({'error': 'वर्तमान सत्र में कोई मतदाता रिकॉर्ड नहीं मिले।'}), 400

        session_id = str(uuid.uuid4())
        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        first_booth = next((v.get('booth_address') for v in voters if v.get('booth_address')), "राजकीय उच्च माध्यमिक विद्यालय")
        
        parsed = {
            'voters': voters,
            'sheet_names': ['सत्र_डेटा'],
            'active_sheet': 'सत्र_डेटा',
            'wards': raw_wards,
            'total_voters': len(voters),
            'default_panchayat_name': res.get('panchayat_name', "ग्राम पंचायत"),
            'default_part_no': raw_wards[0] if raw_wards else "1",
            'default_booth_address': first_booth,
            'default_jila_parishad': res.get('jila_parishad', ''),
            'default_panchayat_samiti': res.get('panchayat_samiti', '')
        }
        
        AGE_CACHE[session_id] = {
            'file_path': None,
            'filename': 'वर्तमान_सत्र_डेटा',
            'parsed': parsed
        }
        
        log_activity(
            feature_id='feature6',
            feature_name='फ़ीचर 6: युवा एवं बुजुर्ग वोटर लिस्ट',
            file_names=['सत्र मतदाता डेटा'],
            action='सत्र से आयु अनुसार लिस्ट लोड',
            details=f"{len(voters)} मतदाता • {len(raw_wards)} वार्ड्स"
        )
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': ['सत्र_डेटा'],
            'active_sheet': 'सत्र_डेटा',
            'wards': raw_wards,
            'total_voters': len(voters),
            'default_panchayat_name': parsed['default_panchayat_name'],
            'default_booth_address': parsed['default_booth_address'],
            'default_jila_parishad': parsed['default_jila_parishad'],
            'default_panchayat_samiti': parsed['default_panchayat_samiti'],
            'voters_sample': voters[:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/age_list/upload_excel', methods=['POST'])
def api_age_upload_excel():
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'कृपया एक वैध एक्सेल फ़ाइल (.xlsx/.xls) चुनें'}), 400
        f = request.files['excel_file']
        if not f or not f.filename:
            return jsonify({'error': 'फ़ाइल उपलब्ध नहीं है'}), 400
            
        file_path = os.path.join(UPLOAD_FOLDER, f"age_{uuid.uuid4().hex[:8]}_{secure_filename(f.filename)}")
        f.save(file_path)
        
        parsed = parse_excel_for_parchi(file_path)
        voters = parsed['voters']
        
        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        parsed['wards'] = raw_wards
        
        session_id = str(uuid.uuid4())
        
        AGE_CACHE[session_id] = {
            'file_path': file_path,
            'filename': f.filename,
            'parsed': parsed
        }

        log_activity(
            feature_id='feature6',
            feature_name='फ़ीचर 6: युवा एवं बुजुर्ग वोटर लिस्ट',
            file_names=[f.filename],
            action='आयु-वार लिस्ट हेतु एक्सेल अपलोड',
            details=f"{parsed['total_voters']} मतदाता • {len(raw_wards)} वार्ड्स"
        )
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': parsed['wards'],
            'total_voters': parsed['total_voters'],
            'default_panchayat_name': parsed.get('default_panchayat_name') or "ग्राम पंचायत",
            'default_booth_address': parsed.get('default_booth_address') or "राजकीय उच्च माध्यमिक विद्यालय",
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', ''),
            'voters_sample': voters[:10]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/age_list/select_sheet', methods=['POST'])
def api_age_select_sheet():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        sheet_name = data.get('sheet_name')
        
        if not session_id or session_id not in AGE_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        file_path = AGE_CACHE[session_id].get('file_path')
        if not file_path:
            return jsonify({'error': 'सत्र फ़ाइल उपलब्ध नहीं है'}), 400
            
        parsed = parse_excel_for_parchi(file_path, sheet_name=sheet_name)
        voters = parsed['voters']
        def extract_w_key(v):
            digits = re.findall(r'\d+', str(v or ''))
            return int(digits[0]) if digits else 999999

        raw_wards = sorted(list(set(str(v.get('part_no', '1')).strip() for v in voters if v.get('part_no'))), key=extract_w_key)
        parsed['wards'] = raw_wards
        AGE_CACHE[session_id]['parsed'] = parsed
        
        return jsonify({
            'session_id': session_id,
            'sheet_names': parsed['sheet_names'],
            'active_sheet': parsed['active_sheet'],
            'wards': parsed['wards'],
            'total_voters': parsed['total_voters'],
            'default_panchayat_name': parsed.get('default_panchayat_name') or "ग्राम पंचायत",
            'default_booth_address': parsed.get('default_booth_address') or "राजकीय उच्च माध्यमिक विद्यालय",
            'default_jila_parishad': parsed.get('default_jila_parishad', ''),
            'default_panchayat_samiti': parsed.get('default_panchayat_samiti', '')
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/age_list/preview_html', methods=['POST'])
def api_age_preview_html():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'ग्राम पंचायत')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', '')
        list_type = data.get('list_type', 'young')
        min_age = data.get('min_age')
        max_age = data.get('max_age')
        sort_order = data.get('sort_order')
        if min_age is not None and str(min_age).strip() != '':
            try: min_age = int(min_age)
            except: min_age = None
        else: min_age = None
        if max_age is not None and str(max_age).strip() != '':
            try: max_age = int(max_age)
            except: max_age = None
        else: max_age = None

        selected_ward = data.get('selected_ward', 'all')
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
        page = int(data.get('page', 1))
        
        if not session_id or session_id not in AGE_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = AGE_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or AGE_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or AGE_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        AGE_CACHE[session_id]['config'] = {
            'ward_title': ward_title,
            'booth_address': booth_address,
            'list_title': list_title,
            'list_type': list_type,
            'min_age': min_age,
            'max_age': max_age,
            'sort_order': sort_order,
            'selected_ward': selected_ward,
            'filter_active_only': filter_active_only,
            'jila_parishad': jp,
            'panchayat_samiti': ps
        }
        
        html, total_pages, total_matched, page_meta = generate_age_html(
            voters=voters,
            ward_title=ward_title,
            booth_address=booth_address,
            list_title=list_title,
            list_type=list_type,
            min_age=min_age,
            max_age=max_age,
            sort_order=sort_order,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            for_preview=True,
            preview_page=page,
            rows_per_page=31,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        return jsonify({
            'html': html,
            'total_pages': total_pages,
            'total_matched': total_matched,
            'current_page': page,
            'current_ward': page_meta.get('current_ward', 'समस्त'),
            'page_in_ward': page_meta.get('page_in_ward', 1),
            'total_in_ward': page_meta.get('total_in_ward', 1),
            'ward_voters': page_meta.get('ward_voters', total_matched),
            'list_type': page_meta.get('list_type', list_type),
            'min_age': page_meta.get('min_age'),
            'max_age': page_meta.get('max_age'),
            'sort_order': page_meta.get('sort_order')
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/age_list/export_pdf', methods=['POST'])
def api_age_export_pdf():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'ग्राम पंचायत')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', '')
        list_type = data.get('list_type', 'young')
        min_age = data.get('min_age')
        max_age = data.get('max_age')
        sort_order = data.get('sort_order')
        if min_age is not None and str(min_age).strip() != '':
            try: min_age = int(min_age)
            except: min_age = None
        else: min_age = None
        if max_age is not None and str(max_age).strip() != '':
            try: max_age = int(max_age)
            except: max_age = None
        else: max_age = None

        selected_ward = data.get('selected_ward', 'all')
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
            
        if not session_id or session_id not in AGE_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = AGE_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or AGE_CACHE[session_id].get('config', {}).get('jila_parishad') or AGE_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or AGE_CACHE[session_id].get('config', {}).get('panchayat_samiti') or AGE_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        pdf_bytes = generate_age_pdf(
            voters=voters,
            ward_title=ward_title,
            booth_address=booth_address,
            list_title=list_title,
            list_type=list_type,
            min_age=min_age,
            max_age=max_age,
            sort_order=sort_order,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        safe_title = re.sub(r'[^\w\s-]', '', ward_title).strip() or "Age_Voter_List"
        type_str = "Young_18-26" if list_type == "young" else ("Senior_120-70" if list_type == "senior" else "Custom_Age")
        file_name = f"Geam_Digital_{safe_title}_{type_str}_List.pdf"

        log_activity(
            feature_id='feature6',
            feature_name='फ़ीचर 6: युवा एवं बुजुर्ग वोटर लिस्ट',
            file_names=[file_name],
            action=f'आयु सूची PDF जनरेट ({type_str})',
            details=f"{len(voters)} कुल मतदाता • {ward_title}"
        )
        
        return send_file(
            BytesIO(pdf_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/age_list/export_excel', methods=['POST'])
def api_age_export_excel():
    try:
        data = request.get_json() or {}
        session_id = data.get('session_id')
        ward_title = data.get('ward_title', 'ग्राम पंचायत')
        booth_address = data.get('booth_address', '')
        list_title = data.get('list_title', '')
        list_type = data.get('list_type', 'young')
        min_age = data.get('min_age')
        max_age = data.get('max_age')
        sort_order = data.get('sort_order')
        if min_age is not None and str(min_age).strip() != '':
            try: min_age = int(min_age)
            except: min_age = None
        else: min_age = None
        if max_age is not None and str(max_age).strip() != '':
            try: max_age = int(max_age)
            except: max_age = None
        else: max_age = None

        selected_ward = data.get('selected_ward', 'all')
        filter_active_only = data.get('filter_active_only', True)
        if isinstance(filter_active_only, str):
            filter_active_only = (filter_active_only.lower() in ['true', '1', 'yes'])
            
        if not session_id or session_id not in AGE_CACHE:
            return jsonify({'error': 'सत्र समाप्त हो गया है। कृपया फ़ाइल पुनः लोड करें।'}), 404
            
        voters = AGE_CACHE[session_id]['parsed']['voters']
        
        jp = data.get('jila_parishad') or AGE_CACHE[session_id].get('config', {}).get('jila_parishad') or AGE_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        ps = data.get('panchayat_samiti') or AGE_CACHE[session_id].get('config', {}).get('panchayat_samiti') or AGE_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')
        
        excel_bytes = generate_age_excel(
            voters=voters,
            ward_title=ward_title,
            booth_address=booth_address,
            list_title=list_title,
            list_type=list_type,
            min_age=min_age,
            max_age=max_age,
            sort_order=sort_order,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        safe_title = re.sub(r'[^\w\s-]', '', ward_title).strip() or "Age_Voter_List"
        type_str = "Young_18-26" if list_type == "young" else ("Senior_120-70" if list_type == "senior" else "Custom_Age")
        file_name = f"Geam_Digital_{safe_title}_{type_str}_List.xlsx"

        log_activity(
            feature_id='feature6',
            feature_name='फ़ीचर 6: युवा एवं बुजुर्ग वोटर लिस्ट',
            file_names=[file_name],
            action=f'आयु सूची Excel जनरेट ({type_str})',
            details=f"{len(voters)} कुल मतदाता • {ward_title}"
        )
        
        return send_file(
            BytesIO(excel_bytes),
            as_attachment=True,
            download_name=file_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/age_list/print', methods=['GET', 'POST'])
def age_print_standalone():
    try:
        session_id = request.args.get('session_id') or (request.form.get('session_id') if request.method == 'POST' else None)
        ward_title = request.args.get('ward_title', 'ग्राम पंचायत')
        booth_address = request.args.get('booth_address', '')
        list_title = request.args.get('list_title', '')
        list_type = request.args.get('list_type', 'young')
        min_age = request.args.get('min_age')
        max_age = request.args.get('max_age')
        sort_order = request.args.get('sort_order')
        if min_age is not None and str(min_age).strip() != '':
            try: min_age = int(min_age)
            except: min_age = None
        else: min_age = None
        if max_age is not None and str(max_age).strip() != '':
            try: max_age = int(max_age)
            except: max_age = None
        else: max_age = None

        selected_ward = request.args.get('selected_ward', 'all')
        filter_active_only = (request.args.get('filter_active_only', 'true').lower() == 'true')
        jp = request.args.get('jila_parishad') or (request.form.get('jila_parishad') if request.method == 'POST' else None)
        ps = request.args.get('panchayat_samiti') or (request.form.get('panchayat_samiti') if request.method == 'POST' else None)
        
        if not session_id or session_id not in AGE_CACHE:
            return "<h3>सत्र समाप्त हो गया है। कृपया मुख्य पोर्टल से पुनः प्रिंट बटन दबाएं।</h3>", 404
            
        voters = AGE_CACHE[session_id]['parsed']['voters']
        
        if 'config' in AGE_CACHE[session_id]:
            c = AGE_CACHE[session_id]['config']
            ward_title = c.get('ward_title', ward_title)
            booth_address = c.get('booth_address', booth_address)
            list_title = c.get('list_title', list_title)
            list_type = c.get('list_type', list_type)
            min_age = c.get('min_age', min_age)
            max_age = c.get('max_age', max_age)
            sort_order = c.get('sort_order', sort_order)
            selected_ward = c.get('selected_ward', selected_ward)
            filter_active_only = c.get('filter_active_only', filter_active_only)
            jp = c.get('jila_parishad', jp)
            ps = c.get('panchayat_samiti', ps)

        if not jp:
            jp = AGE_CACHE[session_id]['parsed'].get('default_jila_parishad', '')
        if not ps:
            ps = AGE_CACHE[session_id]['parsed'].get('default_panchayat_samiti', '')

        html, total_pages, total_matched, _ = generate_age_html(
            voters=voters,
            ward_title=ward_title,
            booth_address=booth_address,
            list_title=list_title,
            list_type=list_type,
            min_age=min_age,
            max_age=max_age,
            sort_order=sort_order,
            filter_active_only=filter_active_only,
            selected_ward=selected_ward,
            for_preview=False,
            rows_per_page=31,
            jila_parishad=jp,
            panchayat_samiti=ps
        )
        
        auto_print_script = """
        <script>
            window.addEventListener('DOMContentLoaded', () => {
                setTimeout(() => {
                    window.print();
                }, 500);
            });
        </script>
        """
        html_with_print = html.replace('</body>', f'{auto_print_script}</body>')
        return html_with_print
    except Exception as e:
        return f"<h3>त्रुटि: {str(e)}</h3>", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Starting PDF to Excel Studio on http://127.0.0.1:{port}")
    app.run(host='127.0.0.1', port=port, debug=False)
