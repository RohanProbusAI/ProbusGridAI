# app.py
import os
import uuid
from flask import Flask, render_template, request, jsonify, send_from_directory, g
from werkzeug.utils import secure_filename
import liberty_logic
from dotenv import load_dotenv
import pandas as pd
import json
import tata as tata_logic
import zuno as zuno_logic
from bajaj import bajaj_ai_logic as bajaj_ai
from bajaj import bajaj_manual_logic as bajaj_manual

load_dotenv()

app = Flask(__name__)

# --- Configuration ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
CONVERTED_FOLDER = os.path.join(BASE_DIR, 'converted')
PROCESSED_FOLDER = os.path.join(BASE_DIR, 'processed')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'pdf', 'xlsx', 'xls', 'json'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- CLEANUP HELPERS ---
def safe_delete(filepath):
    """Delete a file safely with logging"""
    try:
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
            print(f"✓ Deleted: {filepath}")
            return True
    except Exception as e:
        print(f"✗ Failed to delete {filepath}: {e}")
    return False

def cleanup_session_files(upload_path=None, converted_path=None):
    """Delete upload and converted files"""
    if upload_path:
        safe_delete(upload_path)
    if converted_path:
        safe_delete(converted_path)

# Track files for cleanup after download
download_cleanup = {}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/select_processor', methods=['POST'])
def select_processor():
    data = request.json
    filename = data.get('filename', '').lower()
    is_bajaj = 'bajaj' in filename
    
    if is_bajaj:
        return jsonify({
            'status': 'processor_selection_required',
            'company': 'Bajaj',
            'processors': [
                {
                    'id': 'bajaj_ai',
                    'name': 'Bajaj AI Processor',
                    'description': 'Uses OpenAI to intelligently extract data',
                    'features': ['Multi-fuel support', 'Smart extraction', 'Complex layouts'],
                    'image': '/static/bajaj_ai_preview.png'
                },
                {
                    'id': 'bajaj_manual',
                    'name': 'Bajaj Manual Processor',
                    'description': 'Interactive column-based extraction',
                    'features': ['Manual mapping', 'Segment detection', 'Quick processing'],
                    'image': 'static/bajaj_manual_preview.png'
                }
            ]
        })
    return jsonify({'status': 'use_liberty', 'company': 'Liberty/Other'})

@app.route('/process_bajaj_ai', methods=['POST'])
def process_bajaj_ai():
    upload_path = None
    try:
        data = request.json
        upload_path = data.get('upload_path')
        company_name = data.get('company_name', 'BAJAJ')
        
        if not upload_path or not os.path.exists(upload_path):
            return jsonify({'detail': 'File not found'}), 400
        
        print(f"Processing Bajaj AI: {upload_path}")
        result = bajaj_ai.parse_file_with_ai(file_path=upload_path, user_company=company_name)
        
        if "error" in result:
            cleanup_session_files(upload_path)
            return jsonify({'detail': result['error']}), 500
        
        worksheets_data = result.get('worksheets', {})
        if not worksheets_data:
            cleanup_session_files(upload_path)
            return jsonify({'status': 'success_no_data', 'message': 'No data extracted'})
        
        results = []
        output_filename = f'bajaj_output_{uuid.uuid4().hex[:8]}.xlsx'
        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            total_entries = 0
            for sheet_name, entries in worksheets_data.items():
                if not entries:
                    continue
                
                output_data = {
                    "COMPANY NAME": [], "SEGMENT": [], "POLICY TYPE": [], "LOCATION": [],
                    "FUEL TYPE": [], "PAYIN": [], "PAYOUT": [], "REMARK": [], "CALCULATION EXPLANATION": []
                }
                
                for entry in entries:
                    output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
                    output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
                    output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
                    output_data["LOCATION"].append(entry.get("LOCATION", ""))
                    output_data["FUEL TYPE"].append(entry.get("FUEL TYPE", ""))
                    output_data["PAYIN"].append(entry.get("PAYIN", ""))
                    output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
                    output_data["REMARK"].append("; ".join(entry.get("REMARK", [])) if isinstance(entry.get("REMARK"), list) else str(entry.get("REMARK", "")))
                    output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
                
                output_df = pd.DataFrame(output_data)
                safe_sheet_name = sheet_name[:31]
                output_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                total_entries += len(entries)
        
        print(f"✓ Created Excel: {output_filename} ({total_entries} entries)")
        
        # Store paths for cleanup after download
        download_cleanup[output_filename] = {'upload': upload_path}
        
        results.append({
            'filename': output_filename,
            'sheet_name': 'All Sheets',
            'download_url': f'/processed/{output_filename}',
            'entry_count': total_entries
        })
        
        return jsonify({'status': 'success_single_file', 'results': results})
        
    except Exception as e:
        print(f"✗ Error in Bajaj AI: {e}")
        if upload_path:
            cleanup_session_files(upload_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500

@app.route('/detect_bajaj_segment', methods=['POST'])
def detect_bajaj_segment():
    try:
        data = request.json
        temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], data['temp_json_id'])
        
        with open(temp_json_path, 'r') as f:
            json_data = json.load(f)
        
        sheet_names = list(json_data['sheets'].keys())
        sheet_index = int(data['sheet_index'])
        sheet_name = sheet_names[sheet_index]
        
        detected_segment, confidence = bajaj_manual.detect_segment_from_worksheet_name(sheet_name)
        
        return jsonify({
            'status': 'success',
            'detected_segment': detected_segment,
            'confidence': confidence
        })
    except Exception as e:
        print(f"✗ Segment detection error: {e}")
        return jsonify({'status': 'error', 'detected_segment': None, 'confidence': 'NONE'})

@app.route('/process_bajaj_manual', methods=['POST'])
def process_bajaj_manual():
    upload_path = None
    try:
        data = request.json
        upload_path = data.get('upload_path')
        company_name = data.get('company_name', 'BAJAJ')
        
        if not upload_path or not os.path.exists(upload_path):
            return jsonify({'detail': 'File not found'}), 400
        
        print(f"Starting Bajaj Manual: {upload_path}")
        json_data = bajaj_manual.excel_to_json_converter(upload_path)
        
        temp_json_id = f"bajaj_manual_{uuid.uuid4().hex[:8]}.json"
        temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
        
        with open(temp_json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        
        sheets_info = []
        for idx, (sheet_name, sheet_data) in enumerate(json_data['sheets'].items()):
            headers = []
            if sheet_data:
                first_row = sheet_data[0]
                for key, value in first_row.items():
                    headers.append({'key': key, 'value': str(value) if value is not None else None})
            
            sheets_info.append({
                'index': idx,
                'name': sheet_name,
                'record_count': len(sheet_data),
                'headers': headers
            })
        
        # Store for cleanup later
        if temp_json_id not in download_cleanup:
            download_cleanup[temp_json_id] = {}
        download_cleanup[temp_json_id]['upload'] = upload_path
        download_cleanup[temp_json_id]['converted'] = temp_json_path
        
        return jsonify({
            'status': 'success_wizard_start',
            'temp_json_id': temp_json_id,
            'filename': os.path.basename(upload_path),
            'company_name': company_name,
            'sheets_info': sheets_info
        })
        
    except Exception as e:
        print(f"✗ Bajaj Manual error: {e}")
        if upload_path:
            cleanup_session_files(upload_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500

@app.route('/process_bajaj_sheet', methods=['POST'])
def process_bajaj_sheet():
    converted_path = None
    upload_path = None
    try:
        payload = request.json
        temp_json_id = payload['temp_json_id']
        temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
        converted_path = temp_json_path
        
        with open(temp_json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        sheet_names = list(json_data['sheets'].keys())
        sheet_index = int(payload['sheet_index'])
        sheet_name = sheet_names[sheet_index]
        sheet_data = json_data['sheets'][sheet_name]
        
        processed_entries = process_bajaj_sheet_data(
            sheet_data, sheet_name, payload['manual_segment_raw'],
            payload['payin_cols'], payload.get('location_col'),
            payload.get('location_type', 'column'), payload.get('location_row'),
            payload.get('policy_type_source', 'auto'), payload.get('policy_col'),
            payload['company_name']
        )
        
        if processed_entries:
            output_filename = f"bajaj_manual_{sheet_name[:20]}_{uuid.uuid4().hex[:8]}.xlsx"
            output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
            
            output_data = {
                "COMPANY NAME": [], "SEGMENT": [], "POLICY TYPE": [], "LOCATION": [],
                "PAYIN": [], "PAYOUT": [], "REMARK": [], "CALCULATION EXPLANATION": []
            }
            
            for entry in processed_entries:
                output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
                output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
                output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
                output_data["LOCATION"].append(entry.get("LOCATION", ""))
                output_data["PAYIN"].append(entry.get("PAYIN", ""))
                output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
                output_data["REMARK"].append("; ".join(entry.get("REMARK", [])) if isinstance(entry.get("REMARK"), list) else str(entry.get("REMARK", "")))
                output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
            
            output_df = pd.DataFrame(output_data)
            output_df.to_excel(output_path, index=False, engine='openpyxl')
            
            # Track files for cleanup
            if temp_json_id in download_cleanup:
                upload_path = download_cleanup[temp_json_id].get('upload')
            download_cleanup[output_filename] = {
                'upload': upload_path,
                'converted': converted_path
            }
            
            return jsonify({
                'status': 'success_sheet_processed',
                'result': {
                    'filename': output_filename,
                    'sheet_name': sheet_name,
                    'download_url': f'/processed/{output_filename}',
                    'entry_count': len(processed_entries)
                }
            })
        else:
            cleanup_session_files(upload_path, converted_path)
            return jsonify({
                'status': 'success_sheet_processed_empty',
                'result': {'filename': None, 'sheet_name': sheet_name, 'entry_count': 0}
            })
        
    except Exception as e:
        print(f"✗ Bajaj sheet error: {e}")
        cleanup_session_files(upload_path, converted_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500

def process_bajaj_sheet_data(sheet_data, sheet_name, manual_segment, payin_cols, 
                             location_col, location_type, location_row, 
                             policy_type_source, policy_col, company_name):
    temp_json = {
        'source_file': sheet_name,
        'total_sheets': 1,
        'sheets': {sheet_name: sheet_data}
    }
    
    result = bajaj_manual.parse_converted_json(
        temp_json, original_excel_path=sheet_name,
        user_company=company_name, skip_prompts=True,
        manual_segment=manual_segment
    )
    

    
    
    worksheets_data = result.get('worksheets', {})
    all_entries = []
    for entries in worksheets_data.values():
        all_entries.extend(entries)
    return all_entries

@app.route('/upload', methods=['POST'])
def upload_file():
    upload_path = None
    try:
        if 'file' not in request.files:
            return jsonify({'detail': 'No file part'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'detail': 'No selected file'}), 400

        if file and allowed_file(file.filename):
            original_filename = secure_filename(file.filename)
            file_ext = original_filename.rsplit('.', 1)[1].lower()
            unique_filename = f"{uuid.uuid4()}_{original_filename}"
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(upload_path)

            is_bajaj = 'bajaj' in original_filename.lower()

            file_type = ""
            if file_ext in {'pdf'}:
                file_type = "pdf"
            elif file_ext in {'png', 'jpg', 'jpeg', 'gif', 'webp'}:
                file_type = "image"
            elif file_ext in {'xlsx', 'xls'}:
                file_type = "excel"
            elif file_ext in {'json'}:
                file_type = "json"

            # PDF/Image - process and cleanup immediately
            if file_type in {"pdf", "image"}:
                print(f"Processing single file: {upload_path}")
                result = liberty_logic.process_pdf_or_image_file(upload_path, app.config['PROCESSED_FOLDER'])
                
                # Store for cleanup after download
                output_filename = result['filename']
                download_cleanup[output_filename] = {'upload': upload_path}
                
                return jsonify({'status': 'success_single_file', 'results': [result]})

            elif file_type in {"excel", "json"}:
                if is_bajaj and file_type == "excel":
                    return jsonify({
                        'status': 'processor_selection_required',
                        'upload_path': upload_path,
                        'filename': original_filename,
                        'company': 'Bajaj',
                        'processors': [
                            {
                                'id': 'bajaj_ai',
                                'name': 'Bajaj AI Processor',
                                'description': 'OpenAI-powered extraction',
                                'features': ['Multi-fuel', 'Smart extraction'],
                                'image': 'static/bajaj_ai_preview.png'
                            },
                            {
                                'id': 'bajaj_manual',
                                'name': 'Bajaj Manual Processor',
                                'description': 'Column-based extraction',
                                'features': ['Manual mapping', 'Quick processing'],
                                'image': 'static/bajaj_manual_preview.png'
                            }
                        ]
                    })
                
                print(f"Starting Excel/JSON conversion: {upload_path}")
                conversion_result = liberty_logic.start_excel_conversion(upload_path, app.config['CONVERTED_FOLDER'])
                suggested_company, _ = liberty_logic.extract_from_filename(original_filename)

                # Store for cleanup
                temp_json_id = conversion_result['temp_json_id']
                temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
                download_cleanup[temp_json_id] = {
                    'upload': upload_path,
                    'converted': temp_json_path
                }

                return jsonify({
                    'status': 'success_wizard_start',
                    'temp_json_id': temp_json_id,
                    'filename': original_filename,
                    'suggested_company': suggested_company or ''
                })
            
            return jsonify({'detail': 'Unsupported file type'}), 400

    except Exception as e:
        print(f"✗ Upload error: {e}")
        if upload_path:
            safe_delete(upload_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500
    
    return jsonify({'detail': 'File type not allowed'}), 400



@app.route('/start_excel_wizard', methods=['POST'])
def start_excel_wizard():
    try:
        data = request.json
        company_name = data.get('company_name')
        temp_json_id = data.get('temp_json_id')

        if not company_name or not temp_json_id:
            return jsonify({'detail': 'Missing required fields'}), 400
        
        # Direct processing for Zuno and Tata
        if company_name.upper() in ['ZUNO', 'TATA']:
            return jsonify({
                'status': f'success_{company_name.lower()}_detected',
                'company_name': company_name,
                'temp_json_id': temp_json_id
            })
        
        # Liberty and others get sheet selection
        json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
        sheet_info = liberty_logic.get_sheet_info_from_json(json_path)
        
        return jsonify({
            'status': 'success_sheets_loaded',
            'company_name': company_name,
            'temp_json_id': temp_json_id,
            'sheets_info': sheet_info
        })

    except Exception as e:
        print(f"✗ Wizard error: {e}")
        return jsonify({'detail': f'Error: {str(e)}'}), 500


@app.route('/process_sheet', methods=['POST'])
def process_sheet():
    converted_path = None
    upload_path = None
    try:
        payload = request.json
        temp_json_id = payload['temp_json_id']
        company_name = payload.get('company_name', '').upper()
        
        # Track cleanup files
        if temp_json_id in download_cleanup:
            upload_path = download_cleanup[temp_json_id].get('upload')
            converted_path = download_cleanup[temp_json_id].get('converted')
        
        # If Tata, use tata_logic
        if company_name == 'TATA':
            temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
            
            with open(temp_json_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            # Get selected sheet
            sheet_names = list(json_data['sheets'].keys())
            sheet_index = int(payload['sheet_index'])
            sheet_name = sheet_names[sheet_index]
            
            # Create temporary JSON with only selected sheet
            selected_sheet_data = {
                'source_file': json_data.get('source_file'),
                'total_sheets': 1,
                'sheets': {sheet_name: json_data['sheets'][sheet_name]}
            }
            
            # Process using tata_logic
            result = tata_logic.parse_converted_json(selected_sheet_data, original_excel_path=json_data.get('source_file'))
            entries = result.get('structured_data', [])
            
            if not entries:
                cleanup_session_files(upload_path, converted_path)
                return jsonify({
                    'status': 'success_sheet_processed_empty',
                    'result': {'filename': None, 'sheet_name': sheet_name, 'entry_count': 0}
                })
            
            output_filename = f'tata_{sheet_name[:20]}_{uuid.uuid4().hex[:8]}.xlsx'
            output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
            
            output_data = {
                "COMPANY NAME": [], "SEGMENT": [], "POLICY TYPE": [], "LOCATION": [],
                "PAYIN": [], "PAYOUT": [], "REMARK": [], "CALCULATION EXPLANATION": []
            }
            
            for entry in entries:
                output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
                output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
                output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
                output_data["LOCATION"].append(entry.get("LOCATION", ""))
                output_data["PAYIN"].append(entry.get("PAYIN", ""))
                output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
                output_data["REMARK"].append("; ".join(entry.get("REMARK", [])) if isinstance(entry.get("REMARK"), list) else str(entry.get("REMARK", "")))
                output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
            
            output_df = pd.DataFrame(output_data)
            output_df.to_excel(output_path, index=False, engine='openpyxl')
            
            download_cleanup[output_filename] = {
                'upload': upload_path,
                'converted': converted_path
            }
            
            return jsonify({
                'status': 'success_sheet_processed',
                'result': {
                    'filename': output_filename,
                    'sheet_name': sheet_name,
                    'download_url': f'/processed/{output_filename}',
                    'entry_count': len(entries)
                }
            })
        
        # Otherwise use liberty_logic for other companies
        result = liberty_logic.process_single_excel_sheet(
            payload,
            app.config['CONVERTED_FOLDER'],
            app.config['PROCESSED_FOLDER']
        )
        
        if result['entry_count'] > 0:
            output_filename = result['filename']
            download_cleanup[output_filename] = {
                'upload': upload_path,
                'converted': converted_path
            }
            return jsonify({'status': 'success_sheet_processed', 'result': result})
        else:
            cleanup_session_files(upload_path, converted_path)
            return jsonify({'status': 'success_sheet_processed_empty', 'result': result})

    except Exception as e:
        print(f"✗ Sheet processing error: {e}")
        cleanup_session_files(upload_path, converted_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500

@app.route('/process_zuno_sheet', methods=['POST'])
def process_zuno_sheet():
    converted_path = None
    upload_path = None
    try:
        data = request.json
        temp_json_id = data.get('temp_json_id')
        temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
        converted_path = temp_json_path
        
        with open(temp_json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        original_excel = json_data.get('source_file')
        result = zuno_logic.parse_file(original_excel)
        entries = result.get('structured_data', [])

        if not entries:
            if temp_json_id in download_cleanup:
                upload_path = download_cleanup[temp_json_id].get('upload')
            cleanup_session_files(upload_path, converted_path)
            return jsonify({
                'status': 'success_sheet_processed_empty',
                'result': {'filename': None, 'sheet_name': 'All Sheets', 'entry_count': 0}
            })
        
        output_filename = f'zuno_output_{uuid.uuid4().hex[:8]}.xlsx'
        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
        
        output_data = {
            "COMPANY NAME": [], "SEGMENT": [], "POLICY TYPE": [], "LOCATION": [],
            "PAYIN": [], "PAYOUT": [], "REMARK": [], "CALCULATION EXPLANATION": []
        }
        
        for entry in entries:
            output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
            output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
            output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
            output_data["LOCATION"].append(entry.get("LOCATION", ""))
            output_data["PAYIN"].append(entry.get("PAYIN", ""))
            output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
            output_data["REMARK"].append("; ".join(entry.get("REMARK", [])) if isinstance(entry.get("REMARK"), list) else str(entry.get("REMARK", "")))
            output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
        
        output_df = pd.DataFrame(output_data)
        output_df.to_excel(output_path, index=False, engine='openpyxl')
        
        # Track for cleanup
        # if temp_json_id in download_cleanup:
        #     upload_path = download_cleanup[temp_json_id].get('upload')
        # download_cleanup[output_filename] = {
        #     'upload': upload_path,
        #     'converted': converted_path
        # }

        # Track ONLY output file, not session files
        download_cleanup[output_filename] = {
            'upload': None,  # Don't delete on individual downloads
            'converted': None
        }

        return jsonify({
            'status': 'success_sheet_processed',
            'result': {
                'filename': output_filename,
                'sheet_name': 'All Sheets',
                'download_url': f'/processed/{output_filename}',
                'entry_count': len(entries)
            }
        })
        
    except Exception as e:
        print(f"✗ Zuno error: {e}")
        cleanup_session_files(upload_path, converted_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500


@app.route('/process_tata_sheet', methods=['POST'])
def process_tata_sheet():
    converted_path = None
    upload_path = None
    try:
        data = request.json
        temp_json_id = data.get('temp_json_id')
        temp_json_path = os.path.join(app.config['CONVERTED_FOLDER'], temp_json_id)
        converted_path = temp_json_path
        
        with open(temp_json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        
        original_excel = json_data.get('source_file')
        result = tata_logic.parse_file(original_excel)
        entries = result.get('structured_data', [])

        if not entries:
            if temp_json_id in download_cleanup:
                upload_path = download_cleanup[temp_json_id].get('upload')
            cleanup_session_files(upload_path, converted_path)
            return jsonify({
                'status': 'success_sheet_processed_empty',
                'result': {'filename': None, 'sheet_name': 'All Sheets', 'entry_count': 0}
            })
        
        output_filename = f'tata_output_{uuid.uuid4().hex[:8]}.xlsx'
        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
        
        output_data = {
            "COMPANY NAME": [], "SEGMENT": [], "POLICY TYPE": [], "LOCATION": [],
            "PAYIN": [], "PAYOUT": [], "REMARK": [], "CALCULATION EXPLANATION": []
        }
        
        for entry in entries:
            output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
            output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
            output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
            output_data["LOCATION"].append(entry.get("LOCATION", ""))
            output_data["PAYIN"].append(entry.get("PAYIN", ""))
            output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
            output_data["REMARK"].append("; ".join(entry.get("REMARK", [])) if isinstance(entry.get("REMARK"), list) else str(entry.get("REMARK", "")))
            output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
        
        output_df = pd.DataFrame(output_data)
        output_df.to_excel(output_path, index=False, engine='openpyxl')
        
        download_cleanup[output_filename] = {
            'upload': None,
            'converted': None
        }

        return jsonify({
            'status': 'success_sheet_processed',
            'result': {
                'filename': output_filename,
                'sheet_name': 'All Sheets',
                'download_url': f'/processed/{output_filename}',
                'entry_count': len(entries)
            }
        })
        
    except Exception as e:
        print(f"✗ Tata error: {e}")
        cleanup_session_files(upload_path, converted_path)
        return jsonify({'detail': f'Error: {str(e)}'}), 500

@app.route('/processed/<filename>')
def download_processed_file(filename):
    """Serves file and deletes it after download"""
    try:
        response = send_from_directory(
            app.config['PROCESSED_FOLDER'],
            filename,
            as_attachment=True
        )
        
        @response.call_on_close
        def cleanup_after_download():
            import time
            time.sleep(0.5)
            
            # Delete the Excel file
            output_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
            safe_delete(output_path)
        
        return response
        
    except FileNotFoundError:
        return jsonify({'detail': 'File not found'}), 404

# @app.route('/cleanup_session', methods=['POST'])
# def cleanup_session():
#     try:
#         data = request.json
#         session_id = data.get('session_id')
#         cleanup_all = data.get('cleanup_all', False)
        
#         if session_id and session_id in download_cleanup:
#             paths = download_cleanup[session_id]
            
#             if cleanup_all:
#                 import time
#                 time.sleep(1)  # Wait for file handles to release
#                 cleanup_session_files(paths.get('upload'), paths.get('converted'))
                
#                 # Also cleanup all processed files for this session
#                 for output_file in list(download_cleanup.keys()):
#                     if output_file.startswith(('zuno_', 'bajaj_', 'liberty_')):
#                         output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_file)
#                         safe_delete(output_path)
            
#             del download_cleanup[session_id]
@app.route('/cleanup_session', methods=['POST'])
def cleanup_session():
    try:
        data = request.json
        session_id = data.get('session_id')
        cleanup_all = data.get('cleanup_all', False)
        
        if session_id and session_id in download_cleanup:
            paths = download_cleanup[session_id]
            
            if cleanup_all:
                import time
                import gc
                time.sleep(1)
                gc.collect()  # Release file handles BEFORE cleanup
                
                # Cleanup upload + converted
                cleanup_session_files(paths.get('upload'), paths.get('converted'))
                
                # Cleanup all processed files
                for output_file in list(download_cleanup.keys()):
                    if output_file.startswith(('zuno_', 'bajaj_', 'liberty_')):
                        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_file)
                        safe_delete(output_path)
                
                # NEW: Empty entire uploads folder
                for filename in os.listdir(app.config['UPLOAD_FOLDER']):
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    safe_delete(file_path)
            
            del download_cleanup[session_id]
            print(f"✓ Session cleanup complete: {session_id}")
        
        return jsonify({'status': 'success'})
    except Exception as e:
        print(f"✗ Cleanup error: {e}")
        return jsonify({'status': 'error', 'detail': str(e)}), 500

if __name__ == '__main__':
    print("Flask server starting...")
    print(f"Uploads: {app.config['UPLOAD_FOLDER']}")
    print(f"Converted: {app.config['CONVERTED_FOLDER']}")
    print(f"Processed: {app.config['PROCESSED_FOLDER']}")
    print("Open http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
