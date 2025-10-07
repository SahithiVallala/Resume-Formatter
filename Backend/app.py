from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import uuid
import os
import traceback

from config import Config
from models.database import TemplateDB
from utils.advanced_template_analyzer import analyze_template
from utils.advanced_resume_parser import parse_resume
from utils.intelligent_formatter import format_resume_intelligent

app = Flask(__name__)
app.config.from_object(Config)
Config.init_app(app)

# Enable CORS for React frontend (allow multiple localhost ports)
CORS(
    app,
    resources={r"/api/*": {
        "origins": [
            "http://localhost:3000",
            "http://localhost:3001",
            "http://127.0.0.1:3000",
            "http://127.0.0.1:3001"
        ]
    }}
)

db = TemplateDB()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/api/templates', methods=['GET'])
def get_templates():
    """Get all templates"""
    templates = db.get_all_templates()
    return jsonify({'success': True, 'templates': templates})

@app.route('/api/templates', methods=['POST'])
def upload_template():
    """Upload and analyze template"""
    try:
        if 'template_file' not in request.files or 'template_name' not in request.form:
            return jsonify({'success': False, 'message': 'Missing file or name'}), 400
        
        file = request.files['template_file']
        name = request.form['template_name'].strip()
        
        if file.filename == '' or not name:
            return jsonify({'success': False, 'message': 'Invalid input'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'message': 'Invalid file type'}), 400
        
        # Save file
        filename = secure_filename(file.filename)
        file_type = filename.rsplit('.', 1)[1].lower()
        template_id = str(uuid.uuid4())
        saved_filename = f"{template_id}_{filename}"
        file_path = os.path.join(Config.TEMPLATE_FOLDER, saved_filename)
        file.save(file_path)
        
        print(f"\n{'='*70}")
        print(f"📤 UPLOADING TEMPLATE: {name}")
        print(f"{'='*70}\n")
        
        # Analyze template with advanced analyzer
        format_data = analyze_template(file_path)
        
        # Save to database
        db.add_template(template_id, name, saved_filename, file_type, format_data)
        
        return jsonify({
            'success': True,
            'id': template_id,
            'name': name,
            'message': 'Template uploaded and analyzed successfully'
        })
    
    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/format', methods=['POST'])
def format_resumes():
    """Format resumes using selected template"""
    try:
        if 'template_id' not in request.form or 'resume_files' not in request.files:
            return jsonify({'success': False, 'message': 'Missing template or files'}), 400
        
        template_id = request.form['template_id']
        template = db.get_template(template_id)
        
        if not template:
            return jsonify({'success': False, 'message': 'Template not found'}), 404
        
        files = request.files.getlist('resume_files')
        formatted_files = []
        
        # Get template analysis
        template_analysis = template['format_data']
        
        # Ensure paths are set
        template_file_path = os.path.join(Config.TEMPLATE_FOLDER, template['filename'])
        template_analysis['template_path'] = template_file_path
        template_analysis['template_type'] = template['file_type']
        
        print(f"\n{'='*70}")
        print(f"🎯 FORMATTING SESSION")
        print(f"{'='*70}")
        print(f"📋 Template: {template['name']}")
        print(f"📁 Template Path: {template_file_path}")
        print(f"✓ Template Exists: {os.path.exists(template_file_path)}")
        print(f"📊 Resumes to Process: {len(files)}")
        print(f"{'='*70}\n")
        
        for idx, file in enumerate(files, 1):
            if file.filename == '' or not allowed_file(file.filename):
                continue
            
            # Save resume
            filename = secure_filename(file.filename)
            file_type = filename.rsplit('.', 1)[1].lower()
            resume_id = str(uuid.uuid4())
            saved_filename = f"{resume_id}_{filename}"
            file_path = os.path.join(Config.RESUME_FOLDER, saved_filename)
            file.save(file_path)
            
            print(f"\n{'─'*70}")
            print(f"📄 Processing Resume {idx}/{len(files)}: {filename}")
            print(f"{'─'*70}")
            
            # Parse resume with advanced parser
            resume_data = parse_resume(file_path, file_type)
            
            if resume_data:
                # Format resume with intelligent formatter
                # Determine output extension based on template type
                if template['file_type'] in ['doc', 'docx']:
                    output_ext = '.docx'
                else:
                    output_ext = '.pdf'
                
                output_filename = f"formatted_{resume_id}{output_ext}"
                output_path = os.path.join(Config.OUTPUT_FOLDER, output_filename)
                
                if format_resume_intelligent(resume_data, template_analysis, output_path):
                    # Check what file was actually created
                    actual_files = []
                    for ext in ['.pdf', '.docx']:
                        check_path = os.path.join(Config.OUTPUT_FOLDER, f"formatted_{resume_id}{ext}")
                        if os.path.exists(check_path):
                            actual_files.append(f"formatted_{resume_id}{ext}")
                    
                    # Use the first file found (prefer PDF if both exist)
                    if actual_files:
                        final_filename = actual_files[0] if '.pdf' not in str(actual_files) else [f for f in actual_files if '.pdf' in f][0] if any('.pdf' in f for f in actual_files) else actual_files[0]
                        formatted_files.append({
                            'filename': final_filename,
                            'original': filename,
                            'name': resume_data['name']
                        })
                        print(f"✅ Successfully formatted: {filename} → {final_filename}\n")
                    else:
                        print(f"⚠️  Formatting completed but output file not found\n")
                else:
                    print(f"❌ Failed to format: {filename}\n")
            else:
                print(f"❌ Failed to parse resume: {filename}\n")
            
            # Cleanup
            try:
                os.remove(file_path)
            except:
                pass
        
        print(f"{'='*70}")
        print(f"✅ FORMATTING COMPLETE: {len(formatted_files)}/{len(files)} successful")
        print(f"{'='*70}\n")
        
        return jsonify({
            'success': True,
            'files': formatted_files,
            'message': f'Formatted {len(formatted_files)} resume(s)'
        })
    
    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/download/<filename>')
def download_file(filename):
    """Download formatted resume"""
    return send_from_directory(Config.OUTPUT_FOLDER, filename, as_attachment=True)

@app.route('/api/templates/<template_id>', methods=['DELETE'])
def delete_template(template_id):
    """Delete template"""
    try:
        template = db.get_template(template_id)
        if template:
            file_path = os.path.join(Config.TEMPLATE_FOLDER, template['filename'])
            if os.path.exists(file_path):
                os.remove(file_path)
            db.delete_template(template_id)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    print("\n" + "="*70)
    print("🎯 RESUME FORMATTER - BACKEND SERVER")
    print("="*70)
    print("✅ API running on http://127.0.0.1:5000")
    print("✅ React frontend: http://localhost:3000")
    print("="*70 + "\n")
    app.run(debug=True, port=5000)
