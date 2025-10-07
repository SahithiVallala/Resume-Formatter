"""
Direct test of the formatting system
"""
import os
import sys
from models.database import TemplateDB
from utils.advanced_resume_parser import parse_resume
from utils.intelligent_formatter import format_resume_intelligent
from config import Config

print("\n" + "="*70)
print("DIRECT FORMATTING TEST")
print("="*70 + "\n")

# Get template
db = TemplateDB()
templates = db.get_all_templates()

if not templates:
    print("❌ No templates in database!")
    sys.exit(1)

template = db.get_template(templates[0]['id'])
print(f"✓ Using template: {template['name']}")

# Set up template analysis
template_analysis = template['format_data']
template_file_path = os.path.join(Config.TEMPLATE_FOLDER, template['filename'])
template_analysis['template_path'] = template_file_path
template_analysis['template_type'] = template['file_type']

print(f"✓ Template path: {template_file_path}")
print(f"✓ Template exists: {os.path.exists(template_file_path)}")
print(f"✓ Template type: {template_analysis['template_type']}")

# Check for a test resume
resume_path = input("\nEnter path to a test resume (PDF or DOCX): ").strip().strip('"')

if not os.path.exists(resume_path):
    print(f"❌ Resume not found: {resume_path}")
    sys.exit(1)

file_type = os.path.splitext(resume_path)[1].lower().replace('.', '')
print(f"\n✓ Resume file: {os.path.basename(resume_path)}")
print(f"✓ File type: {file_type}")

# Parse resume
print("\n" + "-"*70)
print("PARSING RESUME...")
print("-"*70)

try:
    resume_data = parse_resume(resume_path, file_type)
    
    if not resume_data:
        print("❌ Failed to parse resume!")
        sys.exit(1)
    
    print(f"\n✅ Resume parsed successfully!")
    
    # Format resume
    print("\n" + "-"*70)
    print("FORMATTING RESUME...")
    print("-"*70)
    
    output_path = os.path.join(Config.OUTPUT_FOLDER, "test_formatted.pdf")
    
    success = format_resume_intelligent(resume_data, template_analysis, output_path)
    
    if success:
        print(f"\n✅ SUCCESS! Formatted resume saved to:")
        print(f"   {output_path}")
    else:
        print(f"\n❌ Formatting failed!")
        
except Exception as e:
    print(f"\n❌ ERROR: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "="*70 + "\n")
