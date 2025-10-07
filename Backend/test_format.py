"""
Quick test script to debug the formatting issue
"""
import os
import sys
from models.database import TemplateDB
from utils.resume_parser import parse_resume
from utils.formatter import format_resume

# Get the first template
db = TemplateDB()
templates = db.get_all_templates()

if not templates:
    print("ERROR: No templates found in database!")
    sys.exit(1)

template = db.get_template(templates[0]['id'])
print(f"\n=== Template Info ===")
print(f"Name: {template['name']}")
print(f"Filename: {template['filename']}")
print(f"Type: {template['file_type']}")

# Check template file exists
from config import Config
template_path = os.path.join(Config.TEMPLATE_FOLDER, template['filename'])
print(f"Template path: {template_path}")
print(f"Template exists: {os.path.exists(template_path)}")

# Update format data with correct path
template_format = template['format_data']
template_format['template_path'] = template_path
template_format['template_type'] = template['file_type']

print(f"\n=== Format Data ===")
print(f"Template path in format_data: {template_format.get('template_path')}")
print(f"Template type in format_data: {template_format.get('template_type')}")
print(f"Sections: {len(template_format.get('sections', []))}")

# Test with a sample resume (you'll need to provide a path)
print("\n=== Testing Resume Parsing ===")
resume_folder = Config.RESUME_FOLDER
print(f"Resume folder: {resume_folder}")

# List files in resume folder
if os.path.exists(resume_folder):
    files = os.listdir(resume_folder)
    print(f"Files in resume folder: {files}")
else:
    print("Resume folder doesn't exist!")

print("\n=== Test Complete ===")
print("To test formatting, place a resume file in the resumes folder and update this script.")
