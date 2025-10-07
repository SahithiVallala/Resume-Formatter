"""
Check template in database and verify it has required fields
"""
from models.database import TemplateDB
import json

db = TemplateDB()
templates = db.get_all_templates()

print(f"\n{'='*70}")
print(f"TEMPLATE DATABASE CHECK")
print(f"{'='*70}\n")

if not templates:
    print("❌ No templates found in database!")
    print("\n💡 Solution: Upload a new template through the web interface")
else:
    for template in templates:
        print(f"Template: {template['name']}")
        print(f"ID: {template['id']}")
        print(f"File: {template['filename']}")
        print(f"Type: {template['file_type']}")
        
        # Get full template data
        full_template = db.get_template(template['id'])
        format_data = full_template['format_data']
        
        print(f"\nFormat Data Keys: {list(format_data.keys())}")
        
        # Check for required fields
        has_template_path = 'template_path' in format_data
        has_template_type = 'template_type' in format_data
        
        print(f"✓ Has template_path: {has_template_path}")
        print(f"✓ Has template_type: {has_template_type}")
        
        if not has_template_path or not has_template_type:
            print(f"\n⚠️  WARNING: This template was uploaded with the OLD system!")
            print(f"💡 Solution: Delete this template and upload it again")
        else:
            print(f"\n✅ Template is compatible with new system!")
        
        print(f"\n{'─'*70}\n")

print(f"{'='*70}\n")
