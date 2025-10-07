import pdfplumber
import re
from docx import Document
from collections import defaultdict

def parse_resume(file_path, file_type):
    """Extract content from resume"""
    try:
        if file_type == 'pdf':
            return parse_pdf_resume(file_path)
        else:
            return parse_word_resume(file_path)
    except Exception as e:
        print(f"Error parsing resume: {e}")
        return None

def parse_pdf_resume(file_path):
    try:
        with pdfplumber.open(file_path) as pdf:
            text = '\n'.join([p.extract_text() or '' for p in pdf.pages])
        return extract_resume_content(text)
    except:
        return None

def parse_word_resume(file_path):
    try:
        doc = Document(file_path)
        text = '\n'.join([p.text for p in doc.paragraphs])
        return extract_resume_content(text)
    except:
        return None

def extract_resume_content(text):
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    name = ''
    for line in lines[:5]:
        if not has_contact_info(line) and len(line) > 3:
            name = line
            break
    
    contact_text = ' '.join(lines[:15])
    email = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', contact_text)
    phone = re.search(r'\d{3}[-.\s]?\d{3}[-.\s]?\d{4}', contact_text)
    linkedin = re.search(r'linkedin\.com[^\s]*', contact_text, re.IGNORECASE)
    
    sections = extract_sections(lines)
    
    return {
        'name': name,
        'email': email.group(0) if email else '',
        'phone': phone.group(0) if phone else '',
        'linkedin': linkedin.group(0) if linkedin else '',
        'sections': sections
    }

def has_contact_info(text):
    return bool(re.search(r'@|http|linkedin|\d{3}[-.\s]\d{3}', text, re.IGNORECASE))

def extract_sections(lines):
    sections = defaultdict(list)
    current_section = None
    
    for line in lines:
        if len(line) < 60 and not line.startswith('•') and not has_contact_info(line):
            current_section = line.lower().strip()
        elif current_section and line:
            sections[current_section].append(line)
    
    return dict(sections)
