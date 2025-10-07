import pdfplumber
import PyPDF2
import re
from docx import Document
from .font_mapper import normalize_font

def analyze_pdf_template(file_path):
    """Extract formatting details from PDF template"""
    try:
        with pdfplumber.open(file_path) as pdf:
            page = pdf.pages[0]
            chars = page.chars
            
            if not chars:
                format_data = get_default_format()
                format_data['template_path'] = file_path
                format_data['template_type'] = 'pdf'
                return format_data
            
            lines = group_chars_into_lines(chars)
            real_lines = [l for l in lines if not is_placeholder(l['text'])]
            
            page_width = page.width
            page_height = page.height
            
            margins = detect_margins(real_lines, page_width, page_height)
            name_style = detect_name_style(real_lines, page_width)
            sections = detect_sections(real_lines)
            body_style = detect_body_style(real_lines)
            
            return {
                'template_path': file_path,
                'template_type': 'pdf',
                'page': {
                    'width': page_width,
                    'height': page_height,
                    'margins': margins
                },
                'name': name_style,
                'sections': sections,
                'body': body_style
            }
    except Exception as e:
        print(f"Error analyzing PDF: {e}")
        format_data = get_default_format()
        format_data['template_path'] = file_path
        format_data['template_type'] = 'pdf'
        return format_data

def get_default_format():
    return {
        'template_path': None,
        'template_type': None,
        'page': {
            'width': 612,
            'height': 792,
            'margins': {'top': 54, 'bottom': 54, 'left': 54, 'right': 54}
        },
        'name': {
            'font': 'Helvetica-Bold',
            'size': 14,
            'alignment': 'center'
        },
        'sections': [],
        'body': {
            'font': 'Helvetica',
            'size': 10,
            'line_spacing': 14
        }
    }

def group_chars_into_lines(chars):
    lines = []
    current_line = []
    current_y = None
    
    for char in sorted(chars, key=lambda c: (c['top'], c['x0'])):
        if current_y is None or abs(char['top'] - current_y) <= 2:
            current_line.append(char)
            current_y = char['top'] if current_y is None else current_y
        else:
            if current_line:
                lines.append(create_line(current_line))
            current_line = [char]
            current_y = char['top']
    
    if current_line:
        lines.append(create_line(current_line))
    
    return lines

def create_line(chars):
    text = ''.join([c['text'] for c in chars])
    return {
        'text': text.strip(),
        'x': chars[0]['x0'],
        'y': chars[0]['top'],
        'font': chars[0].get('fontname', 'Helvetica'),
        'font_size': chars[0].get('size', 10)
    }

def is_placeholder(text):
    placeholders = ['your name', 'your degree', 'phone number', 'email address',
                   'insert', 'replace', 'template', 'example']
    return any(p in text.lower() for p in placeholders) or len(text) > 120

def detect_margins(lines, width, height):
    if not lines:
        return {'top': 54, 'bottom': 54, 'left': 54, 'right': 54}
    
    xs = [l['x'] for l in lines]
    ys = [l['y'] for l in lines]
    
    return {
        'top': min(ys) if ys else 54,
        'bottom': height - max(ys) if ys else 54,
        'left': min(xs) if xs else 54,
        'right': width - max(xs) if xs else 54
    }

def detect_name_style(lines, page_width):
    if not lines:
        return {'font': 'Helvetica-Bold', 'size': 14, 'alignment': 'center'}
    
    first_line = lines[0]
    alignment = 'center' if abs(first_line['x'] - page_width/2) < 50 else 'left'
    
    return {
        'font': normalize_font(first_line['font']),
        'size': first_line['font_size'],
        'alignment': alignment
    }

def detect_sections(lines):
    sections = []
    seen = set()
    
    for i, line in enumerate(lines):
        text = line['text'].strip()
        
        if len(text) > 60 or len(text) < 3:
            continue
        
        if text.lower() in seen:
            continue
        
        if line['font_size'] >= 10 and not text.startswith('•'):
            sections.append({
                'heading': text,
                'font': normalize_font(line['font']),
                'size': line['font_size'],
                'has_underline': True
            })
            seen.add(text.lower())
    
    return sections[:10]

def detect_body_style(lines):
    body_lines = [l for l in lines if l['font_size'] < 12]
    
    if body_lines:
        sizes = [l['font_size'] for l in body_lines]
        common_size = max(set(sizes), key=sizes.count)
        return {
            'font': normalize_font(body_lines[0]['font']),
            'size': common_size,
            'line_spacing': 14
        }
    
    return {'font': 'Helvetica', 'size': 10, 'line_spacing': 14}

def analyze_word_template(file_path):
    """Extract formatting from Word template"""
    try:
        # Try to open as .docx first, then fall back to other methods
        doc = Document(file_path)
        
        sections = []
        for para in doc.paragraphs[:20]:
            if para.text.strip() and len(para.text) < 60:
                sections.append({
                    'heading': para.text.strip(),
                    'font': 'Helvetica-Bold',
                    'size': 11,
                    'has_underline': True
                })
        
        return {
            'template_path': file_path,
            'template_type': 'docx',
            'page': {
                'width': 612,
                'height': 792,
                'margins': {'top': 54, 'bottom': 54, 'left': 54, 'right': 54}
            },
            'name': {'font': 'Helvetica-Bold', 'size': 14, 'alignment': 'center'},
            'sections': sections[:8],
            'body': {'font': 'Helvetica', 'size': 10, 'line_spacing': 14}
        }
    except Exception as e:
        print(f"Warning: Could not analyze Word template '{file_path}': {e}")
        print("Using default formatting instead.")
        format_data = get_default_format()
        format_data['template_path'] = file_path
        format_data['template_type'] = 'docx'
        return format_data
