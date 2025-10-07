import re

FONT_MAP = {
    'calibri': 'Helvetica',
    'calibri-bold': 'Helvetica-Bold',
    'calibri-italic': 'Helvetica-Oblique',
    'calibri-bolditalic': 'Helvetica-BoldOblique',
    'arial': 'Helvetica',
    'arial-bold': 'Helvetica-Bold',
    'times': 'Times-Roman',
    'times-bold': 'Times-Bold',
    'times-italic': 'Times-Italic',
    'courier': 'Courier',
}

def normalize_font(font_name):
    """Convert any font to ReportLab-compatible font"""
    if not font_name:
        return 'Helvetica'
    
    font_clean = re.sub(r'^[A-Z]+\+', '', font_name).lower()
    
    if font_clean in FONT_MAP:
        return FONT_MAP[font_clean]
    
    if 'calibri' in font_clean or 'arial' in font_clean:
        if 'bold' in font_clean and 'italic' in font_clean:
            return 'Helvetica-BoldOblique'
        elif 'bold' in font_clean:
            return 'Helvetica-Bold'
        elif 'italic' in font_clean:
            return 'Helvetica-Oblique'
        return 'Helvetica'
    
    if 'times' in font_clean:
        if 'bold' in font_clean and 'italic' in font_clean:
            return 'Times-BoldItalic'
        elif 'bold' in font_clean:
            return 'Times-Bold'
        elif 'italic' in font_clean:
            return 'Times-Italic'
        return 'Times-Roman'
    
    if 'courier' in font_clean:
        if 'bold' in font_clean:
            return 'Courier-Bold'
        return 'Courier'
    
    return 'Helvetica'
