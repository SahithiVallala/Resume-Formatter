import os

class Config:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
    TEMPLATE_FOLDER = os.path.join(UPLOAD_FOLDER, 'templates')
    RESUME_FOLDER = os.path.join(UPLOAD_FOLDER, 'resumes')
    OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')
    DATABASE = os.path.join(BASE_DIR, 'templates.db')
    
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    
    ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx'}
    
    @staticmethod
    def init_app(app):
        for folder in [Config.TEMPLATE_FOLDER, Config.RESUME_FOLDER, Config.OUTPUT_FOLDER]:
            os.makedirs(folder, exist_ok=True)
