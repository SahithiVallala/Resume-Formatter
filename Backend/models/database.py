import sqlite3
import json
from datetime import datetime
import os
import sys

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import Config

class TemplateDB:
    def __init__(self):
        self.db_path = Config.DATABASE
        self.init_db()
    
    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                filename TEXT NOT NULL,
                file_type TEXT NOT NULL,
                upload_date TEXT NOT NULL,
                format_data TEXT NOT NULL
            )
        ''')
        conn.commit()
        conn.close()
    
    def add_template(self, template_id, name, filename, file_type, format_data):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO templates (id, name, filename, file_type, upload_date, format_data)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (template_id, name, filename, file_type, datetime.now().isoformat(), json.dumps(format_data)))
        conn.commit()
        conn.close()
    
    def get_all_templates(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT id, name, filename, file_type, upload_date FROM templates')
        templates = [{'id': row[0], 'name': row[1], 'filename': row[2], 
                     'file_type': row[3], 'upload_date': row[4]} for row in cursor.fetchall()]
        conn.close()
        return templates
    
    def get_template(self, template_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM templates WHERE id = ?', (template_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return {
                'id': row[0],
                'name': row[1],
                'filename': row[2],
                'file_type': row[3],
                'upload_date': row[4],
                'format_data': json.loads(row[5])
            }
        return None
    
    def delete_template(self, template_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM templates WHERE id = ?', (template_id,))
        conn.commit()
        conn.close()
