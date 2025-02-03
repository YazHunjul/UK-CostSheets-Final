import sqlite3
from datetime import datetime
import json
import pandas as pd
import os

class DatabaseManager:
    def __init__(self):
        self.db_path = 'data/canopy_submissions.db'
        # Ensure data directory exists
        os.makedirs('data', exist_ok=True)
        self.init_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS submissions
            (id INTEGER PRIMARY KEY AUTOINCREMENT,
             timestamp TEXT,
             project_name TEXT,
             
             -- Customer Details
             customer_name TEXT,
             site_address TEXT,
             contact_name TEXT,
             contact_number TEXT,
             email TEXT,
             
             -- Canopy Details
             canopy_type TEXT,
             canopy_length REAL,
             canopy_width REAL,
             canopy_height REAL,
             canopy_color TEXT,
             
             -- Additional Options
             guttering TEXT,
             side_panels TEXT,
             door_panels TEXT,
             
             -- Delivery & Installation
             delivery_location TEXT,
             delivery_lift_qty INTEGER,
             plant_hires TEXT,
             quantities TEXT,
             strip_out TEXT,
             
             -- Costs
             total_cost REAL,
             excel_path TEXT)
        ''')
        conn.commit()
        conn.close()

    def save_submission(self, data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''
            INSERT INTO submissions
            (timestamp, project_name, customer_name, site_address, 
             contact_name, contact_number, email, canopy_type, 
             canopy_length, canopy_width, canopy_height, canopy_color,
             guttering, side_panels, door_panels, delivery_location, 
             delivery_lift_qty, plant_hires, quantities, strip_out,
             total_cost, excel_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data.get('project_name', 'Unnamed Project'),
            data.get('customer_name', ''),
            data.get('site_address', ''),
            data.get('contact_name', ''),
            data.get('contact_number', ''),
            data.get('email', ''),
            data.get('canopy_type', ''),
            data.get('canopy_length', 0),
            data.get('canopy_width', 0),
            data.get('canopy_height', 0),
            data.get('canopy_color', ''),
            data.get('guttering', ''),
            data.get('side_panels', ''),
            data.get('door_panels', ''),
            data.get('delivery_location', ''),
            data.get('delivery_lift_qty', 0),
            json.dumps(data.get('plant_hires', {})),
            json.dumps(data.get('quantities', {})),
            data.get('strip_out', ''),
            data.get('total_cost', 0.0),
            data.get('excel_path', '')
        ))
        conn.commit()
        conn.close()

    def get_submissions(self, limit=50, search_term=None, date_from=None, date_to=None):
        conn = sqlite3.connect(self.db_path)
        query = 'SELECT * FROM submissions WHERE 1=1'
        params = []

        if search_term:
            query += ''' AND (
                project_name LIKE ? OR 
                delivery_location LIKE ? OR 
                plant_hires LIKE ?
            )'''
            search_pattern = f'%{search_term}%'
            params.extend([search_pattern, search_pattern, search_pattern])

        if date_from:
            query += ' AND timestamp >= ?'
            params.append(date_from)
        if date_to:
            query += ' AND timestamp <= ?'
            params.append(date_to)

        query += ' ORDER BY timestamp DESC LIMIT ?'
        params.append(limit)

        submissions = pd.read_sql_query(query, conn, params=params)
        conn.close()
        return submissions

    def delete_submission(self, submission_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('DELETE FROM submissions WHERE id = ?', (submission_id,))
        conn.commit()
        conn.close()

    def export_to_excel(self, filename='submissions_export.xlsx'):
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql_query('SELECT * FROM submissions', conn)
        conn.close()
        
        # Convert JSON strings to readable format
        df['plant_hires'] = df['plant_hires'].apply(json.loads)
        df['quantities'] = df['quantities'].apply(json.loads)
        
        # Export to Excel
        df.to_excel(filename, index=False)
        return filename 