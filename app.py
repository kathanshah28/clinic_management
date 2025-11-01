import os
import json
from datetime import datetime
from flask import Flask, jsonify, request, render_template
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
import string # <--- ADDED THIS IMPORT!

# Load .env
load_dotenv()

# --- CONFIGURATION ---
SERVICE_ACCOUNT_INFO = json.loads(os.environ['GOOGLE_CREDENTIALS'])
SHEET_ID = os.environ['SHEET_ID']
SHEET_NAME = os.environ['SHEET_NAME']

# Google Sheets credentials setup
credentials = service_account.Credentials.from_service_account_info(
    SERVICE_ACCOUNT_INFO,
    scopes=['https://www.googleapis.com/auth/spreadsheets']
)

service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()

app = Flask(__name__)

# --- Helper Function ---
def column_to_letter(col_index):
    """Convert 0-based column index to A1 letter notation."""
    letter = ''
    while col_index >= 0:
        remainder = col_index % 26
        letter = string.ascii_uppercase[remainder] + letter
        col_index = col_index // 26 - 1
    return letter

# ------------------- API Routes -------------------

# Get all patients
@app.route('/api/patients', methods=['GET'])
def get_patients():
    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify([])
        headers = values[0]
        patients = [dict(zip(headers, row)) for row in values[1:]]
        return jsonify(patients)
    except Exception as e:
        print(f"Error getting all patients: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/patients/today', methods=['GET'])
def get_today_patients():
    """Fetches today's patients and ensures the Patient_ID key is present for the frontend."""
    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify([])

        headers = values[0]
        patients = [dict(zip(headers, row)) for row in values[1:]]

        today_day = datetime.today().strftime('%A').lower()
        today_patients = []

        # Find the index for 'Patient ID' header
        patient_id_key = 'Patient ID' # Assuming your header is 'Patient ID' with a space
        
        for p in patients:
            visit_days_str = p.get('visit days', '') or '' 
            visit_days = [d.strip().lower() for d in visit_days_str.split(',') if d.strip()]
            
            if today_day in visit_days or 'daily' in visit_days:
                if 'Visit Count' not in p or not p['Visit Count']:
                    p['Visit Count'] = '0'
                
                # --- CRITICAL FIX FOR FRONTEND ---
                # The frontend expects 'Patient_ID' (with underscore)
                # We map the actual header ('Patient ID' with space) to the frontend key
                p['Patient_ID'] = p.get(patient_id_key, '') 
                # ---------------------------------
                
                today_patients.append(p)

        return jsonify(today_patients)
    except Exception as e:
        print(f"Error loading today's patients: {e}")
        return jsonify({'status': 'error', 'message': 'Failed to load data'}), 500


# Add new patient
@app.route('/api/patients', methods=['POST'])
def add_patient():
    data = request.json
    row = [
        data.get('Patient_ID',''),
        data.get('Name',''),
        data.get('number',''),
        data.get('Age',''),
        data.get('Gender',''),
        data.get('Occupation',''),
        data.get('Ref. by',''),
        data.get('Address',''),
        data.get('Date of joining',''),
        data.get('conditions',''),
        data.get('Time',''),
        ','.join(data.get('visit days',[])), 
        data.get('Visit Count','0'),
        'No'
    ]
    try:
        sheet.values().append(
            spreadsheetId=SHEET_ID,
            range=SHEET_NAME,
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values':[row]}
        ).execute()
        return jsonify({'status':'success'})
    except Exception as e:
        print(f"Error adding patient: {e}")
        return jsonify({'status':'error', 'error': str(e)}), 500


@app.route('/api/patients/<patient_id>/attend', methods=['PUT'])
def mark_attendance(patient_id):
    """
    Finds the patient by ID, increments their 'Visit Count', and updates 
    only that specific cell in the Google Sheet using dynamic range calculation.
    """
    data = request.json
    action = data.get('action', '').lower()

    if action != 'confirm':
        return jsonify({'status': 'ignored'})

    try:
        # 1. Fetch all values (including headers)
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify({'status': 'not found', 'message': 'No data in sheet'}), 404

        headers = values[0]
        data_rows = values[1:] 

        # 2. Determine column indexes (CRITICAL: Match these strings exactly to your headers)
        try:
            # Assumes your header is 'Patient ID' with a space
            patient_id_col_index = headers.index('Patient ID') 
            visit_count_col_index = headers.index('Visit Count')
        except ValueError as e:
            print(f"Header Error: Column not found in sheet: {e}")
            return jsonify({'status': 'error', 'message': f"Missing required column headers in sheet: {e}"}), 500

        updated = False
        
        # 3. Find the patient row
        for i, row in enumerate(data_rows, start=2): # i is the 1-based sheet row number (Row 2 is index 0)
            
            # Safely check for Patient ID match
            current_patient_id = str(row[patient_id_col_index]) if len(row) > patient_id_col_index else None

            if current_patient_id == patient_id:
                
                # Safely get current count, default to 0
                current_count_str = row[visit_count_col_index] if len(row) > visit_count_col_index else '0'
                
                try:
                    current_count = int(current_count_str or '0')
                except ValueError:
                    print(f"Warning: Non-numeric value found for Visit Count: '{current_count_str}'. Resetting to 0.")
                    current_count = 0 
                
                new_visit_count = current_count + 1

                # 4. Calculate the update range (e.g., 'Sheet1!C5')
                col_letter = column_to_letter(visit_count_col_index)
                range_to_update = f'{SHEET_NAME}!{col_letter}{i}'
                
                # 5. Update only the specific cell
                sheet.values().update(
                    spreadsheetId=SHEET_ID,
                    range=range_to_update,
                    valueInputOption='USER_ENTERED',
                    body={'values':[[new_visit_count]]}
                ).execute()
                
                updated = True
                break
        
        return jsonify({'status':'updated' if updated else 'not found'})

    except Exception as e:
        print(f"--- FATAL ERROR IN MARK_ATTENDANCE ---\n{e}\n--------------------------------------")
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ------------------- Frontend Routes -------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/today')
def today_page():
    return render_template('today_patients.html')

@app.route('/add')
def add_patient_page():
    return render_template('add_patients.html')

@app.route('/history')
def history_page():
    return render_template('all_patients.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)