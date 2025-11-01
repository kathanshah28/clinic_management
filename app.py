import os
import ssl
import certifi
import string
from datetime import datetime
from flask import Flask, request, jsonify, render_template
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

# ------------------- SSL Fix (Windows Certificate Issue) -------------------
ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())

# ------------------- Load Environment Variables -------------------
load_dotenv()

GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")

print("Service account file path:", GOOGLE_CREDENTIALS_FILE)
if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
    raise FileNotFoundError("Service account file not found. Check your .env or file path.")
print("Google Sheet ID:", SHEET_ID)
print("Worksheet name:", SHEET_NAME)

# ------------------- Google Sheets Setup -------------------
credentials = service_account.Credentials.from_service_account_file(
    GOOGLE_CREDENTIALS_FILE,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
service = build("sheets", "v4", credentials=credentials)
sheet = service.spreadsheets()

# ------------------- Flask App -------------------
app = Flask(__name__)

# ------------------- Helper Functions -------------------
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


# Get today's patients
@app.route('/api/patients/today', methods=['GET'])
def get_today_patients():
    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify([])

        headers = values[0]
        patients = [dict(zip(headers, row)) for row in values[1:]]

        today_full = datetime.today().strftime('%A').lower()  # e.g., saturday
        today_short = today_full[:3]  # e.g., sat

        today_patients = []

        for p in patients:
            visit_days_str = (p.get('Visit Days', '') or '').lower().replace('\n', ' ').replace(';', ',')
            visit_days = [d.strip() for d in visit_days_str.split(',') if d.strip()]

            # Flexible match: allow short/full form or 'daily'
            if any([
                today_full in visit_days,
                today_short in visit_days,
                any(today_full.startswith(v[:3]) for v in visit_days),
                'daily' in visit_days
            ]):
                p['Visit Count'] = p.get('Visit Count', '0')
                p['Patient_ID'] = p.get('Patient_ID', '')
                today_patients.append(p)

        print(f"✅ Today's patients count: {len(today_patients)}")
        return jsonify(today_patients)

    except Exception as e:
        print(f"❌ Error loading today's patients: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


# Add new patient (auto-clean Visit Days)
@app.route('/api/patients', methods=['POST'])
def add_patient():
    data = request.json

    # Clean Visit Days
    raw_days = data.get('Visit Days', [])
    visit_days_clean = [v.strip().capitalize() for v in raw_days if v.strip()]

    row = [
        data.get('Patient_ID', ''),
        data.get('Name', ''),
        data.get('Number', ''),
        data.get('Age', ''),
        data.get('Gender', ''),
        data.get('Occupation', ''),
        data.get('Ref.by', ''),
        data.get('Address', ''),
        data.get('Date of joining', ''),
        data.get('Conditions', ''),
        data.get('Time', ''),
        ', '.join(visit_days_clean),
        data.get('Visit Count', '0'),
    ]

    try:
        sheet.values().append(
            spreadsheetId=SHEET_ID,
            range=SHEET_NAME,
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [row]}
        ).execute()
        return jsonify({'status': 'success'})
    except Exception as e:
        print(f"Error adding patient: {e}")
        return jsonify({'status': 'error', 'error': str(e)}), 500


# Mark attendance / increment visit count
@app.route('/api/patients/<patient_id>/attend', methods=['PUT'])
def mark_attendance(patient_id):
    data = request.json
    action = data.get('action', '').lower()
    if action != 'confirm':
        return jsonify({'status': 'ignored'})

    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify({'status': 'not found', 'message': 'No data in sheet'}), 404

        headers = values[0]
        data_rows = values[1:]

        try:
            patient_id_col_index = headers.index('Patient_ID')
            visit_count_col_index = headers.index('Visit Count')
        except ValueError as e:
            return jsonify({'status': 'error', 'message': f"Missing required column headers: {e}"}), 500

        updated = False

        for i, row in enumerate(data_rows, start=2):  # Row 2 = first data row
            current_patient_id = str(row[patient_id_col_index]) if len(row) > patient_id_col_index else None
            if current_patient_id == patient_id:
                current_count_str = row[visit_count_col_index] if len(row) > visit_count_col_index else '0'
                try:
                    current_count = int(current_count_str or '0')
                except ValueError:
                    current_count = 0
                new_visit_count = current_count + 1

                col_letter = column_to_letter(visit_count_col_index)
                range_to_update = f'{SHEET_NAME}!{col_letter}{i}'

                sheet.values().update(
                    spreadsheetId=SHEET_ID,
                    range=range_to_update,
                    valueInputOption='USER_ENTERED',
                    body={'values': [[new_visit_count]]}
                ).execute()
                updated = True
                break

        return jsonify({'status': 'updated' if updated else 'not found'})

    except Exception as e:
        print(f"Error in mark_attendance: {e}")
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

# ------------------- Run App -------------------
if __name__ == '__main__':

    port = int(os.environ.get('PORT', 5000))
    app.run(host='127.0.0.1', port=port, debug=False)

