from flask import Flask, request, render_template, redirect
import pandas as pd
from datetime import datetime
import os
from flask_cors import CORS
 
app = Flask(__name__)
CORS(app)
EXCEL_FILE = 'csr_log.xlsx'
 
# Initialize Excel if not present
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=['APM ID', 'Department', 'CSR Type', 'In-Time', 'Out-Time', 'Duration'])
    df.to_excel(EXCEL_FILE, index=False)
 
@app.route('/')
def scan_qr():
    return jsonify({"message":"Backend is Running!"})
 
@app.route('/submit', methods=['POST'])
def submit():
    data=request.get_json()
    apm_id = data.form['apm_id']
    department = data.form['department']
    csr_type = data.form['csr_type']
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
 
    df = pd.read_excel(EXCEL_FILE)
 
    # Check if entry with APM ID exists without out-time
    matching_entries = df[(df['APM ID'] == apm_id) & 
                          (df['Department'] == department) & 
                          (df['CSR Type'] == csr_type) & 
                          (df['Out-Time'].isna())]

    if not matching_entries.empty:
        # Update the latest matching entry with Out-Time and Duration
        idx = matching_entries.index[-1]
        in_time = datetime.strptime(df.loc[idx, 'In-Time'], '%Y-%m-%d %H:%M:%S')
        out_time = datetime.strptime(current_time, '%Y-%m-%d %H:%M:%S')
        duration = str(out_time - in_time)

        df.at[idx, 'Out-Time'] = current_time
        df.at[idx, 'Duration'] = duration
    else:
        # Create a new row for a fresh check-in
        new_entry = pd.DataFrame([[apm_id, department, csr_type, current_time, '', '']],
                                 columns=['APM ID', 'Department', 'CSR Type', 'In-Time', 'Out-Time', 'Duration'])
        df = pd.concat([df, new_entry], ignore_index=True)

 
    df.to_excel(EXCEL_FILE, index=False)
    return jsonify({"message":"Entry submitted successfully!"})
 
if __name__ == '__main__':
    app.run(debug=True)