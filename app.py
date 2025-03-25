from flask import Flask, request, render_template, redirect
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
EXCEL_FILE = 'csr_log.xlsx'

# Initialize Excel if not present
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=['APM ID', 'Department', 'CSR Type', 'In-Time', 'Out-Time', 'Duration'])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def scan_qr():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    apm_id = request.form['apm_id']
    department = request.form['department']
    csr_type = request.form['csr_type']
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    df = pd.read_excel(EXCEL_FILE)

    # Check if entry with APM ID exists without out-time
    if apm_id in df['APM ID'].values:
        idx = df[df['APM ID'] == apm_id].last_valid_index()
        if pd.isna(df.loc[idx, 'Out-Time']):
            in_time = datetime.strptime(df.loc[idx, 'In-Time'], '%Y-%m-%d %H:%M:%S')
            out_time = datetime.strptime(current_time, '%Y-%m-%d %H:%M:%S')
            duration = str(out_time - in_time)
            df.loc[idx, 'Out-Time'] = current_time
            df.loc[idx, 'Duration'] = duration
        else:
            # Start a new row for new check-in
            df.loc[len(df)] = [apm_id, department, csr_type, current_time, '', '']
    else:
        # New Entry
        df.loc[len(df)] = [apm_id, department, csr_type, current_time, '', '']

    df.to_excel(EXCEL_FILE, index=False)
    return "Entry submitted successfully!"

if __name__ == '__main__':
    app.run(debug=True)
