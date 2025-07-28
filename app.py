# app.py
from flask import Flask, render_template, request, redirect, send_file
import os
import pandas as pd
from datetime import datetime
from utils.export_excel import export_daily_report, merge_weekly_reports

app = Flask(__name__)
DAILY_FOLDER = 'daily_reports'
WEEKLY_FOLDER = 'weekly_reports'
os.makedirs(DAILY_FOLDER, exist_ok=True)
os.makedirs(WEEKLY_FOLDER, exist_ok=True)

@app.route('/')
def index():
    today = datetime.today().strftime('%Y-%m-%d')
    file_path = os.path.join(DAILY_FOLDER, f'Daily_Report_{today}.xlsx')
    existing_data = []

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        df = df.fillna('')  # <-- this line replaces NaN with ''
        existing_data = df.to_dict(orient='records')

    return render_template('index.html', data=existing_data)


@app.route('/submit', methods=['POST'])
def submit():
    data = {
        'SN': request.form.getlist('sn'),
        'Location Building': request.form.getlist('location_building'),
        'Location Floor': request.form.getlist('location_floor'),
        'Activity': request.form.getlist('activity'),
        'Actual Qty Done': request.form.getlist('actual_qty'),
        'Current Resources Trade': request.form.getlist('trade'),
        'Current Resources Qty': request.form.getlist('resource_qty'),
        'Current Resources Hrs.': request.form.getlist('resource_hrs'),
        'Current Resources Total Hrs.': request.form.getlist('resource_total_hrs'),
        'Remarks': request.form.getlist('remarks'),
    }
    df = pd.DataFrame(data)
    today = datetime.today().strftime('%Y-%m-%d')
    file_path = os.path.join(DAILY_FOLDER, f'Daily_Report_{today}.xlsx')
    export_daily_report(df, file_path)
    return f"Report saved as {file_path}"

@app.route('/merge_weekly')
def merge():
    file_path = merge_weekly_reports(DAILY_FOLDER, WEEKLY_FOLDER)
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
