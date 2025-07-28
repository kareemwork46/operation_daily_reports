import os
import pandas as pd
from datetime import datetime

def export_daily_report(df, file_path):
    df.to_excel(file_path, index=False)

def merge_weekly_reports(daily_folder, weekly_folder):
    today = datetime.today()
    week_number = today.strftime('%Y-W%U')
    merged_df = pd.DataFrame()

    for filename in os.listdir(daily_folder):
        if filename.endswith('.xlsx'):
            path = os.path.join(daily_folder, filename)
            daily_df = pd.read_excel(path)
            daily_df['Source File'] = filename
            merged_df = pd.concat([merged_df, daily_df], ignore_index=True)

    output_path = os.path.join(weekly_folder, f'Weekly_Report_{week_number}.xlsx')
    merged_df.to_excel(output_path, index=False)
    return output_path
