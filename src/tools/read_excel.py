import pandas as pd
import os

folder_path = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品'
excel_files = ['26AA0788_御見積書.xlsx', 'QTKG20260106A12-1.XLSX']

for file in excel_files:
    file_path = os.path.join(folder_path, file)
    print(f"--- Reading {file} ---")
    try:
        # Load all sheets
        xls = pd.ExcelFile(file_path)
        print(f"Sheet names: {xls.sheet_names}")
        
        for sheet_name in xls.sheet_names:
            print(f"\nSheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Print first 20 rows to get an idea of content
            print(df.head(20).to_string())
            print("-" * 20)
    except Exception as e:
        print(f"Error reading {file}: {e}")
