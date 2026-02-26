import pandas as pd
import os

file_path = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品\QTKG20260106A12-1.XLSX'

print(f"Reading {os.path.basename(file_path)}")
try:
    df = pd.read_excel(file_path, header=None)
    # Flatten the dataframe and print non-null values
    print("--- Extracted Text ---")
    for val in df.values.flatten():
         if pd.notna(val):
             print(val)
except Exception as e:
    print(e)
