import pandas as pd
import os

target_file = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品リスト.xlsx'

print(f"Reading {os.path.basename(target_file)}")
try:
    df = pd.read_excel(target_file)
    print("--- First 10 rows (including updated ones) ---")
    print(df.head(15).to_markdown())
except Exception as e:
    print(f"Error: {e}")
