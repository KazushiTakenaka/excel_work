import pandas as pd
import os

target_file = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品リスト.xlsx'

print(f"Reading headers from: {os.path.basename(target_file)}")
try:
    df = pd.read_excel(target_file, nrows=5)
    print("--- Columns ---")
    print(df.columns.tolist())
    print("\n--- First few rows ---")
    print(df.head().to_markdown(index=False))
except Exception as e:
    print(f"Error: {e}")
