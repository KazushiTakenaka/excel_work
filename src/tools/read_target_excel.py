
import pandas as pd
import os
import sys

if len(sys.argv) > 1:
    excel_path = sys.argv[1]
else:
    excel_path = os.path.join('見積書', 'QTKG20260106A12-1.XLSX')

if not os.path.exists(excel_path):
    print(f"Error: File not found at {excel_path}")
    exit(1)

print(f"--- Extracting text from {excel_path} ---")

try:
    # Read all sheets
    xls = pd.ExcelFile(excel_path)
    for sheet_name in xls.sheet_names:
        print(f"\n--- Sheet: {sheet_name} ---")
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # Print non-empty rows
        for index, row in df.iterrows():
            # Filter out completely empty rows
            if not row.dropna().empty:
                # Convert row to string representation with handling for NaN
                row_str = " | ".join([str(val) if pd.notna(val) else "" for val in row])
                print(f"Row {index}: {row_str}")

except Exception as e:
    print(f"Error reading Excel: {e}")
