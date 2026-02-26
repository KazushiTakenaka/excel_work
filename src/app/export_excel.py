import pandas as pd
import os

folder_path = os.path.join('見積書', '加工品')
excel_files = ['26AA0788_御見積書.xlsx', 'QTKG20260106A12-1.XLSX']
output_file = 'excel_content.md'

with open(output_file, 'w', encoding='utf-8') as f:
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        f.write(f"# File: {file}\n\n")
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                f.write(f"## Sheet: {sheet_name}\n\n")
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                # Drop fully empty columns/rows to clean up
                df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
                f.write(df.to_string(index=False))
                f.write("\n\n")
        except Exception as e:
            f.write(f"Error reading {file}: {e}\n\n")

print(f"Exported to {output_file}")
