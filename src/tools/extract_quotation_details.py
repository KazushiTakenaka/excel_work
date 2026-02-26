import os
import pandas as pd
import pdfplumber
import easyocr
import numpy as np
import re
import logging

# Suppress warnings
logging.getLogger('easyocr').setLevel(logging.ERROR)

TARGET_DIR = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品'

# OCR Reader (initialized lazily)
_ocr_reader = None

def get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        print("Initializing EasyOCR...")
        _ocr_reader = easyocr.Reader(['ja', 'en'])
    return _ocr_reader

def clean_text(text):
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
    return text.strip().replace('　', ' ')

def extract_from_excel(file_path):
    print(f"\nProcessing Excel: {os.path.basename(file_path)}")
    try:
        df = pd.read_excel(file_path, header=None)
        
        # Determine header row
        header_row_idx = -1
        col_map = {}
        
        # Searching for header keywords
        keywords = {
            'part_no': [r'図番', r'図\s*番', r'品番'],
            'name': [r'品名', r'品\s*名', r'名称', r'商品名'],
            'unit_price': [r'単価', r'単\s*価'],
            'amount': [r'金額', r'金\s*額', r'小計']
        }
        
        for idx, row in df.iterrows():
            row_text = [clean_text(str(val)) for val in row.values]
            found_cols = {}
            
            for col_idx, text in enumerate(row_text):
                for key, patterns in keywords.items():
                    for pattern in patterns:
                        if re.search(pattern, text):
                            found_cols[key] = col_idx
                            break
            
            # If we matched at least 3 distinct keys, assume this is the header
            if len(set(found_cols.keys())) >= 2: # Relaxed to 2 for better hit rate
                header_row_idx = idx
                col_map = found_cols
                break
        
        if header_row_idx == -1:
            print("  Could not detect header row.")
            return []
            
        print(f"  Header detected at row {header_row_idx+1}. Columns: {col_map}")
        
        results = []
        # Iterate rows after header
        for idx in range(header_row_idx + 1, len(df)):
            row = df.iloc[idx]
            
            item = {}
            # Extract available columns
            for key, col_idx in col_map.items():
                val = row[col_idx]
                if pd.notna(val):
                    item[key] = val
            
            # Validate if it looks like a data row (must have at least one key field)
            if item.get('part_no') or item.get('name'):
                 results.append(item)
        
        return results

    except Exception as e:
        print(f"  Error reading Excel: {e}")
        return []

def extract_from_pdf(file_path):
    print(f"\nProcessing PDF: {os.path.basename(file_path)}")
    reader = get_ocr_reader()
    results = []
    
    try:
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                # Using OCR as default for this specific task since we know it's likely image-based
                # Only if extract_text fails or returns little text
                text = page.extract_text()
                if not text or len(text) < 50:
                    print(f"  Page {i+1}: Running OCR...")
                    im = page.to_image(resolution=300).original
                    img_np = np.array(im)
                    ocr_result = reader.readtext(img_np, detail=0)
                    
                    # Simple heuristic parsing for the specific format seen in logs
                    # Expected line format roughly: "TEM2521... Name... Material... Qty... UnitPrice... Amount"
                    # But OCR splits lines arbitrarily.
                    
                    # Let's look for lines that start with typical part number patterns
                    # or lines that appear to be data rows.
                    
                    # Combine all text to single string to handle multiline splits if needed, 
                    # but line-by-line check is safer for tabular data if OCR respects layout.
                    
                    current_item = {}
                    
                    for line in ocr_result:
                        line = line.strip()
                        # Heuristic: Part number often starts with alphabets and contains numbers
                        # e.g., TEM2521, TKM...
                        
                        # Case 1: Row starting with Part Number
                        # Regex for Part No: [A-Z]+[0-9]+.*
                        if re.match(r'^[A-Z]+\d+', line):
                            # This line likely contains the part number and maybe more
                            parts = line.split()
                            current_item['part_no'] = parts[0]
                            # Sometimes name is attached
                            if len(parts) > 1:
                                current_item['name_fragment'] = " ".join(parts[1:])
                        
                        # Case 2: Price / Amount (Numbers with commas or endings like "-")
                        # e.g., 9,000-
                        match_price = re.findall(r'[\d,]+-?', line)
                        if match_price:
                            # Filter logical numbers strings
                            nums = [p for p in match_price if any(c.isdigit() for c in p)]
                            if nums:
                                # Assuming last two numbers in a data block are Unit Price and Amount
                                # This is weak but might work for this specific document
                                pass
                        
                        # NOTE: Parsing raw OCR list is hard without spatial info.
                        # Ideally, we should use detail=1 to get bounding boxes.
                        
                    # Re-implementing with spatial awareness if standard readtext fails to capture structure
                    # For now, let's just dump the OCR lines that look like data rows
                    
                    # Pattern based extraction from list
                    data_rows = []
                    for line in ocr_result:
                         # Filter for lines that look like our data
                         # e.g. TEM2521...
                         if re.match(r'^[A-Z0-9_\-]+', line) and len(line) > 5:
                             data_rows.append(line)
                    
                    # Since we can't perfectly parse without more distinct rules, 
                    # we will return the raw rows that look like items for now.
                    for row in data_rows:
                        # Attempt to split roughly
                        # This is highly specific to the sample provided in logs
                        # "TEM2521_70-POO1lカバー01" -> Part: TEM2521_70-POO1, Name: lカバー01 (OCR error likely)
                        
                        # Very naive split based on underscores or spaces
                        # Checking user's specific file log: "TEM2521_70-POO1lカバー01" seems concatenated?
                        
                        item = {'raw_line': row}
                        results.append(item)
                        
                else:
                    # Text based PDF parsing
                    lines = text.split('\n')
                    for line in lines:
                        # Similar heuristic
                        pass

    except Exception as e:
        print(f"  Error reading PDF: {e}")
    
    return results

def main():
    if not os.path.exists(TARGET_DIR):
        print(f"Directory not found: {TARGET_DIR}")
        return

    files = os.listdir(TARGET_DIR)
    
    all_data = {}

    for file in files:
        path = os.path.join(TARGET_DIR, file)
        if file.lower().endswith(('.xlsx', '.xls')):
            data = extract_from_excel(path)
            all_data[file] = data
        elif file.lower().endswith('.pdf'):
            data = extract_from_pdf(path)
            all_data[file] = data

    print("\n\n--- Extraction Results ---")
    for filename, items in all_data.items():
        print(f"\nFile: {filename}")
        if not items:
            print("  No items extracted.")
            continue
            
        # Convert to DataFrame for pretty printing
        try:
            df = pd.DataFrame(items)
            # Reorder cols if present
            cols = [c for c in ['part_no', 'name', 'unit_price', 'amount', 'raw_line'] if c in df.columns]
            print(df[cols].to_markdown(index=False))
        except:
             print(items)

if __name__ == "__main__":
    main()
