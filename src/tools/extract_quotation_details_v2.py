import os
import pandas as pd
import pdfplumber
import easyocr
import numpy as np
import re
import logging
import warnings

# Suppress warnings
warnings.simplefilter(action='ignore', category=UserWarning)
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

def parse_price(text):
    # Extract number from string like "¥1,000", "1,000-", "1000"
    match = re.search(r'[\d,]+', text)
    if match:
        return int(match.group().replace(',', ''))
    return 0

def extract_from_excel(file_path):
    print(f"\nProcessing Excel: {os.path.basename(file_path)}")
    try:
        df = pd.read_excel(file_path, header=None)
        
        # Determine header row
        header_row_idx = -1
        col_map = {}
        
        # Searching for header keywords - Expanded keywords
        keywords = {
            'part_no': [r'図番', r'図\s*番', r'品番', r'図面', r'製品番号'],
            'name': [r'品名', r'品\s*名', r'名称', r'商品名', r'件名'],
            'unit_price': [r'単価', r'単\s*価', r'価格'],
            'amount': [r'金額', r'金\s*額', r'小計', r'合計']
        }
        
        for idx, row in df.iterrows():
            row_text = [clean_text(str(val)) for val in row.values]
            found_cols = {}
            
            for col_idx, text in enumerate(row_text):
                for key, patterns in keywords.items():
                    for pattern in patterns:
                        if re.search(pattern, text):
                            # Prioritize first match? or keep looking?
                            if key not in found_cols:
                                found_cols[key] = col_idx
                            break
            
            # If we matched at least 'unit_price' and 'amount', and maybe 'part_no' or 'name'
            score = len(found_cols)
            if score >= 2:
                # Check specifics: Do we have price info?
                if 'unit_price' in found_cols or 'amount' in found_cols:
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
            has_data = False
            
            # Extract available columns
            for key, col_idx in col_map.items():
                val = row[col_idx]
                if pd.notna(val):
                    val_str = clean_text(str(val))
                    if val_str:
                        item[key] = val_str
                        has_data = True
            
            # Normalize numeric values
            if 'unit_price' in item:
                item['unit_price'] = parse_price(item['unit_price'])
            if 'amount' in item:
                item['amount'] = parse_price(item['amount'])
            
            # Validate: must have something resembling a part number or name, AND some price info usually
            # But let's be lenient for now.
            if has_data and (item.get('part_no') or item.get('name')):
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
                # Always use OCR for consistency on this specific file type as seen in logs
                im = page.to_image(resolution=300).original
                img_np = np.array(im)
                ocr_result = reader.readtext(img_np, detail=0)
                
                print(f"  Page {i+1}: OCR lines: {len(ocr_result)}")
                
                # State machine for parsing
                current_item = {}
                
                for line in ocr_result:
                    line = line.strip()
                    if not line: continue
                    
                    # Start of item: Part Number Pattern
                    # Matches "TEM..." or "TKM..." or typical code structure
                    # Heuristic: Uppercase letters followed by numbers, at least 5 chars
                    if re.match(r'^[A-Z]+\d+[_\-0-9A-Z]*', line):
                        # If we have a previous item with enough info, save it
                        if current_item and ('part_no' in current_item or 'amount' in current_item):
                            results.append(current_item)
                            current_item = {}
                        
                        # Parse this line for PartNO and Name
                        # Issue: "TEM2521_70-POO1lカバー01" -> PartNo + Name merged
                        # Try to split by non-ascii?
                        # Part No is usually ASCII. Name often contains Kana/Kanji.
                        
                        match = re.match(r'^([A-Z0-9_\-]+)(.*)', line)
                        if match:
                            current_item['part_no'] = match.group(1)
                            rest = match.group(2).strip()
                            if rest:
                                current_item['name'] = rest
                        else:
                            current_item['part_no'] = line
                            
                    # Price / Amount
                    # Looks like "9,000-" or "9,000"
                    elif re.search(r'[\d,]+-?$', line):
                        # Filter out dates or phone numbers
                        if re.search(r'\d{4}年', line) or re.search(r'\d{2,4}-\d{2,4}-\d{4}', line):
                            continue
                            
                        val = parse_price(line)
                        if val > 0:
                            if 'unit_price' not in current_item:
                                current_item['unit_price'] = val
                            elif 'amount' not in current_item:
                                current_item['amount'] = val
                            else:
                                # Start new item? Or ignore?
                                pass
                                
                    # Name supplement
                    # If line is Japanese and we have a part no but no name (or short name)
                    elif re.search(r'[ぁ-んァ-ン一-龥]', line):
                        if 'part_no' in current_item and 'name' not in current_item:
                             current_item['name'] = line
                        elif 'part_no' in current_item:
                             # Append to name?
                             current_item['name'] = current_item.get('name', '') + " " + line

                # Append last item
                if current_item and ('part_no' in current_item or 'amount' in current_item):
                    results.append(current_item)

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
            cols = [c for c in ['part_no', 'name', 'unit_price', 'amount'] if c in df.columns]
            print(df[cols].to_markdown(index=False))
        except:
             print(items)

if __name__ == "__main__":
    main()
