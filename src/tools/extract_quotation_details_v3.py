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
    text = str(text)
    # Extract number from string like "¥1,000", "1,000-", "1000"
    # remove dates or phone numbers
    if re.search(r'\d{4}[-/年]\d{1,2}[-/月]', text): return 0
    if re.search(r'\d{2,4}-\d{2,4}-\d{4}', text): return 0
    
    match = re.search(r'[\d,]+', text)
    if match:
        try:
            return int(match.group().replace(',', ''))
        except:
            return 0
    return 0

def extract_from_excel(file_path):
    print(f"\nProcessing Excel: {os.path.basename(file_path)}")
    try:
        df = pd.read_excel(file_path, header=None)
        
        # Determine header row with scoring
        best_header_row_idx = -1
        best_score = 0
        best_col_map = {}
        
        keywords = {
            'part_no': [r'図番', r'図\s*番', r'品番', r'図面', r'製品番号'],
            'name': [r'品名', r'品\s*名', r'名称', r'商品名', r'件名'],
            'unit_price': [r'単価', r'単\s*価', r'価格'],
            'amount': [r'金額', r'金\s*額', r'小計', r'合計']
        }
        
        for idx, row in df.iterrows():
            row_text = [clean_text(str(val)) for val in row.values]
            found_cols = {}
            found_keys = set()
            
            for col_idx, text in enumerate(row_text):
                for key, patterns in keywords.items():
                    # If this key is already found for this row, skip (or overwrite?)
                    if key in found_keys: continue
                    
                    for pattern in patterns:
                        if re.search(pattern, text):
                            found_cols[key] = col_idx
                            found_keys.add(key)
                            break
            
            score = len(found_keys)
            # We want headers that have at least 2 distinct types of info
            if score > best_score and score >= 2:
                best_score = score
                best_header_row_idx = idx
                best_col_map = found_cols
                
        if best_header_row_idx == -1:
            print("  Could not detect header row.")
            return []
            
        print(f"  Best Header detected at row {best_header_row_idx+1} (Score: {best_score}). Columns: {best_col_map}")
        
        results = []
        # Iterate rows after header
        for idx in range(best_header_row_idx + 1, len(df)):
            row = df.iloc[idx]
            
            item = {}
            has_data = False
            
            # Extract available columns
            for key, col_idx in best_col_map.items():
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
                
            # Heuristic: Valid row should have a part number OR amount
            if has_data and (item.get('part_no') or item.get('amount')):
                # Filter out rows that are just comments or emptyish
                if item.get('part_no') == 'nan' or item.get('name') == 'nan': continue
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
                im = page.to_image(resolution=300).original
                img_np = np.array(im)
                ocr_result = reader.readtext(img_np, detail=0)
                
                # Context buffer to associate price with previous part
                current_item = {}
                
                for line in ocr_result:
                    line = line.strip()
                    if not line: continue
                    
                    # 1. Check for Part Number pattern (e.g. TEM2521...)
                    # Uppercase followed by digits
                    part_match = re.search(r'([A-Z]+\d+[_\-0-9A-Z]*)', line)
                    
                    if part_match and len(part_match.group(1)) > 4:
                        # If we have a pending item, save it
                        if current_item:
                            results.append(current_item)
                            current_item = {}
                            
                        # Split part no and potential name
                        # Assuming structure "PartNo Name" or "PartNoName"
                        # We use the match end to split
                        part_no = part_match.group(1)
                        current_item['part_no'] = part_no
                        
                        # The rest of the line is likely the name
                        rest = line.replace(part_no, '').strip()
                        # If rest starts with special chars, clean
                        rest = re.sub(r'^[_:|\- ]+', '', rest)
                        if rest:
                            current_item['name'] = rest
                        
                        continue

                    # 2. Check for Price / Amount
                    # Valid price should be number, maybe with comma, maybe with -
                    # Avoid phone numbers (06-..., 072-...)
                    if re.search(r'0\d{1,4}-\d{1,4}-\d{4}', line): continue
                    
                    # Look for prices
                    # 9,000- or 9000
                    price_match = re.search(r'([\d,]+)-?$', line)
                    if price_match:
                        price_val = parse_price(price_match.group(1))
                        if price_val > 100: # Filter out small numbers like qty "1" or "2"
                            if 'unit_price' not in current_item:
                                current_item['unit_price'] = price_val
                            elif 'amount' not in current_item:
                                current_item['amount'] = price_val
                    
                    # 3. Check for specific Kana (Name candidate if not set)
                    if 'part_no' in current_item and 'name' not in current_item:
                        if re.search(r'[ぁ-んァ-ン一-龥]{2,}', line):
                             current_item['name'] = line

                # Append last item
                if current_item:
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

    print("\n\n--- Final Extraction Results ---")
    for filename, items in all_data.items():
        print(f"\nFile: {filename}")
        if not items:
            print("  No items extracted.")
            continue
            
        try:
            df = pd.DataFrame(items)
            cols = [c for c in ['part_no', 'name', 'unit_price', 'amount'] if c in df.columns]
            print(df[cols].to_markdown(index=False))
        except:
             print(items)

if __name__ == "__main__":
    main()
