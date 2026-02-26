import os
import pandas as pd
import re
from src.lib.pdf_reader import PDFReader
import pdfplumber

# Configuration
INPUT_DIR = '見積書'
OUTPUT_FILE = 'quotation_summary.xlsx'
KEYWORDS = {
    'product': ['品名', '商品名', '商品', '名称', '銘柄'],
    'quantity': ['数量', '数'],
    'unit': ['単位'],
    'unit_price': ['単価', '単価(税別)', '単価（税別）'],
    'amount': ['金額', '合計金額(税別)', '合計金額']
}

def find_header_row(df, keywords):
    """Find the index of the header row based on keywords."""
    for idx, row in df.iterrows():
        row_str = " ".join([str(x) for x in row if pd.notna(x)])
        # Check if at least one keyword from each mandatory category exists
        if any(k in row_str for k in KEYWORDS['product']) and \
           any(k in row_str for k in KEYWORDS['amount']):
            return idx
    return None

def extract_from_excel(file_path):
    """Extract data from Excel file."""
    try:
        # Read all sheets (or specific one? defaulting to first usually works but let's check)
        # Using openpyxl engine
        df_raw = pd.read_excel(file_path, header=None)
        
        header_idx = find_header_row(df_raw, KEYWORDS)
        if header_idx is None:
            return []

        # Reload with correct header
        df = pd.read_excel(file_path, header=header_idx)
        
        # Identify columns
        cols = df.columns
        col_map = {}
        for col in cols:
            col_str = str(col)
            for key, words in KEYWORDS.items():
                if any(word == col_str for word in words) or any(word in col_str for word in words):
                    if key not in col_map: 
                        col_map[key] = col
        
        # specific handling for '図番'/'型番' which might be in '図番' column or separate
        # If not found, look for '図番' specific keywords
        if 'model_number' not in col_map:
             for col in cols:
                if '図番' in str(col) or '型番' in str(col):
                    col_map['model_number'] = col
                    break

        if 'product' not in col_map:
            return []

        results = []
        for _, row in df.iterrows():
            if pd.isna(row[col_map['product']]): continue
            
            # Stop if "Total" or similar is found in product name (heuristic)
            if '合計' in str(row[col_map['product']]): break
            
            item = {
                'ファイル名': os.path.basename(file_path),
                '品名': row[col_map['product']],
                '図番/型番': row[col_map.get('model_number')] if 'model_number' in col_map else '',
                '数量': row[col_map.get('quantity')] if 'quantity' in col_map else 0,
                '単位': row[col_map.get('unit')] if 'unit' in col_map else '',
                '単価': row[col_map.get('unit_price')] if 'unit_price' in col_map else 0,
                '金額': row[col_map.get('amount')] if 'amount' in col_map else 0,
            }
            results.append(item)
        return results

    except Exception as e:
        print(f"Error processing Excel {file_path}: {e}")
        return []

        return results

    except Exception as e:
        print(f"Error processing PDF {file_path}: {e}")
        return []

def parse_ocr_text(text, filename):
    """Parse OCR text using multiple strategies."""
    results = []
    lines = text.split('\n')
    lines = [line.strip() for line in lines if line.strip()] # Clean empty lines

    # Strategy 1: Horizontal Line Parsing (e.g., "Product 1 Unit 1000 1000")
    # Regex: Product (non-space) | Qty (digits) | Unit | Price | Amount
    # Note: Product might contain spaces, but usually last 4 tokens are Qty, Unit, Price, Amount
    horizontal_items = []
    
    # Simple regex for lines ending with two numbers (Price, Amount)
    # Allows for '¥', ',', '-', '.'
    price_pattern = r'[¥￥]?[\d,]+(?:\.\d+)?[-－]?'
    
    # Regex to find line ending with: Qty Unit Price Amount
    # \s+ is separator
    # Group 1: Product (everything before)
    # Group 2: Qty
    # Group 3: Unit
    # Group 4: Unit Price
    # Group 5: Amount
    line_regex = re.compile(fr'(.*?)\s+(\d+)\s+([^\s\d]+)\s+({price_pattern})\s+({price_pattern})$')

    for line in lines:
        match = line_regex.search(line)
        if match:
            # Validate if it looks like a valid line (e.g. price and amount are numbers)
            try:
                p_name = match.group(1).strip()
                qty = match.group(2)
                unit = match.group(3)
                u_price = match.group(4)
                amount = match.group(5)
                
                # Filter out likely false positives (e.g. date strings)
                if len(p_name) > 1:
                    horizontal_items.append({
                        'ファイル名': filename,
                        '品名': p_name,
                        '図番/型番': '',
                        '数量': float(qty),
                        '単位': unit,
                        '単価': u_price,
                        '金額': amount
                    })
            except:
                pass
    
    if len(horizontal_items) > 0:
        return horizontal_items

    # Strategy 2: Vertical/Block Parsing (Specific for fragmented OCR like in 注文No...pdf)
    # Pattern: Name -> (Material ->) Qty -> Unit -> Price -> Amount
    # We look for a sequence of numbers identifying Price and Amount
    vertical_items = []
    
    def is_price(s):
        return re.match(f'^{price_pattern}$', s) is not None
    
    def is_qty(s):
        return re.match(r'^\d+$', s) is not None

    i = 0
    while i < len(lines) - 1:
        # Check current line (Price) and next line (Amount)
        # Scan for Amount (last element)
        if i + 1 < len(lines) and is_price(lines[i]) and is_price(lines[i+1]):
             # Candidate for Price -> Amount
             u_price = lines[i]
             amount = lines[i+1]
             
             # Look backwards for Unit and Qty
             # Expected: ... Qty, Unit, Price, Amount
             # But OCR might split them differently.
             # Pattern in debug: Qty, Unit, Price, Amount
             # Text: "1", "個", "50,100-", "50,100"
             
             qty = None
             unit = None
             name_candidates = []
             
             # Look back 1 step: Unit?
             if i > 0:
                 unit = lines[i-1]
                 # Look back 2 steps: Qty?
                 if i > 1 and is_qty(lines[i-2]):
                     qty = lines[i-2]
                     # Name is everything before
                     # heuristic: take up to 2 lines before Qty as name/model
                     start_name = max(0, i - 4)
                     name_candidates = lines[start_name : i-2]
                 else:
                     # Maybe Unit and Qty are on same line? "1 | 個" -> "1 | 個"
                     pass
             
             if qty and unit:
                 name = " ".join(name_candidates)
                 vertical_items.append({
                    'ファイル名': filename,
                    '品名': name,
                    '図番/型番': '',
                    '数量': float(qty),
                    '単位': unit,
                    '単価': u_price,
                    '金額': amount
                 })
                 i += 2 # Skip Price and Amount
                 continue
        i += 1
        
    return vertical_items

def extract_from_pdf(file_path, reader):
    """Extract data from PDF using PDFReader (text/OCR) + Parsing."""
    try:
        filename = os.path.basename(file_path)
        full_text = reader.extract_text(file_path)
        
        # Try to parse the text
        parsed_data = parse_ocr_text(full_text, filename)
        if parsed_data:
            return parsed_data
            
        return [] # Return empty if no strategy worked (or implement table extraction fallback here if needed)

    except Exception as e:
        print(f"Error processing PDF {file_path}: {e}")
        return []

def main():
    all_data = []
    reader = PDFReader() # Initialize OCR reader once

    files = [f for f in os.listdir(INPUT_DIR) if f.startswith('~$') is False] # Skip temp files
    
    for filename in files:
        filepath = os.path.join(INPUT_DIR, filename)
        print(f"Processing {filename}...")
        
        if filename.lower().endswith(('.xlsx', '.xls')):
            data = extract_from_excel(filepath)
            all_data.extend(data)
        elif filename.lower().endswith('.pdf'):
            data = extract_from_pdf(filepath, reader)
            all_data.extend(data)
    
    if all_data:
        df_result = pd.DataFrame(all_data)
        # Reorder columns
        cols = ['ファイル名', '品名', '図番/型番', '数量', '単位', '単価', '金額']
        # Add missing columns
        for c in cols:
            if c not in df_result.columns: df_result[c] = ''
        
        df_result = df_result[cols]
        df_result.to_excel(OUTPUT_FILE, index=False)
        print(f"Successfully saved to {OUTPUT_FILE}")
    else:
        print("No data extracted.")

if __name__ == "__main__":
    main()
