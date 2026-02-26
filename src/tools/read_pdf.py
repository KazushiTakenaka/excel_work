import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..'))

try:
    from pypdf import PdfReader
except ImportError:
    print("pypdf not installed. Trying to install...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
    from pypdf import PdfReader

file_path = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品\注文No.20260105MT05.pdf'

try:
    reader = PdfReader(file_path)
    print(f"--- Reading {os.path.basename(file_path)} ---")
    print(f"Number of pages: {len(reader.pages)}")
    
    text_content = ""
    for i, page in enumerate(reader.pages):
        print(f"\nPage {i+1}:")
        text = page.extract_text()
        if text:
            print(text)
            text_content += text
        else:
            print("[No text extracted - likely image-based PDF]")
    
    if not text_content.strip():
        print("\nWARNING: No text content found in PDF. It might be scanned/image-only.")

except Exception as e:
    print(f"Error reading PDF: {e}")
