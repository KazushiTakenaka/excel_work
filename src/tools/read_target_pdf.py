
import pdfplumber
import os

import sys

if len(sys.argv) > 1:
    pdf_path = sys.argv[1]
else:
    pdf_path = os.path.join('見積書', '20260106TKエンジニアリング御中.pdf')

if not os.path.exists(pdf_path):
    print(f"Error: File not found at {pdf_path}")
    exit(1)

try:
    with pdfplumber.open(pdf_path) as pdf:
        print(f"--- Extracting text from {pdf_path} ---")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            print(f"\n--- Page {i+1} ---")
            if text:
                print(text)
            else:
                print("(No text found on this page. It might be an image.)")
except Exception as e:
    print(f"Error reading PDF: {e}")
