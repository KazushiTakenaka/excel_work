
import pdfplumber
import easyocr
import sys
import os
import numpy as np

# Suppress easyocr warnings
import logging
logging.getLogger('easyocr').setLevel(logging.ERROR)

if len(sys.argv) > 1:
    pdf_path = sys.argv[1]
else:
    pdf_path = os.path.join('見積書', '注文No.20260105MT05.pdf')

if not os.path.exists(pdf_path):
    print(f"Error: File not found at {pdf_path}")
    exit(1)

print(f"Initializing OCR reader... (This may take a moment)")
reader = easyocr.Reader(['ja', 'en'])

try:
    with pdfplumber.open(pdf_path) as pdf:
        print(f"--- Extracting text (OCR) from {pdf_path} ---")
        for i, page in enumerate(pdf.pages):
            print(f"\n--- Page {i+1} ---")
            
            # Convert page to image
            # resolution=300 is usually good for OCR
            im = page.to_image(resolution=300).original
            
            # Convert PIL image to numpy array for easyocr
            im_np = np.array(im)
            
            # Perform OCR
            result = reader.readtext(im_np)
            
            # Print results
            for (bbox, text, prob) in result:
                if prob > 0.3: # Filter low confidence
                    print(text)
            
except Exception as e:
    print(f"Error reading PDF: {e}")
