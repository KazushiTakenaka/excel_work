import os
import pandas as pd
import pdfplumber
import easyocr
import numpy as np
from PIL import Image
import sys
import logging

# Suppress EasyOCR warnings
logging.getLogger('easyocr').setLevel(logging.ERROR)

# Set directory path
directory = r'c:\Users\taman\OneDrive\デスクトップ\作業中\104.ai_workspaces\excel_work\見積書\加工品'

class PDFReader:
    def __init__(self, languages=['ja', 'en']):
        """
        Initialize the PDFReader with OCR languages.
        :param languages: List of languages for OCR (default: ['ja', 'en'])
        """
        print("Initializing EasyOCR...")
        # gpu=False if you don't have a CUDA-compatible GPU, or leave it to auto-detect
        self.reader = easyocr.Reader(languages)

    def extract_text(self, pdf_path):
        """
        Extract text from a PDF file. Automatically switches to OCR if text is sparse.
        :param pdf_path: Path to the PDF file.
        :return: Extracted text as a string (pages separated by newlines).
        """
        full_text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    # Try standard text extraction first
                    text = page.extract_text()
                    
                    # Heuristic: If text is very short/empty (likely scanned), try OCR
                    if not text or len(text.strip()) < 50:
                        print(f"[PDFReader] Page {i+1}: Text sparse/empty. Attempting OCR...")
                        
                        try:
                            # Convert page to image for OCR
                            # resolution=300 is usually good for OCR
                            im = page.to_image(resolution=300).original
                            
                            # Convert PIL image to numpy array for EasyOCR
                            img_np = np.array(im)
                            
                            # Perform OCR
                            ocr_result = self.reader.readtext(img_np, detail=0)
                            text = "\n".join(ocr_result)
                            print(f"[PDFReader] Page {i+1}: Used OCR.")
                        except Exception as e_ocr:
                            print(f"[PDFReader] Page {i+1}: OCR failed: {e_ocr}")
                    else:
                        print(f"[PDFReader] Page {i+1}: Used Text Extraction.")

                    full_text += f"\n--- Page {i+1} ---\n{text}\n"
                    
        except Exception as e:
            return f"Error reading PDF: {e}"

        return full_text

def read_excel(file_path):
    print(f"\n==========================================")
    print(f"Reading Excel: {os.path.basename(file_path)}")
    print(f"==========================================")
    try:
        # Read the first sheet
        df = pd.read_excel(file_path)
        print("\n[Columns]")
        print(df.columns.tolist())
        
        print("\n[First 10 rows]")
        try:
            print(df.head(10).to_markdown(index=False))
        except ImportError:
            print(df.head(10))
            
        # Also print non-null counts to give an idea of content
        print("\n[Summary Info]")
        df.info()
    except Exception as e:
        print(f"Error reading Excel {file_path}: {e}")

def main():
    if not os.path.exists(directory):
        print(f"Directory not found: {directory}")
        return

    files = os.listdir(directory)
    print(f"Files found: {files}")
    
    # Initialize PDF reader once
    pdf_reader = None

    for file in files:
        file_path = os.path.join(directory, file)
        if file.lower().endswith(('.xlsx', '.xls')):
            read_excel(file_path)
        elif file.lower().endswith('.pdf'):
            if pdf_reader is None:
                pdf_reader = PDFReader()
            
            print(f"\n==========================================")
            print(f"Reading PDF: {os.path.basename(file_path)}")
            print(f"==========================================")
            content = pdf_reader.extract_text(file_path)
            print(content)
        else:
            print(f"Skipping file: {file}")

if __name__ == "__main__":
    main()
