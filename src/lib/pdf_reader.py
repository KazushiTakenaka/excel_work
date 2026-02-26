import pdfplumber
import easyocr
import os
import io

class PDFReader:
    def __init__(self, languages=['ja', 'en']):
        """
        Initialize the PDFReader with OCR languages.
        :param languages: List of languages for OCR (default: ['ja', 'en'])
        """
        self.reader = easyocr.Reader(languages) if easyocr else None

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
                    # Try text extraction first
                    text = page.extract_text()
                    
                    # Heuristic: If text is very short (likely just header/footer or empty), try OCR
                    if not text or len(text.strip()) < 50:
                        # Convert page to image for OCR
                        # pdfplumber to_image returns a PageImage, .original gives PIL Image
                        im = page.to_image(resolution=300).original
                        
                        # Convert PIL image to bytes for easyocr (or pass numpy array if compatible, 
                        # but easyocr accepts bytes or file path. It also accepts numpy array)
                        # EasyOCR readtext accepts: file path, url, byte, numpy array
                        import numpy as np
                        img_np = np.array(im)
                        
                        ocr_result = self.reader.readtext(img_np, detail=0)
                        text = "\n".join(ocr_result)
                        print(f"[PDFReader] Page {i+1}: Used OCR.")
                    else:
                         print(f"[PDFReader] Page {i+1}: Used Text Extraction.")

                    full_text += f"\n--- Page {i+1} ---\n{text}"
                    
        except Exception as e:
            return f"Error reading PDF: {e}"

        return full_text

if __name__ == "__main__":
    # Test execution
    import sys
    if len(sys.argv) > 1:
        reader = PDFReader()
        print(reader.extract_text(sys.argv[1]))
    else:
        print("Usage: python pdf_reader.py <pdf_path>")
