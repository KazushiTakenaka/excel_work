---
name: read_pdf
description: Extracts text from PDF files using pdfplumber (for text) and EasyOCR (for images).
---

# PDF Reading Skill with OCR

This skill provides a robust method to extract text from PDF files, automatically handling both standard text-based PDFs and image-based (scanned) PDFs using OCR.

## Prerequisites

Ensure you have Python installed and the following libraries:

1.  **Install Dependencies**
    Create or update your `requirements.txt` with:
    ```
    pdfplumber
    easyocr
    opencv-python-headless
    numpy
    ```
    
    Then run:
    ```bash
    pip install -r requirements.txt
    ```

    *Note: The first time you run EasyOCR, it will download the necessary model files (approx. several hundred MBs).*

## Implementation

Create a file named `utils/pdf_reader.py` (or similar) with the following content. This class handles the logic to switch between text extraction and OCR.

```python
import pdfplumber
import easyocr
import numpy as np
from PIL import Image

class PDFReader:
    def __init__(self, languages=['ja', 'en']):
        """
        Initialize the PDFReader with OCR languages.
        :param languages: List of languages for OCR (default: ['ja', 'en'])
        """
        # Initialize EasyOCR reader
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
                        
                        # Convert page to image for OCR
                        # resolution=300 is usually good for OCR
                        im = page.to_image(resolution=300).original
                        
                        # Convert PIL image to numpy array for EasyOCR
                        img_np = np.array(im)
                        
                        # Perform OCR
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
    import sys
    if len(sys.argv) > 1:
        reader = PDFReader()
        print(reader.extract_text(sys.argv[1]))
    else:
        print("Usage: python utils/pdf_reader.py <pdf_path>")
```

## Usage Example

You can use the `PDFReader` class in your scripts like this:

```python
from utils.pdf_reader import PDFReader

# Initialize
reader = PDFReader(languages=['ja', 'en'])

# Read a PDF file
text_content = reader.extract_text("path/to/your/document.pdf")

# Output the result
print(text_content)
```
