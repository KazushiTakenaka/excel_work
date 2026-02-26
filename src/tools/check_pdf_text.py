import pdfplumber
import os

pdf_dir = '見積書'
results = []

print(f"Checking PDFs in {pdf_dir}...")

for filename in os.listdir(pdf_dir):
    if filename.lower().endswith('.pdf'):
        path = os.path.join(pdf_dir, filename)
        try:
            with pdfplumber.open(path) as pdf:
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                if text and len(text.strip()) > 50:
                    results.append((filename, "Text-based (OK)", len(text)))
                else:
                    results.append((filename, "Image-based or Empty (Might need OCR)", 0))
        except Exception as e:
            results.append((filename, f"Error: {e}", 0))

print("\n--- Results ---")
for res in results:
    print(f"{res[0]}: {res[1]}")
