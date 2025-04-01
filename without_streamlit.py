import pytesseract
from pdf2image import convert_from_bytes
import os

def ocr_pdf_to_text(pdf_file):
    pdf_content = pdf_file.read()
    images = convert_from_bytes(pdf_content, dpi=300)
    full_text = ""
    for i, image in enumerate(images):
        text = pytesseract.image_to_string(image)
        full_text += f"\n--- Page {i + 1} ---\n{text}"
        print(f"OCR extracted text from page {i + 1}:\n{text[:500]}...")
    return full_text

# Test with your PDF
with open("MHCutOOff_10-pages__trimmed.pdf", "rb") as f:  # Replace with your PDF path
    text = ocr_pdf_to_text(f)
    print("Full text:", text[:1000])