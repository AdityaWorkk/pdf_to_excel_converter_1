#this file is for extracting the first page of the pdf as the origan pdf is too lengthy and the pdf to excel code can be tested on this first

import fitz  # PyMuPDF

# Define the PDF file path
pdf_path = 'output_first_page.pdf'

# Open the PDF file
pdf_document = fitz.open(pdf_path)

# Loop through each page in the PDF
for page_num in range(len(pdf_document)):
    page = pdf_document.load_page(page_num)
    text = page.get_text("text")
    print(f"--- Page {page_num + 1} ---")
    print(text)
    print("\n")

# Close the PDF file
pdf_document.close()
