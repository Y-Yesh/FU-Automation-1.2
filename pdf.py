import fitz  # PyMuPDF

# Load the PDF
pdf_path = r"C:\Users\GOOD\Downloads\Air Assessment Summary Report_designed 250210.pdf"  # Replace with your actual file path
doc = fitz.open(pdf_path)
print(doc)