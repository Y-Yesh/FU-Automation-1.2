from docx import Document
import win32com.client
from docx2pdf import convert

from time import sleep
from csv_extraction import participant_dict, company_dict, country_dict

# Load the Word document
doc_path = r"C:\Users\GOOD\Downloads\AIR.docx"  # Replace with your actual file path
doc = Document(doc_path)






date = '16/02/2025' # Initialize data variable with current date

#print(doc.tables)

'''for index, table in enumerate(doc.tables):
    print(f"Table {index}: {table}")
    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]  # Extract text from each cell
        #print(row_data)
        for cell in row.cells:
            #print(cell.text)
            cell.text = cell.text.replace('name' ,'yeshwanth')
        #print("\t".join(row_data))  # Print the row with tab spacing'''


for each in participant_dict.keys():

# ---------------------- Modifying Table 1 (Summary Table) ----------------------

    summary_table = doc.tables[0]

    for index, row in enumerate(summary_table.rows,start=1):
        if index == 1:   # Row 1 is 'Name' and 'Date'
            for index,cell in enumerate(row.cells, start=1):
                if index == 2: # Cell 2 is 'Name'
                    cell.text = each
                if index == 4 :
                        cell.text = date
        row_data = [cell.text.strip() for cell in row.cells]  # Extract text from each cell
        print(row_data)


    # Save the updated document as a DOCX file (using python-docx)
    updated_doc_path = fr"C:\Users\GOOD\Downloads\Test\{each}.docx"
    pdf_file = fr"C:\Users\GOOD\Downloads\Test\{each}.pdf"

    doc.save(updated_doc_path)

    convert(updated_doc_path, pdf_file)

'''
    # Convert the saved DOCX to PDF using win32com
    word = win32com.client.Dispatch("Word.Application")
    #word.Visible = False  # Run in background

    #sleep(20)

    # Use a separate variable for the COM document
    com_doc = word.Documents.Open(updated_doc_path)
    com_doc.SaveAs(pdf_file, FileFormat=17)  # 17 is the PDF format code
    com_doc.Close()
    word.Quit()'''

