from docx import Document
from docx2pdf import convert
from utils.map_level import map_level


from time import sleep
from csv_extraction import participant_dict, company_dict, country_dict
from definitions import definitions

# Load the Word document
doc_path = r"C:\Users\GOOD\Downloads\AIR.docx"  # Replace with your actual file path
doc = Document(doc_path)


date = '16/02/2025' # Initialize data variable with current date


summary_table = doc.tables[0]
'''for each in summary_table.rows:
    print(len(summary_table.rows))
    row_data = [cell.text for cell in each.cells]  # Extract text from each cell
    print(row_data)'''


#print(doc.tables)


for each,details in participant_dict.items():

    print(f'Name ----------> {each}')
    #print(f'Details -------> {details['Total Score']}')
    #print(f'Level -------> {[map_level(details['Total Score'])]}')
    #print(f'Definition ------> {definitions["Total Score"][map_level(details["Total Score"])]}')
# ---------------------- Modifying Table 1 (Summary Table) ----------------------

    # Row 1 is 'Name' and 'Date' 
    row0 = summary_table.rows[0] 
    row0.cells[1].text = each
    row0.cells[3].text = date   
                
    row_data = [cell.text.strip() for cell in row0.cells]  # Extract text from each cell
    print(row_data)


    # Row 2 is 'Company' and 'Definition'
    row1 = summary_table.rows[1] 
    row1.cells[1].text = details['Company']
    row1.cells[3].text = definitions['Total Score'][map_level(details['Total Score'])]
    row_data1 = [cell.text.strip() for cell in row1.cells]  # Extract text from each cell
    print(row_data1)

    #Row 3 is 'Country' and 'Total Score'
    row2 = summary_table.rows[2] 
    row2.cells[1].text = details['Country']
    row2.cells[2].text = str(details['Total Score'])
                
    row_data2 = [cell.text.strip() for cell in row2.cells]  # Extract text from each cell
    print(row_data2)

    #Row 4 is 'Company' and 'Total Score'
    row3 = summary_table.rows[3] 
    row3.cells[1].text = details['Department']
    row3.cells[2].text = str(details['Total Score'])
                
    row_data3 = [cell.text.strip() for cell in row3.cells]  # Extract text from each cell
    print(row_data3)

    #Row 5 is 'Company' and 'Total Score'
    row4 = summary_table.rows[4] 
    row3.cells[1].text = details['Position']
    row3.cells[2].text = map_level(details['Total Score'])
                
    row_data4 = [cell.text.strip() for cell in row4.cells]  # Extract text from each cell
    print(row_data4)







