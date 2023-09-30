from connectionExecution import connectionResults
headerFooterResult , rowResult = connectionResults(80)

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn

finalDataList  = []
counter = 1
for rRow in rowResult:
    internalList  = [str(rRow[7]) , str(rRow[5])  ,str(rRow[9]) ,str(rRow[5]) ,str(rRow[6])+str(rRow[7]),str(rRow[1]) ,str(rRow[3]) ,str(rRow[2]) ,"1"]
    counter =  counter + 1 
    if counter == 9:
        break
    finalDataList.append(internalList)

document = Document('Template1.docx')

style = document.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(8)

dummy_data =finalDataList
# Access the first table in the document
if len(document.tables) > 0:
    print(len(document.tables))
    table = document.tables[1] # 1 // ! 
    num_rows = len(table.rows)
    num_cols = len(table.columns)
    if num_rows >= len(dummy_data) and num_cols >= len(dummy_data[0]):
        for i in range(len(dummy_data)):
            for j in range(len(dummy_data[i])):
                cell = table.cell(i, j)
                # set_cell_format(cell)
                cell.text = dummy_data[i][j]
    else:
        print("Table does not have enough rows or columns for the dummy data.")
else:
    print("No table found in the document.")

# # Save the modified document
document.save('Tabular_Filled.docx')

