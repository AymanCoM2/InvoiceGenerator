from docx import Document
from docx.shared import Pt
from connectionExecution import connectionResults

headerFooterResult, rowResult = connectionResults(80)
headerFooterList = headerFooterResult[0]
finalDataList = []

counter = 1
for rRow in rowResult:
    internalList = [str(rRow[7]), str(rRow[5]), str(rRow[9]), str(rRow[5]), str(
        rRow[6]) + str(rRow[7]), str(rRow[1]), str(rRow[3]), str(rRow[2]), "1"]
    counter = counter + 1
    if counter == 9:
        break
    finalDataList.append(internalList)

document = Document('Template1.docx')

style = document.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(8)

dummy_data = finalDataList
header_data = ["", "", "", str(headerFooterList[5]), str(
    headerFooterList[5]), "", str(headerFooterList[2]), str(headerFooterList[1])]
footer_data = [str(headerFooterList[8]), str(headerFooterList[9]), str(
    headerFooterList[10]), str(headerFooterList[11]), str(headerFooterList[12])]

# if len(document.tables) > 0:
#     table = document.tables[0]
#     for i in range(len(header_data)):
#         for j in range(len(table.columns)):
#             cell = table.cell(i, j)
#             cell.text = header_data[i]

if len(document.tables) > 1:
    table = document.tables[1]
    num_rows = len(table.rows)
    num_cols = len(table.columns)
    if num_rows >= len(dummy_data) and num_cols >= len(dummy_data[0]):
        for i in range(len(dummy_data)):
            for j in range(len(dummy_data[i])):
                cell = table.cell(i, j)
                cell.text = dummy_data[i][j]
    else:
        print("Table 1 does not have enough rows or columns for the dummy data.")

if len(document.tables) > 2:
    table = document.tables[2]
    for i in range(len(footer_data)):
        for j in range(len(table.columns)):
            cell = table.cell(i, j)
            cell.text = footer_data[i]

document.save('Tabular_Filled.docx')
