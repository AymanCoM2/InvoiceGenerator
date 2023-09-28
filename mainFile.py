from connectionExecution import fillResults
headerFooterResult , rowResult = fillResults(50)


from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn


# for hfRow in headerFooterResult:
    # print(hfRow) 
    # Prints the Tuple Of the Data 

# print("----------------------------") 
# Seperation Of Rows From Headers and Footer
finalDataList  = []

for rRow in rowResult:
    internalList  = [str(rRow[7]) , str(rRow[5])  ,str(rRow[9]) ,str(rRow[5]) ,str(rRow[6])+str(rRow[7]),str(rRow[1]) ,str(rRow[3]) ,str(rRow[2]) ,"1"]
    finalDataList.append(internalList)
    # print(rRow) 
    # Print Each Tuple Of Items In FATOORA 
    # print("****************")



# Load the existing Word document
document = Document('Tabular.docx')

# Function to set cell formatting
def set_cell_format(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(12)  # Set font size
            run.font.name = 'Arial'  # Set font name
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Center-align text
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')  # Set font for non-Latin characters
            run._element.rPr.rFonts.set(qn('w:cs'), 'Arial')  # Set complex script font
            run._element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')  # Set high ANSI font
            run.element.rPr.rFonts.set(qn('w:hint'), 'eastAsia')  # Hint for non-Latin characters
            run.element.rPr.rFonts.set(qn('w:hint'), 'cs')  # Hint for complex script
            run.element.rPr.rFonts.set(qn('w:hint'), 'hAnsi')  # Hint for high ANSI
            run.element.rPr.rFonts.set(qn('w:cs'), 'Arial')
            run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
            run.element.rPr.rFonts.set(qn('w:hAnsi'), 'Arial')

# Dummy data (you can replace this with your actual data)
dummy_data =finalDataList
# [
#     ['Header1', 'Header2', 'Header3', 'Header4', 'Header5', 'Header6', 'Header7', 'Header8', 'Header9'],
#     ['Data1', 'Data2', 'Data3', 'Data4', 'Data5', 'Data6', 'الله اكبر ولله الحمد هذا والله اعلى واعلم يا اخوان في هذه الاشياء', 'Data8', 'Data9'],
#     ['Data1', 'Data2', 'Data3', 'Data4', 'Data5', 'Data6', 'الله اكبر ولله الحمد هذا والله اعلى واعلم يا اخوان في هذه الاشياء', 'Data8', 'Data9'],
#     ['Data1', 'Data2', 'Data3', 'Data4', 'Data5', 'Data6', 'الله اكبر ولله الحمد هذا والله اعلى واعلم يا اخوان في هذه الاشياء', 'Data8', 'Data9'],
# ]

# Access the first table in the document
if len(document.tables) > 0:
    table = document.tables[0]
    num_rows = len(table.rows)
    num_cols = len(table.columns)
    # Check if the table has enough rows and columns for the dummy data
    if num_rows >= len(dummy_data) and num_cols >= len(dummy_data[0]):
        # Populate the table with dummy data
        for i in range(len(dummy_data)):
            for j in range(len(dummy_data[i])):
                cell = table.cell(i, j)
                set_cell_format(cell)
                cell.text = dummy_data[i][j]
    else:
        print("Table does not have enough rows or columns for the dummy data.")
else:
    print("No table found in the document.")

# Save the modified document
document.save('Tabular_Filled.docx')