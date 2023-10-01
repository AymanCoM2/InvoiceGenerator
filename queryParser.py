from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from connectionExecution import connectionResults
from fillFooter import fillingFooterFunction
from merge import combineParts

# Assuming you have a function connectionResults(80) to retrieve data
headerFooterResult, rowResult = connectionResults(10)

print(headerFooterResult)