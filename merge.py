from docxcompose.composer import Composer
from docx import Document as Document_compose
from docx2pdf import convert
from pypdf import PdfMerger
pdfFiles = []


def combineParts(fileNamesList):
    for docx_file in fileNamesList:
        convert(docx_file)
        pdf_file = docx_file.replace('.docx', '.pdf')
        print(f"Converted {docx_file} to {pdf_file}")
        pdfFiles.append(pdf_file)


    merger = PdfMerger()
    for pdf in pdfFiles:
        merger.append(pdf)
    merger.write("result.pdf")
    merger.close()
