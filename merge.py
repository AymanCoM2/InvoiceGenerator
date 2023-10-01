from docxcompose.composer import Composer
from docx import Document as Document_compose


def combineParts(fileNamesList):
    master = Document_compose(fileNamesList[0])
    composer = Composer(master)
    for x in range(1, len(fileNamesList)):
        doc2 = Document_compose(fileNamesList[x])
        composer.append(doc2)
    composer.save("combined.docx")
