from docx import Document


def readDocument(filePath):
    doc = Document(filePath)
    return doc, doc.core_properties
