from docx import Document
from PyPDF2 import PdfReader


def read_docx(path: str) -> str:
    doc = Document(path)

    return "".join([p.text for p in doc.paragraphs])


def read_pdf(path: str) -> str:
    reader = PdfReader(path)

    return "".join([p.extract_text() for p in reader.pages])
