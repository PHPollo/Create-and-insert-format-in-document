import os

from docx import Document


def create_Word(name_document : str, direction):

    document = Document()
    
    document.save(os.path.join(direction, f'{name_document}.docx'))
