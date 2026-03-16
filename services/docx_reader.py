from pathlib import Path

from docx import Document


class DocxReader:
    @staticmethod
    def load(file_path: str | Path) -> Document:
        return Document(str(file_path))
