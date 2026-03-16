from pathlib import Path

from docx import Document


class Exporter:
    @staticmethod
    def build_output_path(source_path: str | Path) -> Path:
        src = Path(source_path)
        return src.with_name(f"{src.stem}_格式调整后{src.suffix}")

    @staticmethod
    def save(document: Document, source_path: str | Path) -> Path:
        output_path = Exporter.build_output_path(source_path)
        document.save(str(output_path))
        return output_path
