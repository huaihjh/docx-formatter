from pathlib import Path

from docx import Document


class ExportError(RuntimeError):
    pass


class Exporter:
    @staticmethod
    def build_output_path(source_path: str | Path) -> Path:
        src = Path(source_path)
        if src.suffix.lower() != ".docx":
            raise ValueError("仅支持 .docx 导出")
        return src.with_name(f"{src.stem}_格式调整后{src.suffix}")

    @staticmethod
    def _next_available_path(path: Path) -> Path:
        if not path.exists():
            return path

        for idx in range(2, 10000):
            candidate = path.with_name(f"{path.stem}_{idx}{path.suffix}")
            if not candidate.exists():
                return candidate

        raise ExportError("导出失败：无法生成可用输出文件名")

    @staticmethod
    def save(document: Document, source_path: str | Path, allow_overwrite: bool = False) -> Path:
        output_path = Exporter.build_output_path(source_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if output_path.exists() and not allow_overwrite:
            output_path = Exporter._next_available_path(output_path)

        try:
            document.save(str(output_path))
        except PermissionError as exc:
            raise ExportError("导出失败：无写入权限或目标文件被占用") from exc
        except OSError as exc:
            raise ExportError(f"导出失败：{exc}") from exc

        return output_path
