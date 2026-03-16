import tempfile
import unittest
from pathlib import Path

from docx import Document

from services.exporter import Exporter


class ExporterTests(unittest.TestCase):
    def test_build_output_path_requires_docx(self) -> None:
        with self.assertRaises(ValueError):
            Exporter.build_output_path('demo.txt')

    def test_save_uses_incremented_name_when_target_exists(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'input.docx'
            src.write_bytes(b'placeholder')

            first_output = Path(tmp) / 'input_格式调整后.docx'
            first_output.write_bytes(b'existing')

            doc = Document()
            doc.add_paragraph('hello')

            saved = Exporter.save(doc, src)
            self.assertEqual(saved.name, 'input_格式调整后_2.docx')
            self.assertTrue(saved.exists())


if __name__ == '__main__':
    unittest.main()
