import tempfile
import unittest
from pathlib import Path

from docx import Document

from models.structure import AnalyzerConfig
from services.diagnostic_reporter import DiagnosticReporter, ReportWriteError
from services.docx_reader import DocxReader
from services.structure_analyzer import StructureAnalyzer


class DiagnosticReporterTests(unittest.TestCase):
    def test_build_report_contains_meta(self) -> None:
        doc = Document()
        doc.add_paragraph('正文测试内容。')

        analysis = StructureAnalyzer.analyze(doc, config=AnalyzerConfig(), debug=False)
        payload = DiagnosticReporter.build_report(
            analysis,
            {'applied_rule': {}},
            source_path='input.docx',
            output_path='output.docx',
        )

        self.assertIn('meta', payload)
        self.assertEqual(payload['meta']['source_path'], 'input.docx')
        self.assertEqual(payload['meta']['output_path'], 'output.docx')
        self.assertIn('generated_at', payload['meta'])

    def test_write_json_validates_required_keys(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            report_path = Path(tmp) / 'bad.json'
            with self.assertRaises(ReportWriteError):
                DiagnosticReporter.write_json(report_path, {'summary': {}})


if __name__ == '__main__':
    unittest.main()
