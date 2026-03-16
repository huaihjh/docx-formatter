import os
import unittest
from pathlib import Path

from models.structure import AnalyzerConfig
from services.docx_reader import DocxReader
from services.structure_analyzer import StructureAnalyzer


def _core_sample_dir() -> Path:
    env_dir = os.getenv('DOCX_SAMPLE_DIR', '').strip()
    if env_dir:
        return Path(env_dir)
    return Path(r'e:\Download\docx_samples')


class CoreRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.sample_dir = _core_sample_dir()
        if not cls.sample_dir.exists():
            raise unittest.SkipTest(f'core sample dir not found: {cls.sample_dir}')

    def _analyze_by_prefix(self, prefix: str):
        candidates = [
            p
            for p in self.sample_dir.glob(f'{prefix}*.docx')
            if '_格式调整后' not in p.stem and not p.name.startswith('~$')
        ]
        if not candidates:
            self.skipTest(f'no sample file found for prefix: {prefix}')
        doc_path = sorted(candidates)[0]
        analysis = StructureAnalyzer.analyze(
            DocxReader.load(doc_path),
            config=AnalyzerConfig(),
            debug=False,
        )
        return doc_path, analysis

    def test_sample3_multilevel_numbering_not_unknown(self) -> None:
        _, analysis = self._analyze_by_prefix('sample_3_')
        target = next((b for b in analysis.blocks if b.clean_text == '1.1 研究背景'), None)
        self.assertIsNotNone(target, 'missing target block: 1.1 研究背景')
        self.assertEqual(target.semantic_label, 'sub_heading')

    def test_sample4_bracket_caption_recognized(self) -> None:
        _, analysis = self._analyze_by_prefix('sample_4_')
        target = next((b for b in analysis.blocks if b.clean_text == '[ 图 1：'), None)
        self.assertIsNotNone(target, 'missing target block: [ 图 1：')
        self.assertEqual(target.semantic_label, 'caption')

    def test_no_unknown_in_core_samples(self) -> None:
        prefixes = ('sample_1_', 'sample_2_', 'sample_3_', 'sample_4_', 'sample_5_')
        for prefix in prefixes:
            with self.subTest(sample=prefix):
                path, analysis = self._analyze_by_prefix(prefix)
                unknown_count = sum(1 for b in analysis.blocks if b.semantic_label == 'unknown')
                self.assertEqual(
                    unknown_count,
                    0,
                    f'unknown blocks found in {path.name}: {unknown_count}',
                )


if __name__ == '__main__':
    unittest.main()
