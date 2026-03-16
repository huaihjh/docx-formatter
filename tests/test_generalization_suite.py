import json
import os
import unittest
from pathlib import Path

from models.structure import AnalyzerConfig
from services.docx_reader import DocxReader
from services.structure_analyzer import StructureAnalyzer


def _generalization_dir() -> Path | None:
    env_dir = os.getenv('DOCX_GENERALIZATION_DIR', '').strip()
    if not env_dir:
        return None
    return Path(env_dir)


def _report_path() -> Path:
    p = os.getenv('DOCX_GENERALIZATION_REPORT_PATH', '').strip()
    if p:
        return Path(p)
    return Path('memory-bank/baseline/generalization_summary.json')


class GeneralizationSuiteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.sample_dir = _generalization_dir()
        if cls.sample_dir is None:
            raise unittest.SkipTest('DOCX_GENERALIZATION_DIR is not set')
        if not cls.sample_dir.exists():
            raise unittest.SkipTest(f'generalization dir not found: {cls.sample_dir}')

        cls.files = sorted(
            [
                p
                for p in cls.sample_dir.glob('*.docx')
                if '_格式调整后' not in p.stem and not p.name.startswith('~$')
            ]
        )
        if not cls.files:
            raise unittest.SkipTest('no .docx files found in generalization dir')

    def test_unknown_ratio_below_threshold(self) -> None:
        threshold = float(os.getenv('DOCX_GENERALIZATION_MAX_UNKNOWN_RATIO', '0.03'))

        total_blocks = 0
        total_unknown = 0
        per_file: list[dict[str, object]] = []

        for p in self.files:
            analysis = StructureAnalyzer.analyze(
                DocxReader.load(p),
                config=AnalyzerConfig(),
                debug=False,
            )
            block_count = len(analysis.blocks)
            unknown_count = sum(1 for b in analysis.blocks if b.semantic_label == 'unknown')
            unknown_ratio = (unknown_count / block_count) if block_count else 0.0

            total_blocks += block_count
            total_unknown += unknown_count

            per_file.append(
                {
                    'file_name': p.name,
                    'file_path': str(p),
                    'block_count': block_count,
                    'unknown_count': unknown_count,
                    'unknown_ratio': round(unknown_ratio, 6),
                    'summary': analysis.summary(),
                }
            )

            with self.subTest(sample=p.name):
                self.assertGreater(block_count, 0, f'no blocks parsed for {p.name}')

        ratio = (total_unknown / total_blocks) if total_blocks else 0.0

        report = {
            'meta': {
                'sample_dir': str(self.sample_dir),
                'threshold_unknown_ratio': threshold,
                'total_files': len(self.files),
            },
            'aggregate': {
                'total_blocks': total_blocks,
                'total_unknown': total_unknown,
                'unknown_ratio': round(ratio, 6),
            },
            'files': per_file,
        }

        rp = _report_path()
        rp.parent.mkdir(parents=True, exist_ok=True)
        rp.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

        self.assertLessEqual(
            ratio,
            threshold,
            f'unknown ratio {ratio:.4f} exceeds threshold {threshold:.4f} '
            f'(unknown={total_unknown}, blocks={total_blocks})',
        )


if __name__ == '__main__':
    unittest.main()
