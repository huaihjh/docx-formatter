# Docx Formatter

Desktop app (PySide6 + python-docx) to auto-adjust `.docx` formatting from natural-language rules and export both a new document and a diagnostic report.

## Goal
Build a generic pipeline:
1. Parse docx structure
2. Split logical blocks
3. Extract features
4. Classify with rules + scoring
5. Apply formatting by class
6. Export doc + report

## Labels
- `semantic_label`: `main_heading`, `sub_heading`, `inline_subheading`, `body`, `list_item`, `caption`, `unknown`
- `location_type`: `paragraph`, `table_cell`

## Main Modules
- `main.py`: app entry
- `ui/main_window.py`: UI and workflow
- `services/rule_parser.py`: NL rule parsing
- `services/structure_analyzer.py`: block split + feature + classification
- `services/formatter.py`: formatting application
- `services/diagnostic_reporter.py`: JSON report generation
- `services/exporter.py`: document export
- `models/format_rule.py`, `models/structure.py`: data models

## Run
```bash
pip install -r requirements.txt
python main.py
```

## Tests
```bash
python -m unittest discover -s tests -v
```

Test layers:
- `core_regression`: fixed critical samples (anti-regression).
  - default sample dir: `e:\Download\docx_samples`
  - override via `DOCX_SAMPLE_DIR`
- `generalization_suite`: broader sample-set evaluation.
  - requires `DOCX_GENERALIZATION_DIR`
  - skipped automatically if not set

Run only generalization suite:
```bash
set DOCX_GENERALIZATION_DIR=E:\path\to\generalization_samples
set DOCX_GENERALIZATION_MAX_UNKNOWN_RATIO=0.03
set DOCX_GENERALIZATION_REPORT_PATH=memory-bank\baseline\generalization_summary.json
python -m unittest tests.test_generalization_suite -v
```

One-click script:
```bash
scripts\run_generalization.bat
```


`generalization_suite` writes a JSON summary to:
- default: `memory-bank/baseline/generalization_summary.json`
- override: `DOCX_GENERALIZATION_REPORT_PATH`

## Docs
- `docs/PROJECT_DOCUMENTATION.md`
