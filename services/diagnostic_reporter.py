import json
from datetime import datetime, timezone
from pathlib import Path

from models.structure import StructureAnalysisResult


class ReportBuildError(RuntimeError):
    pass


class ReportWriteError(RuntimeError):
    pass


class DiagnosticReporter:
    REQUIRED_KEYS = (
        "baseline",
        "summary",
        "paragraph_semantic_labels",
        "paragraph_location_types",
        "blocks",
    )

    @staticmethod
    def build_report(
        analysis: StructureAnalysisResult,
        applied_result: dict,
        source_path: str | Path | None = None,
        output_path: str | Path | None = None,
    ) -> dict:
        if analysis is None:
            raise ReportBuildError("识别报告构建失败：analysis 不能为空")
        if not isinstance(applied_result, dict):
            raise ReportBuildError("识别报告构建失败：applied_result 必须为 dict")

        applied_map = applied_result.get("applied_rule", applied_result)
        split_map = applied_result.get("paragraph_split_applied", {})
        bold_map = applied_result.get("runs_bold_cleared", {})

        if not isinstance(applied_map, dict):
            raise ReportBuildError("识别报告构建失败：applied_rule 必须为 dict")
        if not isinstance(split_map, dict) or not isinstance(bold_map, dict):
            raise ReportBuildError("识别报告构建失败：split/bold 映射必须为 dict")

        blocks = []
        for block in analysis.blocks:
            blocks.append(
                {
                    "block_id": block.block_id,
                    "paragraph_key": block.paragraph_key,
                    "source": block.source,
                    "source_type": block.source_type,
                    "location_type": block.location_type,
                    "raw_text": block.raw_text,
                    "clean_text": block.clean_text,
                    "semantic_label": block.semantic_label,
                    "scores": block.scores,
                    "reasons": block.reasons,
                    "body_baseline_match_score": block.body_baseline_match_score,
                    "final_confidence": block.final_confidence,
                    "in_table": block.features.in_table,
                    "from_soft_break": block.from_soft_break,
                    "from_inline_split": block.from_inline_split,
                    "features": block.features.to_dict(),
                    "applied_rule": applied_map.get(block.paragraph_key, "skip"),
                    "paragraph_split_applied": split_map.get(block.paragraph_key, False),
                    "runs_bold_cleared": bold_map.get(block.paragraph_key, False),
                }
            )

        report = {
            "meta": {
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "source_path": str(source_path) if source_path is not None else None,
                "output_path": str(output_path) if output_path is not None else None,
            },
            "baseline": analysis.baseline.__dict__,
            "summary": analysis.summary(),
            "paragraph_semantic_labels": analysis.paragraph_semantic_labels,
            "paragraph_location_types": analysis.paragraph_location_types,
            "blocks": blocks,
        }

        DiagnosticReporter._validate_payload(report)
        return report

    @staticmethod
    def write_json(path: str | Path, payload: dict) -> Path:
        p = Path(path)
        DiagnosticReporter._validate_payload(payload)
        p.parent.mkdir(parents=True, exist_ok=True)
        try:
            p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError as exc:
            raise ReportWriteError(f"识别报告写入失败：{exc}") from exc
        return p

    @staticmethod
    def _validate_payload(payload: dict) -> None:
        if not isinstance(payload, dict):
            raise ReportWriteError("识别报告写入失败：payload 必须为 dict")

        for key in DiagnosticReporter.REQUIRED_KEYS:
            if key not in payload:
                raise ReportWriteError(f"识别报告写入失败：缺少字段 {key}")

        if not isinstance(payload["blocks"], list):
            raise ReportWriteError("识别报告写入失败：blocks 必须为 list")
