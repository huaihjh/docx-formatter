import json
from pathlib import Path

from models.structure import StructureAnalysisResult


class DiagnosticReporter:
    @staticmethod
    def build_report(
        analysis: StructureAnalysisResult,
        applied_result: dict,
    ) -> dict:
        applied_map = applied_result.get("applied_rule", applied_result)
        split_map = applied_result.get("paragraph_split_applied", {})
        bold_map = applied_result.get("runs_bold_cleared", {})

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

        return {
            "baseline": analysis.baseline.__dict__,
            "summary": analysis.summary(),
            "paragraph_semantic_labels": analysis.paragraph_semantic_labels,
            "paragraph_location_types": analysis.paragraph_location_types,
            "blocks": blocks,
        }

    @staticmethod
    def write_json(path: str | Path, payload: dict) -> Path:
        p = Path(path)
        p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return p
