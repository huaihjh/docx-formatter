from dataclasses import asdict, dataclass, field
from typing import Any


SemanticLabel = str
LocationType = str


@dataclass
class AnalyzerConfig:
    split_soft_breaks: bool = True
    split_inline_subheading: bool = True
    keep_list_item_integrity: bool = True
    unknown_score_threshold: float = 2.0
    unknown_margin_threshold: float = 0.35
    body_marker_tokens: tuple[str, ...] = (
        "本",
        "该",
        "通过",
        "针对",
        "为了",
        "其",
    )


@dataclass
class DocumentBaseline:
    font_name: str | None = None
    font_size_pt: float | None = None
    alignment: str | None = None
    paragraph_length: float | None = None
    bold_ratio: float | None = None


@dataclass
class BlockFeatures:
    raw_text: str
    clean_text: str
    text_length: int
    is_empty: bool
    style_name: str
    in_table: bool
    alignment: str
    bold_ratio: float
    italic_ratio: float
    font_size_pt: float | None
    has_soft_break: bool
    has_main_numbering: bool
    has_sub_numbering: bool
    list_item_like: bool
    caption_like: bool
    ends_with_colon: bool
    ends_with_period: bool
    chinese_ratio: float
    punctuation_density: float
    long_sentence_like: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class LogicalBlock:
    block_id: int
    paragraph_key: str
    raw_text: str
    clean_text: str
    paragraph: Any
    paragraph_index: int
    line_index: int
    split_index: int
    source_type: str
    location_type: LocationType
    table_path: str | None
    row_index: int | None
    col_index: int | None
    from_soft_break: bool
    from_inline_split: bool
    features: BlockFeatures
    semantic_label: SemanticLabel = "unknown"
    scores: dict[str, float] = field(default_factory=dict)
    reasons: list[str] = field(default_factory=list)
    body_baseline_match_score: float = 0.0
    final_confidence: float = 0.0

    @property
    def source(self) -> str:
        if self.table_path is not None:
            return (
                f"table[{self.table_path}]"
                f"/row[{self.row_index}]"
                f"/col[{self.col_index}]"
                f"/p[{self.paragraph_index}]"
                f"/line[{self.line_index}]"
                f"/split[{self.split_index}]"
            )
        return f"document/p[{self.paragraph_index}]/line[{self.line_index}]/split[{self.split_index}]"


@dataclass
class StructureAnalysisResult:
    baseline: DocumentBaseline
    blocks: list[LogicalBlock]
    paragraph_semantic_labels: dict[str, SemanticLabel]
    paragraph_location_types: dict[str, LocationType]

    def summary(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for block in self.blocks:
            key = block.semantic_label
            counts[key] = counts.get(key, 0) + 1
        return counts
