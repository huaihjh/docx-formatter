import re

from models.format_rule import FormatRule, SectionRule


class RuleParseError(Exception):
    pass


class RuleParser:
    CN_SIZE_MAP = {
        "小六": 6.5,
        "六号": 7.5,
        "小五": 9,
        "五号": 10.5,
        "小四": 12,
        "四号": 14,
        "小三": 15,
        "三号": 16,
        "小二": 18,
        "二号": 22,
        "小一": 24,
        "一号": 26,
    }

    FONT_CANDIDATES = ["宋体", "黑体", "微软雅黑", "楷体", "仿宋"]

    TARGET_KEYWORDS = {
        "title": ["标题"],
        "body": ["正文", "段落"],
        "table": ["表格", "表内", "表中", "单元格"],
    }

    @classmethod
    def parse(cls, text: str) -> FormatRule:
        content = (text or "").strip()
        if not content:
            raise RuleParseError("格式要求不能为空")

        clauses = cls._split_clauses(content)
        target_clause_map = cls._group_clauses_by_target(clauses)

        rule = FormatRule()
        for target, joined_text in target_clause_map.items():
            if joined_text.strip():
                setattr(rule, target, cls._parse_section(joined_text))

        if not any(rule.to_dict().values()):
            raise RuleParseError("未识别到有效的格式规则，请检查输入")

        return rule

    @staticmethod
    def _split_clauses(text: str) -> list[str]:
        parts = re.split(r"[，,。；;\n]+", text)
        return [part.strip() for part in parts if part.strip()]

    @classmethod
    def _group_clauses_by_target(cls, clauses: list[str]) -> dict[str, str]:
        grouped: dict[str, list[str]] = {"title": [], "body": [], "table": []}

        active_targets = ["body"]

        for clause in clauses:
            explicit_targets = cls._targets_in_clause(clause)
            if explicit_targets:
                active_targets = explicit_targets

            for target in active_targets:
                grouped[target].append(clause)

        return {target: "，".join(texts) for target, texts in grouped.items()}

    @classmethod
    def _targets_in_clause(cls, clause: str) -> list[str]:
        targets: list[str] = []
        for target, keywords in cls.TARGET_KEYWORDS.items():
            if any(keyword in clause for keyword in keywords):
                targets.append(target)
        return targets

    @classmethod
    def _parse_section(cls, text: str) -> SectionRule:
        section = SectionRule()
        lower_text = text.lower()

        font = cls._parse_font(text)
        if font:
            section.font_name = font

        size = cls._parse_size(text, lower_text)
        if size:
            section.font_size = size

        bold = cls._parse_bold(text)
        if bold is not None:
            section.bold = bold

        alignment = cls._parse_alignment(text)
        if alignment:
            section.alignment = alignment

        line_spacing = cls._parse_line_spacing(text, lower_text)
        if line_spacing:
            section.line_spacing = line_spacing

        indent = cls._parse_first_line_indent(text, lower_text)
        if indent is not None:
            section.first_line_indent = indent

        space_before = cls._parse_space(text, lower_text, "前")
        if space_before is not None:
            section.space_before = space_before

        space_after = cls._parse_space(text, lower_text, "后")
        if space_after is not None:
            section.space_after = space_after

        return section

    @classmethod
    def _parse_font(cls, text: str) -> str | None:
        for font in cls.FONT_CANDIDATES:
            if font in text:
                return font
        return None

    @classmethod
    def _parse_size(cls, text: str, lower_text: str) -> float | None:
        for cn_size, pt in cls.CN_SIZE_MAP.items():
            if cn_size in text:
                return pt

        pt_match = re.search(r"(?<!段前)(?<!段后)(\d+(?:\.\d+)?)\s*pt", lower_text)
        if pt_match:
            return float(pt_match.group(1))

        return None

    @staticmethod
    def _parse_bold(text: str) -> bool | None:
        negative_patterns = [
            "不加粗",
            "不要加粗",
            "不需要加粗",
            "取消加粗",
            "不粗体",
            "不要粗体",
        ]
        if any(p in text for p in negative_patterns):
            return False
        if "加粗" in text or "粗体" in text:
            return True
        return None

    @staticmethod
    def _parse_alignment(text: str) -> str | None:
        if "居中" in text:
            return "center"
        if "左对齐" in text:
            return "left"
        if "右对齐" in text:
            return "right"
        return None

    @staticmethod
    def _parse_line_spacing(text: str, lower_text: str) -> float | None:
        if "单倍" in text:
            return 1.0
        if "1.5倍" in text or "1.5 倍" in text:
            return 1.5
        if "2倍" in text or "2 倍" in text:
            return 2.0

        match = re.search(r"(\d+(?:\.\d+)?)\s*倍", lower_text)
        if match:
            return float(match.group(1))

        return None

    @staticmethod
    def _parse_first_line_indent(text: str, lower_text: str) -> float | None:
        patterns = [
            r"首行缩进\s*(\d+(?:\.\d+)?)\s*字?符?",
            r"段落缩进\s*(\d+(?:\.\d+)?)\s*字?符?",
            r"缩进\s*(\d+(?:\.\d+)?)\s*字?符?",
        ]
        for p in patterns:
            m = re.search(p, text)
            if m:
                return float(m.group(1))

        lower_patterns = [
            r"首行缩进\s*(\d+(?:\.\d+)?)",
            r"段落缩进\s*(\d+(?:\.\d+)?)",
            r"缩进\s*(\d+(?:\.\d+)?)",
        ]
        for p in lower_patterns:
            m = re.search(p, lower_text)
            if m:
                return float(m.group(1))

        return None

    @staticmethod
    def _parse_space(text: str, lower_text: str, direction: str) -> float | None:
        pattern = rf"段{direction}\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)?"
        match = re.search(pattern, lower_text)
        if match:
            return float(match.group(1))

        pattern_cn = rf"段{direction}\s*(\d+(?:\.\d+)?)"
        match_cn = re.search(pattern_cn, text)
        if match_cn:
            return float(match_cn.group(1))

        return None
