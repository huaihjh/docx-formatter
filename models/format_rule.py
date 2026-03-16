from dataclasses import dataclass, field
from typing import Any


@dataclass
class SectionRule:
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    alignment: str | None = None
    line_spacing: float | None = None
    first_line_indent: float | None = None
    space_before: float | None = None
    space_after: float | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            key: value
            for key, value in {
                "font_name": self.font_name,
                "font_size": self.font_size,
                "bold": self.bold,
                "alignment": self.alignment,
                "line_spacing": self.line_spacing,
                "first_line_indent": self.first_line_indent,
                "space_before": self.space_before,
                "space_after": self.space_after,
            }.items()
            if value is not None
        }


@dataclass
class FormatRule:
    title: SectionRule = field(default_factory=SectionRule)
    body: SectionRule = field(default_factory=SectionRule)
    table: SectionRule = field(default_factory=SectionRule)
    list_item: SectionRule = field(default_factory=SectionRule)
    inline_subheading: SectionRule = field(default_factory=SectionRule)
    caption: SectionRule = field(default_factory=SectionRule)

    def normalize(self) -> None:
        body_defaults = self.body.to_dict()

        if not self.table.to_dict():
            defaults = dict(body_defaults)
            defaults["first_line_indent"] = None
            self.table = SectionRule(**defaults)

        if not self.list_item.to_dict():
            defaults = dict(body_defaults)
            defaults["first_line_indent"] = None
            self.list_item = SectionRule(**defaults)

        if not self.inline_subheading.to_dict():
            defaults = dict(body_defaults)
            defaults.update(
                {
                    "bold": self.body.bold if self.body.bold is not None else True,
                    "alignment": "left",
                    "first_line_indent": None,
                }
            )
            self.inline_subheading = SectionRule(**defaults)

        if not self.caption.to_dict():
            defaults = dict(body_defaults)
            defaults["first_line_indent"] = None
            self.caption = SectionRule(**defaults)

    def to_dict(self) -> dict[str, dict[str, Any]]:
        return {
            "title": self.title.to_dict(),
            "body": self.body.to_dict(),
            "table": self.table.to_dict(),
            "list_item": self.list_item.to_dict(),
            "inline_subheading": self.inline_subheading.to_dict(),
            "caption": self.caption.to_dict(),
        }
