import unittest

from models.format_rule import FormatRule, SectionRule
from services.formatter import Formatter


class FormatterTests(unittest.TestCase):
    def test_rule_for_prefers_table_location(self) -> None:
        rule = FormatRule(
            title=SectionRule(font_name="TitleFont"),
            body=SectionRule(font_name="BodyFont"),
            table=SectionRule(font_name="TableFont"),
            list_item=SectionRule(font_name="ListFont"),
            inline_subheading=SectionRule(font_name="InlineFont"),
            caption=SectionRule(font_name="CaptionFont"),
        )

        self.assertIs(Formatter._rule_for("main_heading", "table_cell", rule), rule.table)
        self.assertIs(Formatter._rule_for("main_heading", "paragraph", rule), rule.title)
        self.assertIs(Formatter._rule_for("caption", "paragraph", rule), rule.caption)
        self.assertIsNone(Formatter._rule_for("unknown", "paragraph", rule))

    def test_table_rule_inherits_body_bold_when_table_bold_missing(self) -> None:
        rule = FormatRule(
            body=SectionRule(bold=False),
            table=SectionRule(font_name="宋体", font_size=10.5),
        )
        section = Formatter._rule_for("main_heading", "table_cell", rule)
        self.assertIsNotNone(section)
        self.assertIs(section.bold, False)

    def test_table_body_inherits_body_indent_when_table_indent_missing(self) -> None:
        rule = FormatRule(
            body=SectionRule(first_line_indent=2.0),
            table=SectionRule(font_name="宋体", font_size=10.5),
        )
        section = Formatter._rule_for("body", "table_cell", rule)
        self.assertIsNotNone(section)
        self.assertEqual(section.first_line_indent, 2.0)


if __name__ == "__main__":
    unittest.main()
