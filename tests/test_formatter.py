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


if __name__ == "__main__":
    unittest.main()
