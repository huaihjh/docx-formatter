import unittest

from services.rule_parser import RuleParseError, RuleParser


class RuleParserTests(unittest.TestCase):
    def test_parse_basic_cn_rule(self) -> None:
        text = "标题黑体三号居中，正文宋体小四首行缩进2字符，1.5倍行距，正文不要加粗，表格宋体五号"
        rule = RuleParser.parse(text)
        rule.normalize()

        self.assertEqual(rule.title.font_name, "黑体")
        self.assertEqual(rule.title.font_size, 16)
        self.assertEqual(rule.title.alignment, "center")

        self.assertEqual(rule.body.font_name, "宋体")
        self.assertEqual(rule.body.font_size, 12)
        self.assertEqual(rule.body.first_line_indent, 2.0)
        self.assertEqual(rule.body.line_spacing, 1.5)
        self.assertIs(rule.body.bold, False)

        self.assertEqual(rule.table.font_name, "宋体")
        self.assertEqual(rule.table.font_size, 10.5)

    def test_parse_empty_raises(self) -> None:
        with self.assertRaises(RuleParseError):
            RuleParser.parse("   ")

    def test_negative_bold_phrase(self) -> None:
        rule = RuleParser.parse("正文不要加粗")
        self.assertIs(rule.body.bold, False)

    def test_negative_center_phrase_maps_to_left(self) -> None:
        rule = RuleParser.parse("\u6807\u9898\u9ed1\u4f53\u4e09\u53f7\u4e0d\u5c45\u4e2d")
        self.assertEqual(rule.title.alignment, "left")

    def test_negative_center_overrides_positive_center(self) -> None:
        rule = RuleParser.parse("\u6807\u9898\u5c45\u4e2d\u4f46\u4e0d\u5c45\u4e2d")
        self.assertEqual(rule.title.alignment, "left")


if __name__ == "__main__":
    unittest.main()
