import unittest

from docx import Document

from services.structure_analyzer import StructureAnalyzer


class StructureAnalyzerTests(unittest.TestCase):
    def test_caption_regex_accepts_bracket_and_colon(self) -> None:
        self.assertTrue(StructureAnalyzer.CAPTION_RE.match("[ 图 1："))
        self.assertTrue(StructureAnalyzer.CAPTION_RE.match("图1："))
        self.assertTrue(StructureAnalyzer.CAPTION_RE.match("表 2.3"))

    def test_list_like_distinguishes_multilevel_numbering(self) -> None:
        doc = Document()

        p_multilevel = doc.add_paragraph("1.1 研究背景")
        p_simple_list = doc.add_paragraph("1. 列表项")

        self.assertFalse(StructureAnalyzer._list_like(p_multilevel.text, p_multilevel))
        self.assertTrue(StructureAnalyzer._list_like(p_simple_list.text, p_simple_list))


if __name__ == "__main__":
    unittest.main()
