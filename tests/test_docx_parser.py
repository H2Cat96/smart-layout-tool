import unittest

from teaching_layout.docx_parser import (
    ANSWER_LINE_SENTINEL,
    clean_text,
    is_underlined_blank_paragraph,
    split_practice_and_answers,
)


class FakeRun:
    def __init__(self, text: str, xml: str = ""):
        self.text = text
        self._element = type("Element", (), {"xml": xml})()


class FakeParagraph:
    def __init__(self, runs):
        self.runs = runs


class DocxParserTests(unittest.TestCase):
    def test_clean_text_preserves_meaningful_inner_spaces(self):
        self.assertEqual(clean_text("1.下列对文章内容的理解（    ）"), "1.下列对文章内容的理解（    ）")
        self.assertEqual(clean_text("阅读\t《一匹骆驼》\n完成小题"), "阅读 《一匹骆驼》 完成小题")

    def test_underlined_blank_paragraph_becomes_answer_line(self):
        paragraph = FakeParagraph([FakeRun("       ", "<w:r><w:rPr><w:u w:val=\"single\"/></w:rPr></w:r>")])

        self.assertTrue(is_underlined_blank_paragraph(paragraph))

    def test_nonblank_underlined_paragraph_is_not_answer_line(self):
        paragraph = FakeParagraph([FakeRun("答案", "<w:r><w:rPr><w:u w:val=\"single\"/></w:rPr></w:r>")])

        self.assertFalse(is_underlined_blank_paragraph(paragraph))

    def test_split_practice_and_answers_keeps_answer_line_before_answers(self):
        paragraphs = [
            "标题",
            "1.题目",
            ANSWER_LINE_SENTINEL,
            "参考答案与解析",
            "答案：",
        ]

        practice, answers = split_practice_and_answers(paragraphs)

        self.assertEqual(practice, ["标题", "1.题目", ANSWER_LINE_SENTINEL])
        self.assertEqual(answers, ["参考答案与解析", "答案："])
