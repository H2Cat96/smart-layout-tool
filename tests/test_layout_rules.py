import unittest

from teaching_layout.layout_rules import FORBIDDEN_LINE_START, FORBIDDEN_LINE_END, wrap_cjk_text


def char_width(text: str) -> float:
    return float(len(text))


class LayoutRulesTests(unittest.TestCase):
    def test_wrap_cjk_text_avoids_closing_quote_at_line_start_after_period(self):
        lines = wrap_cjk_text("包在怀里。”我一直留神", max_width=6, width_fn=char_width)

        self.assertEqual(lines[0], "包在怀里。”")
        self.assertEqual(lines[1], "我一直留神")
        self.assertFalse(any(line[0] in FORBIDDEN_LINE_START for line in lines if line))

    def test_wrap_cjk_text_avoids_left_quote_at_line_end(self):
        lines = wrap_cjk_text("他说：“可以胶的", max_width=4, width_fn=char_width)

        self.assertFalse(any(line[-1] in FORBIDDEN_LINE_END for line in lines if line))
