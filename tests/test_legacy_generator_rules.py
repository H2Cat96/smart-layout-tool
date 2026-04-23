import importlib.util
import tempfile
import unittest
from pathlib import Path

from docx import Document


GENERATOR_PATH = (
    Path(__file__).resolve().parents[1]
    / "templates"
    / "gonggu-neiye"
    / "legacy"
    / "generate_print_pdf.py"
)


def load_generator():
    spec = importlib.util.spec_from_file_location("legacy_generator_under_test", GENERATOR_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class LegacyGeneratorRulesTests(unittest.TestCase):
    def test_load_docx_paragraphs_preserves_tables_and_underlined_runs(self):
        generator = load_generator()
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "sample.docx"
            doc = Document()
            p = doc.add_paragraph()
            p.add_run("普通文字")
            p.add_run("画线文字").underline = True
            table = doc.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "事件"
            table.cell(0, 1).text = "作者感情"
            table.cell(1, 0).text = "听到消息"
            table.cell(1, 1).text = "伤感"
            doc.save(docx_path)

            blocks = generator.load_docx_paragraphs(docx_path)

            self.assertEqual(blocks[0]["type"], "paragraph")
            self.assertEqual(blocks[0]["text"], "普通文字画线文字")
            self.assertEqual(blocks[0]["underline_ranges"], [[4, 8]])
            self.assertEqual(blocks[1]["type"], "table")
            self.assertEqual(blocks[1]["rows"], [["事件", "作者感情"], ["听到消息", "伤感"]])

    def test_gonggu_specific_style_fixes(self):
        generator = load_generator()
        fonts = {
            "title": "Title",
            "title_mid": "TitleMid",
            "title_bold": "TitleBold",
            "question": "Question",
            "body": "Body",
            "footer": "Footer",
        }

        article_title = generator.paragraph_style("泉", 3, False, fonts, {})
        second_article_title = generator.paragraph_style("一匹骆驼", 39, False, fonts, {})
        second_reading_prompt = generator.paragraph_style("阅读《一匹骆驼》，完成下面的小题。", 38, False, fonts, {})
        author = generator.paragraph_style("贾平凹", 4, False, fonts, {})
        source = generator.paragraph_style("（2011-2012北京顺义九上期末）", 15, False, fonts, {})
        answer_label = generator.paragraph_style("【答案】", 2, True, fonts, {})
        answer_section = generator.paragraph_style("【练习二】", 1, True, fonts, {})

        self.assertEqual(article_title["font"], "TitleMid")
        self.assertEqual(second_article_title["kind"], "article_title")
        self.assertEqual(second_article_title["font"], "TitleMid")
        self.assertEqual(second_reading_prompt["kind"], "reading_prompt")
        self.assertEqual(author["font"], "Body")
        self.assertEqual(author["size"], 14)
        self.assertEqual(source["align"], "left")
        self.assertEqual(source["color"], "#898989")
        self.assertEqual(answer_label["kind"], "answer_label")
        self.assertFalse(answer_label["bar"])
        self.assertNotIn("bold_rule", answer_label)
        self.assertEqual(answer_section["kind"], "section")
        self.assertTrue(answer_section["bar"])

    def test_question_wrapping_aligns_content_with_title(self):
        generator = load_generator()
        style = {
            "font": "Helvetica",
            "size": 10,
            "leading": 12,
            "left_indent": 28,
            "first_line_indent": 0,
        }

        lines = generator.wrap_text_by_widths(
            "请从人物描写的角度，赏析下列句子的表达效果。",
            style,
            [30, 30],
        )

        self.assertGreater(len(lines), 1)
        self.assertEqual(generator.line_x_offset(style, first_line=True), 28)
        self.assertEqual(generator.line_x_offset(style, first_line=False), 28)

    def test_template_shell_draws_side_strips_on_white_background_with_section_color(self):
        generator = load_generator()

        class FakeCanvas:
            def __init__(self):
                self.rects = []
                self.colors = []

            def setFillColor(self, color):
                self.color = color
                self.colors.append((round(color.red, 6), round(color.green, 6), round(color.blue, 6)))

            def rect(self, x, y, w, h, fill, stroke):
                self.rects.append((x, y, w, h, fill, stroke))

            def setFont(self, font, size):
                pass

            def drawCentredString(self, x, y, text):
                pass

        canvas = FakeCanvas()
        spread = {
            "spread_index": 2,
            "slots": [
                {
                    "role": "side_strip",
                    "bbox_pt": {"x": 0, "y": 0, "w": 20, "h": 100},
                    "fill_rgb": "#ffd9ed",
                }
            ],
        }
        rules = {"colors": {"practice_bar": "#fce5e4"}}

        generator.paint_template_shell(canvas, spread, 200, 100, "Footer", "white", rules, page_number_start=1)

        self.assertEqual(canvas.rects, [(0, 0, 20, 100, 1, 0)])
        self.assertEqual(canvas.colors[0], (0.988235, 0.898039, 0.894118))

    def test_paragraph_style_uses_layout_rule_overrides(self):
        generator = load_generator()
        fonts = {
            "title": "Title",
            "title_mid": "TitleMid",
            "title_bold": "TitleBold",
            "question": "Question",
            "body": "Body",
            "footer": "Footer",
        }
        rules = {
            "styles": {
                "option": {"font": "OptionFont", "size": 15, "leading": 28, "left_indent": 31},
                "section_title": {"font": "SectionFont", "size": 17, "leading": 25},
            },
            "colors": {
                "practice_bar": "#eeeeee",
                "answer_gray": "#777777",
            },
        }

        option = generator.paragraph_style("A. 选项文字", 10, False, fonts, rules)
        practice_section = generator.paragraph_style("【练习二】", 1, False, fonts, rules)
        answer_section = generator.paragraph_style("【练习二】", 1, True, fonts, rules)

        self.assertEqual(option["font"], "OptionFont")
        self.assertEqual(option["size"], 15)
        self.assertEqual(option["leading"], 28)
        self.assertEqual(option["left_indent"], 31)
        self.assertEqual(practice_section["font"], "SectionFont")
        self.assertEqual(practice_section["bar_color"], "#eeeeee")
        self.assertEqual(answer_section["bar_color"], "#777777")

    def test_load_svg_assets_uses_asset_map_filenames_and_color_overrides(self):
        generator = load_generator()
        with tempfile.TemporaryDirectory() as tmp:
            svg_dir = Path(tmp)
            (svg_dir / "custom-corner.svg").write_text(
                '<svg viewBox="0 0 30.05 30.05"><path fill="#111111" d="M0 0h1v1z"/></svg>',
                encoding="utf-8",
            )
            (svg_dir / "custom-answer.svg").write_text(
                '<svg viewBox="0 0 30.05 30.05"><path fill="#222222" d="M0 0h1v1z"/></svg>',
                encoding="utf-8",
            )
            (svg_dir / "custom-bar.svg").write_text(
                '<svg viewBox="0 0 481.89 28.35"><path fill="#333333" d="M0 0h1v1z"/></svg>',
                encoding="utf-8",
            )
            (svg_dir / "custom-badge.svg").write_text(
                '<svg viewBox="0 0 22.68 12.76"><path fill="#444444" d="M0 0h1v1z"/></svg>',
                encoding="utf-8",
            )
            asset_map = {
                "assets": {
                    "practice_title_corner": {"file": "custom-corner.svg", "color": "#aaaaaa"},
                    "answer_title_corner": {"file": "custom-answer.svg", "color": "#bbbbbb"},
                    "section_bar": {"file": "custom-bar.svg", "practice_color": "#cccccc"},
                    "question_badge": {"file": "custom-badge.svg", "color": "#dddddd"},
                }
            }

            assets = generator.load_svg_assets(svg_dir, asset_map)

            self.assertEqual(Path(assets["标题角标"]["path"]).name, "custom-corner.svg")
            self.assertEqual(assets["标题角标"]["fill"], "#aaaaaa")
            self.assertEqual(Path(assets["参考答案"]["path"]).name, "custom-answer.svg")
            self.assertEqual(assets["参考答案"]["fill"], "#bbbbbb")
            self.assertEqual(Path(assets["标题"]["path"]).name, "custom-bar.svg")
            self.assertEqual(assets["标题"]["fill"], "#cccccc")
            self.assertEqual(Path(assets["序号"]["path"]).name, "custom-badge.svg")
            self.assertEqual(assets["序号"]["fill"], "#dddddd")
