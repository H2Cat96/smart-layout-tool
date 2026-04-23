import importlib.util
import tempfile
import unittest
from pathlib import Path


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
