import unittest
from pathlib import Path

from teaching_layout.legacy_adapter import TemplatePaths


class CliAdapterTests(unittest.TestCase):
    def test_template_paths_resolve_known_files(self):
        root = Path("/tmp/example-template")
        paths = TemplatePaths.from_root(root)

        self.assertEqual(paths.generator, root / "模板抽取输出" / "generate_print_pdf.py")
        self.assertEqual(paths.template_json, root / "模板抽取输出" / "template.json")
        self.assertEqual(paths.font_map, root / "模板抽取输出" / "font-map.json")
        self.assertEqual(paths.svg_dir, root / "Document fonts" / "SVG")
