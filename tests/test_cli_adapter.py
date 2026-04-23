import unittest
from pathlib import Path
from unittest.mock import patch

from teaching_layout.cli import main
from teaching_layout.legacy_adapter import TemplatePaths


class CliAdapterTests(unittest.TestCase):
    def test_template_paths_resolve_known_files(self):
        root = Path("/tmp/example-template")
        paths = TemplatePaths.from_root(root)

        self.assertEqual(paths.generator, root / "模板抽取输出" / "generate_print_pdf.py")
        self.assertEqual(paths.template_json, root / "模板抽取输出" / "template.json")
        self.assertEqual(paths.font_map, root / "模板抽取输出" / "font-map.json")
        self.assertEqual(paths.svg_dir, root / "Document fonts" / "SVG")

    def test_import_template_command_calls_importer(self):
        with patch("teaching_layout.cli.import_template") as importer:
            importer.return_value = {
                "schema_version": "0.1",
                "template_id": "new-template",
                "status": "needs_layout_reference",
                "package_path": "/tmp/templates/new-template",
                "missing": ["assets/references/pdf_or_idml"],
            }

            main(
                [
                    "import-template",
                    "--id",
                    "new-template",
                    "--name",
                    "新模板",
                    "--source",
                    "/tmp/source",
                    "--templates-dir",
                    "/tmp/templates",
                ]
            )

            importer.assert_called_once_with(
                template_id="new-template",
                name="新模板",
                source_dir=Path("/tmp/source"),
                templates_dir=Path("/tmp/templates"),
                overwrite=False,
            )

    def test_inspect_template_command_calls_inspector(self):
        with patch("teaching_layout.cli.inspect_template") as inspector:
            inspector.return_value = {
                "template": "示例模板",
                "status": "portable",
                "missing": [],
                "absolute_paths": [],
            }

            main(["inspect-template", "--template", "示例模板", "--templates-dir", "/tmp/templates"])

            inspector.assert_called_once_with("示例模板", templates_dir=Path("/tmp/templates"))

    def test_build_command_defaults_to_single_page_mode(self):
        with patch("teaching_layout.cli.build_with_legacy_generator") as builder:
            builder.return_value = {"outputs": {"pdf": "/tmp/out.pdf"}}

            main(
                [
                    "build",
                    "--template",
                    "人文-课后巩固",
                    "--docx",
                    "/tmp/input.docx",
                    "--out",
                    "/tmp/out",
                ]
            )

            self.assertEqual(builder.call_args.kwargs["page_mode"], "single")
            self.assertEqual(builder.call_args.kwargs["color_mode"], "cmyk")
