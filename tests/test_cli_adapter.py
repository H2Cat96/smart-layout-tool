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
