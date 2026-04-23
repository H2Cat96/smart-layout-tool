import json
import tempfile
import unittest
from pathlib import Path

from teaching_layout.template_inspector import inspect_template


class TemplateInspectorTests(unittest.TestCase):
    def test_inspect_template_resolves_relative_fonts_and_reports_portable(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package = root / "sample"
            (package / "legacy").mkdir(parents=True)
            (package / "extracted").mkdir()
            (package / "assets" / "fonts").mkdir(parents=True)
            (package / "assets" / "svg").mkdir(parents=True)
            (package / "legacy" / "generate.py").write_text("# stub", encoding="utf-8")
            (package / "extracted" / "template.json").write_text("{}", encoding="utf-8")
            (package / "assets" / "fonts" / "Body.ttf").write_bytes(b"font")
            (package / "template-package.json").write_text(
                json.dumps(
                    {
                        "id": "sample",
                        "name": "示例模板",
                        "engine": "legacy-generator",
                        "source_root": ".",
                        "paths": {
                            "generator": "legacy/generate.py",
                            "template_json": "extracted/template.json",
                            "font_map": "extracted/font-map.json",
                            "svg_dir": "assets/svg",
                        },
                    }
                ),
                encoding="utf-8",
            )
            (package / "extracted" / "font-map.json").write_text(
                json.dumps(
                    {
                        "fonts": {
                            "Body": {
                                "path": "../assets/fonts/Body.ttf",
                                "reportlab_name": "Body",
                                "status": "exact",
                            }
                        },
                        "style_defaults": {"body": "Body"},
                    }
                ),
                encoding="utf-8",
            )

            result = inspect_template("sample", templates_dir=root)

            self.assertEqual(result["status"], "portable")
            self.assertEqual(result["missing"], [])
            self.assertEqual(result["absolute_paths"], [])
            self.assertEqual(result["fonts"][0]["resolved_path"], str((package / "assets" / "fonts" / "Body.ttf").resolve()))

    def test_inspect_template_reports_absolute_paths_and_missing_fonts(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package = root / "sample"
            (package / "legacy").mkdir(parents=True)
            (package / "extracted").mkdir()
            (package / "assets" / "svg").mkdir(parents=True)
            (package / "legacy" / "generate.py").write_text("# stub", encoding="utf-8")
            (package / "extracted" / "template.json").write_text(
                json.dumps({"source": {"pdf": "/Users/tal/Desktop/ref.pdf"}}),
                encoding="utf-8",
            )
            (package / "template-package.json").write_text(
                json.dumps(
                    {
                        "id": "sample",
                        "name": "示例模板",
                        "engine": "legacy-generator",
                        "source_root": ".",
                        "paths": {
                            "generator": "legacy/generate.py",
                            "template_json": "extracted/template.json",
                            "font_map": "extracted/font-map.json",
                            "svg_dir": "assets/svg",
                        },
                    }
                ),
                encoding="utf-8",
            )
            (package / "extracted" / "font-map.json").write_text(
                json.dumps(
                    {
                        "fonts": {
                            "Body": {
                                "path": "/Users/tal/Desktop/fonts/Body.ttf",
                                "reportlab_name": "Body",
                                "status": "exact",
                            }
                        },
                        "style_defaults": {"body": "Body"},
                    }
                ),
                encoding="utf-8",
            )

            result = inspect_template("sample", templates_dir=root)

            self.assertEqual(result["status"], "not_portable")
            self.assertIn("font:Body", result["missing"])
            self.assertEqual(
                {(item["file"], item["json_path"]) for item in result["absolute_paths"]},
                {
                    ("extracted/font-map.json", "fonts.Body.path"),
                    ("extracted/template.json", "source.pdf"),
                },
            )
