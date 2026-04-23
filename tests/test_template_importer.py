import json
import tempfile
import unittest
from pathlib import Path

from teaching_layout.template_importer import import_template


class TemplateImporterTests(unittest.TestCase):
    def test_import_template_copies_user_assets_and_generates_package_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source"
            templates_dir = root / "templates"
            (source / "assets" / "svg").mkdir(parents=True)
            (source / "assets" / "images").mkdir(parents=True)
            (source / "assets" / "fonts").mkdir(parents=True)
            (source / "assets" / "references").mkdir(parents=True)
            (source / "assets" / "svg" / "标题角标.svg").write_text("<svg/>", encoding="utf-8")
            (source / "assets" / "images" / "preview.png").write_bytes(b"png")
            (source / "assets" / "fonts" / "FZKTK.TTF").write_bytes(b"font")
            (source / "assets" / "references" / "sample.pdf").write_bytes(b"%PDF")

            result = import_template(
                template_id="sample-template",
                name="示例模板",
                source_dir=source,
                templates_dir=templates_dir,
            )

            package_root = templates_dir / "sample-template"
            self.assertEqual(result["status"], "draft_template_ready")
            self.assertEqual(result["template_id"], "sample-template")
            self.assertEqual(Path(result["package_path"]), package_root)
            self.assertTrue((package_root / "assets" / "svg" / "标题角标.svg").exists())
            self.assertTrue((package_root / "assets" / "images" / "preview.png").exists())
            self.assertTrue((package_root / "assets" / "fonts" / "FZKTK.TTF").exists())
            self.assertTrue((package_root / "assets" / "references" / "sample.pdf").exists())

            package = json.loads((package_root / "template-package.json").read_text(encoding="utf-8"))
            self.assertEqual(package["schema_version"], "0.1")
            self.assertEqual(package["id"], "sample-template")
            self.assertEqual(package["engine"]["type"], "template-skeleton")
            self.assertEqual(package["paths"]["asset_map"], "asset-map.json")

            asset_map = json.loads((package_root / "asset-map.json").read_text(encoding="utf-8"))
            self.assertEqual(asset_map["assets"]["标题角标"]["file"], "标题角标.svg")
            self.assertEqual(asset_map["assets"]["标题角标"]["source"], "assets/svg/标题角标.svg")

            font_map = json.loads((package_root / "extracted" / "font-map.json").read_text(encoding="utf-8"))
            self.assertEqual(font_map["fonts"]["FZKTK"]["path"], "../assets/fonts/FZKTK.TTF")

    def test_import_template_marks_missing_pdf_or_idml_reference(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source"
            (source / "assets" / "svg").mkdir(parents=True)

            result = import_template(
                template_id="needs-reference",
                name="缺少参考模板",
                source_dir=source,
                templates_dir=root / "templates",
            )

            self.assertEqual(result["status"], "needs_layout_reference")
            self.assertEqual(result["missing"], ["assets/references/pdf_or_idml"])
            package_root = Path(result["package_path"])
            layout_rules = json.loads((package_root / "layout-rules.json").read_text(encoding="utf-8"))
            self.assertEqual(layout_rules["status"], "needs_layout_reference")

    def test_import_template_refuses_to_overwrite_existing_package_by_default(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source"
            package = root / "templates" / "existing"
            source.mkdir()
            package.mkdir(parents=True)

            with self.assertRaises(FileExistsError):
                import_template(
                    template_id="existing",
                    name="Existing",
                    source_dir=source,
                    templates_dir=root / "templates",
                )
