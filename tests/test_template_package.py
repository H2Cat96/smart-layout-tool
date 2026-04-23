import json
import tempfile
import unittest
from pathlib import Path

from teaching_layout.legacy_adapter import resolve_template_paths
from teaching_layout.template_package import TemplatePackage


class TemplatePackageTests(unittest.TestCase):
    def test_load_template_package_resolves_source_relative_paths(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source-materials"
            package = root / "templates" / "sample"
            package.mkdir(parents=True)
            (package / "template-package.json").write_text(
                json.dumps(
                    {
                        "id": "sample",
                        "name": "Sample Template",
                        "engine": "legacy-generator",
                        "source_root": str(source),
                        "paths": {
                            "generator": "extracted/generate.py",
                            "template_json": "extracted/template.json",
                            "font_map": "extracted/font-map.json",
                            "svg_dir": "fonts/SVG",
                        },
                        "package_paths": {
                            "layout_rules": "layout-rules.json",
                            "asset_map": "asset-map.json"
                        },
                    }
                ),
                encoding="utf-8",
            )

            loaded = TemplatePackage.from_root(package)

            self.assertEqual(loaded.id, "sample")
            self.assertEqual(loaded.generator, source / "extracted" / "generate.py")
            self.assertEqual(loaded.template_json, source / "extracted" / "template.json")
            self.assertEqual(loaded.font_map, source / "extracted" / "font-map.json")
            self.assertEqual(loaded.svg_dir, source / "fonts" / "SVG")
            self.assertEqual(loaded.layout_rules, package / "layout-rules.json")
            self.assertEqual(loaded.asset_map, package / "asset-map.json")

    def test_resolve_template_paths_supports_named_template_package(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source-materials"
            package = root / "sample"
            package.mkdir()
            (package / "template-package.json").write_text(
                json.dumps(
                    {
                        "id": "sample",
                        "name": "Sample Template",
                        "engine": "legacy-generator",
                        "source_root": str(source),
                        "paths": {
                            "generator": "extracted/generate.py",
                            "template_json": "extracted/template.json",
                            "font_map": "extracted/font-map.json",
                            "svg_dir": "fonts/SVG",
                        },
                        "package_paths": {
                            "layout_rules": "layout-rules.json",
                            "asset_map": "asset-map.json"
                        },
                    }
                ),
                encoding="utf-8",
            )

            paths = resolve_template_paths("sample", templates_dir=root)

            self.assertEqual(paths.root, source)
            self.assertEqual(paths.generator, source / "extracted" / "generate.py")
            self.assertEqual(paths.layout_rules, package / "layout-rules.json")
            self.assertEqual(paths.asset_map, package / "asset-map.json")

    def test_builtin_gonggu_template_resolves_repo_local_files(self):
        repo_root = Path(__file__).resolve().parents[1]

        paths = resolve_template_paths("gonggu-neiye")

        self.assertEqual(paths.root, repo_root / "templates" / "gonggu-neiye")
        self.assertEqual(paths.generator, paths.root / "legacy" / "generate_print_pdf.py")
        self.assertEqual(paths.template_json, paths.root / "extracted" / "template.json")
        self.assertEqual(paths.font_map, paths.root / "extracted" / "font-map.json")
        self.assertEqual(paths.svg_dir, paths.root / "assets" / "svg")
