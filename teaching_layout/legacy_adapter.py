from __future__ import annotations

import argparse
import importlib.util
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Any

from .template_package import DEFAULT_TEMPLATES_DIR, TemplatePackage


@dataclass(frozen=True)
class TemplatePaths:
    root: Path
    generator: Path
    template_json: Path
    font_map: Path
    svg_dir: Path
    layout_rules: Path | None = None
    asset_map: Path | None = None

    @classmethod
    def from_root(cls, root: Path) -> "TemplatePaths":
        return cls(
            root=root,
            generator=root / "模板抽取输出" / "generate_print_pdf.py",
            template_json=root / "模板抽取输出" / "template.json",
            font_map=root / "模板抽取输出" / "font-map.json",
            svg_dir=root / "Document fonts" / "SVG",
            layout_rules=None,
            asset_map=None,
        )

    @classmethod
    def from_package(cls, package: TemplatePackage) -> "TemplatePaths":
        if package.engine != "legacy-generator":
            raise ValueError(f"Unsupported template engine: {package.engine}")
        return cls(
            root=package.source_root,
            generator=package.generator,
            template_json=package.template_json,
            font_map=package.font_map,
            svg_dir=package.svg_dir,
            layout_rules=package.layout_rules,
            asset_map=package.asset_map,
        )

    def validate(self) -> None:
        for path in (self.generator, self.template_json, self.font_map, self.svg_dir):
            if not path.exists():
                raise FileNotFoundError(f"Template asset missing: {path}")
        for optional_path in (self.layout_rules, self.asset_map):
            if optional_path is not None and not optional_path.exists():
                raise FileNotFoundError(f"Template config missing: {optional_path}")


def resolve_template_paths(template_ref: str | Path, templates_dir: Path = DEFAULT_TEMPLATES_DIR) -> TemplatePaths:
    template_text = str(template_ref)
    candidate = Path(template_text).expanduser()
    if candidate.exists():
        if (candidate / "template-package.json").exists():
            return TemplatePaths.from_package(TemplatePackage.from_root(candidate))
        return TemplatePaths.from_root(candidate)

    named_candidate = templates_dir / template_text
    if named_candidate.exists():
        if (named_candidate / "template-package.json").exists():
            return TemplatePaths.from_package(TemplatePackage.from_root(named_candidate))
        return TemplatePaths.from_root(named_candidate)

    for package_config in sorted(templates_dir.glob("*/template-package.json")):
        package = TemplatePackage.from_root(package_config.parent)
        if template_text in {package.id, package.name, *package.aliases}:
            return TemplatePaths.from_package(package)

    raise FileNotFoundError(
        f"Template not found as path or named package: {template_text} (templates_dir={templates_dir})"
    )


def load_generator(generator_path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location("teaching_layout_legacy_generator", generator_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load generator: {generator_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def build_with_legacy_generator(
    template_ref: str | Path,
    docx: Path,
    output_dir: Path,
    pdf_name: str,
    preview_name: str,
    title: str,
    background_mode: str = "white",
    page_mode: str = "single",
    color_mode: str = "cmyk",
    templates_dir: Path = DEFAULT_TEMPLATES_DIR,
) -> dict[str, Any]:
    paths = resolve_template_paths(template_ref, templates_dir=templates_dir)
    paths.validate()
    output_dir.mkdir(parents=True, exist_ok=True)

    generator = load_generator(paths.generator)
    args = argparse.Namespace(
        template=str(paths.template_json),
        font_map=str(paths.font_map),
        docx=str(docx),
        output_dir=str(output_dir),
        background_dir=None,
        background_mode=background_mode,
        svg_dir=str(paths.svg_dir),
        background_dpi=220,
        pdf_name=pdf_name,
        preview_name=preview_name,
        title=title,
        page_mode=page_mode,
        color_mode=color_mode,
        layout_rules=str(paths.layout_rules) if paths.layout_rules else None,
        asset_map=str(paths.asset_map) if paths.asset_map else None,
    )
    return generator.build_pdf(args)
