from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATES_DIR = PROJECT_ROOT / "templates"


def _resolve_path(base: Path, value: str) -> Path:
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    return base / path


@dataclass(frozen=True)
class TemplatePackage:
    id: str
    name: str
    root: Path
    engine: str
    source_root: Path
    generator: Path
    template_json: Path
    font_map: Path
    svg_dir: Path
    layout_rules: Path | None = None
    asset_map: Path | None = None

    @classmethod
    def from_root(cls, package_root: Path) -> "TemplatePackage":
        config_path = package_root / "template-package.json"
        if not config_path.exists():
            raise FileNotFoundError(f"Template package config missing: {config_path}")
        data = json.loads(config_path.read_text(encoding="utf-8"))
        return cls.from_config(package_root, data)

    @classmethod
    def from_config(cls, package_root: Path, data: dict[str, Any]) -> "TemplatePackage":
        source_root = _resolve_path(package_root, data["source_root"])
        paths = data.get("paths", {})
        package_paths = data.get("package_paths", {})
        return cls(
            id=data["id"],
            name=data.get("name", data["id"]),
            root=package_root,
            engine=data.get("engine", "legacy-generator"),
            source_root=source_root,
            generator=_resolve_path(source_root, paths.get("generator", "模板抽取输出/generate_print_pdf.py")),
            template_json=_resolve_path(source_root, paths.get("template_json", "模板抽取输出/template.json")),
            font_map=_resolve_path(source_root, paths.get("font_map", "模板抽取输出/font-map.json")),
            svg_dir=_resolve_path(source_root, paths.get("svg_dir", "Document fonts/SVG")),
            layout_rules=_resolve_optional(package_root, package_paths.get("layout_rules")),
            asset_map=_resolve_optional(package_root, package_paths.get("asset_map")),
        )


def _resolve_optional(base: Path, value: str | None) -> Path | None:
    if not value:
        return None
    return _resolve_path(base, value)
