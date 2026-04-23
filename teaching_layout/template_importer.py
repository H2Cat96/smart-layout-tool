from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any


TEMPLATE_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]*$")
ASSET_SUBDIRS = ("svg", "images", "fonts", "references")
REFERENCE_EXTENSIONS = {".pdf", ".idml"}
FONT_EXTENSIONS = {".ttf", ".otf", ".ttc"}


def import_template(
    template_id: str,
    name: str,
    source_dir: Path,
    templates_dir: Path,
    overwrite: bool = False,
) -> dict[str, Any]:
    source_dir = Path(source_dir)
    templates_dir = Path(templates_dir)
    package_root = templates_dir / template_id

    if not TEMPLATE_ID_PATTERN.match(template_id):
        raise ValueError(f"Invalid template id: {template_id}")
    if not source_dir.exists():
        raise FileNotFoundError(f"Template source missing: {source_dir}")
    if package_root.exists() and not overwrite:
        raise FileExistsError(f"Template package already exists: {package_root}")
    if package_root.exists():
        shutil.rmtree(package_root)

    _copy_asset_tree(source_dir, package_root)
    extracted_dir = package_root / "extracted"
    engines_dir = package_root / "engines"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    engines_dir.mkdir(parents=True, exist_ok=True)

    references = _find_files(package_root / "assets" / "references")
    has_layout_reference = any(path.suffix.lower() in REFERENCE_EXTENSIONS for path in references)
    status = "draft_template_ready" if has_layout_reference else "needs_layout_reference"
    missing = [] if has_layout_reference else ["assets/references/pdf_or_idml"]

    _write_json(package_root / "template-package.json", _package_json(template_id, name))
    _write_json(package_root / "asset-map.json", _asset_map(package_root))
    _write_json(package_root / "layout-rules.json", _layout_rules(status, missing))
    _write_json(extracted_dir / "font-map.json", _font_map(package_root))
    _write_json(extracted_dir / "template.json", _template_data(status))
    _write_json(extracted_dir / "style-map.json", _style_map(status))
    _write_json(extracted_dir / "geometry-map.json", _geometry_map(status))
    _write_render_stub(engines_dir / "render.py")
    _write_readme(package_root / "README.md", template_id, name, status, missing)

    return {
        "schema_version": "0.1",
        "template_id": template_id,
        "name": name,
        "status": status,
        "missing": missing,
        "package_path": str(package_root),
        "copied_assets": {
            subdir: len(_find_files(package_root / "assets" / subdir))
            for subdir in ASSET_SUBDIRS
        },
    }


def _copy_asset_tree(source_dir: Path, package_root: Path) -> None:
    for subdir in ASSET_SUBDIRS:
        src = source_dir / "assets" / subdir
        dst = package_root / "assets" / subdir
        if src.exists():
            shutil.copytree(src, dst)
        else:
            dst.mkdir(parents=True, exist_ok=True)


def _find_files(root: Path) -> list[Path]:
    if not root.exists():
        return []
    return sorted(path for path in root.rglob("*") if path.is_file())


def _relative(path: Path, base: Path) -> str:
    return path.relative_to(base).as_posix()


def _write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _package_json(template_id: str, name: str) -> dict[str, Any]:
    return {
        "schema_version": "0.1",
        "id": template_id,
        "name": name,
        "description": "Imported template package skeleton. Layout rules are template-defined.",
        "engine": {
            "type": "template-skeleton",
            "entry": "engines/render.py",
        },
        "capabilities": {
            "input": ["docx", "yach-doc"],
            "output": ["pdf", "preview-png", "manifest"],
            "page_model": "template-defined",
        },
        "paths": {
            "layout_rules": "layout-rules.json",
            "asset_map": "asset-map.json",
            "template_data": "extracted/template.json",
            "font_map": "extracted/font-map.json",
            "style_map": "extracted/style-map.json",
            "geometry_map": "extracted/geometry-map.json",
            "svg_dir": "assets/svg",
            "image_dir": "assets/images",
            "font_dir": "assets/fonts",
            "reference_dir": "assets/references",
        },
        "runtime": {
            "requires_fonts": True,
            "requires_network": False,
        },
    }


def _asset_map(package_root: Path) -> dict[str, Any]:
    assets: dict[str, dict[str, str]] = {}
    for svg in _find_files(package_root / "assets" / "svg"):
        key = svg.stem
        assets[key] = {
            "file": svg.name,
            "source": _relative(svg, package_root),
            "role": "template-defined",
        }
    for image in _find_files(package_root / "assets" / "images"):
        key = image.stem
        assets.setdefault(
            key,
            {
                "file": image.name,
                "source": _relative(image, package_root),
                "role": "template-defined",
            },
        )
    return {
        "schema_version": "0.1",
        "description": "Generated asset inventory. Roles need template-specific review.",
        "assets": assets,
    }


def _font_map(package_root: Path) -> dict[str, Any]:
    fonts: dict[str, dict[str, str]] = {}
    for font in _find_files(package_root / "assets" / "fonts"):
        if font.suffix.lower() not in FONT_EXTENSIONS:
            continue
        key = font.stem
        fonts[key] = {
            "path": f"../assets/fonts/{font.name}",
            "reportlab_name": key,
            "status": "imported",
        }
    return {
        "schema_version": "0.1",
        "source_folder": "../assets/fonts",
        "fonts": fonts,
        "style_defaults": {},
        "notes": [
            "Generated from imported font files.",
            "Review reportlab_name and style_defaults before production rendering.",
        ],
    }


def _layout_rules(status: str, missing: list[str]) -> dict[str, Any]:
    return {
        "schema_version": "0.1",
        "status": status,
        "missing": missing,
        "description": "Template-defined layout rules skeleton. Fill after PDF/IDML/reference analysis.",
        "page_flow": {
            "mode": "template-defined",
        },
        "styles": {},
        "geometry": {},
    }


def _template_data(status: str) -> dict[str, Any]:
    return {
        "schema_version": "0.1",
        "status": status,
        "spreads": [],
        "flow_sequences": [],
    }


def _style_map(status: str) -> dict[str, Any]:
    return {
        "schema_version": "0.1",
        "status": status,
        "styles": {},
    }


def _geometry_map(status: str) -> dict[str, Any]:
    return {
        "schema_version": "0.1",
        "status": status,
        "regions": {},
    }


def _write_render_stub(path: Path) -> None:
    path.write_text(
        "\n".join(
            [
                "from __future__ import annotations",
                "",
                "",
                "def build_pdf(*args, **kwargs):",
                "    raise NotImplementedError(",
                "        \"This imported template package needs a renderer after layout calibration.\"",
                "    )",
                "",
            ]
        ),
        encoding="utf-8",
    )


def _write_readme(path: Path, template_id: str, name: str, status: str, missing: list[str]) -> None:
    missing_text = ", ".join(missing) if missing else "none"
    path.write_text(
        "\n".join(
            [
                f"# {name}",
                "",
                f"- Template id: `{template_id}`",
                f"- Import status: `{status}`",
                f"- Missing: `{missing_text}`",
                "",
                "This package was generated from user-provided raw template assets.",
                "Review layout rules, asset roles, font mapping, and renderer behavior before production use.",
                "",
            ]
        ),
        encoding="utf-8",
    )
