from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .legacy_adapter import resolve_template_paths
from .template_package import DEFAULT_TEMPLATES_DIR


INSPECTED_JSON_FILES = (
    "template-package.json",
    "extracted/font-map.json",
    "extracted/template.json",
    "layout-rules.json",
    "asset-map.json",
)


def inspect_template(template_ref: str | Path, templates_dir: Path = DEFAULT_TEMPLATES_DIR) -> dict[str, Any]:
    paths = resolve_template_paths(template_ref, templates_dir=templates_dir)
    package_root = _find_package_root(paths.root)
    font_map = _read_json(paths.font_map)

    missing: list[str] = []
    fonts = _inspect_fonts(font_map, paths.font_map.parent, missing)
    required_files = _inspect_required_files(paths, missing)
    absolute_paths = _find_absolute_paths(package_root)

    status = "portable" if not missing and not absolute_paths else "not_portable"
    return {
        "template": str(template_ref),
        "status": status,
        "package_path": str(package_root),
        "missing": missing,
        "absolute_paths": absolute_paths,
        "required_files": required_files,
        "fonts": fonts,
    }


def _find_package_root(root: Path) -> Path:
    current = root.resolve()
    for candidate in (current, *current.parents):
        if (candidate / "template-package.json").exists():
            return candidate
    return root


def _read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def _inspect_required_files(paths: Any, missing: list[str]) -> list[dict[str, Any]]:
    required = {
        "generator": paths.generator,
        "template_json": paths.template_json,
        "font_map": paths.font_map,
        "svg_dir": paths.svg_dir,
    }
    records = []
    for name, path in required.items():
        exists = path.exists()
        if not exists:
            missing.append(name)
        records.append({"name": name, "path": str(path), "exists": exists})
    for name, path in {"layout_rules": paths.layout_rules, "asset_map": paths.asset_map}.items():
        if path is None:
            continue
        exists = path.exists()
        if not exists:
            missing.append(name)
        records.append({"name": name, "path": str(path), "exists": exists})
    return records


def _inspect_fonts(font_map: dict[str, Any], base_dir: Path, missing: list[str]) -> list[dict[str, Any]]:
    records = []
    for source_font, item in font_map.get("fonts", {}).items():
        raw_path = item.get("path", "")
        resolved = _resolve_json_path(raw_path, base_dir)
        exists = resolved.exists()
        if not exists:
            missing.append(f"font:{source_font}")
        records.append(
            {
                "source_font": source_font,
                "reportlab_name": item.get("reportlab_name", ""),
                "path": raw_path,
                "resolved_path": str(resolved),
                "exists": exists,
                "is_absolute": Path(raw_path).expanduser().is_absolute(),
            }
        )
    return records


def _resolve_json_path(value: str, base_dir: Path) -> Path:
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    return (base_dir / path).resolve()


def _find_absolute_paths(package_root: Path) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for relative in INSPECTED_JSON_FILES:
        path = package_root / relative
        if not path.exists():
            continue
        data = _read_json(path)
        _collect_absolute_values(data, relative, "", records)
    return records


def _collect_absolute_values(value: Any, file_name: str, json_path: str, records: list[dict[str, str]]) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{json_path}.{key}" if json_path else str(key)
            _collect_absolute_values(child, file_name, child_path, records)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            child_path = f"{json_path}.{index}" if json_path else str(index)
            _collect_absolute_values(child, file_name, child_path, records)
    elif isinstance(value, str):
        expanded = Path(value).expanduser()
        if expanded.is_absolute():
            records.append({"file": file_name, "json_path": json_path, "value": value})
