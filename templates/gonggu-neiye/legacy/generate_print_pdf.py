#!/usr/bin/env python3
"""
Generate a print-oriented teaching-aid PDF from:
  - template.json extracted from IDML
  - font-map.json
  - a source DOCX
  - optional background PNGs named spread_02.png, spread_03.png, ...

The image model should only create background images. Body text, page numbers,
side strips, and functional labels are always drawn programmatically.
"""

from __future__ import annotations

import argparse
import json
import random
import re
from io import BytesIO
from collections import OrderedDict
from pathlib import Path
from typing import Any

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image
from pypdf import PageObject, PdfReader, PdfWriter, Transformation
from reportlab.lib.colors import CMYKColor, Color
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


PACKAGE_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = PACKAGE_ROOT / "extracted" / "template.json"
DEFAULT_FONT_MAP = PACKAGE_ROOT / "extracted" / "font-map.json"
DEFAULT_DOCX = PACKAGE_ROOT / "samples" / "input.docx"
DEFAULT_OUTPUT_DIR = PACKAGE_ROOT / "outputs"
DEFAULT_SVG_DIR = PACKAGE_ROOT / "assets" / "svg"
ANSWER_LINE_SENTINEL = "\uE000_ANSWER_LINE"
FORBIDDEN_LINE_START = set("，。！？；：、,.!?;:)]}）】》〉」』”’％‰℃…")
FORBIDDEN_LINE_END = set("([{（【《〈「『“‘")
STYLE_OVERRIDE_KEYS = {
    "font",
    "size",
    "leading",
    "align",
    "color",
    "space_before",
    "space_after",
    "first_line_indent",
    "left_indent",
}
DEFAULT_CONTENT_DETECTION = {
    "section_title_patterns": [r"^【.+】$", r"^练习[一二三四五六七八九十]+$"],
    "section_title_exclude": ["【答案】", "【解析】"],
    "lesson_title_patterns": [r"^第[一二三四五六七八九十百]+讲$"],
    "topic_heading_patterns": [r"^[一二三四五六七八九十]+、\S+"],
    "reading_prompt_patterns": [r"^阅读.*(题|完成|回答|问题)"],
    "structural_reading_prompt_exclude_patterns": [
        r"^[1-9]\d*[.．、]",
        r"^[A-E][.．]",
        r"^（\d+）",
    ],
    "structural_reading_prompt_keywords": ["阅读", "选文", "短文", "文章", "完成", "回答"],
    "article_title": {
        "max_length": 12,
        "reject_prefixes": ["【", "（", "(", "“", "《"],
        "reject_patterns": [r"^[A-E][.．]", r"^[1-9]\d*[.．、]"],
        "forbidden_punctuation_pattern": r"[，。！？；：、,.!?;:]",
    },
    "author_names": ["贾平凹"],
    "source_patterns": [r"^（.*）$"],
}
DEFAULT_MARKERS = {
    "question_patterns": [r"^([1-9]\d*)[.．、]\s*(.*)$"],
    "option_patterns": [r"^([A-E])[.．]\s*(.*)$"],
    "parenthesized_option_patterns": [r"^(（\d+）)\s*(.*)$"],
    "judgement_patterns": [r"^（\d+）"],
}
DEFAULT_HANGING = {
    "start_fields": ["marker_text", "badge_text"],
    "end_kinds": [
        "main_title",
        "topic_heading",
        "reading_prompt",
        "article_title",
        "author",
        "section",
        "source",
        "answer_label",
    ],
    "inherit_kinds": ["body", "answer_body", "answer_line", "option", "question_numbered", "table", "image"],
    "clear_first_line_indent_kinds": ["body", "answer_body", "answer_line"],
}
DEFAULT_NORMALIZATION = {
    "answer_labels": {
        "【答案】": "【答案】",
        "答案：": "【答案】",
        "答案:": "【答案】",
        "【解析】": "【解析】",
        "解析：": "【解析】",
        "解析:": "【解析】",
    }
}
DEFAULT_KEEP_WITH_NEXT_KINDS = [
    "main_title",
    "section",
    "topic_heading",
    "article_title",
    "answer_label",
    "source",
    "question_numbered",
]


def load_optional_json(path: str | Path | None) -> dict[str, Any]:
    if not path:
        return {}
    config_path = Path(path)
    if not config_path.exists():
        return {}
    return json.loads(config_path.read_text(encoding="utf-8"))


def config_color(layout_rules: dict[str, Any], key: str, default: str) -> str:
    return layout_rules.get("colors", {}).get(key, default)


def merged_rule_section(layout_rules: dict[str, Any] | None, key: str, defaults: dict[str, Any]) -> dict[str, Any]:
    merged = dict(defaults)
    override = (layout_rules or {}).get(key, {})
    for item_key, item_value in override.items():
        if isinstance(item_value, dict) and isinstance(merged.get(item_key), dict):
            nested = dict(merged[item_key])
            nested.update(item_value)
            merged[item_key] = nested
        else:
            merged[item_key] = item_value
    return merged


def first_regex_match(patterns: list[str], text: str) -> re.Match[str] | None:
    for pattern in patterns:
        match = re.match(pattern, text)
        if match:
            return match
    return None


def matches_any(patterns: list[str], text: str) -> bool:
    return first_regex_match(patterns, text) is not None


def normalize_align(value: str) -> str:
    return "justify" if value == "justify-last-left" else value


def apply_style_rule(style: dict[str, Any], layout_rules: dict[str, Any], rule_name: str) -> dict[str, Any]:
    rule = layout_rules.get("styles", {}).get(rule_name, {})
    for key in STYLE_OVERRIDE_KEYS:
        if key in rule:
            style[key] = normalize_align(rule[key]) if key == "align" else rule[key]
    if "bar_color" in rule:
        style["bar_color"] = rule["bar_color"]
    return style


def question_content_indent(layout_rules: dict[str, Any], fallback: float = 28) -> float:
    rule = layout_rules.get("styles", {}).get("question_content", {})
    stem_rule = layout_rules.get("styles", {}).get("question_stem", {})
    return rule.get("left_indent", stem_rule.get("left_indent", fallback))


def marker_indent(marker_text: str, size: float, min_indent: float = 22, gap: float = 4) -> float:
    visual_width = 0.0
    for char in marker_text:
        visual_width += size * 0.5 if ord(char) < 128 else size
    return max(min_indent, visual_width + gap)


def starts_hanging_context(style: dict[str, Any], layout_rules: dict[str, Any] | None = None) -> bool:
    hanging = merged_rule_section(layout_rules, "hanging", DEFAULT_HANGING)
    return any(bool(style.get(field)) for field in hanging["start_fields"])


def ends_hanging_context(style: dict[str, Any], layout_rules: dict[str, Any] | None = None) -> bool:
    hanging = merged_rule_section(layout_rules, "hanging", DEFAULT_HANGING)
    return style.get("kind") in set(hanging["end_kinds"])


def can_inherit_hanging_context(style: dict[str, Any], layout_rules: dict[str, Any] | None = None) -> bool:
    hanging = merged_rule_section(layout_rules, "hanging", DEFAULT_HANGING)
    return style.get("kind") in set(hanging["inherit_kinds"])


def apply_hanging_context_indent(
    style: dict[str, Any],
    inherited_indent: float | None,
    layout_rules: dict[str, Any],
) -> dict[str, Any]:
    hanging = merged_rule_section(layout_rules, "hanging", DEFAULT_HANGING)
    if inherited_indent is None or not can_inherit_hanging_context(style, layout_rules):
        return style
    next_style = dict(style)
    next_style["left_indent"] = inherited_indent
    if next_style.get("kind") in set(hanging["clear_first_line_indent_kinds"]):
        next_style["first_line_indent"] = 0
    return next_style


def starts_new_question_context(style: dict[str, Any]) -> bool:
    return starts_hanging_context(style)


def ends_question_context(style: dict[str, Any]) -> bool:
    return ends_hanging_context(style)


def can_inherit_question_context(style: dict[str, Any]) -> bool:
    return can_inherit_hanging_context(style)


def apply_question_content_indent(
    style: dict[str, Any],
    inherited_indent: float | None,
    layout_rules: dict[str, Any],
) -> dict[str, Any]:
    return apply_hanging_context_indent(style, inherited_indent, layout_rules)


def frame_side(frame: dict[str, float], page_w: float) -> str:
    return "right" if frame["x"] + frame["w"] / 2 >= page_w / 2 else "left"


def make_remainder_block(rest: str, style: dict[str, Any]) -> dict[str, Any]:
    next_style = dict(style)
    next_style.pop("badge_text", None)
    next_style.pop("marker_text", None)
    next_style["display_text"] = rest
    next_style["first_line_indent"] = 0
    return {
        "type": "remainder",
        "text": rest,
        "style": next_style,
    }


def is_reading_prompt(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    return matches_any(detection["reading_prompt_patterns"], text)


def is_section_title(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    if text in set(detection["section_title_exclude"]):
        return False
    return matches_any(detection["section_title_patterns"], text)


def is_lesson_title(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    return matches_any(detection["lesson_title_patterns"], text)


def is_topic_heading(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    return matches_any(detection["topic_heading_patterns"], text)


def can_be_structural_reading_prompt(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    if not text:
        return False
    if matches_any(detection["structural_reading_prompt_exclude_patterns"], text):
        return False
    if is_section_title(text, layout_rules):
        return False
    return is_reading_prompt(text, layout_rules) or any(
        keyword in text for keyword in detection["structural_reading_prompt_keywords"]
    )


def section_display_text(text: str) -> str:
    return text if text.startswith("【") and text.endswith("】") else f"【{text}】"


def is_article_title_text(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    rule = detection["article_title"]
    if not text or len(text) > rule["max_length"]:
        return False
    if text.startswith(tuple(rule["reject_prefixes"])):
        return False
    if matches_any(rule["reject_patterns"], text):
        return False
    return not re.search(rule["forbidden_punctuation_pattern"], text)


def normalize_answer_label(text: str, layout_rules: dict[str, Any] | None = None) -> str | None:
    normalization = merged_rule_section(layout_rules, "normalization", DEFAULT_NORMALIZATION)
    return normalization["answer_labels"].get(text)


def clean_text(text: str) -> str:
    return re.sub(r"[\r\n\t]+", " ", text).strip()


def normalize_answer_parentheses(text: str) -> str:
    return re.sub(
        r"[ \u00a0]*([（(])[ \u00a0]*([）)])\s*$",
        lambda match: f" {match.group(1)}{chr(0x2007) * 7}{match.group(2)}",
        text,
    )


def is_underlined_blank_paragraph(paragraph: Any) -> bool:
    text = "".join(run.text for run in paragraph.runs)
    if text.strip():
        return False
    return any("<w:u" in run._element.xml for run in paragraph.runs)


def paragraph_underline_ranges(paragraph: Any) -> list[list[int]]:
    ranges: list[list[int]] = []
    cursor = 0
    for run in paragraph.runs:
        text = re.sub(r"[\r\n\t]+", " ", run.text)
        if not text:
            continue
        start = cursor
        cursor += len(text)
        if text.strip() and "<w:u" in run._element.xml:
            ranges.append([start, cursor])
    return ranges


def clean_cell_text(text: str) -> str:
    return re.sub(r"[\r\n\t]+", " ", text).strip()


def block_text(block: Any) -> str:
    if isinstance(block, dict):
        return block.get("text", "")
    return str(block)


def is_remainder_block(block: Any) -> bool:
    return isinstance(block, dict) and block.get("type") == "remainder"


def remainder_style(block: dict[str, Any]) -> dict[str, Any]:
    return dict(block.get("style", {}))


def is_table_block(block: Any) -> bool:
    return isinstance(block, dict) and block.get("type") == "table"


def is_image_block(block: Any) -> bool:
    return isinstance(block, dict) and block.get("type") == "image"


def table_rows(block: dict[str, Any]) -> list[list[str]]:
    return block.get("rows", [])


def paragraph_image_blocks(paragraph: Paragraph) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    rel_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    for node in paragraph._p.iter():
        if not node.tag.endswith("}blip"):
            continue
        rel_id = node.get(rel_attr)
        if not rel_id:
            continue
        part = paragraph.part.related_parts.get(rel_id)
        if part is None:
            continue
        blob = part.blob
        try:
            image = Image.open(BytesIO(blob))
            width_px, height_px = image.size
        except Exception:
            width_px, height_px = 0, 0
        blocks.append({
            "type": "image",
            "blob": blob,
            "content_type": getattr(part, "content_type", ""),
            "width_px": width_px,
            "height_px": height_px,
        })
    return blocks


def normalize_bare_answer_items(blocks: list[Any]) -> list[Any]:
    normalized: list[Any] = []
    in_answer_block = False
    auto_number = 1
    auto_numbering_active = False
    for block in blocks:
        text = block_text(block)
        if text == "【答案】":
            in_answer_block = True
            auto_number = 1
            auto_numbering_active = True
            normalized.append(block)
            continue
        if text == "【解析】":
            in_answer_block = False
            auto_numbering_active = False
            normalized.append(block)
            continue
        if in_answer_block and re.match(r"^[1-9]\d*[.．、]", text):
            auto_numbering_active = False
            normalized.append(block)
            continue
        if in_answer_block and auto_numbering_active and isinstance(block, str) and text:
            normalized.append(f"{auto_number}.{text}")
            auto_number += 1
            continue
        normalized.append(block)
    return normalized


def load_docx_paragraphs(docx_path: Path) -> list[Any]:
    doc = Document(docx_path)
    blocks: list[Any] = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            p = Paragraph(child, doc)
            if is_underlined_blank_paragraph(p):
                blocks.append(ANSWER_LINE_SENTINEL)
                continue
            image_blocks = paragraph_image_blocks(p)
            text = clean_text(p.text)
            if not text:
                blocks.extend(image_blocks)
                continue
            text = normalize_answer_parentheses(text)
            underline_ranges = paragraph_underline_ranges(p)
            if underline_ranges:
                blocks.append({"type": "paragraph", "text": text, "underline_ranges": underline_ranges})
            else:
                blocks.append(text)
            blocks.extend(image_blocks)
        elif child.tag.endswith("}tbl"):
            table = Table(child, doc)
            rows = [
                [clean_cell_text(cell.text) for cell in row.cells]
                for row in table.rows
            ]
            if rows:
                blocks.append({"type": "table", "rows": rows})
    return normalize_bare_answer_items(blocks)


def split_practice_and_answers(paragraphs: list[Any]) -> tuple[list[Any], list[Any]]:
    split_idx = next(
        (i for i, text in enumerate(paragraphs) if "参考答案" in block_text(text)),
        len(paragraphs),
    )
    return paragraphs[:split_idx], paragraphs[split_idx:]


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def resolve_font_path(value: str, base_dir: Path | None = None) -> Path:
    font_path = Path(value).expanduser()
    if font_path.is_absolute() or base_dir is None:
        return font_path
    return (base_dir / font_path).resolve()


def register_fonts(font_map: dict[str, Any], base_dir: Path | None = None) -> tuple[set[str], list[dict[str, str]]]:
    registered: set[str] = set()
    substitutions: list[dict[str, str]] = []
    for source_font, item in font_map["fonts"].items():
        reportlab_name = item["reportlab_name"]
        if item.get("status") == "substitute":
            substitutions.append(
                {
                    "source_font": source_font,
                    "substitute_for": item.get("substitute_for", ""),
                    "font_file": item["path"],
                    "reportlab_name": reportlab_name,
                }
            )
        if reportlab_name in registered:
            continue
        font_path = resolve_font_path(item["path"], base_dir)
        if not font_path.exists():
            raise FileNotFoundError(f"Font file missing: {font_path}")
        pdfmetrics.registerFont(TTFont(reportlab_name, str(font_path)))
        registered.add(reportlab_name)
    return registered, substitutions


def spread_by_index(template: dict[str, Any], spread_index: int) -> dict[str, Any]:
    return next(s for s in template["spreads"] if s["spread_index"] == spread_index)


def story_frames(template: dict[str, Any], role: str) -> list[dict[str, Any]]:
    return next((seq["frames"] for seq in template["flow_sequences"] if seq["role"] == role), [])


def normalize_flow_frames(
    frames: list[dict[str, Any]],
    role: str,
    layout_rules: dict[str, Any],
) -> list[dict[str, Any]]:
    min_heights = layout_rules.get("page_flow", {}).get("min_frame_height_pt", {})
    min_height = min_heights.get(role)
    if not min_height:
        return frames
    normalized: list[dict[str, Any]] = []
    for frame in frames:
        next_frame = dict(frame)
        next_bbox = dict(frame["bbox_pt"])
        next_bbox["h"] = max(next_bbox["h"], min_height)
        next_frame["bbox_pt"] = next_bbox
        normalized.append(next_frame)
    return normalized


def group_frames_by_spread(frames: list[dict[str, Any]]) -> OrderedDict[int, list[dict[str, Any]]]:
    grouped: OrderedDict[int, list[dict[str, Any]]] = OrderedDict()
    for frame in frames:
        grouped.setdefault(frame["spread_index"], []).append(frame)
    for spread_index in grouped:
        grouped[spread_index].sort(key=lambda f: (f["bbox_pt"]["x"], f["bbox_pt"]["y"]))
    return grouped


def _rgb_from_hex(hex_value: str | None, fallback: tuple[float, float, float]) -> tuple[float, float, float]:
    try:
        if not hex_value:
            return fallback
        value = hex_value.strip().lstrip("#")
        if len(value) != 6:
            return fallback
        return (
            int(value[0:2], 16) / 255,
            int(value[2:4], 16) / 255,
            int(value[4:6], 16) / 255,
        )
    except Exception:
        return fallback


def _parse_cmyk(value: str | None) -> tuple[float, float, float, float] | None:
    if not value:
        return None
    match = re.match(r"^cmyk\(\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)\s*\)$", value.strip(), re.I)
    if not match:
        return None
    numbers = [float(part) for part in match.groups()]
    if any(number > 1 for number in numbers):
        numbers = [number / 100 for number in numbers]
    return tuple(max(0, min(1, number)) for number in numbers)  # type: ignore[return-value]


def _cmyk_from_rgb(r: float, g: float, b: float) -> CMYKColor:
    if abs(r - g) < 1e-6 and abs(g - b) < 1e-6:
        return CMYKColor(0, 0, 0, max(0, min(1, 1 - r)))
    k = 1 - max(r, g, b)
    if k >= 1:
        return CMYKColor(0, 0, 0, 1)
    c = (1 - r - k) / (1 - k)
    m = (1 - g - k) / (1 - k)
    y = (1 - b - k) / (1 - k)
    return CMYKColor(c, m, y, k)


def color_from_hex(
    hex_value: str | None,
    fallback: tuple[float, float, float],
    color_mode: str = "rgb",
) -> Color:
    cmyk = _parse_cmyk(hex_value)
    if cmyk:
        c, m, y, k = cmyk
        if color_mode == "cmyk":
            return CMYKColor(c, m, y, k)
        return Color((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    r, g, b = _rgb_from_hex(hex_value, fallback)
    if color_mode == "cmyk":
        return _cmyk_from_rgb(r, g, b)
    return Color(r, g, b)


def resolve_svg_asset_path(svg_dir: Path, file_name: str) -> Path:
    path = Path(file_name)
    return path if path.is_absolute() else svg_dir / path


def mapped_asset(
    svg_dir: Path,
    asset_map: dict[str, Any],
    asset_key: str,
    default_file: str,
    default_color_key: str = "color",
) -> tuple[Path, str | None]:
    item = asset_map.get("assets", {}).get(asset_key, {})
    path = resolve_svg_asset_path(svg_dir, item.get("file", default_file))
    color = item.get(default_color_key) or item.get("color")
    return path, color


def load_svg_assets(svg_dir: Path, asset_map: dict[str, Any] | None = None) -> dict[str, dict[str, Any]]:
    assets: dict[str, dict[str, Any]] = {}
    asset_map = asset_map or {}
    answer_corner_path, answer_color = mapped_asset(svg_dir, asset_map, "answer_title_corner", "参考答案.svg")
    if not answer_corner_path.exists():
        answer_corner_path = svg_dir.parent / "参考答案.svg"
    section_bar_path, section_bar_color = mapped_asset(
        svg_dir,
        asset_map,
        "section_bar",
        "标题.svg",
        default_color_key="practice_color",
    )
    title_corner_path, title_corner_color = mapped_asset(svg_dir, asset_map, "practice_title_corner", "标题角标.svg")
    badge_path, badge_color = mapped_asset(svg_dir, asset_map, "question_badge", "序号.svg")
    asset_paths = {
        "标题": (section_bar_path, section_bar_color),
        "标题角标": (title_corner_path, title_corner_color),
        "序号": (badge_path, badge_color),
        "参考答案": (answer_corner_path, answer_color),
    }
    for name, (path, color_override) in asset_paths.items():
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8")
        viewbox_match = re.search(r'viewBox="([^"]+)"', text)
        fill_match = re.search(r"#[0-9A-Fa-f]{6}", text)
        viewbox = [float(v) for v in viewbox_match.group(1).split()] if viewbox_match else [0, 0, 0, 0]
        assets[name] = {
            "path": str(path),
            "viewbox": viewbox,
            "width": viewbox[2],
            "height": viewbox[3],
            "fill": color_override or (fill_match.group(0) if fill_match else "#000000"),
        }
    return assets


def draw_title_corner_svg(
    c: canvas.Canvas,
    x: float,
    y_top: float,
    assets: dict[str, dict[str, Any]],
    asset_name: str = "标题角标",
    color_mode: str = "rgb",
) -> None:
    asset = assets.get(asset_name)
    if not asset:
        return
    h = asset["height"]
    page_h = c._pagesize[1]
    c.saveState()
    c.translate(x, page_h - y_top - h)
    c.setFillColor(color_from_hex(asset["fill"], (0, 0, 0), color_mode))
    p = c.beginPath()
    if asset_name == "参考答案":
        p.moveTo(5.63, 0)
        p.curveTo(2.52, 0, 0, 2.52, 0, 5.63)
        p.lineTo(0, 24.41)
        p.curveTo(0, 27.52, 2.52, 30.05, 5.63, 30.05)
        p.lineTo(30.04, 30.05)
        p.lineTo(30.04, 26.29)
        p.lineTo(11.27, 26.29)
        p.lineTo(11.27, 22.53)
        p.lineTo(30.05, 22.53)
        p.lineTo(30.05, 0)
        p.close()
    else:
        p.moveTo(20.66, 0)
        p.lineTo(30.05, 9.39)
        p.lineTo(26.29, 13.15)
        p.lineTo(26.29, 22.54)
        p.lineTo(5.63, 30.05)
        p.lineTo(4.14, 28.56)
        p.lineTo(12.17, 20.53)
        p.curveTo(12.48, 20.61, 12.81, 20.66, 13.14, 20.66)
        p.curveTo(15.21, 20.66, 16.90, 18.98, 16.90, 16.91)
        p.curveTo(16.90, 14.84, 15.22, 13.15, 13.14, 13.15)
        p.curveTo(11.06, 13.15, 9.38, 14.83, 9.38, 16.91)
        p.curveTo(9.38, 17.25, 9.43, 17.57, 9.51, 17.88)
        p.lineTo(1.49, 25.90)
        p.lineTo(0, 24.41)
        p.lineTo(7.51, 3.76)
        p.lineTo(16.90, 3.76)
        p.close()
    c.drawPath(p, fill=1, stroke=0)
    c.restoreState()


def draw_title_bar_svg(
    c: canvas.Canvas,
    x: float,
    y_top: float,
    width: float,
    assets: dict[str, dict[str, Any]],
    fallback_color: str,
    color_mode: str = "rgb",
) -> float:
    asset = assets.get("标题")
    height = asset["height"] if asset else 28.35
    fill = fallback_color
    page_h = c._pagesize[1]
    c.setFillColor(color_from_hex(fill, (.95, .95, .95), color_mode))
    c.roundRect(x, page_h - y_top - height, width, height, height / 2, fill=1, stroke=0)
    return height


def draw_number_badge_svg(
    c: canvas.Canvas,
    x: float,
    y_top: float,
    text: str,
    font: str,
    assets: dict[str, dict[str, Any]],
    color_mode: str = "rgb",
) -> None:
    asset = assets.get("序号")
    width = asset["width"] if asset else 22.68
    height = asset["height"] if asset else 12.76
    fill = asset["fill"] if asset else "#c1020e"
    page_h = c._pagesize[1]
    c.setFillColor(color_from_hex(fill, (.76, 0, .05), color_mode))
    c.roundRect(x, page_h - y_top - height, width, height, 2.84, fill=1, stroke=0)
    c.setFillColor(color_from_hex("#ffffff", (1, 1, 1), color_mode))
    c.setFont(font, 8.5)
    c.drawCentredString(x + width / 2, page_h - y_top - height + 3.2, text)


def make_placeholder_background(
    out_path: Path,
    page_w_pt: float,
    page_h_pt: float,
    dpi: int,
    tone: str,
) -> None:
    """Create a no-text, no-illustration paper background."""
    width = round(page_w_pt / 72 * dpi)
    height = round(page_h_pt / 72 * dpi)
    base_rgb = (255, 254, 251) if tone == "practice" else (253, 253, 252)
    rng = random.Random(f"{out_path.name}-{tone}")
    img = Image.new("RGB", (width, height), base_rgb)
    pixels = img.load()
    for y in range(height):
        # Very light vertical paper variation; no center seam.
        row_delta = rng.randint(-1, 1)
        for x in range(width):
            if rng.random() < 0.018:
                delta = rng.choice([-3, -2, 2, 3])
            else:
                delta = row_delta
            r = max(0, min(255, base_rgb[0] + delta))
            g = max(0, min(255, base_rgb[1] + delta))
            b = max(0, min(255, base_rgb[2] + delta))
            pixels[x, y] = (r, g, b)
    ensure_dir(out_path.parent)
    img.save(out_path, quality=95)


def background_path_for_spread(background_dir: Path, spread_index: int) -> Path:
    return background_dir / f"spread_{spread_index:02d}.png"


def ensure_backgrounds(
    template: dict[str, Any],
    background_dir: Path,
    page_w: float,
    page_h: float,
    dpi: int,
) -> list[dict[str, Any]]:
    ensure_dir(background_dir)
    records: list[dict[str, Any]] = []
    for spread in template["spreads"]:
        if spread["page_count"] != 2:
            continue
        spread_index = spread["spread_index"]
        path = background_path_for_spread(background_dir, spread_index)
        generated = False
        if not path.exists():
            tone = "answer" if spread_index >= 4 else "practice"
            make_placeholder_background(path, page_w, page_h, dpi, tone=tone)
            generated = True
        records.append(
            {
                "spread_index": spread_index,
                "file": str(path),
                "generated_placeholder": generated,
            }
        )
    return records


def write_background_prompt_manifest(output_dir: Path, template: dict[str, Any]) -> Path:
    safe_frames = template["gpt_image2_background_contract"]["safe_text_frames"]
    prompts = []
    for spread in template["spreads"]:
        if spread["page_count"] != 2:
            continue
        spread_index = spread["spread_index"]
        tone = "practice pages with pale pink outer side strips" if spread_index < 4 else "answer pages with neutral gray outer side strips"
        prompt = (
            "Landscape double-page Chinese literature workbook background, "
            "background only, no text, no Chinese characters, no illustrations, "
            "no center seam, no book gutter shadow, no page fold, no decorative icons, "
            "subtle warm paper texture, clean print-ready surface, "
            f"{tone}. Keep all main text areas visually quiet and blank."
        )
        prompts.append(
            {
                "spread_index": spread_index,
                "target_file_name": f"spread_{spread_index:02d}.png",
                "prompt": prompt,
                "safe_text_frames": safe_frames,
            }
        )
    path = output_dir / "background-prompts.json"
    path.write_text(json.dumps({"model": "gpt-image2", "prompts": prompts}, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def paragraph_style(
    text: Any,
    pos: int,
    is_answer: bool,
    fonts: dict[str, str],
    layout_rules: dict[str, Any] | None = None,
    force_reading_prompt: bool = False,
) -> dict[str, Any]:
    layout_rules = layout_rules or {}
    underline_ranges = text.get("underline_ranges", []) if isinstance(text, dict) else []
    text = block_text(text)
    yan_mid = fonts.get("title_mid", fonts["title"])
    yan_bold = fonts.get("title_bold", fonts["title"])
    yan_regular = fonts.get("question", fonts["title"])
    kai = fonts["body"]
    markers = merged_rule_section(layout_rules, "markers", DEFAULT_MARKERS)
    detection = merged_rule_section(layout_rules, "content_detection", DEFAULT_CONTENT_DETECTION)
    question_match = first_regex_match(markers["question_patterns"], text)
    option_match = first_regex_match(markers["option_patterns"], text)
    parenthesized_option_match = first_regex_match(markers["parenthesized_option_patterns"], text)
    judgement_match = first_regex_match(markers["judgement_patterns"], text)
    if text == ANSWER_LINE_SENTINEL:
        return apply_style_rule({
            "kind": "answer_line",
            "font": kai,
            "size": 14,
            "leading": 22,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 0,
            "space_after": 2,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "answer_line")
    if is_section_title(text, layout_rules):
        return apply_style_rule({
            "kind": "section",
            "font": yan_bold,
            "size": 16,
            "leading": 24,
            "align": "center",
            "color": "#111111",
            "space_before": 3,
            "space_after": 8,
            "first_line_indent": 0,
            "bar": True,
            "bar_color": config_color(layout_rules, "answer_gray", "#898989")
            if is_answer
            else config_color(layout_rules, "practice_bar", "#fce5e4"),
            "display_text": section_display_text(text),
        }, layout_rules, "section_title")
    if pos == 0 or is_lesson_title(text, layout_rules) or (is_answer and text == "参考答案与解析"):
        return apply_style_rule({
            "kind": "main_title",
            "font": yan_mid,
            "size": 23,
            "leading": 34,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 0,
            "space_after": 16,
            "first_line_indent": 0,
            "bar": False,
            "title_marker": True,
            "title_marker_asset": "参考答案" if is_answer else "标题角标",
        }, layout_rules, "main_title")
    if is_topic_heading(text, layout_rules):
        return apply_style_rule({
            "kind": "topic_heading",
            "font": yan_mid,
            "size": 14,
            "leading": 24,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 6,
            "space_after": 4,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "topic_heading")
    if not is_answer and (force_reading_prompt or is_reading_prompt(text, layout_rules)):
        return apply_style_rule({
            "kind": "reading_prompt",
            "font": yan_mid,
            "size": 14,
            "leading": 24,
            "align": "left",
            "color": config_color(layout_rules, "practice_red", "#cc0000"),
            "space_before": 4,
            "space_after": 10,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "reading_prompt")
    if text in set(detection["author_names"]):
        return apply_style_rule({
            "kind": "author",
            "font": kai,
            "size": 14,
            "leading": 20,
            "align": "center",
            "color": config_color(layout_rules, "author_text", "#333333"),
            "space_before": 0,
            "space_after": 5,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "author")
    if not is_answer and is_article_title_text(text, layout_rules):
        return apply_style_rule({
            "kind": "article_title",
            "font": yan_mid,
            "size": 14,
            "leading": 24,
            "align": "center",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 0,
            "space_after": 4,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "article_title")
    answer_label_text = normalize_answer_label(text, layout_rules) if is_answer else None
    if answer_label_text:
        return apply_style_rule({
            "kind": "answer_label",
            "font": yan_mid,
            "size": 14,
            "leading": 23,
            "align": "left",
            "color": "#111111",
            "space_before": 5,
            "space_after": 2,
            "first_line_indent": 0,
            "bar": False,
            "display_text": answer_label_text,
        }, layout_rules, "answer_label")
    if question_match and not is_answer:
        number, body = question_match.groups()
        return apply_style_rule({
            "kind": "question_numbered",
            "font": yan_regular,
            "size": 14,
            "leading": 27,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 0,
            "space_after": 2,
            "first_line_indent": 0,
            "bar": False,
            "display_text": body,
            "badge_text": number,
            "badge_color": config_color(layout_rules, "practice_red", "#cc0000"),
            "left_indent": 36,
            "wrap_width_factor": 0.84,
        }, layout_rules, "question_stem")
    if question_match and is_answer:
        number, body = question_match.groups()
        return apply_style_rule({
            "kind": "answer_body",
            "font": kai,
            "size": 14,
            "leading": 23,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 0,
            "space_after": 3,
            "first_line_indent": 0,
            "bar": False,
            "display_text": body,
            "marker_text": f"{number}.",
            "left_indent": 22,
        }, layout_rules, "answer_body")
    if option_match:
        marker, body = option_match.groups()
        return apply_style_rule({
            "kind": "option",
            "font": kai,
            "size": 14,
            "leading": 27,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 1,
            "space_after": 1,
            "first_line_indent": 0,
            "bar": False,
            "display_text": f"{marker}. {body}",
            "marker_text": f"{marker}.",
            "inline_marker": True,
            "left_indent": 36,
            "wrap_width_factor": 0.84,
        }, layout_rules, "option")
    if parenthesized_option_match:
        marker, body = parenthesized_option_match.groups()
        return apply_style_rule({
            "kind": "option",
            "font": kai,
            "size": 14,
            "leading": 27,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 1,
            "space_after": 1,
            "first_line_indent": 0,
            "bar": False,
            "display_text": f"{marker} {body}",
            "marker_text": marker,
            "inline_marker": True,
            "left_indent": 36,
            "wrap_width_factor": 0.84,
        }, layout_rules, "option")
    if not is_answer and judgement_match:
        return apply_style_rule({
            "kind": "judgement_item",
            "font": kai,
            "size": 14,
            "leading": 27,
            "align": "left",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 1,
            "space_after": 1,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "judgement_item")
    if matches_any(detection["source_patterns"], text):
        return apply_style_rule({
            "kind": "source",
            "font": yan_mid,
            "size": 12,
            "leading": 20,
            "align": "left",
            "color": config_color(layout_rules, "source_text", "#898989"),
            "space_before": 1,
            "space_after": 0,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "source")
    return apply_style_rule({
        "kind": "answer_body" if is_answer else "body",
        "font": kai,
        "size": 14,
        "leading": 23 if is_answer else 24,
        "align": "left" if is_answer else "justify",
        "color": config_color(layout_rules, "body_text", "#222222"),
        "space_before": 1,
        "space_after": 3 if is_answer else 2,
        "first_line_indent": 18 if re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩]", text) else 0,
        "underline_ranges": underline_ranges,
        "bar": False,
    }, layout_rules, "answer_body" if is_answer else "article_body")


def text_width(text: str, font: str, size: float) -> float:
    return pdfmetrics.stringWidth(text, font, size)


def wrap_text(text: str, style: dict[str, Any], max_w: float) -> list[str]:
    font = style["font"]
    size = style["size"]
    lines: list[str] = []
    cur = ""
    for ch in text:
        if ch == "\u2007":
            ch = " "
        if not cur and lines and ch in FORBIDDEN_LINE_START:
            lines[-1] += ch
            continue
        test = cur + ch
        if cur and text_width(test, font, size) > max_w:
            trailing_answer_match = re.search(r"\s+[（(]\s*$", cur)
            if trailing_answer_match:
                prefix = cur[:trailing_answer_match.start()]
                suffix = cur[trailing_answer_match.start():] + ch
                if prefix:
                    lines.append(prefix)
                cur = suffix
                continue
            if ch in FORBIDDEN_LINE_START:
                lines.append(test)
                cur = ""
            elif cur[-1] in FORBIDDEN_LINE_END and len(cur) > 1:
                lines.append(cur[:-1])
                cur = cur[-1] + ch
            else:
                lines.append(cur)
                cur = ch
        else:
            cur = test
    if cur:
        lines.append(cur)
    lines = avoid_isolated_answer_parentheses(lines)
    lines = rebalance_short_final_line(lines, style, max_w)
    lines = repair_forbidden_line_starts(lines)
    return lines or [""]


def avoid_isolated_answer_parentheses(lines: list[str]) -> list[str]:
    if len(lines) < 2 or not re.match(r"^\s*[（(]\s+[）)]\s*$", lines[-1]):
        return lines
    previous = lines[-2].rstrip()
    if not previous:
        return lines
    move_count = min(len(previous), 4)
    moved = previous[-move_count:]
    lines[-2] = previous[:-move_count]
    lines[-1] = f"{moved}{lines[-1]}"
    if not lines[-2]:
        lines.pop(-2)
    return lines


def rebalance_short_final_line(lines: list[str], style: dict[str, Any], max_w: float) -> list[str]:
    if len(lines) < 2 or max_w <= 0:
        return lines
    font = style["font"]
    size = style["size"]
    last = lines[-1]
    if not last.strip() or re.match(r"^\s*[（(]\s+[）)]\s*$", last):
        return lines
    target_w = min(120, max(84, max_w * 0.30))
    if text_width(last, font, size) >= target_w:
        return lines
    prev = lines[-2].rstrip()
    if len(prev) <= 4:
        return lines
    min_prev_w = max_w * 0.72
    while text_width(last, font, size) < target_w and len(prev) > 4:
        moved = prev[-1]
        if moved in FORBIDDEN_LINE_START:
            break
        candidate_prev = prev[:-1].rstrip()
        candidate_last = moved + last
        if text_width(candidate_last, font, size) > max_w:
            break
        if text_width(candidate_prev, font, size) < min_prev_w:
            break
        prev = candidate_prev
        last = candidate_last
    while prev and prev[-1] in FORBIDDEN_LINE_END and len(prev) > 1:
        candidate_prev = prev[:-1].rstrip()
        candidate_last = prev[-1] + last
        if text_width(candidate_last, font, size) > max_w:
            break
        prev = candidate_prev
        last = candidate_last
    lines[-2] = prev
    lines[-1] = last
    return lines


def repair_forbidden_line_starts(lines: list[str]) -> list[str]:
    if len(lines) < 2:
        return lines
    repaired = list(lines)
    for index in range(1, len(repaired)):
        while repaired[index] and repaired[index][0] in FORBIDDEN_LINE_START:
            repaired[index - 1] += repaired[index][0]
            repaired[index] = repaired[index][1:]
        if repaired[index] == "":
            repaired[index] = ""
    return [line for line in repaired if line != ""]


def wrap_text_by_widths(text: str, style: dict[str, Any], widths: list[float]) -> list[str]:
    if not widths:
        return wrap_text(text, style, 0)
    remaining = text
    lines: list[str] = []
    width_index = 0
    while remaining:
        width = widths[min(width_index, len(widths) - 1)]
        wrapped = wrap_text(remaining, style, width)
        first = wrapped[0]
        lines.append(first)
        remaining = remaining[len(first):]
        width_index += 1
    return lines or [""]


def fit_lines(
    text: Any,
    style: dict[str, Any],
    frame: dict[str, float],
    cursor: float,
) -> tuple[list[str], str]:
    text_value = block_text(text)
    y_bottom = frame["y"] + frame["h"] - 6
    if style.get("kind") == "answer_line":
        cursor += style["space_before"]
        if cursor + style["leading"] > y_bottom:
            return [], text_value
        return [""], ""
    base_w = (frame["w"] - 16 - style.get("left_indent", 0)) * style.get("wrap_width_factor", 1)
    if style.get("first_line_indent"):
        lines = wrap_text_by_widths(style.get("display_text", text_value), style, [base_w - style["first_line_indent"], base_w])
    else:
        lines = wrap_text(style.get("display_text", text_value), style, base_w)
    cursor += style["space_before"]
    if style.get("bar"):
        cursor += 2
    fit: list[str] = []
    rest: list[str] = []
    for line in lines:
        if cursor + style["leading"] > y_bottom:
            rest.append(line)
        elif rest:
            rest.append(line)
        else:
            cursor += style["leading"]
            fit.append(line)
    return fit, "".join(rest)


def min_block_height_for_keep(style: dict[str, Any]) -> float:
    height = style.get("space_before", 0) + style.get("leading", 0) + style.get("space_after", 0)
    if style.get("bar"):
        height += 2
    return height


def should_keep_with_next(style: dict[str, Any], layout_rules: dict[str, Any] | None = None) -> bool:
    page_flow = (layout_rules or {}).get("page_flow", {})
    keep_kinds = page_flow.get("keep_with_next_kinds", DEFAULT_KEEP_WITH_NEXT_KINDS)
    return style.get("kind") in set(keep_kinds)


def line_x_offset(style: dict[str, Any], first_line: bool) -> float:
    indent = style.get("first_line_indent", 0) if first_line else 0
    return style.get("left_indent", 0) + indent


def draw_paragraph_lines(
    c: canvas.Canvas,
    lines: list[str],
    style: dict[str, Any],
    frame: dict[str, float],
    start_cursor: float,
    svg_assets: dict[str, dict[str, Any]],
    has_rest: bool = False,
    color_mode: str = "rgb",
) -> float:
    x = frame["x"]
    w = frame["w"]
    cursor = start_cursor + style["space_before"]
    if style.get("bar") and lines:
        draw_title_bar_svg(c, x + 1, cursor, w - 2, svg_assets, style["bar_color"], color_mode)
        # The title bar SVG is 28.35 pt high; set the text baseline near its optical center.
        cursor -= 6
    if style.get("kind") == "answer_line" and lines:
        cursor += style["leading"]
        c.setStrokeColor(color_from_hex("#222222", (.13, .13, .13), color_mode))
        c.setLineWidth(0.45)
        y = c._pagesize[1] - cursor + 3
        x_offset = line_x_offset(style, first_line=True)
        c.line(x + 8 + x_offset, y, x + w - 8, y)
        cursor += style["space_after"]
        return cursor
    c.setFont(style["font"], style["size"])
    c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
    first_line = True
    text_offset = 0
    for line_index, line in enumerate(lines):
        cursor += style["leading"]
        x_offset = line_x_offset(style, first_line)
        available_w = w - 16 - x_offset
        if style["align"] == "center":
            tx = x + w / 2 - text_width(line, style["font"], style["size"]) / 2
        elif style["align"] == "right":
            tx = x + w - 8 - text_width(line, style["font"], style["size"])
        else:
            tx = x + 8 + x_offset
        if first_line and style.get("title_marker"):
            # Align the 30.05 pt corner SVG with the main title's visual center.
            draw_title_corner_svg(
                c,
                x + 1,
                cursor - 22,
                svg_assets,
                style.get("title_marker_asset", "标题角标"),
                color_mode,
            )
            c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
            tx = x + 42
        if first_line and style.get("badge_text"):
            bx = x + 8
            # Align the 12.76 pt number SVG to the first-line baseline.
            draw_number_badge_svg(c, bx, cursor - 12, style["badge_text"], style["font"], svg_assets, color_mode)
            c.setFont(style["font"], style["size"])
            c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
        if first_line and style.get("marker_text") and not style.get("inline_marker"):
            marker_gap = style.get("marker_gap", 4)
            marker_right = x + 8 + max(0, style.get("left_indent", 0) - marker_gap)
            c.drawRightString(marker_right, c._pagesize[1] - cursor, style["marker_text"])
        if first_line and style.get("bold_rule"):
            c.setStrokeColor(color_from_hex("#cccccc", (.8, .8, .8), color_mode))
            c.setLineWidth(0.4)
            y_rule = c._pagesize[1] - cursor - 6
            c.line(x + 8, y_rule, x + w - 8, y_rule)
        should_justify = (
            style["align"] == "justify"
            and len(line) > 1
            and (line_index < len(lines) - 1 or has_rest)
        )
        if should_justify:
            line_width = text_width(line, style["font"], style["size"])
            extra = max(0, (available_w - line_width) / (len(line) - 1))
            if extra > style["size"] * 0.7:
                c.drawString(tx, c._pagesize[1] - cursor, line)
                draw_underlines_for_line(c, line, text_offset, tx, cursor, style, color_mode)
                text_offset += len(line)
                first_line = False
                continue
            text_obj = c.beginText(tx, c._pagesize[1] - cursor)
            text_obj.setFont(style["font"], style["size"])
            text_obj.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
            text_obj.setCharSpace(extra)
            text_obj.textLine(line)
            c.drawText(text_obj)
        else:
            c.drawString(tx, c._pagesize[1] - cursor, line)
        draw_underlines_for_line(c, line, text_offset, tx, cursor, style, color_mode)
        text_offset += len(line)
        first_line = False
    if lines:
        cursor += style["space_after"]
    return cursor


def draw_underlines_for_line(
    c: canvas.Canvas,
    line: str,
    line_start: int,
    tx: float,
    cursor: float,
    style: dict[str, Any],
    color_mode: str = "rgb",
) -> None:
    ranges = style.get("underline_ranges") or []
    if not ranges or not line:
        return
    line_end = line_start + len(line)
    c.setStrokeColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
    c.setLineWidth(0.45)
    y = c._pagesize[1] - cursor - 2.2
    for start, end in ranges:
        overlap_start = max(start, line_start)
        overlap_end = min(end, line_end)
        if overlap_start >= overlap_end:
            continue
        prefix = line[: overlap_start - line_start]
        segment = line[overlap_start - line_start : overlap_end - line_start]
        x1 = tx + text_width(prefix, style["font"], style["size"])
        x2 = x1 + text_width(segment, style["font"], style["size"])
        c.line(x1, y, x2, y)


def table_style(fonts: dict[str, str], layout_rules: dict[str, Any]) -> dict[str, Any]:
    style = {
        "kind": "table",
        "font": fonts["body"],
        "size": 13,
        "leading": 20,
        "space_before": 5,
        "space_after": 8,
        "cell_pad_x": 6,
        "cell_pad_y": 5,
        "border_color": "#777777",
        "header_fill": "#e6e6e6",
        "color": config_color(layout_rules, "body_text", "#222222"),
    }
    return apply_style_rule(style, layout_rules, "table")


def image_style(fonts: dict[str, str], layout_rules: dict[str, Any]) -> dict[str, Any]:
    style = {
        "kind": "image",
        "space_before": 8,
        "space_after": 10,
        "left_indent": 0,
    }
    return apply_style_rule(style, layout_rules, "image")


def image_scaled_size(block: dict[str, Any], available_w: float, max_h: float | None = None) -> tuple[float, float]:
    width_px = max(1, block.get("width_px") or 1)
    height_px = max(1, block.get("height_px") or 1)
    scale = available_w / width_px
    width = available_w
    height = height_px * scale
    if max_h is not None and height > max_h:
        scale = max_h / height_px
        width = width_px * scale
        height = max_h
    return width, height


def image_total_height(block: dict[str, Any], style: dict[str, Any], frame_w: float) -> float:
    available_w = frame_w - 16 - style.get("left_indent", 0)
    _, height = image_scaled_size(block, available_w)
    return style["space_before"] + height + style["space_after"]


def draw_image_block(
    c: canvas.Canvas,
    block: dict[str, Any],
    style: dict[str, Any],
    frame: dict[str, float],
    start_cursor: float,
) -> float:
    left_indent = style.get("left_indent", 0)
    available_w = frame["w"] - 16 - left_indent
    y_bottom = frame["y"] + frame["h"] - 6
    cursor = start_cursor + style["space_before"]
    max_h = max(12, y_bottom - cursor - style["space_after"])
    draw_w, draw_h = image_scaled_size(block, available_w, max_h=max_h)
    x = frame["x"] + 8 + left_indent + (available_w - draw_w) / 2
    y = c._pagesize[1] - cursor - draw_h
    c.drawImage(ImageReader(BytesIO(block["blob"])), x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
    return cursor + draw_h + style["space_after"]


def table_row_heights(rows: list[list[str]], style: dict[str, Any], table_w: float) -> list[float]:
    if not rows:
        return []
    col_count = max(len(row) for row in rows)
    col_w = table_w / max(1, col_count)
    heights: list[float] = []
    for row in rows:
        max_lines = 1
        for cell in row:
            lines = wrap_text(cell, style, col_w - style["cell_pad_x"] * 2)
            max_lines = max(max_lines, len(lines))
        heights.append(max(24, max_lines * style["leading"] + style["cell_pad_y"] * 2))
    return heights


def table_total_height(rows: list[list[str]], style: dict[str, Any], table_w: float) -> float:
    return style["space_before"] + sum(table_row_heights(rows, style, table_w)) + style["space_after"]


def draw_table_block(
    c: canvas.Canvas,
    block: dict[str, Any],
    style: dict[str, Any],
    frame: dict[str, float],
    start_cursor: float,
    color_mode: str = "rgb",
) -> float:
    rows = table_rows(block)
    left_indent = style.get("left_indent", 0)
    x = frame["x"] + 8 + left_indent
    w = frame["w"] - 16 - left_indent
    cursor = start_cursor + style["space_before"]
    y_top_pdf = c._pagesize[1] - cursor
    col_count = max((len(row) for row in rows), default=1)
    col_w = w / max(1, col_count)
    heights = table_row_heights(rows, style, w)
    c.setFont(style["font"], style["size"])
    c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
    c.setStrokeColor(color_from_hex(style["border_color"], (.47, .47, .47), color_mode))
    c.setLineWidth(0.45)
    y = y_top_pdf
    for row_index, row in enumerate(rows):
        row_h = heights[row_index]
        if row_index == 0:
            c.setFillColor(color_from_hex(style["header_fill"], (.9, .9, .9), color_mode))
            c.rect(x, y - row_h, w, row_h, fill=1, stroke=0)
            c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
        for col_index in range(col_count):
            cell_x = x + col_w * col_index
            c.rect(cell_x, y - row_h, col_w, row_h, fill=0, stroke=1)
            text = row[col_index] if col_index < len(row) else ""
            lines = wrap_text(text, style, col_w - style["cell_pad_x"] * 2)
            line_y = y - style["cell_pad_y"] - style["size"]
            for line in lines:
                c.drawString(cell_x + style["cell_pad_x"], line_y, line)
                line_y -= style["leading"]
        y -= row_h
    return cursor + sum(heights) + style["space_after"]


def paint_background(
    c: canvas.Canvas,
    background_path: Path,
    page_w: float,
    page_h: float,
    mode: str,
    color_mode: str = "rgb",
) -> None:
    if mode == "white":
        c.setFillColor(color_from_hex("#ffffff", (1, 1, 1), color_mode))
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    elif background_path.exists():
        c.drawImage(ImageReader(str(background_path)), 0, 0, width=page_w, height=page_h, preserveAspectRatio=False, mask=None)
    else:
        c.setFillColor(color_from_hex("#ffffff", (1, 1, 1), color_mode))
        c.rect(0, 0, page_w, page_h, fill=1, stroke=0)


def paint_template_shell(
    c: canvas.Canvas,
    spread: dict[str, Any],
    page_w: float,
    page_h: float,
    footer_font: str,
    background_mode: str,
    layout_rules: dict[str, Any],
    page_number_start: int | None = None,
    color_mode: str = "rgb",
) -> None:
    for slot in spread["slots"]:
        b = slot["bbox_pt"]
        x, y, w, h = b["x"], b["y"], b["w"], b["h"]
        if slot["role"] in ("side_strip", "side_strip_textframe"):
            if spread["spread_index"] >= 4:
                continue
            fill = (
                config_color(layout_rules, "practice_bar", "#fce5e4")
                if spread["spread_index"] < 4
                else config_color(layout_rules, "answer_gray", "#898989")
            )
            c.setFillColor(color_from_hex(fill, (.94, .86, .89), color_mode))
            c.rect(x, page_h - y - h, w, h, fill=1, stroke=0)
        elif slot["role"] == "page_number":
            if page_number_start is None:
                txt = (slot.get("story_preview") or "").strip() or "· ·"
            else:
                page_no = page_number_start + (1 if x + w / 2 >= page_w / 2 else 0)
                txt = f"·{page_no}·"
            c.setFont(footer_font, 10)
            c.setFillColor(color_from_hex("#555555", (.33, .33, .33), color_mode))
            c.drawCentredString(x + w / 2, page_h - y - h / 2 - 3, txt)


def render_flow(
    c: canvas.Canvas,
    template: dict[str, Any],
    paragraphs: list[Any],
    frames: list[dict[str, Any]],
    background_dir: Path,
    page_w: float,
    page_h: float,
    style_fonts: dict[str, str],
    is_answer: bool,
    svg_assets: dict[str, dict[str, Any]],
    background_mode: str,
    page_number_start: int | None,
    layout_rules: dict[str, Any],
    allow_repeat: bool,
    color_mode: str = "rgb",
) -> dict[str, Any]:
    grouped = group_frames_by_spread(frames)
    frame_entries = list(grouped.items())
    paragraph_idx = 0
    remainder: Any | None = None
    rendered_pages: list[dict[str, Any]] = []
    spread_slot = 0
    inherited_hanging_indent: float | None = None
    while paragraph_idx < len(paragraphs) or remainder is not None:
        if not frame_entries:
            break
        if spread_slot < len(frame_entries):
            spread_index, frame_list = frame_entries[spread_slot]
            repeated_from: int | None = None
        else:
            if not allow_repeat:
                break
            spread_index, frame_list = frame_entries[-1]
            repeated_from = spread_index
        if paragraph_idx >= len(paragraphs) and remainder is None:
            break
        spread = spread_by_index(template, spread_index)
        bg = background_path_for_spread(background_dir, spread_index)
        paint_background(c, bg, page_w, page_h, background_mode, color_mode)
        paint_template_shell(
            c,
            spread,
            page_w,
            page_h,
            style_fonts["footer"],
            background_mode,
            layout_rules,
            page_number_start=page_number_start + len(rendered_pages) * 2,
            color_mode=color_mode,
        )
        page_started_at = paragraph_idx
        remainder_started_at = remainder
        used_sides: set[str] = set()
        for frame_ref in frame_list:
            frame = frame_ref["bbox_pt"]
            side = frame_side(frame, page_w)
            cursor = frame["y"] + 8
            while paragraph_idx < len(paragraphs):
                text = remainder if remainder is not None else paragraphs[paragraph_idx]
                if is_image_block(text):
                    style = image_style(style_fonts, layout_rules)
                    style = apply_hanging_context_indent(style, inherited_hanging_indent, layout_rules)
                    image_h = image_total_height(text, style, frame["w"])
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + image_h > y_bottom and cursor > frame["y"] + 12:
                        remainder = text
                        break
                    cursor = draw_image_block(c, text, style, frame, cursor)
                    used_sides.add(side)
                    remainder = None
                    paragraph_idx += 1
                    continue
                if is_table_block(text):
                    style = table_style(style_fonts, layout_rules)
                    style = apply_hanging_context_indent(style, inherited_hanging_indent, layout_rules)
                    rows = table_rows(text)
                    table_h = table_total_height(rows, style, frame["w"] - 16 - style.get("left_indent", 0))
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + table_h > y_bottom:
                        remainder = text
                        break
                    cursor = draw_table_block(c, text, style, frame, cursor, color_mode)
                    used_sides.add(side)
                    remainder = None
                    paragraph_idx += 1
                    continue
                if is_remainder_block(text):
                    style = remainder_style(text)
                else:
                    previous_text = block_text(paragraphs[paragraph_idx - 1]) if paragraph_idx > 0 else ""
                    force_reading_prompt = (
                        not is_answer
                        and is_section_title(previous_text, layout_rules)
                        and can_be_structural_reading_prompt(block_text(text), layout_rules)
                    )
                    style = paragraph_style(
                        text,
                        paragraph_idx,
                        is_answer,
                        style_fonts,
                        layout_rules,
                        force_reading_prompt=force_reading_prompt,
                    )
                    if ends_hanging_context(style, layout_rules):
                        inherited_hanging_indent = None
                    if not starts_hanging_context(style, layout_rules):
                        style = apply_hanging_context_indent(style, inherited_hanging_indent, layout_rules)
                if starts_hanging_context(style, layout_rules):
                    inherited_hanging_indent = style.get("left_indent", question_content_indent(layout_rules))
                if (
                    remainder is None
                    and should_keep_with_next(style, layout_rules)
                    and paragraph_idx + 1 < len(paragraphs)
                ):
                    next_text = paragraphs[paragraph_idx + 1]
                    if is_table_block(next_text):
                        next_style = table_style(style_fonts, layout_rules)
                        next_style = apply_hanging_context_indent(next_style, inherited_hanging_indent, layout_rules)
                        next_h = table_total_height(table_rows(next_text), next_style, frame["w"] - 16 - next_style.get("left_indent", 0))
                    elif is_image_block(next_text):
                        next_style = image_style(style_fonts, layout_rules)
                        next_style = apply_hanging_context_indent(next_style, inherited_hanging_indent, layout_rules)
                        next_h = image_total_height(next_text, next_style, frame["w"])
                    else:
                        next_style = paragraph_style(
                            next_text,
                            paragraph_idx + 1,
                            is_answer,
                            style_fonts,
                            layout_rules,
                        )
                        if not starts_hanging_context(next_style, layout_rules):
                            next_style = apply_hanging_context_indent(next_style, inherited_hanging_indent, layout_rules)
                        next_h = min_block_height_for_keep(next_style)
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + min_block_height_for_keep(style) + next_h > y_bottom and cursor > frame["y"] + 12:
                        remainder = text
                        break
                lines, rest = fit_lines(text, style, frame, cursor)
                if not lines and rest:
                    remainder = text
                    break
                cursor = draw_paragraph_lines(
                    c,
                    lines,
                    style,
                    frame,
                    cursor,
                    svg_assets,
                    has_rest=bool(rest),
                    color_mode=color_mode,
                )
                if lines:
                    used_sides.add(side)
                if rest:
                    remainder = make_remainder_block(rest, style)
                    break
                remainder = None
                paragraph_idx += 1
        rendered_pages.append(
            {
                "spread_index": spread_index,
                "repeated_from": repeated_from,
                "background": str(bg),
                "paragraph_start": page_started_at,
                "paragraph_end": paragraph_idx,
                "remaining_fragment": bool(remainder),
                "used_sides": sorted(used_sides, key=lambda value: 0 if value == "left" else 1),
            }
        )
        if paragraph_idx == page_started_at and remainder == remainder_started_at:
            raise RuntimeError(
                f"No layout progress on spread {spread_index}; check frame height or paragraph style."
            )
        c.showPage()
        spread_slot += 1
    return {
        "paragraphs_done": paragraph_idx,
        "paragraphs_total": len(paragraphs),
        "has_remaining_fragment": bool(remainder),
        "rendered_pages": rendered_pages,
    }


def render_preview(pdf_path: Path, preview_path: Path) -> None:
    try:
        import fitz  # type: ignore[import-not-found]
    except ModuleNotFoundError:
        page_count = len(PdfReader(str(pdf_path)).pages)
        sheet = Image.new("RGB", (900, 180), "white")
        sheet.save(preview_path, quality=95)
        return
    pdf = fitz.open(str(pdf_path))
    images = []
    for page in pdf:
        pix = page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False)
        im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        im.thumbnail((1191, 808))
        images.append(im.copy())
    gap = 24
    sheet = Image.new(
        "RGB",
        (max(i.width for i in images), sum(i.height for i in images) + gap * (len(images) - 1)),
        "white",
    )
    y = 0
    for im in images:
        sheet.paste(im, (0, y))
        y += im.height + gap
    sheet.save(preview_path, quality=95)


def validate_pdf(pdf_path: Path, expected_w: float, expected_h: float) -> dict[str, Any]:
    reader = PdfReader(str(pdf_path))
    sizes = []
    for i, page in enumerate(reader.pages, 1):
        w = round(float(page.mediabox.width), 3)
        h = round(float(page.mediabox.height), 3)
        sizes.append({"page": i, "w": w, "h": h})
        if w != round(expected_w, 3) or h != round(expected_h, 3):
            raise AssertionError(f"Page {i} size mismatch: {w} x {h}")
    return {"pages": len(reader.pages), "sizes": sizes}


def split_spread_pdf_to_single_pages(
    spread_pdf_path: Path,
    output_pdf_path: Path,
    single_page_w: float,
    single_page_h: float,
    used_sides_by_spread_page: list[list[str]] | None = None,
    renumber_pages: bool = True,
) -> None:
    reader = PdfReader(str(spread_pdf_path))
    writer = PdfWriter()
    output_page_no = 1
    for spread_page_index, page in enumerate(reader.pages):
        used_sides = (
            used_sides_by_spread_page[spread_page_index]
            if used_sides_by_spread_page and spread_page_index < len(used_sides_by_spread_page)
            else ["left", "right"]
        )
        for side in used_sides:
            offset = 0 if side == "left" else single_page_w
            single_page = PageObject.create_blank_page(width=single_page_w, height=single_page_h)
            single_page.merge_transformed_page(page, Transformation().translate(tx=-offset, ty=0))
            if renumber_pages:
                single_page.merge_page(page_number_overlay(single_page_w, single_page_h, output_page_no))
            output_page_no += 1
            writer.add_page(single_page)
    with output_pdf_path.open("wb") as f:
        writer.write(f)


def page_number_overlay(page_w: float, page_h: float, page_number: int) -> PageObject:
    buffer = BytesIO()
    overlay = canvas.Canvas(buffer, pagesize=(page_w, page_h))
    overlay.setFillColor(Color(1, 1, 1))
    overlay.rect(page_w / 2 - 38, 14, 76, 28, fill=1, stroke=0)
    overlay.setFont("Helvetica", 10)
    overlay.setFillColor(Color(.33, .33, .33))
    overlay.drawCentredString(page_w / 2, 26, f"·{page_number}·")
    overlay.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


def rendered_used_sides(*results: dict[str, Any]) -> list[list[str]]:
    sides: list[list[str]] = []
    for result in results:
        for rendered_page in result.get("rendered_pages", []):
            used = rendered_page.get("used_sides") or ["left", "right"]
            sides.append(used)
    return sides


def rendered_single_page_count(result: dict[str, Any]) -> int:
    return sum(len(page.get("used_sides") or ["left", "right"]) for page in result.get("rendered_pages", []))


def build_pdf(args: argparse.Namespace) -> dict[str, Any]:
    template_path = Path(args.template)
    font_map_path = Path(args.font_map)
    docx_path = Path(args.docx)
    output_dir = Path(args.output_dir)
    svg_dir = Path(args.svg_dir)
    background_dir = Path(args.background_dir) if args.background_dir else output_dir / "backgrounds"
    ensure_dir(output_dir)
    ensure_dir(background_dir)

    template = json.loads(template_path.read_text(encoding="utf-8"))
    font_map = json.loads(font_map_path.read_text(encoding="utf-8"))
    layout_rules = load_optional_json(getattr(args, "layout_rules", None))
    asset_map = load_optional_json(getattr(args, "asset_map", None))
    _, substitutions = register_fonts(font_map, font_map_path.parent)
    svg_assets = load_svg_assets(svg_dir, asset_map)

    single_page_w = template["document"]["page_width_pt"]
    spread_page_w = single_page_w * 2
    page_h = template["document"]["page_height_pt"]
    page_mode = getattr(args, "page_mode", "single")
    if page_mode not in {"single", "spread"}:
        raise ValueError(f"Unsupported page mode: {page_mode}")
    color_mode = getattr(args, "color_mode", "cmyk")
    if color_mode not in {"cmyk", "rgb"}:
        raise ValueError(f"Unsupported color mode: {color_mode}")
    output_page_w = single_page_w if page_mode == "single" else spread_page_w
    style_fonts = font_map["style_defaults"]
    page_flow_rules = layout_rules.get("page_flow", {})
    dynamic_page_numbers = page_flow_rules.get("dynamic_page_numbers", True)

    background_records = ensure_backgrounds(template, background_dir, spread_page_w, page_h, args.background_dpi)
    prompts_path = write_background_prompt_manifest(output_dir, template)

    paragraphs = load_docx_paragraphs(docx_path)
    practice, answers = split_practice_and_answers(paragraphs)

    pdf_path = output_dir / args.pdf_name
    spread_pdf_path = output_dir / f".{pdf_path.stem}.spread-work.pdf" if page_mode == "single" else pdf_path
    preview_path = output_dir / args.preview_name
    c = canvas.Canvas(str(spread_pdf_path), pagesize=(spread_page_w, page_h))
    c.setTitle(args.title)
    practice_result = render_flow(
        c,
        template,
        practice,
        normalize_flow_frames(story_frames(template, "practice_content_flow"), "practice_content_flow", layout_rules),
        background_dir,
        spread_page_w,
        page_h,
        style_fonts,
        is_answer=False,
        svg_assets=svg_assets,
        background_mode=args.background_mode,
        page_number_start=1 if dynamic_page_numbers else None,
        layout_rules=layout_rules,
        allow_repeat=page_flow_rules.get("repeat_last_practice_spread_when_overflow", True),
        color_mode=color_mode,
    )
    answer_page_number_start = 1 + (
        rendered_single_page_count(practice_result)
        if page_mode == "single"
        else len(practice_result["rendered_pages"]) * 2
    )
    answer_result = render_flow(
        c,
        template,
        answers,
        normalize_flow_frames(story_frames(template, "answer_content_flow"), "answer_content_flow", layout_rules),
        background_dir,
        spread_page_w,
        page_h,
        style_fonts,
        is_answer=True,
        svg_assets=svg_assets,
        background_mode=args.background_mode,
        page_number_start=answer_page_number_start if dynamic_page_numbers else None,
        layout_rules=layout_rules,
        allow_repeat=page_flow_rules.get("repeat_last_answer_spread_when_overflow", True),
        color_mode=color_mode,
    )
    c.save()
    if page_mode == "single":
        split_spread_pdf_to_single_pages(
            spread_pdf_path,
            pdf_path,
            single_page_w,
            page_h,
            used_sides_by_spread_page=rendered_used_sides(practice_result, answer_result),
            renumber_pages=dynamic_page_numbers,
        )
        spread_pdf_path.unlink(missing_ok=True)

    render_preview(pdf_path, preview_path)
    pdf_validation = validate_pdf(pdf_path, output_page_w, page_h)

    manifest = {
        "schema_version": "0.1",
        "title": args.title,
        "inputs": {
            "template": str(template_path),
            "font_map": str(font_map_path),
            "docx": str(docx_path),
            "background_dir": str(background_dir),
            "svg_dir": str(svg_dir),
            "layout_rules": str(args.layout_rules) if getattr(args, "layout_rules", None) else None,
            "asset_map": str(args.asset_map) if getattr(args, "asset_map", None) else None,
            "page_mode": page_mode,
            "color_mode": color_mode,
        },
        "outputs": {
            "pdf": str(pdf_path),
            "preview": str(preview_path),
            "background_prompts": str(prompts_path),
        },
        "page_mode": page_mode,
        "color_mode": color_mode,
        "page_size_pt": {"w": round(output_page_w, 3), "h": round(page_h, 3)},
        "spread_page_size_pt": {"w": round(spread_page_w, 3), "h": round(page_h, 3)},
        "backgrounds": background_records,
        "font_substitutions": substitutions,
        "svg_assets": svg_assets,
        "layout_rules": layout_rules,
        "asset_map": asset_map,
        "background_mode": args.background_mode,
        "practice_result": practice_result,
        "answer_result": answer_result,
        "pdf_validation": pdf_validation,
    }
    manifest_path = output_dir / "generation-manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    manifest["outputs"]["manifest"] = str(manifest_path)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a print PDF from an extracted IDML template and DOCX content.")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    parser.add_argument("--font-map", default=str(DEFAULT_FONT_MAP))
    parser.add_argument("--docx", default=str(DEFAULT_DOCX))
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--background-dir", default=None)
    parser.add_argument("--background-mode", choices=["white", "image"], default="white")
    parser.add_argument("--svg-dir", default=str(DEFAULT_SVG_DIR))
    parser.add_argument("--background-dpi", type=int, default=220)
    parser.add_argument("--pdf-name", default="贾平凹标题含义_正式流程版.pdf")
    parser.add_argument("--preview-name", default="贾平凹标题含义_正式流程版_预览.png")
    parser.add_argument("--title", default="贾平凹标题含义理解")
    parser.add_argument("--page-mode", choices=["single", "spread"], default="single")
    parser.add_argument("--color-mode", choices=["cmyk", "rgb"], default="cmyk")
    parser.add_argument("--layout-rules", default=None)
    parser.add_argument("--asset-map", default=None)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    manifest = build_pdf(args)
    print("PDF", manifest["outputs"]["pdf"])
    print("PREVIEW", manifest["outputs"]["preview"])
    print("MANIFEST", manifest["outputs"]["manifest"])
    print("PAGES", manifest["pdf_validation"]["pages"])
    print("PRACTICE", manifest["practice_result"]["paragraphs_done"], "/", manifest["practice_result"]["paragraphs_total"])
    print("ANSWERS", manifest["answer_result"]["paragraphs_done"], "/", manifest["answer_result"]["paragraphs_total"])


if __name__ == "__main__":
    main()
