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
    "cell_pad_x",
    "cell_pad_y",
    "max_height_pt",
    "wrap_width_factor",
    "min_last_line_chars",
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
        "reject_patterns": [r"^[A-E][.．]", r"^[1-9]\d*[.．、]", r"^祝$", r"^[X\d]+年[X\d]+月[X\d]+日$"],
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
    "inherit_kinds": ["body", "answer_body", "answer_line", "option", "question_numbered", "image"],
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


def apply_answer_line_context_spacing(
    style: dict[str, Any],
    previous_style_kind: str | None,
    layout_rules: dict[str, Any],
) -> dict[str, Any]:
    if style.get("kind") != "answer_line" or previous_style_kind == "answer_line":
        return style
    first_space_before = layout_rules.get("styles", {}).get("answer_line", {}).get("first_space_before", 6)
    if not first_space_before:
        return style
    adjusted = dict(style)
    adjusted["space_before"] = adjusted.get("space_before", 0) + first_space_before
    return adjusted


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
    return True


def normalize_answer_label(text: str, layout_rules: dict[str, Any] | None = None) -> str | None:
    normalization = merged_rule_section(layout_rules, "normalization", DEFAULT_NORMALIZATION)
    return normalization["answer_labels"].get(text)


def is_answer_main_title(text: str) -> bool:
    return bool(re.search(r"(?:参考)?答案与解析[:：]?$", text.strip()))


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


def _paragraph_has_fill_in_blanks(paragraph: Any) -> bool:
    """True if this paragraph has any blank (space-only) underlined run — fill-in style."""
    return any(
        not run.text.strip() and run.underline and run.text
        for run in paragraph.runs
    )


def _fill_in_text_units(paragraph: Any) -> list[tuple[str, bool]]:
    units: list[tuple[str, bool]] = []
    for run in paragraph.runs:
        raw = re.sub(r"[\r\n\t]+", " ", run.text)
        if not raw:
            continue
        has_u = "<w:u" in run._element.xml
        is_protected_blank = not raw.strip() and has_u
        units.extend((ch, is_protected_blank) for ch in raw)
    return units


def _fill_in_trim_bounds(paragraph: Any) -> tuple[int, int]:
    units = _fill_in_text_units(paragraph)
    start = 0
    while start < len(units) and units[start][0].isspace() and not units[start][1]:
        start += 1
    end = len(units)
    while end > start and units[end - 1][0].isspace() and not units[end - 1][1]:
        end -= 1
    return start, end


def paragraph_fill_in_text(paragraph: Any) -> str:
    """Build paragraph text preserving underlined blanks and trimming only padding."""
    units = _fill_in_text_units(paragraph)
    start, end = _fill_in_trim_bounds(paragraph)
    return "".join(ch for ch, _ in units[start:end])


def paragraph_underline_ranges(paragraph: Any) -> list[list[int]]:
    """Return character ranges for underline rendering.

    For fill-in paragraphs: positions are relative to paragraph_fill_in_text()
    (blank underlined runs remain spaces so only the vector underline is visible).
    For vocabulary paragraphs: positions are relative to clean_text() output.
    """
    is_fill_in = _paragraph_has_fill_in_blanks(paragraph)
    trim_start, trim_end = _fill_in_trim_bounds(paragraph) if is_fill_in else (0, None)
    ranges: list[list[int]] = []
    cursor = 0
    for run in paragraph.runs:
        raw = re.sub(r"[\r\n\t]+", " ", run.text)
        if not raw:
            continue
        has_u = "<w:u" in run._element.xml
        is_blank = not raw.strip()
        start = cursor
        if is_fill_in and is_blank and has_u:
            cursor += len(raw)
            end = cursor
            visible_start = max(start, trim_start)
            visible_end = min(end, trim_end if trim_end is not None else end)
            if visible_start < visible_end:
                ranges.append([
                    visible_start - trim_start,
                    visible_end - trim_start,
                    visible_end - visible_start,
                ])
        else:
            cursor += len(raw)
            if not is_fill_in and has_u and not is_blank:
                ranges.append([start, cursor])
    return ranges


def paragraph_bold_ranges(paragraph: Any) -> list[list[int]]:
    """Return character ranges where Word run-level bold is explicitly set."""
    ranges: list[list[int]] = []
    cursor = 0
    for run in paragraph.runs:
        text = re.sub(r"[\r\n\t]+", " ", run.text)
        if not text:
            continue
        start = cursor
        cursor += len(text)
        if run.bold is True:
            ranges.append([start, cursor])
    return ranges


def shift_inline_ranges(ranges: list[list[int]], offset: int, max_len: int) -> list[list[int]]:
    shifted: list[list[int]] = []
    for rng in ranges:
        start = max(0, rng[0] - offset)
        end = min(max_len, rng[1] - offset)
        if start >= end:
            continue
        next_rng = [start, end]
        if len(rng) > 2:
            next_rng.append(rng[2])
        shifted.append(next_rng)
    return shifted


def offset_inline_ranges(ranges: list[list[int]], offset: int) -> list[list[int]]:
    shifted: list[list[int]] = []
    for rng in ranges:
        next_rng = [rng[0] + offset, rng[1] + offset]
        if len(rng) > 2:
            next_rng.append(rng[2])
        shifted.append(next_rng)
    return shifted


def map_body_inline_ranges(
    ranges: list[list[int]],
    original_body_start: int,
    display_body_start: int,
    display_len: int,
) -> list[list[int]]:
    delta = display_body_start - original_body_start
    mapped: list[list[int]] = []
    for rng in ranges:
        if rng[1] <= original_body_start:
            start, end = rng[0], rng[1]
        else:
            start = rng[0] + delta if rng[0] >= original_body_start else rng[0]
            end = rng[1] + delta
        start = max(0, min(display_len, start))
        end = max(0, min(display_len, end))
        if start >= end:
            continue
        next_rng = [start, end]
        if len(rng) > 2:
            next_rng.append(rng[2])
        mapped.append(next_rng)
    return mapped


def clean_cell_text(text: str) -> str:
    return re.sub(r"[\r\n\t]+", " ", text).strip()


def cell_text_with_fill_ins(cell: Any) -> str:
    """Like clean_cell_text but converts blank underlined runs to underscore markers.
    Preserves paragraph breaks as \\n so column-width and multi-line rendering work correctly."""
    parts: list[str] = []
    for p in cell.paragraphs:
        if _paragraph_has_fill_in_blanks(p):
            t = paragraph_fill_in_text(p)
        else:
            t = re.sub(r"[\r\n\t]+", " ", p.text).strip()
        if t:
            parts.append(t)
    return "\n".join(parts)


def cell_text_and_underline_ranges(cell: Any) -> tuple[str, list[list[int]]]:
    """Return table-cell text plus underline ranges using original blank widths."""
    parts: list[str] = []
    ranges: list[list[int]] = []
    cursor = 0
    for p in cell.paragraphs:
        para_parts: list[str] = []
        para_cursor = 0
        is_fill_in = _paragraph_has_fill_in_blanks(p)
        for run in p.runs:
            raw = re.sub(r"[\r\n\t]+", " ", run.text)
            if not raw:
                continue
            has_u = "<w:u" in run._element.xml
            is_blank = not raw.strip()
            start = cursor + para_cursor
            para_parts.append(raw)
            para_cursor += len(raw)
            if is_fill_in and is_blank and has_u:
                ranges.append([start, cursor + para_cursor, len(raw)])
            elif not is_fill_in and has_u and not is_blank:
                ranges.append([start, cursor + para_cursor])
        para_text = "".join(para_parts).strip()
        if not para_text:
            continue
        if parts:
            cursor += 1
        parts.append(para_text)
        cursor += len(para_text)
    return "\n".join(parts), ranges


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


def is_image_row_block(block: Any) -> bool:
    return isinstance(block, dict) and block.get("type") == "image_row"


_IMAGE_ROW_LABEL_TOKEN_RE = re.compile(r'^[一-鿿]{1,4}$')


def _split_image_row_labels(text: str) -> list[str] | None:
    """Return label list if text is a row of short CJK labels (甲 乙 丙), else None."""
    tokens = re.split(r'[\s ]+', text.strip())
    if len(tokens) < 2 or len(tokens) > 6:
        return None
    if all(_IMAGE_ROW_LABEL_TOKEN_RE.match(t) for t in tokens):
        return tokens
    return None


def table_rows(block: dict[str, Any]) -> list[list[str]]:
    return block.get("rows", [])


_WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CHINESE_NUMS = "零一二三四五六七八九十"


def _build_numbering_map(doc) -> dict[str, Any]:
    """Parse numbering.xml into {abstract_nums, num_map} for list-number resolution."""
    try:
        num_part = doc.part.numbering_part
    except (AttributeError, NotImplementedError):
        return {}
    if num_part is None:
        return {}
    NS = _WNS
    tree = num_part._element
    abstract_nums: dict[str, dict] = {}
    for abs_num in tree.findall(f"{{{NS}}}abstractNum"):
        abs_id = abs_num.get(f"{{{NS}}}abstractNumId")
        levels: dict[str, dict] = {}
        for lvl in abs_num.findall(f"{{{NS}}}lvl"):
            ilvl = lvl.get(f"{{{NS}}}ilvl")
            fmt_el = lvl.find(f"{{{NS}}}numFmt")
            text_el = lvl.find(f"{{{NS}}}lvlText")
            start_el = lvl.find(f"{{{NS}}}start")
            levels[ilvl] = {
                "numFmt": fmt_el.get(f"{{{NS}}}val") if fmt_el is not None else "decimal",
                "lvlText": text_el.get(f"{{{NS}}}val") if text_el is not None else "%1",
                "start": int(start_el.get(f"{{{NS}}}val", "1")) if start_el is not None else 1,
            }
        abstract_nums[abs_id] = levels
    num_map: dict[str, dict] = {}
    for num in tree.findall(f"{{{NS}}}num"):
        num_id = num.get(f"{{{NS}}}numId")
        abs_id_el = num.find(f"{{{NS}}}abstractNumId")
        if abs_id_el is None:
            continue
        abs_id = abs_id_el.get(f"{{{NS}}}val")
        overrides: dict[str, dict] = {}
        for ov in num.findall(f"{{{NS}}}lvlOverride"):
            ilvl = ov.get(f"{{{NS}}}ilvl")
            so = ov.find(f"{{{NS}}}startOverride")
            if so is not None:
                overrides[ilvl] = {"start": int(so.get(f"{{{NS}}}val", "1"))}
        num_map[num_id] = {"abs_id": abs_id, "overrides": overrides}
    return {"abstract_nums": abstract_nums, "num_map": num_map}


def _format_list_num(n: int, num_fmt: str) -> str:
    if num_fmt == "upperLetter":
        return chr(ord("A") + (n - 1)) if 1 <= n <= 26 else str(n)
    if num_fmt == "lowerLetter":
        return chr(ord("a") + (n - 1)) if 1 <= n <= 26 else str(n)
    if num_fmt in ("chineseCounting", "chineseCountingThousand", "ideographTraditional"):
        if 1 <= n <= 10:
            return _CHINESE_NUMS[n]
    return str(n)


def _paragraph_list_prefix(p, numbering_map: dict, counters: dict) -> str:
    """Return the auto-list prefix (e.g. '（2）') for a paragraph, or '' if none."""
    if not numbering_map:
        return ""
    NS = _WNS
    pPr = p._p.find(f"{{{NS}}}pPr")
    if pPr is None:
        return ""
    num_pr = pPr.find(f"{{{NS}}}numPr")
    if num_pr is None:
        return ""
    ilvl_el = num_pr.find(f"{{{NS}}}ilvl")
    num_id_el = num_pr.find(f"{{{NS}}}numId")
    if ilvl_el is None or num_id_el is None:
        return ""
    ilvl = ilvl_el.get(f"{{{NS}}}val", "0")
    num_id = num_id_el.get(f"{{{NS}}}val", "0")
    if num_id == "0":
        return ""
    num_map = numbering_map.get("num_map", {})
    abstract_nums = numbering_map.get("abstract_nums", {})
    if num_id not in num_map:
        return ""
    entry = num_map[num_id]
    abs_id = entry["abs_id"]
    if abs_id not in abstract_nums:
        return ""
    lvl_def = abstract_nums[abs_id].get(ilvl, {})
    start = entry.get("overrides", {}).get(ilvl, {}).get("start", lvl_def.get("start", 1))
    key = (num_id, ilvl)
    if key not in counters:
        counters[key] = start
    else:
        counters[key] += 1
    n = counters[key]
    lvl_text = lvl_def.get("lvlText", "%1")
    num_fmt = lvl_def.get("numFmt", "decimal")
    return re.sub(r"%\d+", _format_list_num(n, num_fmt), lvl_text)


def _extract_embedded_images(element, part) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    rel_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    for node in element.iter():
        if not node.tag.endswith("}blip"):
            continue
        rel_id = node.get(rel_attr)
        if not rel_id:
            continue
        rel_part = part.related_parts.get(rel_id)
        if rel_part is None:
            continue
        blob = rel_part.blob
        try:
            image = Image.open(BytesIO(blob))
            width_px, height_px = image.size
        except Exception:
            width_px, height_px = 0, 0
        blocks.append({
            "type": "image",
            "blob": blob,
            "content_type": getattr(rel_part, "content_type", ""),
            "width_px": width_px,
            "height_px": height_px,
        })
    return blocks


def paragraph_image_blocks(paragraph: Paragraph) -> list[dict[str, Any]]:
    return _extract_embedded_images(paragraph._p, paragraph.part)


def cell_image_blocks(cell) -> list[dict[str, Any]]:
    return _extract_embedded_images(cell._tc, cell.part)


def flatten_image_table(rows_with_cells: list[list[dict[str, Any]]]) -> list[Any]:
    """Convert a Word table that carries images (not data) into inline image blocks.

    When the table is the common "images on top row, labels on bottom row" shape,
    the label text is attached to its image as a ``caption`` so the pair renders
    atomically and cannot be split across pages.
    """
    if len(rows_with_cells) == 2:
        row_img, row_text = rows_with_cells
        row_img_only_images = all(c["images"] and not c["text"] for c in row_img)
        row_text_only_text = all(not c["images"] for c in row_text)
        if row_img_only_images and row_text_only_text and len(row_img) == len(row_text):
            result: list[Any] = []
            for img_cell, text_cell in zip(row_img, row_text):
                caption = text_cell["text"] or None
                for idx, img in enumerate(img_cell["images"]):
                    paired = dict(img)
                    if caption and idx == len(img_cell["images"]) - 1:
                        paired["caption"] = caption
                    result.append(paired)
            return result
    result = []
    for row in rows_with_cells:
        for cell in row:
            result.extend(cell["images"])
            if cell["text"]:
                result.append(cell["text"])
    return result




STEM_BLANK_RE = re.compile(r"[（(][\s　]*[）)][\s。，]*$")
BARE_OPTION_REJECT_RE = re.compile(
    r"^(?:[（(]\d+[）)]|\d+[.．、]|[①-⑳]|[一二三四五六七八九十百]+[、.．]|【|[（(]\d{4}|第[一二三四五六七八九十百]+)"
)
_SOURCE_IN_PARENS_RE = re.compile(r"^[（(].*\d{4}.*[）)]$")


def is_candidate_bare_option(text: str) -> bool:
    text = (text or "").strip()
    if not text or len(text) > 160:
        return False
    if BARE_OPTION_REJECT_RE.match(text):
        return False
    if STEM_BLANK_RE.search(text):
        return False
    if _SOURCE_IN_PARENS_RE.match(text):
        return False  # source citations like （改编自2024-2025...）
    return True


def infer_bare_options(blocks: list[Any]) -> list[Any]:
    """When a question stem ends with a blank "（ ）" marker, auto-label bare
    body paragraphs that follow as A./B./C./... Supports two shapes:
      - All 4 candidates are bare (assume they start at A).
      - Some are bare and one is already labeled; start letter is back-calculated
        from the labeled position so pre-existing D. maps to A/B/C/D.
    """
    option_pat = re.compile(r"^([A-F])[.．]\s*(.*)$")
    result: list[Any] = []
    i = 0
    while i < len(blocks):
        block = blocks[i]
        text = block_text(block)
        result.append(block)
        if not text or not STEM_BLANK_RE.search(text):
            i += 1
            continue
        candidates: list[tuple[str, Any]] = []
        labeled_idx = -1
        labeled_letter: str | None = None
        max_candidates = 4
        j = i + 1
        while j < len(blocks):
            nb = blocks[j]
            nt = block_text(nb).strip()
            if not nt:
                break
            if is_table_block(nb) or is_image_block(nb):
                break
            m = option_pat.match(nt)
            if m:
                letter = m.group(1)
                if letter in ("E", "F"):
                    max_candidates = max(max_candidates, ord(letter) - ord("A") + 1)
                if len(candidates) >= max_candidates:
                    break
                if labeled_idx == -1:
                    labeled_idx = len(candidates)
                    labeled_letter = letter
                candidates.append(("labeled", nb))
                j += 1
                continue
            if is_candidate_bare_option(nt):
                if len(candidates) >= max_candidates:
                    break
                candidates.append(("bare", nb))
                j += 1
                continue
            break
        has_bare = any(kind == "bare" for kind, _ in candidates)
        if not has_bare:
            i += 1
            continue
        if labeled_letter:
            start_letter = chr(ord(labeled_letter) - labeled_idx)
        elif len(candidates) >= 3:
            start_letter = "A"
        else:
            i += 1
            continue
        if ord(start_letter) < ord("A") or ord(start_letter) + len(candidates) - 1 > ord("F"):
            i += 1
            continue
        for k, (kind, b) in enumerate(candidates):
            if kind == "labeled":
                result.append(b)
            else:
                letter = chr(ord(start_letter) + k)
                result.append(f"{letter}. {block_text(b).strip()}")
        i = j
    return result


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


_LIST_PREFIX_RE = re.compile(r"^(?:[（(]\d+[）)]|[①-⑳]|\d+[.．、])")
# Source citations like （改编自2025年...） or （选自...） should never get list prefixes
_SOURCE_CITATION_RE = re.compile(r"^（[^）]*(?:\d{4}|改编自|改写自|选自|摘自)[^）]*）")

# Map from a base font to its bold-weight equivalent for in-line bold rendering
_BOLD_FONT_MAP: dict[str, str] = {
    "FZKaiGBK": "FZYanSongZhun",
    "FZYanSongZhun": "FZYanSongCu",
    "FZYanSongZhong": "FZYanSongCu",
    "AlibabaPuHuiTiR": "AlibabaPuHuiTiB",
    "AlibabaPuHuiTiL": "AlibabaPuHuiTiB",
    "AlibabaPuHuiTiM": "AlibabaPuHuiTiB",
}

# FZKaiGBK has glyphs for ①–⑩ (U+2460–U+2469) but not ⑪–⑳ (U+246A–U+2473).
_MISSING_CIRCLED_NUMS: frozenset[int] = frozenset(range(0x2460, 0x2474))
_CIRCLED_FALLBACK_FONT: str | None = None
for _fb_path in ("/Library/Fonts/Arial Unicode.ttf", "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"):
    try:
        pdfmetrics.registerFont(TTFont("ArialUnicode", _fb_path))
        _CIRCLED_FALLBACK_FONT = "ArialUnicode"
        break
    except Exception:
        pass


def load_docx_paragraphs(docx_path: Path) -> list[Any]:
    doc = Document(docx_path)
    numbering_map = _build_numbering_map(doc)
    list_counters: dict = {}
    blocks: list[Any] = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            p = Paragraph(child, doc)
            if is_underlined_blank_paragraph(p):
                blocks.append(ANSWER_LINE_SENTINEL)
                continue
            image_blocks = paragraph_image_blocks(p)
            text = paragraph_fill_in_text(p) if _paragraph_has_fill_in_blanks(p) else clean_text(p.text)
            list_prefix = _paragraph_list_prefix(p, numbering_map, list_counters)
            list_prefix_len = 0
            if list_prefix and not _LIST_PREFIX_RE.match(text) and not _SOURCE_CITATION_RE.match(text):
                text = list_prefix + text
                list_prefix_len = len(list_prefix)
            if not text:
                blocks.extend(image_blocks)
                continue
            text = normalize_answer_parentheses(text)
            underline_ranges = paragraph_underline_ranges(p)
            bold_ranges = paragraph_bold_ranges(p)
            if list_prefix_len:
                underline_ranges = offset_inline_ranges(underline_ranges, list_prefix_len)
                bold_ranges = offset_inline_ranges(bold_ranges, list_prefix_len)
            word_align = "right" if p.alignment == 2 else None
            if underline_ranges or bold_ranges or word_align:
                block: dict[str, Any] = {"type": "paragraph", "text": text}
                if underline_ranges:
                    block["underline_ranges"] = underline_ranges
                if bold_ranges:
                    block["bold_ranges"] = bold_ranges
                if word_align:
                    block["word_align"] = word_align
                blocks.append(block)
            else:
                blocks.append(text)
            blocks.extend(image_blocks)
        elif child.tag.endswith("}tbl"):
            table = Table(child, doc)
            rows_with_cells: list[list[dict[str, Any]]] = []
            table_has_images = False
            for row in table.rows:
                row_cells: list[dict[str, Any]] = []
                for cell in row.cells:
                    images = cell_image_blocks(cell)
                    if images:
                        table_has_images = True
                    cell_text, cell_underlines = cell_text_and_underline_ranges(cell)
                    row_cells.append({"text": cell_text, "images": images, "underline_ranges": cell_underlines})
                rows_with_cells.append(row_cells)
            if not rows_with_cells:
                continue
            blocks.append({
                "type": "table",
                "rows": [[cell["text"] for cell in row] for row in rows_with_cells],
                "row_images": [[cell["images"] for cell in row] for row in rows_with_cells],
                "cell_underline_ranges": [[cell["underline_ranges"] for cell in row] for row in rows_with_cells],
            })
    # After "祝" (letter salutation), the immediately following text line is
    # the closing wish and should also be right-aligned.
    propagated: list[Any] = []
    after_zhu = False
    for blk in blocks:
        t = block_text(blk)
        if t == "祝":
            after_zhu = True
        elif after_zhu and t:
            if isinstance(blk, str):
                blk = {"type": "paragraph", "text": blk, "word_align": "right"}
            elif isinstance(blk, dict) and "word_align" not in blk:
                blk = dict(blk)
                blk["word_align"] = "right"
            after_zhu = False
        propagated.append(blk)
    # Merge label-text + N consecutive image blocks into image_row blocks.
    merged: list[Any] = []
    pi = 0
    while pi < len(propagated):
        blk = propagated[pi]
        t = block_text(blk)
        if t:
            labels = _split_image_row_labels(t)
            if labels:
                n = len(labels)
                ahead = propagated[pi + 1: pi + 1 + n]
                if len(ahead) == n and all(is_image_block(b) for b in ahead):
                    merged.append({
                        "type": "image_row",
                        "items": [{"label": lbl, "image": img} for lbl, img in zip(labels, ahead)],
                    })
                    pi += n + 1
                    continue
        merged.append(blk)
        pi += 1
    return infer_bare_options(normalize_bare_answer_items(merged))


def split_practice_and_answers(paragraphs: list[Any]) -> tuple[list[Any], list[Any]]:
    split_idx = next(
        (i for i, text in enumerate(paragraphs) if is_answer_main_title(block_text(text))),
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
    """Create a no-text, no-illustration paper background (numpy-vectorized)."""
    import numpy as np
    width = round(page_w_pt / 72 * dpi)
    height = round(page_h_pt / 72 * dpi)
    base_rgb = (255, 254, 251) if tone == "practice" else (253, 253, 252)
    seed = hash(f"{out_path.name}-{tone}") & 0xFFFF_FFFF
    rng = np.random.default_rng(seed)
    row_deltas = rng.integers(-1, 2, size=height, dtype=np.int16)
    canvas = np.empty((height, width, 3), dtype=np.int16)
    canvas[:] = base_rgb
    canvas += row_deltas[:, np.newaxis, np.newaxis]
    noise_mask = rng.random((height, width)) < 0.018
    noise_vals = rng.choice(np.array([-3, -2, 2, 3], dtype=np.int16), size=(height, width))
    canvas[noise_mask] += noise_vals[noise_mask, np.newaxis]
    np.clip(canvas, 0, 255, out=canvas)
    img = Image.fromarray(canvas.astype(np.uint8), mode="RGB")
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
    template_dir: Path | None = None,
) -> list[dict[str, Any]]:
    ensure_dir(background_dir)
    # Template-level cache: backgrounds are identical across runs, so reuse if available
    template_cache_dir = (template_dir / "backgrounds_cache") if template_dir else None
    if template_cache_dir:
        ensure_dir(template_cache_dir)
    records: list[dict[str, Any]] = []
    for spread in template["spreads"]:
        if spread["page_count"] != 2:
            continue
        spread_index = spread["spread_index"]
        path = background_path_for_spread(background_dir, spread_index)
        generated = False
        if not path.exists():
            tone = "answer" if spread_index >= 4 else "practice"
            # Check template-level cache first
            cache_path = background_path_for_spread(template_cache_dir, spread_index) if template_cache_dir else None
            if cache_path and cache_path.exists():
                import shutil
                shutil.copy2(cache_path, path)
            else:
                make_placeholder_background(path, page_w, page_h, dpi, tone=tone)
                if cache_path:
                    import shutil
                    shutil.copy2(path, cache_path)
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


_POEM_LINE_REJECT_RE = re.compile(
    r"^(?:[1-9]\d*[.．、]|[（(]\d+[）)]|[A-F][.．]|[①-⑳]|【|（\d{4}|参考答案)"
)


def is_candidate_poem_line(text: str, layout_rules: dict[str, Any] | None = None) -> bool:
    text = text.strip()
    max_chars = (layout_rules or {}).get("poem_line_max_chars", 22)
    if not text or len(text) > max_chars:
        return False
    if _POEM_LINE_REJECT_RE.match(text):
        return False
    return True


def paragraph_style(
    text: Any,
    pos: int,
    is_answer: bool,
    fonts: dict[str, str],
    layout_rules: dict[str, Any] | None = None,
    force_reading_prompt: bool = False,
    force_poem_line: bool = False,
) -> dict[str, Any]:
    layout_rules = layout_rules or {}
    underline_ranges = text.get("underline_ranges", []) if isinstance(text, dict) else []
    bold_ranges = text.get("bold_ranges", []) if isinstance(text, dict) else []
    word_align = text.get("word_align", None) if isinstance(text, dict) else None
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
    if pos == 0 or is_lesson_title(text, layout_rules) or (is_answer and is_answer_main_title(text)):
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
    if not is_answer and text == "祝":
        return apply_style_rule({
            "kind": "letter_closing",
            "font": kai,
            "size": 14,
            "leading": 24,
            "align": "right",
            "color": config_color(layout_rules, "body_text", "#222222"),
            "space_before": 2,
            "space_after": 0,
            "first_line_indent": 0,
            "bar": False,
        }, layout_rules, "letter_closing")
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
        body_start = question_match.start(2)
        display_start = body_start
        for rng in underline_ranges:
            if rng[0] < body_start <= rng[1]:
                display_start = min(display_start, rng[0])
        display_body = text[display_start:]
        shifted_underlines = shift_inline_ranges(underline_ranges, display_start, len(display_body))
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
            "display_text": display_body,
            "badge_text": number,
            "badge_color": config_color(layout_rules, "practice_red", "#cc0000"),
            "left_indent": 36,
            "wrap_width_factor": 1.0 if shifted_underlines else 0.84,
            "underline_ranges": shifted_underlines,
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
        display_text = f"{marker} {body}"
        body_start = parenthesized_option_match.start(2)
        display_body_start = len(marker) + 1
        display_underlines = map_body_inline_ranges(
            underline_ranges, body_start, display_body_start, len(display_text)
        )
        display_bolds = map_body_inline_ranges(
            bold_ranges, body_start, display_body_start, len(display_text)
        )
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
            "display_text": display_text,
            "marker_text": marker,
            "inline_marker": True,
            "left_indent": 36,
            "wrap_width_factor": 0.84,
            "underline_ranges": display_underlines,
            "bold_ranges": display_bolds,
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
    if not is_answer and force_poem_line:
        if is_candidate_poem_line(text, layout_rules):
            return _poem_line_style(kai, layout_rules)
        # Dynasty+author line like "【唐】王维" — not a poem line but keeps poem context
        if re.match(r"^【[^】]{1,4}】.{1,8}$", text):
            return apply_style_rule({
                "kind": "author",
                "font": kai,
                "size": 13,
                "leading": 24,
                "align": "center",
                "color": config_color(layout_rules, "body_text", "#222222"),
                "space_before": 0,
                "space_after": 6,
                "first_line_indent": 0,
                "bar": False,
            }, layout_rules, "author")
    _align = "right" if word_align == "right" else ("left" if is_answer else "justify")
    style = apply_style_rule({
        "kind": "answer_body" if is_answer else "body",
        "font": kai,
        "size": 14,
        "leading": 23 if is_answer else 24,
        "align": _align,
        "color": config_color(layout_rules, "body_text", "#222222"),
        "space_before": 1,
        "space_after": 3 if is_answer else 2,
        "first_line_indent": 0 if word_align == "right" else (18 if re.match(r"^[①-⑳]", text) else (0 if is_answer else 28)),
        "underline_ranges": underline_ranges,
        "bold_ranges": bold_ranges,
        "bar": False,
    }, layout_rules, "answer_body" if is_answer else "article_body")
    return style


def _poem_line_style(kai: str, layout_rules: dict[str, Any]) -> dict[str, Any]:
    return apply_style_rule({
        "kind": "poem_line",
        "font": kai,
        "size": 14,
        "leading": 28,
        "align": "center",
        "color": config_color(layout_rules, "body_text", "#222222"),
        "space_before": 0,
        "space_after": 2,
        "first_line_indent": 0,
        "bar": False,
    }, layout_rules, "poem_line")


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
            trailing_answer_match = re.search(r"[\s ]+[（(][\s ]*$", cur)
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
    if len(lines) < 2:
        return lines
    last = lines[-1]
    BLANK_SP = r"[\s ]"
    is_isolated = bool(re.match(rf"^\s*[（(]{BLANK_SP}+[）)]\s*$", last))
    # Also handle the case where a blank was split: line ends with "   ）" (tail only)
    is_tail = bool(re.match(rf"^{BLANK_SP}+[）)]\s*$", last)) and not is_isolated
    if not is_isolated and not is_tail:
        return lines
    previous = lines[-2].rstrip()
    if not previous:
        return lines
    if is_tail:
        # Find the opening paren in previous and reunite the blank
        for k in range(len(previous) - 1, max(-1, len(previous) - 12), -1):
            if previous[k] in "（(":
                fragment = previous[k:] + last  # preserve spaces inside blank
                lines[-2] = previous[:k].rstrip()
                lines[-1] = " " + fragment.lstrip()
                if not lines[-2]:
                    lines.pop(-2)
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
    BLANK_SP = r"[\s\u2007]"  # whitespace + figure space (U+2007)
    if not last.strip() or re.match(rf"^\s*[（(]{BLANK_SP}+[）)]\s*$", last):
        return lines
    if re.search(rf"[（(]{BLANK_SP}+[）)]", last):
        return lines  # last line contains an answer blank — don't rebalance
    min_chars = style.get("min_last_line_chars", 2)
    target_w = size * min_chars
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


def _split_line_by_bold(
    line: str, line_start: int, bold_ranges: list[list[int]]
) -> list[tuple[str, bool]]:
    """Split a line string into (segment, is_bold) pairs based on character ranges."""
    line_end = line_start + len(line)
    segments: list[tuple[str, bool]] = []
    pos = 0
    for start, end in sorted(bold_ranges):
        seg_start = max(start - line_start, 0)
        seg_end = min(end - line_start, len(line))
        if seg_start >= len(line) or seg_end <= 0 or seg_start >= seg_end:
            continue
        if pos < seg_start:
            segments.append((line[pos:seg_start], False))
        segments.append((line[seg_start:seg_end], True))
        pos = seg_end
    if pos < len(line):
        segments.append((line[pos:], False))
    return segments or [(line, False)]


def _split_for_circled_fallback(text: str) -> list[tuple[str, bool]]:
    """Split text into (segment, needs_fallback) pairs at ⑪–⑳ characters."""
    segments: list[tuple[str, bool]] = []
    buf = ""
    for ch in text:
        if ord(ch) in _MISSING_CIRCLED_NUMS:
            if buf:
                segments.append((buf, False))
                buf = ""
            segments.append((ch, True))
        else:
            buf += ch
    if buf:
        segments.append((buf, False))
    return segments or [(text, False)]


def _draw_circled_number_fallback(
    c: canvas.Canvas, x: float, y: float, ch: str, font: str, size: float
) -> float:
    """Draw ⑪–⑳ using ArialUnicode fallback font."""
    char_w = pdfmetrics.stringWidth(ch, font, size)
    if _CIRCLED_FALLBACK_FONT:
        c.saveState()
        c.setFont(_CIRCLED_FALLBACK_FONT, size)
        c.drawString(x, y, ch)
        c.restoreState()
    return char_w


def _draw_justified_with_circled_fallback(
    c: canvas.Canvas,
    line: str,
    tx: float,
    y: float,
    style: dict[str, Any],
    extra: float,
    color_mode: str,
) -> None:
    """Draw a justified line char-by-char, switching to fallback font for ⑪–⑳."""
    base_font = style["font"]
    size = style["size"]
    c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
    cx = tx
    for i, ch in enumerate(line):
        if ord(ch) in _MISSING_CIRCLED_NUMS and _CIRCLED_FALLBACK_FONT:
            c.setFont(_CIRCLED_FALLBACK_FONT, size)
        else:
            c.setFont(base_font, size)
        c.drawString(cx, y, ch)
        char_w = pdfmetrics.stringWidth(ch, base_font, size)
        if i < len(line) - 1:
            cx += char_w + extra
    c.setFont(base_font, size)


def draw_line_with_format(
    c: canvas.Canvas,
    line: str,
    line_start: int,
    tx: float,
    y: float,
    style: dict[str, Any],
) -> None:
    """Draw a text line, using the bold font variant for bold_ranges segments."""
    bold_ranges = style.get("bold_ranges") or []
    has_fallback = any(ord(ch) in _MISSING_CIRCLED_NUMS for ch in line)
    if not bold_ranges and not has_fallback:
        c.drawString(tx, y, line)
        return
    base_font = style["font"]
    bold_font = _BOLD_FONT_MAP.get(base_font, base_font)
    size = style["size"]
    cx = tx
    segs = _split_line_by_bold(line, line_start, bold_ranges) if bold_ranges else [(line, False)]
    for seg, is_bold in segs:
        font = bold_font if is_bold else base_font
        c.setFont(font, size)
        for subseg, is_fallback in _split_for_circled_fallback(seg):
            if is_fallback:
                cx += _draw_circled_number_fallback(c, cx, y, subseg, font, size)
            else:
                c.drawString(cx, y, subseg)
                cx += text_width(subseg, font, size)
    c.setFont(base_font, size)


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
            and not re.search(r"_{2,}", line)
            and not style.get("underline_ranges")
            and not style.get("bold_ranges")
        )
        if should_justify:
            line_width = text_width(line, style["font"], style["size"])
            extra = max(0, (available_w - line_width) / (len(line) - 1))
            if extra > style["size"] * 0.7:
                draw_line_with_format(c, line, text_offset, tx, c._pagesize[1] - cursor, style)
                draw_underlines_for_line(c, line, text_offset, tx, cursor, style, color_mode)
                text_offset += len(line)
                first_line = False
                continue
            if _CIRCLED_FALLBACK_FONT and any(ord(ch) in _MISSING_CIRCLED_NUMS for ch in line):
                _draw_justified_with_circled_fallback(
                    c, line, tx, c._pagesize[1] - cursor, style, extra, color_mode
                )
            else:
                text_obj = c.beginText(tx, c._pagesize[1] - cursor)
                text_obj.setFont(style["font"], style["size"])
                text_obj.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
                text_obj.setCharSpace(extra)
                text_obj.textLine(line)
                c.drawText(text_obj)
                # Reset Tc to 0: character spacing set inside a BT/ET block persists in PDF
                # graphics state, causing subsequent drawString calls to render wider.
                _tc_reset = c.beginText(0, 0)
                _tc_reset.setCharSpace(0)
                c.drawText(_tc_reset)
        else:
            draw_line_with_format(c, line, text_offset, tx, c._pagesize[1] - cursor, style)
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
    y = c._pagesize[1] - cursor - 4.0
    for rng in ranges:
        start, end = rng[0], rng[1]
        n_spaces = rng[2] if len(rng) > 2 else None
        overlap_start = max(start, line_start)
        overlap_end = min(end, line_end)
        if overlap_start >= overlap_end:
            continue
        prefix = line[: overlap_start - line_start]
        segment = line[overlap_start - line_start : overlap_end - line_start]
        x1 = tx + text_width(prefix, style["font"], style["size"])
        if n_spaces is not None:
            x2 = x1 + text_width(" " * (overlap_end - overlap_start), style["font"], style["size"])
        else:
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
        "max_height_pt": None,
    }
    return apply_style_rule(style, layout_rules, "image")


def image_scaled_size(
    block: dict[str, Any],
    available_w: float,
    max_h: float | None = None,
    style_max_h: float | None = None,
) -> tuple[float, float]:
    width_px = max(1, block.get("width_px") or 1)
    height_px = max(1, block.get("height_px") or 1)
    user_scale = block.get("scale", 1.0)
    scale = available_w / width_px
    width = available_w * user_scale
    height = height_px * scale * user_scale
    cap = style_max_h if style_max_h else None
    if max_h is not None:
        cap = min(cap, max_h) if cap else max_h
    if cap is not None and height > cap:
        scale = cap / height_px
        width = width_px * scale
        height = cap
    return width, height


IMAGE_CAPTION_LEADING = 22
IMAGE_CAPTION_SPACE_BEFORE = 2


def image_caption_height(block: dict[str, Any]) -> float:
    if not block.get("caption"):
        return 0
    return IMAGE_CAPTION_SPACE_BEFORE + IMAGE_CAPTION_LEADING


def image_total_height(block: dict[str, Any], style: dict[str, Any], frame_w: float) -> float:
    available_w = frame_w - 16 - style.get("left_indent", 0)
    _, height = image_scaled_size(block, available_w, style_max_h=style.get("max_height_pt"))
    return style["space_before"] + height + image_caption_height(block) + style["space_after"]


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
    max_h = max(12, y_bottom - cursor - style["space_after"] - image_caption_height(block))
    draw_w, draw_h = image_scaled_size(block, available_w, max_h=max_h, style_max_h=style.get("max_height_pt"))
    x = frame["x"] + 8 + left_indent + (available_w - draw_w) / 2
    y = c._pagesize[1] - cursor - draw_h
    c.drawImage(ImageReader(BytesIO(block["blob"])), x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
    cursor += draw_h
    caption = block.get("caption")
    if caption:
        cursor += IMAGE_CAPTION_SPACE_BEFORE + IMAGE_CAPTION_LEADING
        caption_font = style.get("caption_font", "FZYanSongZhong")
        caption_size = style.get("caption_size", 12)
        c.setFont(caption_font, caption_size)
        c.setFillColor(color_from_hex(style.get("caption_color", "cmyk(0,0,0,0.75)"), (.3, .3, .3), "cmyk"))
        cx = frame["x"] + 8 + left_indent + available_w / 2
        c.drawCentredString(cx, c._pagesize[1] - cursor + 6, caption)
    return cursor + style["space_after"]


IMAGE_ROW_LABEL_LEADING = 20
IMAGE_ROW_LABEL_SPACE_AFTER = 4
IMAGE_ROW_COL_GAP = 8


def image_row_total_height(block: dict[str, Any], style: dict[str, Any], frame_w: float) -> float:
    items = block.get("items", [])
    n = max(1, len(items))
    available_w = frame_w - 16 - style.get("left_indent", 0)
    col_w = (available_w - (n - 1) * IMAGE_ROW_COL_GAP) / n
    max_img_h = 0.0
    for item in items:
        _, h = image_scaled_size(item["image"], col_w, style_max_h=style.get("max_height_pt"))
        max_img_h = max(max_img_h, h)
    return style["space_before"] + IMAGE_ROW_LABEL_LEADING + IMAGE_ROW_LABEL_SPACE_AFTER + max_img_h + style["space_after"]


def draw_image_row(
    c: canvas.Canvas,
    block: dict[str, Any],
    style: dict[str, Any],
    frame: dict[str, float],
    start_cursor: float,
) -> float:
    items = block.get("items", [])
    n = max(1, len(items))
    left_indent = style.get("left_indent", 0)
    available_w = frame["w"] - 16 - left_indent
    col_w = (available_w - (n - 1) * IMAGE_ROW_COL_GAP) / n
    cursor = start_cursor + style["space_before"]
    x0 = frame["x"] + 8 + left_indent
    label_font = style.get("caption_font", "FZYanSongZhong")
    label_size = style.get("caption_size", 12)
    c.setFont(label_font, label_size)
    c.setFillColor(color_from_hex(style.get("caption_color", "cmyk(0,0,0,0.75)"), (.3, .3, .3), "cmyk"))
    for j, item in enumerate(items):
        cx = x0 + j * (col_w + IMAGE_ROW_COL_GAP) + col_w / 2
        c.drawCentredString(cx, c._pagesize[1] - cursor - label_size + 3, item["label"])
    cursor += IMAGE_ROW_LABEL_LEADING + IMAGE_ROW_LABEL_SPACE_AFTER
    max_img_h = 0.0
    for j, item in enumerate(items):
        img = item["image"]
        draw_w, draw_h = image_scaled_size(img, col_w, style_max_h=style.get("max_height_pt"))
        x = x0 + j * (col_w + IMAGE_ROW_COL_GAP) + (col_w - draw_w) / 2
        y = c._pagesize[1] - cursor - draw_h
        c.drawImage(ImageReader(BytesIO(img["blob"])), x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
        max_img_h = max(max_img_h, draw_h)
    cursor += max_img_h
    return cursor + style["space_after"]


def _cell_image_scaled_height(images: list[dict], col_content_w: float) -> float:
    """Return height of tallest image when scaled to fit col_content_w width."""
    max_h = 0.0
    for img in images:
        wpx = max(1, img.get("width_px") or 1)
        hpx = max(1, img.get("height_px") or 1)
        h = hpx * col_content_w / wpx
        max_h = max(max_h, h)
    return max_h


def _proportional_col_widths(rows: list[list[str]], table_w: float, font: str, size: float, pad: float) -> list[float]:
    """Compute column widths: equal base share + proportional to total content.

    Each column gets an equal half-share as a guaranteed floor, then the
    remaining half is distributed proportional to the sum of all line widths
    in that column. This ensures the most-content column gets the most width
    while narrow columns always stay readable.
    """
    if not rows:
        return []
    col_count = max(len(row) for row in rows)
    col_content: list[float] = []
    min_unit = text_width("一", font, size)
    for col in range(col_count):
        total_w = 0.0
        for row in rows:
            if col < len(row):
                for line in row[col].split("\n"):
                    total_w += text_width(line, font, size)
        col_content.append(max(total_w, min_unit))
    base = table_w / (2 * col_count)
    remaining = table_w - base * col_count
    total_content = sum(col_content)
    return [base + remaining * (w / total_content) for w in col_content]


def table_row_heights(
    rows: list[list[str]],
    style: dict[str, Any],
    table_w: float,
    row_images: list[list[list[dict]]] | None = None,
) -> list[float]:
    if not rows:
        return []
    col_count = max(len(row) for row in rows)
    # Proportional widths based on content
    col_widths = _proportional_col_widths(rows, table_w, style["font"], style["size"], style["cell_pad_x"])
    col_w = table_w / max(1, col_count)  # fallback equal
    col_content_w = col_w - style["cell_pad_x"] * 2  # will be overridden per-col below
    heights: list[float] = []
    for row_idx, row in enumerate(rows):
        max_cell_h = 0.0
        for col_idx in range(max(len(row), col_count)):
            cw = (col_widths[col_idx] if col_idx < len(col_widths) else col_w) - style["cell_pad_x"] * 2
            images = (row_images[row_idx][col_idx]
                      if row_images and row_idx < len(row_images) and col_idx < len(row_images[row_idx])
                      else [])
            if images:
                img_h = _cell_image_scaled_height(images, cw)
                cell_h = img_h + style["cell_pad_y"] * 2
            else:
                text = row[col_idx] if col_idx < len(row) else ""
                lines = [l for para in text.split("\n") for l in (wrap_text(para, style, cw) or [""])]
                cell_h = len(lines) * style["leading"] + style["cell_pad_y"] * 2
            max_cell_h = max(max_cell_h, cell_h)
        heights.append(max(24, max_cell_h))
    return heights


def table_total_height(rows: list[list[str]], style: dict[str, Any], table_w: float, row_images: list | None = None) -> float:
    return style["space_before"] + sum(table_row_heights(rows, style, table_w, row_images)) + style["space_after"]


def draw_table_block(
    c: canvas.Canvas,
    block: dict[str, Any],
    style: dict[str, Any],
    frame: dict[str, float],
    start_cursor: float,
    color_mode: str = "rgb",
) -> float:
    rows = table_rows(block)
    row_images: list[list[list[dict]]] = block.get("row_images", [])
    cell_underline_ranges: list[list[list[list[int]]]] = block.get("cell_underline_ranges", [])
    left_indent = style.get("left_indent", 0)
    x = frame["x"] + 8 + left_indent
    w = frame["w"] - 16 - left_indent
    cursor = start_cursor + style["space_before"]
    y_top_pdf = c._pagesize[1] - cursor
    col_count = max((len(row) for row in rows), default=1)
    col_widths = _proportional_col_widths(rows, w, style["font"], style["size"], style["cell_pad_x"])
    col_w = w / max(1, col_count)  # fallback equal width
    heights = table_row_heights(rows, style, w, row_images if row_images else None)
    c.setFont(style["font"], style["size"])
    c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
    c.setStrokeColor(color_from_hex(style["border_color"], (.47, .47, .47), color_mode))
    c.setLineWidth(0.45)
    y = y_top_pdf
    for row_index, row in enumerate(rows):
        row_h = heights[row_index]
        has_any_image = any(
            row_images[row_index][ci] if row_index < len(row_images) and ci < len(row_images[row_index]) else []
            for ci in range(col_count)
        ) if row_images and row_index < len(row_images) else False
        if row_index == 0 and not has_any_image:
            c.setFillColor(color_from_hex(style["header_fill"], (.9, .9, .9), color_mode))
            c.rect(x, y - row_h, w, row_h, fill=1, stroke=0)
            c.setFillColor(color_from_hex(style["color"], (.13, .13, .13), color_mode))
        cell_x_offset = 0.0
        for col_index in range(col_count):
            cw = col_widths[col_index] if col_index < len(col_widths) else col_w
            col_content_w = cw - style["cell_pad_x"] * 2
            cell_x = x + cell_x_offset
            cell_x_offset += cw
            c.rect(cell_x, y - row_h, cw, row_h, fill=0, stroke=1)
            images = (row_images[row_index][col_index]
                      if row_images and row_index < len(row_images) and col_index < len(row_images[row_index])
                      else [])
            if images:
                img = images[0]
                wpx = max(1, img.get("width_px") or 1)
                hpx = max(1, img.get("height_px") or 1)
                scale = col_content_w / wpx
                draw_w = col_content_w
                draw_h = hpx * scale
                # Center image in cell
                img_x = cell_x + style["cell_pad_x"] + (col_content_w - draw_w) / 2
                img_y = y - style["cell_pad_y"] - draw_h
                c.drawImage(ImageReader(BytesIO(img["blob"])), img_x, img_y,
                            width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
            else:
                text = row[col_index] if col_index < len(row) else ""
                underline_ranges = (
                    cell_underline_ranges[row_index][col_index]
                    if row_index < len(cell_underline_ranges)
                    and col_index < len(cell_underline_ranges[row_index])
                    else []
                )
                line_entries: list[tuple[str, int]] = []
                text_offset = 0
                for para in text.split("\n"):
                    para_lines = wrap_text(para, style, col_content_w) or [""]
                    para_offset = text_offset
                    for line in para_lines:
                        line_entries.append((line, para_offset))
                        para_offset += len(line)
                    text_offset += len(para) + 1
                line_y = y - style["cell_pad_y"] - style["size"]
                underline_style = dict(style)
                underline_style["underline_ranges"] = underline_ranges
                for line, line_start in line_entries:
                    if line:
                        c.drawString(cell_x + style["cell_pad_x"], line_y, line)
                        draw_underlines_for_line(
                            c,
                            line,
                            line_start,
                            cell_x + style["cell_pad_x"],
                            c._pagesize[1] - line_y,
                            underline_style,
                            color_mode,
                        )
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
    _prev_style_kind: str | None = None
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
                if is_image_row_block(text):
                    style = image_style(style_fonts, layout_rules)
                    style = apply_hanging_context_indent(style, inherited_hanging_indent, layout_rules)
                    row_h = image_row_total_height(text, style, frame["w"])
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + row_h > y_bottom and cursor > frame["y"] + 12:
                        remainder = text
                        break
                    cursor = draw_image_row(c, text, style, frame, cursor)
                    used_sides.add(side)
                    remainder = None
                    paragraph_idx += 1
                    continue
                if is_image_block(text):
                    style = image_style(style_fonts, layout_rules)
                    style = apply_hanging_context_indent(style, inherited_hanging_indent, layout_rules)
                    image_h = image_total_height(text, style, frame["w"])
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + image_h > y_bottom and cursor > frame["y"] + 12:
                        remainder = text
                        break
                    if (
                        should_keep_with_next(style, layout_rules)
                        and paragraph_idx + 1 < len(paragraphs)
                        and cursor > frame["y"] + 12
                    ):
                        next_text = paragraphs[paragraph_idx + 1]
                        if is_image_block(next_text):
                            next_h = image_total_height(next_text, style, frame["w"])
                        elif is_table_block(next_text):
                            next_style = table_style(style_fonts, layout_rules)
                            next_style = apply_hanging_context_indent(next_style, inherited_hanging_indent, layout_rules)
                            next_h = table_total_height(table_rows(next_text), next_style, frame["w"] - 16 - next_style.get("left_indent", 0))
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
                        if cursor + image_h + next_h > y_bottom:
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
                    _ri = text.get("row_images") if isinstance(text, dict) else None
                    left_indent = style.get("left_indent", 0)
                    tw = frame["w"] - 16 - left_indent
                    table_h = table_total_height(rows, style, tw, _ri)
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + table_h > y_bottom:
                        rh = table_row_heights(rows, style, tw, _ri)
                        avail = y_bottom - cursor - style["space_before"] - style["space_after"]
                        fit_count = 0
                        accum = 0.0
                        for h in rh:
                            if accum + h > avail:
                                break
                            accum += h
                            fit_count += 1
                        min_rows = 2 if len(rows) > 1 else 1
                        if fit_count >= min_rows:
                            fit_block = {"type": "table", "rows": rows[:fit_count]}
                            if _ri:
                                fit_block["row_images"] = _ri[:fit_count]
                            cursor = draw_table_block(c, fit_block, style, frame, cursor, color_mode)
                            used_sides.add(side)
                            rest_rows = [rows[0]] + rows[fit_count:] if fit_count > 0 else rows[fit_count:]
                            rest_block: dict[str, Any] = {"type": "table", "rows": rest_rows}
                            if _ri:
                                rest_block["row_images"] = [_ri[0]] + _ri[fit_count:] if fit_count > 0 else _ri[fit_count:]
                            remainder = rest_block
                            break
                        else:
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
                    previous_style_kind = _prev_style_kind
                    force_reading_prompt = (
                        not is_answer
                        and is_section_title(previous_text, layout_rules)
                        and can_be_structural_reading_prompt(block_text(text), layout_rules)
                    )
                    force_poem_line = (
                        not is_answer
                        and _prev_style_kind in ("article_title", "poem_line", "author")
                    )
                    style = paragraph_style(
                        text,
                        paragraph_idx,
                        is_answer,
                        style_fonts,
                        layout_rules,
                        force_reading_prompt=force_reading_prompt,
                        force_poem_line=force_poem_line,
                    )
                    style = apply_answer_line_context_spacing(style, previous_style_kind, layout_rules)
                    _prev_style_kind = style.get("kind")
                    if ends_hanging_context(style, layout_rules):
                        inherited_hanging_indent = None
                    if not starts_hanging_context(style, layout_rules):
                        style = apply_hanging_context_indent(style, inherited_hanging_indent, layout_rules)
                if starts_hanging_context(style, layout_rules):
                    inherited_hanging_indent = style.get("left_indent", question_content_indent(layout_rules))
                lines, rest = fit_lines(text, style, frame, cursor)
                if (
                    remainder is None
                    and should_keep_with_next(style, layout_rules)
                    and paragraph_idx + 1 < len(paragraphs)
                ):
                    next_text = paragraphs[paragraph_idx + 1]
                    if is_table_block(next_text):
                        next_style = table_style(style_fonts, layout_rules)
                        next_style = apply_hanging_context_indent(next_style, inherited_hanging_indent, layout_rules)
                        _tw = frame["w"] - 16 - next_style.get("left_indent", 0)
                        _rh = table_row_heights(table_rows(next_text), next_style, _tw)
                        _min_rows = min(2, len(_rh))
                        next_h = next_style["space_before"] + sum(_rh[:_min_rows]) + next_style["space_after"]
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
                    if style.get("kind") == "source":
                        next_h = max(next_h, 80)
                    actual_h = style.get("space_before", 0) + max(len(lines), 1) * style.get("leading", 0) + style.get("space_after", 0)
                    y_bottom = frame["y"] + frame["h"] - 6
                    if cursor + actual_h + next_h > y_bottom and cursor > frame["y"] + 12:
                        remainder = text
                        break
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
            # When all content is consumed and the last rendered block was an
            # answer line, fill the remaining frame space with extra lines so
            # the page doesn't leave a large blank gap.
            y_bottom = frame["y"] + frame["h"] - 6
            if (
                remainder is None
                and paragraph_idx >= len(paragraphs)
                and _prev_style_kind == "answer_line"
                and not is_answer
                and cursor > frame["y"] + 12
            ):
                fill_style = paragraph_style(
                    ANSWER_LINE_SENTINEL, paragraph_idx, is_answer, style_fonts, layout_rules
                )
                fill_style = apply_hanging_context_indent(fill_style, inherited_hanging_indent, layout_rules)
                line_h = fill_style["leading"] + fill_style.get("space_after", 0)
                _fill_count = 0
                while cursor + line_h <= y_bottom and _fill_count < 3:
                    lines, _ = fit_lines(ANSWER_LINE_SENTINEL, fill_style, frame, cursor)
                    if not lines:
                        break
                    cursor = draw_paragraph_lines(
                        c, lines, fill_style, frame, cursor, svg_assets, has_rest=False, color_mode=color_mode
                    )
                    used_sides.add(side)
                    _fill_count += 1
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
    page_number_start: int = 1,
) -> None:
    reader = PdfReader(str(spread_pdf_path))
    writer = PdfWriter()
    output_page_no = page_number_start
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

    background_records = ensure_backgrounds(template, background_dir, spread_page_w, page_h, args.background_dpi, template_dir=PACKAGE_ROOT)
    prompts_path = write_background_prompt_manifest(output_dir, template)

    paragraphs = load_docx_paragraphs(docx_path)

    overrides_path = Path(docx_path).parent / "overrides.json"
    if overrides_path.exists():
        import json as _json
        _overrides = _json.loads(overrides_path.read_text("utf-8"))
        for idx_str, scale in _overrides.get("image_scales", {}).items():
            idx = int(idx_str)
            if idx < len(paragraphs) and is_image_block(paragraphs[idx]):
                paragraphs[idx]["scale"] = scale

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
        page_number_start=getattr(args, "page_number_start", 1) if dynamic_page_numbers else None,
        layout_rules=layout_rules,
        allow_repeat=page_flow_rules.get("repeat_last_practice_spread_when_overflow", True),
        color_mode=color_mode,
    )
    _pn_start = getattr(args, "page_number_start", 1)
    answer_page_number_start = _pn_start + (
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
            page_number_start=_pn_start,
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
