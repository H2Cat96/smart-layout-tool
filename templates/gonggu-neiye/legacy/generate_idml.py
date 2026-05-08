"""Generate an IDML file by injecting Word content into the template IDML.

Usage:
    python3 generate_idml.py --docx input.docx --output output.idml
"""
from __future__ import annotations

import argparse
import copy
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

# Import shared functions from the PDF generator
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from generate_print_pdf import (
    load_docx_paragraphs,
    paragraph_style,
    block_text,
    split_practice_and_answers,
    load_optional_json,
    load_shared_rules,
    _deep_merge,
    _init_inline_formatting,
    ANSWER_LINE_SENTINEL,
)

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_FONT_MAP = PACKAGE_ROOT / "extracted" / "font-map.json"
DEFAULT_LAYOUT_RULES = PACKAGE_ROOT / "layout-rules.json"

NO_CHAR_STYLE = 'CharacterStyle/$ID/[No character style]'

# Module-level font name map, initialized from layout-rules at runtime
IDML_FONT_NAMES: dict[str, str] = {
    "FZYanSongZhong": "方正颜宋简体_中",
    "FZYanSongZhun": "方正颜宋简体_准",
    "FZYanSongCu": "方正颜宋简体_粗",
    "FZKaiGBK": "方正楷体_GBK",
}

# Character-level overrides per kind (font, size, leading, color)
KIND_CHAR_OVERRIDES: dict[str, dict[str, Any]] = {
    "main_title": {"font": "FZYanSongZhong", "size": 23, "leading": 34, "color": "Color/Black"},
    "section": {"font": "FZYanSongCu", "size": 16, "leading": 24},
    "topic_heading": {"font": "FZYanSongZhun", "size": 14, "leading": 27},
    "body": {"font": "FZKaiGBK", "size": 14, "leading": 24},
    "poem_line": {"font": "FZKaiGBK", "size": 14, "leading": 24},
    "question_numbered": {"font": "FZYanSongZhun", "size": 14, "leading": 27},
    "option": {"font": "FZKaiGBK", "size": 14, "leading": 27},
    "source": {"font": "FZYanSongZhun", "size": 12, "leading": 20},
    "reading_prompt": {"font": "FZKaiGBK", "size": 14, "leading": 24},
    "article_title": {"font": "FZYanSongZhun", "size": 14, "leading": 27},
    "author": {"font": "FZKaiGBK", "size": 12, "leading": 20},
    "answer_body": {"font": "FZKaiGBK", "size": 14, "leading": 23},
    "answer_label": {"font": "FZYanSongZhun", "size": 14, "leading": 23},
    "answer_line": {"font": "FZKaiGBK", "size": 14, "leading": 23},
    "judgement_item": {"font": "FZKaiGBK", "size": 14, "leading": 27},
    "table": {"font": "FZKaiGBK", "size": 14, "leading": 18},
}


_PARAGRAPH_STYLE_MAPPING: dict[str, str] = {}
_DEFAULT_PARAGRAPH_STYLE = "ParagraphStyle/阅读文章正文"


def kind_to_paragraph_style(kind: str) -> str:
    return _PARAGRAPH_STYLE_MAPPING.get(kind, _DEFAULT_PARAGRAPH_STYLE)


def make_content_element(text: str) -> etree._Element:
    el = etree.Element("Content")
    el.text = text
    return el


def make_br_element() -> etree._Element:
    return etree.Element("Br")


def make_char_range(text: str, attrs: dict[str, str] | None = None,
                    char_overrides: dict[str, Any] | None = None) -> etree._Element:
    csr = etree.Element("CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)
    if attrs:
        for k, v in attrs.items():
            csr.set(k, v)
    if char_overrides:
        if "size" in char_overrides:
            csr.set("PointSize", str(char_overrides["size"]))
        if "color" in char_overrides:
            csr.set("FillColor", char_overrides["color"])
        props = etree.SubElement(csr, "Properties")
        if "leading" in char_overrides:
            ld = etree.SubElement(props, "Leading")
            ld.set("type", "unit")
            ld.text = str(char_overrides["leading"])
        if "font" in char_overrides:
            af = etree.SubElement(props, "AppliedFont")
            af.set("type", "string")
            af.text = IDML_FONT_NAMES.get(char_overrides["font"], char_overrides["font"])
    content = make_content_element(text)
    csr.append(content)
    return csr


def make_br_char_range(char_overrides: dict[str, Any] | None = None) -> etree._Element:
    csr = etree.Element("CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)
    if char_overrides:
        if "size" in char_overrides:
            csr.set("PointSize", str(char_overrides["size"]))
        if "color" in char_overrides:
            csr.set("FillColor", char_overrides["color"])
        props = etree.SubElement(csr, "Properties")
        if "leading" in char_overrides:
            ld = etree.SubElement(props, "Leading")
            ld.set("type", "unit")
            ld.text = str(char_overrides["leading"])
        if "font" in char_overrides:
            af = etree.SubElement(props, "AppliedFont")
            af.set("type", "string")
            af.text = IDML_FONT_NAMES.get(char_overrides["font"], char_overrides["font"])
    csr.append(make_br_element())
    return csr


def _split_underlines(text: str) -> list[tuple[str, bool]]:
    """Split text into segments, separating runs of 3+ underscores."""
    import re
    parts: list[tuple[str, bool]] = []
    last = 0
    for m in re.finditer(r'_{3,}', text):
        if m.start() > last:
            parts.append((text[last:m.start()], False))
        # Replace underscores with spaces of same length
        parts.append((" " * len(m.group()), True))
        last = m.end()
    if last < len(text):
        parts.append((text[last:], False))
    return parts if parts else [(text, False)]


def _apply_ruby_annotations(psr: etree._Element, ruby_annotations: list[dict[str, Any]]) -> None:
    """Post-process a PSR to add InDesign Ruby attributes for pinyin annotations."""
    if not ruby_annotations:
        return

    # Build a map: char_index -> annotation
    ruby_map: dict[int, dict[str, Any]] = {}
    for ann in ruby_annotations:
        start = ann["start"]
        chars = ann["chars"]
        for ci in range(len(chars)):
            ruby_map[start + ci] = ann

    # Walk through CSR nodes, track text offset, split where needed
    csr_nodes = [el for el in psr if el.tag == "CharacterStyleRange"]
    offset = 0
    for csr in csr_nodes:
        content = csr.find("Content")
        if content is None or not content.text:
            br = csr.find("Br")
            if br is None:
                offset += 0
            continue

        csr_text = content.text
        csr_start = offset
        offset += len(csr_text)

        # Check if any ruby annotation falls in this CSR
        matching_anns = []
        for ann in ruby_annotations:
            ann_start = ann["start"]
            ann_end = ann_start + len(ann["chars"])
            if ann_start < csr_start + len(csr_text) and ann_end > csr_start:
                matching_anns.append(ann)

        if not matching_anns:
            continue

        # Split this CSR into parts: before-ruby, ruby-chars, after-ruby
        # For simplicity, handle one annotation per CSR (most common case)
        parent_idx = list(psr).index(csr)
        psr.remove(csr)

        remaining_text = csr_text
        local_offset = csr_start
        insert_idx = parent_idx
        for ann in sorted(matching_anns, key=lambda a: a["start"]):
            ann_start = ann["start"]
            ann_end = ann_start + len(ann["chars"])
            # Clip to this CSR range
            clip_start = max(ann_start, local_offset) - local_offset + (csr_start - csr_start)
            rel_start = ann_start - local_offset
            rel_end = ann_end - local_offset

            if rel_start < 0:
                rel_start = 0
            if rel_end > len(remaining_text):
                rel_end = len(remaining_text)

            # Before annotation
            if rel_start > 0:
                before_csr = copy.deepcopy(csr)
                bc = before_csr.find("Content")
                bc.text = remaining_text[:rel_start]
                before_csr.attrib.pop("RubyFlag", None)
                psr.insert(insert_idx, before_csr)
                insert_idx += 1

            # Annotated chars
            ruby_csr = copy.deepcopy(csr)
            rc = ruby_csr.find("Content")
            rc.text = remaining_text[rel_start:rel_end]
            ruby_csr.set("RubyFlag", "1")
            ruby_csr.set("RubyString", ann["pinyin"])
            ruby_csr.set("RubyType", "GroupRuby")
            ruby_csr.set("RubyFontSize", "7")
            ruby_csr.set("RubyAlignment", "RubyJIS")
            ruby_csr.set("RubyOverhang", "true")
            ruby_props = ruby_csr.find("Properties")
            if ruby_props is None:
                ruby_props = etree.SubElement(ruby_csr, "Properties")
            rf = etree.SubElement(ruby_props, "RubyFont")
            rf.set("type", "string")
            rf.text = "RjPinyin"
            psr.insert(insert_idx, ruby_csr)
            insert_idx += 1

            remaining_text = remaining_text[rel_end:]
            local_offset = ann_end

        # After all annotations
        if remaining_text:
            after_csr = copy.deepcopy(csr)
            ac = after_csr.find("Content")
            ac.text = remaining_text
            after_csr.attrib.pop("RubyFlag", None)
            psr.insert(insert_idx, after_csr)


def build_paragraph_node(
    text: str,
    style: dict[str, Any],
) -> etree._Element:
    """Build a ParagraphStyleRange element for a single paragraph."""
    kind = style.get("kind", "body")
    para_style = kind_to_paragraph_style(kind)
    char_overrides = KIND_CHAR_OVERRIDES.get(kind)

    psr = etree.Element("ParagraphStyleRange")
    psr.set("AppliedParagraphStyle", para_style)
    if kind in ("article_title", "author", "poem_line"):
        psr.set("Justification", "CenterAlign")

    bold_ranges = style.get("bold_ranges") or []
    superscript_ranges = style.get("superscript_ranges") or []
    emphasis_ranges = style.get("emphasis_ranges") or []

    # Split by underline blanks first
    underline_parts = _split_underlines(text)
    has_underlines = any(is_ul for _, is_ul in underline_parts)

    if not has_underlines:
        if not bold_ranges and not superscript_ranges and not emphasis_ranges:
            csr = make_char_range(text, char_overrides=char_overrides)
            psr.append(csr)
        else:
            segments = _split_text_by_formatting(text, bold_ranges, superscript_ranges, emphasis_ranges)
            for seg_text, seg_attrs in segments:
                csr = make_char_range(seg_text, seg_attrs, char_overrides=char_overrides)
                psr.append(csr)
    else:
        offset = 0
        for seg_text, is_underline in underline_parts:
            if is_underline:
                csr = make_char_range(seg_text, {"Underline": "true"}, char_overrides=char_overrides)
                psr.append(csr)
            elif bold_ranges or superscript_ranges or emphasis_ranges:
                sub_segments = _split_text_by_formatting(
                    seg_text,
                    [[r[0] - offset, r[1] - offset] for r in bold_ranges],
                    [[r[0] - offset, r[1] - offset] for r in superscript_ranges],
                    [[r[0] - offset, r[1] - offset] for r in emphasis_ranges],
                )
                for sub_text, sub_attrs in sub_segments:
                    csr = make_char_range(sub_text, sub_attrs, char_overrides=char_overrides)
                    psr.append(csr)
            else:
                csr = make_char_range(seg_text, char_overrides=char_overrides)
                psr.append(csr)
            offset += len(seg_text)

    # Apply ruby (pinyin) annotations
    ruby_annotations = style.get("ruby_annotations") or []
    if ruby_annotations:
        _apply_ruby_annotations(psr, ruby_annotations)

    psr.append(make_br_char_range(char_overrides))
    return psr


def _split_text_by_formatting(
    text: str,
    bold_ranges: list[list[int]],
    superscript_ranges: list[list[int]],
    emphasis_ranges: list[list[int]],
) -> list[tuple[str, dict[str, str] | None]]:
    """Split text into segments with their formatting attributes."""
    if not text:
        return [("", None)]

    boundaries = set()
    boundaries.add(0)
    boundaries.add(len(text))
    for rng in bold_ranges + superscript_ranges + emphasis_ranges:
        boundaries.add(max(0, rng[0]))
        boundaries.add(min(len(text), rng[1]))

    sorted_bounds = sorted(boundaries)
    segments: list[tuple[str, dict[str, str] | None]] = []

    for i in range(len(sorted_bounds) - 1):
        start, end = sorted_bounds[i], sorted_bounds[i + 1]
        seg = text[start:end]
        if not seg:
            continue

        attrs: dict[str, str] = {}

        for rng in bold_ranges:
            if rng[0] <= start and end <= rng[1]:
                attrs["StrokeWeight"] = "0.2"
                attrs["StrokeColor"] = "Color/Black"
                break

        for rng in superscript_ranges:
            if rng[0] <= start and end <= rng[1]:
                attrs["Position"] = "Superscript"
                break

        for rng in emphasis_ranges:
            if rng[0] <= start and end <= rng[1]:
                attrs["KentenKind"] = "KentenSmallBlackCircle"
                attrs["KentenPosition"] = "BelowLeft"
                break

        segments.append((seg, attrs if attrs else None))

    return segments if segments else [("", None)]


def _extract_decoration_prototypes(template_path: Path, practice_story_id: str, answer_story_id: str) -> dict[str, etree._Element]:
    """Extract inline decoration prototypes from both practice and answer stories."""
    import zipfile as _zf
    with _zf.ZipFile(str(template_path), "r") as zf:
        practice_xml = zf.read(f"Stories/Story_{practice_story_id}.xml")
        answer_xml = zf.read(f"Stories/Story_{answer_story_id}.xml")

    prototypes: dict[str, etree._Element] = {}

    # Practice story decorations
    proot = etree.fromstring(practice_xml)
    psr_list = list(proot.iter("ParagraphStyleRange"))

    # Index 0: main_title with red Polygon
    if len(psr_list) > 0:
        for csr in psr_list[0].iter("CharacterStyleRange"):
            for child in csr:
                if child.tag == "Polygon":
                    prototypes["practice_title_polygon"] = copy.deepcopy(child)
                    break

    # Index 1: section bar (pink TextFrame with section title inside)
    if len(psr_list) > 1:
        prototypes["practice_section_bar_psr"] = copy.deepcopy(psr_list[1])
        for csr in psr_list[1].iter("CharacterStyleRange"):
            for child in csr:
                if child.tag == "TextFrame":
                    prototypes["practice_section_bar_tf"] = copy.deepcopy(child)
                    break

    # Index 7: question background TextFrame
    if len(psr_list) > 7:
        for csr in psr_list[7].iter("CharacterStyleRange"):
            for child in csr:
                if child.tag == "TextFrame":
                    prototypes["practice_question_bg"] = copy.deepcopy(child)
                    break

    # Answer story decorations
    aroot = etree.fromstring(answer_xml)
    apsr_list = list(aroot.iter("ParagraphStyleRange"))

    # Index 0: answer title with gray Polygon
    if len(apsr_list) > 0:
        for csr in apsr_list[0].iter("CharacterStyleRange"):
            for child in csr:
                if child.tag == "Polygon":
                    prototypes["answer_title_polygon"] = copy.deepcopy(child)
                    break

    # Index 1: answer section bar (gray TextFrame)
    if len(apsr_list) > 1:
        prototypes["answer_section_bar_psr"] = copy.deepcopy(apsr_list[1])
        for csr in apsr_list[1].iter("CharacterStyleRange"):
            for child in csr:
                if child.tag == "TextFrame":
                    prototypes["answer_section_bar_tf"] = copy.deepcopy(child)
                    break

    return prototypes


_DECORATION_CACHE: dict[str, etree._Element] | None = None
_DECORATION_CACHE_KEY: tuple[str, str, str] | None = None


def _get_decorations(template_path: Path, practice_story_id: str, answer_story_id: str) -> dict[str, etree._Element]:
    global _DECORATION_CACHE, _DECORATION_CACHE_KEY
    key = (str(template_path), practice_story_id, answer_story_id)
    if _DECORATION_CACHE is None or _DECORATION_CACHE_KEY != key:
        _DECORATION_CACHE = _extract_decoration_prototypes(template_path, practice_story_id, answer_story_id)
        _DECORATION_CACHE_KEY = key
    return _DECORATION_CACHE


def _clone_decoration(proto: etree._Element, id_counter: list[int]) -> etree._Element:
    """Clone a decoration element with fresh Self IDs."""
    clone = copy.deepcopy(proto)
    for el in clone.iter():
        if el.get("Self"):
            id_counter[0] += 1
            el.set("Self", f"udec{id_counter[0]:04x}")
        if el.get("ParentStory"):
            id_counter[0] += 1
            el.set("ParentStory", f"udec{id_counter[0]:04x}")
    return clone


def _build_question_badge_story(story_id: str, badge_text: str) -> bytes:
    """Build a sub-story for a question badge (white number inside red frame)."""
    NSMAP = {"idPkg": "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging"}
    root = etree.Element("{http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging}Story",
                         nsmap=NSMAP)
    root.set("DOMVersion", "19.5")

    story = etree.SubElement(root, "Story")
    story.set("Self", story_id)
    story.set("UserText", "true")
    story.set("IsEndnoteStory", "false")
    story.set("AppliedTOCStyle", "n")
    story.set("TrackChanges", "false")
    story.set("StoryTitle", "$ID/")
    story.set("AppliedNamedGrid", "n")

    sp = etree.SubElement(story, "StoryPreference")
    sp.set("OpticalMarginAlignment", "false")
    sp.set("OpticalMarginSize", "12")
    sp.set("FrameType", "TextFrameType")
    sp.set("StoryOrientation", "Horizontal")
    sp.set("StoryDirection", "LeftToRightDirection")

    ice = etree.SubElement(story, "InCopyExportOption")
    ice.set("IncludeGraphicProxies", "true")
    ice.set("IncludeAllResources", "false")

    psr = etree.SubElement(story, "ParagraphStyleRange")
    psr.set("AppliedParagraphStyle", "ParagraphStyle/出处")
    psr.set("Justification", "CenterAlign")

    csr = etree.SubElement(psr, "CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)
    csr.set("FillColor", "Color/Paper")

    props = etree.SubElement(csr, "Properties")
    af = etree.SubElement(props, "AppliedFont")
    af.set("type", "string")
    af.text = IDML_FONT_NAMES["FZYanSongZhun"]

    content = etree.SubElement(csr, "Content")
    content.text = badge_text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                          standalone=True, pretty_print=True)


def _build_section_bar_story(story_id: str, text: str) -> bytes:
    """Build a sub-story XML for a section bar TextFrame (e.g. 【练习一】)."""
    NSMAP = {"idPkg": "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging"}
    root = etree.Element("{http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging}Story",
                         nsmap=NSMAP)
    root.set("DOMVersion", "19.5")

    story = etree.SubElement(root, "Story")
    story.set("Self", story_id)
    story.set("UserText", "true")
    story.set("IsEndnoteStory", "false")
    story.set("AppliedTOCStyle", "n")
    story.set("TrackChanges", "false")
    story.set("StoryTitle", "$ID/")
    story.set("AppliedNamedGrid", "n")

    sp = etree.SubElement(story, "StoryPreference")
    sp.set("OpticalMarginAlignment", "false")
    sp.set("OpticalMarginSize", "12")
    sp.set("FrameType", "TextFrameType")
    sp.set("StoryOrientation", "Horizontal")
    sp.set("StoryDirection", "LeftToRightDirection")

    ice = etree.SubElement(story, "InCopyExportOption")
    ice.set("IncludeGraphicProxies", "true")
    ice.set("IncludeAllResources", "false")

    psr = etree.SubElement(story, "ParagraphStyleRange")
    psr.set("AppliedParagraphStyle", "ParagraphStyle/问题")
    psr.set("Justification", "CenterAlign")

    csr = etree.SubElement(psr, "CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)
    csr.set("FillColor", "Color/Black")
    csr.set("PointSize", "16")

    props = etree.SubElement(csr, "Properties")
    ld = etree.SubElement(props, "Leading")
    ld.set("type", "enumeration")
    ld.text = "Auto"
    af = etree.SubElement(props, "AppliedFont")
    af.set("type", "string")
    af.text = IDML_FONT_NAMES["FZYanSongCu"]

    content = etree.SubElement(csr, "Content")
    content.text = text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                          standalone=True, pretty_print=True)


def _make_section_bar_psr(
    tf_proto: etree._Element,
    sub_story_id: str,
    id_counter: list[int],
    is_answer: bool,
) -> etree._Element:
    """Create a PSR containing an inline section bar TextFrame pointing to a sub-story."""
    psr = etree.Element("ParagraphStyleRange")
    psr.set("AppliedParagraphStyle", "ParagraphStyle/$ID/NormalParagraphStyle")
    psr.set("SpaceBefore", "11.338582677165356")
    psr.set("SpaceAfter", "17.007874015748033")

    csr = etree.SubElement(psr, "CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)

    tf = copy.deepcopy(tf_proto)
    # Assign fresh Self ID
    id_counter[0] += 1
    tf.set("Self", f"udec{id_counter[0]:04x}")
    # Point to the new sub-story
    tf.set("ParentStory", sub_story_id)
    # Rename any nested elements with Self
    for el in tf.iter():
        if el is tf:
            continue
        if el.get("Self"):
            id_counter[0] += 1
            el.set("Self", f"udec{id_counter[0]:04x}")

    csr.append(tf)
    csr.append(make_br_element())

    return psr


def _build_table_psr(rows: list[list[str]], id_counter: list[int]) -> etree._Element:
    """Build a PSR containing an IDML Table element."""
    num_rows = len(rows)
    num_cols = max(len(r) for r in rows) if rows else 1
    total_width = 455.88  # approximate text frame width in pt
    col_width = total_width / num_cols
    row_height = 28.35  # ~10mm

    psr = etree.Element("ParagraphStyleRange")
    psr.set("AppliedParagraphStyle", "ParagraphStyle/阅读文章正文")

    csr = etree.SubElement(psr, "CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)

    id_counter[0] += 1
    table_id = f"utbl{id_counter[0]:04x}"

    table = etree.SubElement(csr, "Table")
    table.set("Self", table_id)
    table.set("HeaderRowCount", "0")
    table.set("FooterRowCount", "0")
    table.set("BodyRowCount", str(num_rows))
    table.set("ColumnCount", str(num_cols))
    table.set("TableDirection", "LeftToRightDirection")

    # Define rows
    for r in range(num_rows):
        row_el = etree.SubElement(table, "Row")
        row_el.set("Self", f"{table_id}i{r}")
        row_el.set("Name", str(r))
        row_el.set("SingleRowHeight", str(row_height))

    # Define columns
    for c in range(num_cols):
        col_el = etree.SubElement(table, "Column")
        col_el.set("Self", f"{table_id}c{c}")
        col_el.set("Name", str(c))
        col_el.set("SingleColumnWidth", str(col_width))

    # Define cells (column-major: col:row)
    for r_idx, row_data in enumerate(rows):
        for c_idx in range(num_cols):
            cell_text = row_data[c_idx] if c_idx < len(row_data) else ""
            cell = etree.SubElement(table, "Cell")
            cell.set("Self", f"{table_id}i{c_idx}x{r_idx}")
            cell.set("Name", f"{c_idx}:{r_idx}")
            cell.set("RowSpan", "1")
            cell.set("ColumnSpan", "1")
            cell.set("CellType", "TextTypeCell")
            cell.set("AppliedCellStyle", "CellStyle/$ID/[None]")
            cell.set("TopEdgeStrokeWeight", "0.5")
            cell.set("BottomEdgeStrokeWeight", "0.5")
            cell.set("LeftEdgeStrokeWeight", "0.5")
            cell.set("RightEdgeStrokeWeight", "0.5")
            cell.set("TopEdgeStrokeColor", "Color/Black")
            cell.set("BottomEdgeStrokeColor", "Color/Black")
            cell.set("LeftEdgeStrokeColor", "Color/Black")
            cell.set("RightEdgeStrokeColor", "Color/Black")

            # Cell content
            cell_psr = etree.SubElement(cell, "ParagraphStyleRange")
            cell_psr.set("AppliedParagraphStyle", "ParagraphStyle/阅读文章正文")
            cell_csr = etree.SubElement(cell_psr, "CharacterStyleRange")
            cell_csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)
            props = etree.SubElement(cell_csr, "Properties")
            af = etree.SubElement(props, "AppliedFont")
            af.set("type", "string")
            af.text = IDML_FONT_NAMES["FZKaiGBK"]
            content = etree.SubElement(cell_csr, "Content")
            content.text = cell_text or ""

    return psr


def _build_image_psr(
    img_blob: bytes, content_type: str, width_px: int, height_px: int,
    id_counter: list[int], image_files: list[tuple[str, bytes]],
    max_width_pt: float = 455.0, max_height_pt: float = 260.0,
) -> etree._Element:
    """Build a PSR containing an inline anchored image."""
    # Convert px to pt (assume 72dpi for now; actual images may differ)
    dpi = 150.0
    width_pt = width_px * 72.0 / dpi
    height_pt = height_px * 72.0 / dpi

    # Scale to fit constraints
    scale = min(1.0, max_width_pt / width_pt, max_height_pt / height_pt)
    display_w = width_pt * scale
    display_h = height_pt * scale

    # Determine file extension
    ext = "png" if "png" in content_type else "tiff" if "tiff" in content_type else "jpg"
    id_counter[0] += 1
    img_name = f"image_{id_counter[0]:04x}.{ext}"
    link_path = f"Links/{img_name}"
    image_files.append((link_path, img_blob))

    id_counter[0] += 1
    rect_id = f"uimg{id_counter[0]:04x}"
    id_counter[0] += 1
    img_id = f"uimg{id_counter[0]:04x}"
    id_counter[0] += 1
    link_id = f"uimg{id_counter[0]:04x}"

    psr = etree.Element("ParagraphStyleRange")
    psr.set("AppliedParagraphStyle", "ParagraphStyle/阅读文章正文")
    psr.set("Justification", "CenterAlign")

    csr = etree.SubElement(psr, "CharacterStyleRange")
    csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)

    rect = etree.SubElement(csr, "Rectangle")
    rect.set("Self", rect_id)
    rect.set("ContentType", "GraphicType")
    rect.set("ItemTransform", f"1 0 0 1 0 0")
    rect.set("AppliedObjectStyle", "ObjectStyle/$ID/[None]")
    rect.set("Visible", "true")

    # Path geometry defines the frame size
    props = etree.SubElement(rect, "Properties")
    pg = etree.SubElement(props, "PathGeometry")
    gpt = etree.SubElement(pg, "GeometryPathType")
    gpt.set("PathOpen", "false")
    ppa = etree.SubElement(gpt, "PathPointArray")
    for x, y in [(0, 0), (display_w, 0), (display_w, display_h), (0, display_h)]:
        pp = etree.SubElement(ppa, "PathPointType")
        pp.set("Anchor", f"{x} {y}")
        pp.set("LeftDirection", f"{x} {y}")
        pp.set("RightDirection", f"{x} {y}")

    # Image element
    image = etree.SubElement(rect, "Image")
    image.set("Self", img_id)
    image.set("ImageRenderingIntent", "UseColorSettings")
    image.set("ImageTypeName", "$ID/Importedimage")
    image.set("Space", "$ID/#Links_702F")

    img_props = etree.SubElement(image, "Properties")
    profile = etree.SubElement(img_props, "Profile")
    profile.set("type", "string")
    profile.text = "$ID/None"
    gb = etree.SubElement(img_props, "GraphicBounds")
    gb.set("Left", "0")
    gb.set("Top", "0")
    gb.set("Right", str(width_pt))
    gb.set("Bottom", str(height_pt))

    link = etree.SubElement(image, "Link")
    link.set("Self", link_id)
    link.set("LinkResourceURI", f"file:Links/{img_name}")
    link.set("StoredState", "Normal")
    link.set("LinkClassID", "35906")
    link.set("LinkResourceFormat", "$ID/Importedimage")

    # AboveLine: image on its own line above anchor point
    aos = etree.SubElement(rect, "AnchoredObjectSetting")
    aos.set("AnchoredPosition", "AboveLine")
    aos.set("HorizontalAlignment", "CenterAlign")
    aos.set("SpaceBefore", "6")
    aos.set("SpaceAfter", "6")

    # 上下型绕排
    twp = etree.SubElement(rect, "TextWrapPreference")
    twp.set("TextWrapMode", "JumpObjectTextWrap")
    twp.set("TextWrapSide", "BothSides")

    # Add Br
    br_csr = etree.SubElement(psr, "CharacterStyleRange")
    br_csr.set("AppliedCharacterStyle", NO_CHAR_STYLE)
    br_csr.append(make_br_element())

    return psr


def build_story_xml(story_id: str, blocks: list[Any], fonts: dict[str, str],
                    layout_rules: dict[str, Any], is_answer: bool,
                    template_path: Path | None = None) -> tuple[bytes, list[tuple[str, bytes]], list[tuple[str, bytes]]]:
    """Build a complete Story XML document from content blocks.

    Returns:
        Tuple of (story_xml_bytes, sub_stories list, image_files list)
    """
    NSMAP = {"idPkg": "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging"}
    root = etree.Element("{http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging}Story",
                         nsmap=NSMAP)
    root.set("DOMVersion", "19.5")

    story = etree.SubElement(root, "Story")
    story.set("Self", story_id)
    story.set("UserText", "true")
    story.set("IsEndnoteStory", "false")
    story.set("AppliedTOCStyle", "n")
    story.set("TrackChanges", "false")
    story.set("StoryTitle", "$ID/")
    story.set("AppliedNamedGrid", "n")

    sp = etree.SubElement(story, "StoryPreference")
    sp.set("OpticalMarginAlignment", "false")
    sp.set("OpticalMarginSize", "12")
    sp.set("FrameType", "TextFrameType")
    sp.set("StoryOrientation", "Horizontal")
    sp.set("StoryDirection", "LeftToRightDirection")

    ice = etree.SubElement(story, "InCopyExportOption")
    ice.set("IncludeGraphicProxies", "true")
    ice.set("IncludeAllResources", "false")

    idml_cfg = layout_rules.get("idml", {})
    _tpl = template_path or (PACKAGE_ROOT / idml_cfg.get("template_file", "assets/references/巩固内页排版样式.idml"))
    _p_sid = idml_cfg.get("practice_story_id", "u3713")
    _a_sid = idml_cfg.get("answer_story_id", "u3b07")
    decorations = _get_decorations(_tpl, _p_sid, _a_sid)
    dec_counter = [0x7000]

    # Select decoration set based on story type
    title_poly_key = "answer_title_polygon" if is_answer else "practice_title_polygon"
    section_bar_tf_key = "answer_section_bar_tf" if is_answer else "practice_section_bar_tf"
    question_bg_key = "practice_question_bg"

    sub_stories: list[tuple[str, bytes]] = []
    image_files: list[tuple[str, bytes]] = []
    sec_prefix = "uans" if is_answer else "usec"
    section_counter = [0x8000]

    def _insert_deco_before_content(csr_el: etree._Element, deco: etree._Element):
        """Insert decoration after Properties but before Content."""
        content_idx = None
        for idx, child in enumerate(csr_el):
            if child.tag == "Content":
                content_idx = idx
                break
        if content_idx is not None:
            csr_el.insert(content_idx, deco)
        else:
            csr_el.insert(0, deco)

    _prev_kind = None
    for i, block in enumerate(blocks):
        if block == ANSWER_LINE_SENTINEL:
            psr = etree.Element("ParagraphStyleRange")
            psr.set("AppliedParagraphStyle", "ParagraphStyle/阅读文章正文")
            csr = make_char_range("　" * 32, {"Underline": "true"})
            psr.append(csr)
            psr.append(make_br_char_range())
            story.append(psr)
            continue
        if isinstance(block, dict) and block.get("type") == "table":
            rows = block.get("rows", [])
            if rows:
                table_psr = _build_table_psr(rows, dec_counter)
                story.append(table_psr)
            continue
        if isinstance(block, dict) and block.get("type") == "image":
            blob = block.get("blob", b"")
            if blob:
                img_psr = _build_image_psr(
                    blob, block.get("content_type", "image/png"),
                    block.get("width_px", 100), block.get("height_px", 100),
                    dec_counter, image_files,
                )
                story.append(img_psr)
            continue

        text_val = block_text(block)
        if not text_val:
            continue

        force_poem = (
            not is_answer
            and _prev_kind in ("article_title", "poem_line", "author")
        )
        style = paragraph_style(
            block, i, is_answer, fonts,
            layout_rules=layout_rules,
            force_poem_line=force_poem,
        )

        # Always use original block's inline formatting (style may have shifted offsets)
        if isinstance(block, dict):
            if block.get("emphasis_ranges"):
                style["emphasis_ranges"] = block["emphasis_ranges"]
            if block.get("superscript_ranges"):
                style["superscript_ranges"] = block["superscript_ranges"]
            if block.get("bold_ranges"):
                style["bold_ranges"] = block["bold_ranges"]
            if block.get("ruby_annotations"):
                style["ruby_annotations"] = block["ruby_annotations"]

        kind = style.get("kind", "body")
        _prev_kind = kind

        # Inject title polygon decoration for main_title
        if kind == "main_title" and title_poly_key in decorations:
            psr = build_paragraph_node("  " + text_val, style)
            first_csr = psr.find("CharacterStyleRange")
            if first_csr is not None:
                if not first_csr.get("FillColor"):
                    first_csr.set("FillColor", "Color/Black")
                poly = _clone_decoration(decorations[title_poly_key], dec_counter)
                _insert_deco_before_content(first_csr, poly)
            story.append(psr)
            continue

        # Section title -> inline TextFrame with colored background
        if kind == "section" and section_bar_tf_key in decorations:
            section_counter[0] += 1
            sub_story_id = f"{sec_prefix}{section_counter[0]:04x}"
            sub_story_xml = _build_section_bar_story(sub_story_id, text_val)
            sub_stories.append((f"Stories/Story_{sub_story_id}.xml", sub_story_xml))
            section_psr = _make_section_bar_psr(
                decorations[section_bar_tf_key], sub_story_id, dec_counter, is_answer
            )
            story.append(section_psr)
            continue

        # Inject question background for reading_prompt or numbered questions (practice only)
        if kind in ("reading_prompt", "question_numbered") and not is_answer and question_bg_key in decorations:
            badge = style.get("badge_text", "")
            if badge:
                section_counter[0] += 1
                qbg_story_id = f"uqbg{section_counter[0]:04x}"
                qbg_story_xml = _build_question_badge_story(qbg_story_id, badge)
                sub_stories.append((f"Stories/Story_{qbg_story_id}.xml", qbg_story_xml))
                bg = copy.deepcopy(decorations[question_bg_key])
                dec_counter[0] += 1
                bg.set("Self", f"udec{dec_counter[0]:04x}")
                bg.set("ParentStory", qbg_story_id)
                for el in bg.iter():
                    if el is bg:
                        continue
                    if el.get("Self"):
                        dec_counter[0] += 1
                        el.set("Self", f"udec{dec_counter[0]:04x}")
            else:
                bg = _clone_decoration(decorations[question_bg_key], dec_counter)
            # Strip leading number prefix since badge already shows it
            display_text = text_val
            if badge:
                import re as _re
                display_text = _re.sub(r'^\d+[.．、]\s*', '', text_val)
            psr = build_paragraph_node(" " + display_text, style)
            first_csr = psr.find("CharacterStyleRange")
            if first_csr is not None:
                _insert_deco_before_content(first_csr, bg)
            story.append(psr)
            continue

        psr = build_paragraph_node(text_val, style)
        story.append(psr)

    story_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                 standalone=True, pretty_print=True)
    return story_bytes, sub_stories, image_files


def _next_id(counter: list[int]) -> str:
    counter[0] += 1
    return f"ugen{counter[0]:04x}"


def duplicate_spreads(
    zf: zipfile.ZipFile,
    spread_template_path: str,
    story_id: str,
    num_extra_spreads: int,
    id_counter: list[int],
) -> tuple[list[str], list[tuple[str, str]]]:
    """Duplicate a spread template multiple times, creating new TextFrame chains.

    Returns:
        - List of new spread file paths (for designmap.xml)
        - List of (frame_id, spread_path) for all new frames to chain
    """
    if num_extra_spreads <= 0:
        return [], []

    template_xml = zf.read(spread_template_path)
    template_root = etree.fromstring(template_xml)

    new_spread_paths: list[str] = []
    new_frames: list[tuple[str, str]] = []

    for _ in range(num_extra_spreads):
        spread_root = copy.deepcopy(template_root)
        spread_el = spread_root.find("Spread")
        new_spread_id = _next_id(id_counter)
        spread_el.set("Self", new_spread_id)

        for tf in spread_root.iter("TextFrame"):
            if tf.get("ParentStory") == story_id:
                new_frame_id = _next_id(id_counter)
                tf.set("Self", new_frame_id)
                tf.set("PreviousTextFrame", "n")
                tf.set("NextTextFrame", "n")
                new_frames.append((new_frame_id, new_spread_id))

        for el in spread_root.iter():
            if el.tag != "TextFrame" and el.get("Self"):
                tag = el.tag
                if tag in ("Spread", "Page"):
                    continue
                old_id = el.get("Self")
                if old_id and not old_id.startswith("ugen"):
                    el.set("Self", _next_id(id_counter))

        spread_path = f"Spreads/Spread_{new_spread_id}.xml"
        new_spread_paths.append(spread_path)

        xml_bytes = etree.tostring(spread_root, xml_declaration=True,
                                   encoding="UTF-8", standalone=True, pretty_print=True)
        new_spread_paths[-1] = (spread_path, xml_bytes)

    return new_spread_paths, new_frames


def chain_text_frames(
    zf_contents: dict[str, bytes],
    spread_template_path: str,
    story_id: str,
    extra_count: int,
    id_counter: list[int],
) -> dict[str, bytes]:
    """Add extra spreads and chain all TextFrames for a story."""
    if extra_count <= 0:
        return zf_contents

    template_xml = zf_contents[spread_template_path]
    template_root = etree.fromstring(template_xml)

    # Find existing frames for this story, ordered by following the chain
    existing_frames: list[str] = []
    all_story_frames: dict[str, tuple[str, str, str]] = {}  # frame_id -> (prev, next, path)
    for path, content in zf_contents.items():
        if not path.startswith("Spreads/"):
            continue
        try:
            root = etree.fromstring(content)
        except Exception:
            continue
        for tf in root.iter("TextFrame"):
            if tf.get("ParentStory") == story_id:
                fid = tf.get("Self")
                all_story_frames[fid] = (tf.get("PreviousTextFrame", "n"),
                                         tf.get("NextTextFrame", "n"), path)
    # Walk the chain from the first frame (PreviousTextFrame == "n")
    first_frame = None
    for fid, (prev, nxt, _) in all_story_frames.items():
        if prev == "n":
            first_frame = fid
            break
    if first_frame:
        cur = first_frame
        while cur and cur in all_story_frames:
            existing_frames.append((cur, all_story_frames[cur][2]))
            nxt = all_story_frames[cur][1]
            cur = nxt if nxt != "n" else None
    else:
        for fid, (_, _, path) in all_story_frames.items():
            existing_frames.append((fid, path))

    # Generate new spreads
    all_new_frames: list[tuple[str, str]] = []
    for i in range(extra_count):
        spread_root = copy.deepcopy(template_root)
        spread_el = spread_root.find("Spread")
        new_spread_id = _next_id(id_counter)
        spread_el.set("Self", new_spread_id)

        # Remove page number TextFrames (non-content-story frames)
        if spread_el is not None:
            for tf in list(spread_el.iter("TextFrame")):
                if tf.get("ParentStory") and tf.get("ParentStory") != story_id:
                    parent = tf.getparent()
                    if parent is not None:
                        parent.remove(tf)

        # Rename all elements to avoid ID conflicts
        for el in spread_root.iter():
            old_self = el.get("Self")
            if old_self and el.tag != "Spread":
                new_id = _next_id(id_counter)
                el.set("Self", new_id)
                if el.tag == "TextFrame" and el.get("ParentStory") == story_id:
                    el.set("PreviousTextFrame", "n")
                    el.set("NextTextFrame", "n")
                    all_new_frames.append((new_id, new_spread_id))

        # Rename Page elements
        for page in spread_root.iter("Page"):
            page.set("Self", _next_id(id_counter))

        spread_path = f"Spreads/Spread_{new_spread_id}.xml"
        xml_bytes = etree.tostring(spread_root, xml_declaration=True,
                                   encoding="UTF-8", standalone=True, pretty_print=True)
        zf_contents[spread_path] = xml_bytes

    # Now chain all frames: existing + new
    all_frames = [f[0] for f in existing_frames] + [f[0] for f in all_new_frames]

    # Update threading in all spread XMLs
    for path, content in list(zf_contents.items()):
        if not path.startswith("Spreads/"):
            continue
        try:
            root = etree.fromstring(content)
        except Exception:
            continue
        modified = False
        for tf in root.iter("TextFrame"):
            if tf.get("ParentStory") == story_id:
                frame_id = tf.get("Self")
                if frame_id in all_frames:
                    idx = all_frames.index(frame_id)
                    prev_id = all_frames[idx - 1] if idx > 0 else "n"
                    next_id = all_frames[idx + 1] if idx < len(all_frames) - 1 else "n"
                    tf.set("PreviousTextFrame", prev_id)
                    tf.set("NextTextFrame", next_id)
                    modified = True
        if modified:
            zf_contents[path] = etree.tostring(root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True, pretty_print=True)

    return zf_contents


def update_designmap(zf_contents: dict[str, bytes], practice_story_id: str = "u3713", answer_story_id: str = "u3b07") -> dict[str, bytes]:
    """Update designmap.xml to include all new spreads, preserving element order and PIs."""
    designmap_xml = zf_contents["designmap.xml"]

    # Parse preserving processing instructions
    parser = etree.XMLParser(remove_pis=False)
    tree = etree.parse(
        __import__("io").BytesIO(designmap_xml), parser
    )
    root = tree.getroot()

    NS = "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging"
    spread_tag = f"{{{NS}}}Spread"

    # Find insertion point: index of first existing spread ref
    existing_spread_refs = root.findall(spread_tag)
    if not existing_spread_refs:
        return zf_contents

    insert_idx = list(root).index(existing_spread_refs[0])

    # Remove all existing spread refs
    for ref in existing_spread_refs:
        root.remove(ref)

    # Order spreads: cover first, then practice, then answer
    # Derive order from TextFrame chain (which spread each frame lives in)
    practice_spreads: list[str] = []
    answer_spreads: list[str] = []
    other_spreads: list[str] = []

    for path in zf_contents:
        if not path.startswith("Spreads/"):
            continue
        try:
            sroot = etree.fromstring(zf_contents[path])
        except Exception:
            other_spreads.append(path)
            continue
        has_practice = False
        has_answer = False
        for tf in sroot.iter("TextFrame"):
            ps = tf.get("ParentStory", "")
            if ps == practice_story_id:
                has_practice = True
            elif ps == answer_story_id:
                has_answer = True
        if has_practice:
            practice_spreads.append(path)
        elif has_answer:
            answer_spreads.append(path)
        else:
            other_spreads.append(path)

    # Order spreads by following each story's TextFrame chain
    def _order_spreads_by_chain(spreads: list[str], story_id: str) -> list[str]:
        """Order spreads according to TextFrame chain for a given story."""
        frame_to_spread: dict[str, str] = {}
        frame_chain: dict[str, tuple[str, str]] = {}  # frame_id -> (prev, next)
        for path in spreads:
            try:
                sroot = etree.fromstring(zf_contents[path])
            except Exception:
                continue
            for tf in sroot.iter("TextFrame"):
                if tf.get("ParentStory") == story_id:
                    fid = tf.get("Self")
                    frame_to_spread[fid] = path
                    frame_chain[fid] = (tf.get("PreviousTextFrame", "n"),
                                        tf.get("NextTextFrame", "n"))
        # Find chain head
        first = None
        for fid, (prev, _) in frame_chain.items():
            if prev == "n":
                first = fid
                break
        if not first:
            return sorted(spreads)
        # Walk chain, collect spread order (deduplicated)
        ordered: list[str] = []
        seen: set[str] = set()
        cur = first
        while cur and cur in frame_chain:
            sp = frame_to_spread.get(cur)
            if sp and sp not in seen:
                ordered.append(sp)
                seen.add(sp)
            nxt = frame_chain[cur][1]
            cur = nxt if nxt != "n" else None
        # Append any remaining (generated spreads not yet in chain)
        for sp in spreads:
            if sp not in seen:
                ordered.append(sp)
        return ordered

    practice_spreads = _order_spreads_by_chain(practice_spreads, practice_story_id)
    answer_spreads = _order_spreads_by_chain(answer_spreads, answer_story_id)
    other_spreads.sort()

    all_spreads = other_spreads + practice_spreads + answer_spreads
    for i, spread_path in enumerate(all_spreads):
        ref = etree.Element(spread_tag)
        ref.set("src", spread_path)
        root.insert(insert_idx + i, ref)

    # Add Story refs for sub-stories (section bars) that aren't already registered
    story_tag = f"{{{NS}}}Story"
    existing_story_srcs = set()
    for ref in root.findall(story_tag):
        existing_story_srcs.add(ref.get("src", ""))
    # Find last story ref to insert after it
    story_refs = root.findall(story_tag)
    if story_refs:
        story_insert_idx = list(root).index(story_refs[-1]) + 1
    else:
        story_insert_idx = len(list(root))
    for path in sorted(zf_contents.keys()):
        if path.startswith("Stories/") and path not in existing_story_srcs:
            ref = etree.Element(story_tag)
            ref.set("src", path)
            root.insert(story_insert_idx, ref)
            story_insert_idx += 1

    zf_contents["designmap.xml"] = etree.tostring(tree, xml_declaration=True,
                                                   encoding="UTF-8", standalone=True,
                                                   pretty_print=True)
    return zf_contents


def generate_idml(docx_path: str, output_path: str, layout_rules_path: str | None = None) -> str:
    """Main entry point: generate IDML from Word document."""
    # Load rules
    shared_rules = load_shared_rules()
    template_rules = load_optional_json(layout_rules_path or str(DEFAULT_LAYOUT_RULES))
    layout_rules = _deep_merge(shared_rules, template_rules)
    _init_inline_formatting(layout_rules)

    # Extract IDML-specific config from layout rules
    idml_cfg = layout_rules.get("idml", {})
    practice_story_id = idml_cfg.get("practice_story_id", "u3713")
    answer_story_id = idml_cfg.get("answer_story_id", "u3b07")
    practice_spread_tpl = idml_cfg.get("practice_spread_template", "Spreads/Spread_u3695.xml")
    answer_spread_tpl = idml_cfg.get("answer_spread_template", "Spreads/Spread_u3a1e.xml")
    practice_master = idml_cfg.get("practice_master", "u3c8")
    template_file = PACKAGE_ROOT / idml_cfg.get("template_file", "assets/references/巩固内页排版样式.idml")

    # Update module-level mappings from config
    global IDML_FONT_NAMES, _PARAGRAPH_STYLE_MAPPING, _DEFAULT_PARAGRAPH_STYLE
    if "font_names" in idml_cfg:
        IDML_FONT_NAMES = idml_cfg["font_names"]
    if "paragraph_style_mapping" in idml_cfg:
        _PARAGRAPH_STYLE_MAPPING = idml_cfg["paragraph_style_mapping"]
    if "default_paragraph_style" in idml_cfg:
        _DEFAULT_PARAGRAPH_STYLE = idml_cfg["default_paragraph_style"]

    # Load font map for style resolution
    font_map = json.loads(DEFAULT_FONT_MAP.read_text(encoding="utf-8"))
    fonts = font_map["style_defaults"]

    # Parse Word document
    paragraphs = load_docx_paragraphs(docx_path)
    practice_blocks, answer_blocks = split_practice_and_answers(paragraphs, layout_rules)

    print(f"Practice blocks: {len(practice_blocks)}")
    print(f"Answer blocks: {len(answer_blocks)}")

    # Build Story XMLs
    practice_story_xml, practice_sub_stories, practice_images = build_story_xml(
        practice_story_id, practice_blocks, fonts, layout_rules, is_answer=False,
        template_path=template_file,
    )
    answer_story_xml, answer_sub_stories, answer_images = build_story_xml(
        answer_story_id, answer_blocks, fonts, layout_rules, is_answer=True,
        template_path=template_file,
    )

    # Read template IDML into memory
    zf_contents: dict[str, bytes] = {}
    with zipfile.ZipFile(str(template_file), "r") as zf:
        for name in zf.namelist():
            zf_contents[name] = zf.read(name)

    # Remove hardcoded page number TextFrames from all spreads
    content_stories = {practice_story_id, answer_story_id}
    for path in list(zf_contents.keys()):
        if not path.startswith("Spreads/"):
            continue
        try:
            root = etree.fromstring(zf_contents[path])
        except Exception:
            continue
        modified = False
        for tf in list(root.iter("TextFrame")):
            ps = tf.get("ParentStory", "")
            if ps and ps not in content_stories:
                parent = tf.getparent()
                if parent is not None:
                    parent.remove(tf)
                    modified = True
        if modified:
            zf_contents[path] = etree.tostring(root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True, pretty_print=True)

    # Replace story content
    zf_contents[f"Stories/Story_{practice_story_id}.xml"] = practice_story_xml
    zf_contents[f"Stories/Story_{answer_story_id}.xml"] = answer_story_xml

    # Add section bar sub-stories
    for story_path, story_xml in practice_sub_stories + answer_sub_stories:
        zf_contents[story_path] = story_xml

    # Add image files
    for img_path, img_blob in practice_images + answer_images:
        zf_contents[img_path] = img_blob

    # Estimate extra spreads needed (~25 blocks per spread, template has 2 spreads each)
    # Better to have a few extra empty pages than lose content (overset text is hidden)
    id_counter = [0x5000]
    practice_extra = max(0, len(practice_blocks) // 25 - 2)
    answer_extra = max(0, len(answer_blocks) // 25 - 2)

    zf_contents = chain_text_frames(
        zf_contents, practice_spread_tpl, practice_story_id, practice_extra, id_counter
    )
    zf_contents = chain_text_frames(
        zf_contents, answer_spread_tpl, answer_story_id, answer_extra, id_counter
    )

    # Apply master spread with decorative patterns to all document pages
    for path in list(zf_contents.keys()):
        if not path.startswith("Spreads/"):
            continue
        try:
            root = etree.fromstring(zf_contents[path])
        except Exception:
            continue
        modified = False
        for page in root.iter("Page"):
            if page.get("AppliedMaster") != practice_master:
                page.set("AppliedMaster", practice_master)
                modified = True
        if modified:
            zf_contents[path] = etree.tostring(root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True,
                                               pretty_print=True)

    # Update designmap
    zf_contents = update_designmap(zf_contents, practice_story_id, answer_story_id)

    # Normalize XML declarations from single quotes to double quotes (InDesign requirement)
    for name in list(zf_contents.keys()):
        val = zf_contents[name]
        if isinstance(val, bytes) and val.startswith(b"<?xml version='1.0'"):
            zf_contents[name] = val.replace(
                b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>",
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                1,
            )

    # Separate image files from IDML (they go to external Links/ folder)
    all_images = practice_images + answer_images
    for img_path, _ in all_images:
        zf_contents.pop(img_path, None)

    # Create output package folder
    output = Path(output_path)
    package_dir = output if output.suffix == "" else output.with_suffix("")
    package_dir.mkdir(parents=True, exist_ok=True)

    # Write IDML file
    idml_name = package_dir.name + ".idml"
    idml_path = package_dir / idml_name
    with zipfile.ZipFile(str(idml_path), "w", zipfile.ZIP_DEFLATED) as out_zf:
        # mimetype must be first and uncompressed
        if "mimetype" in zf_contents:
            out_zf.writestr("mimetype", zf_contents.pop("mimetype"), compress_type=zipfile.ZIP_STORED)
        # designmap.xml must be second
        if "designmap.xml" in zf_contents:
            out_zf.writestr("designmap.xml", zf_contents.pop("designmap.xml"))
        for name in sorted(zf_contents.keys()):
            out_zf.writestr(name, zf_contents[name])

    # Write Links folder with images (external, InDesign will show missing link)
    links_dir = package_dir / "Links"
    links_dir.mkdir(exist_ok=True)
    for img_path, img_blob in all_images:
        img_file = links_dir / Path(img_path).name
        img_file.write_bytes(img_blob)

    # Copy Document Fonts from template assets
    fonts_dir = package_dir / "Document Fonts"
    fonts_dir.mkdir(exist_ok=True)
    fonts_src = PACKAGE_ROOT / "assets" / "fonts"
    font_count = 0
    if fonts_src.is_dir():
        import shutil
        for font_file in fonts_src.iterdir():
            if font_file.suffix.lower() in (".ttf", ".otf"):
                shutil.copy2(str(font_file), str(fonts_dir / font_file.name))
                font_count += 1

    print(f"Package output: {package_dir}/")
    print(f"  IDML: {idml_path}")
    print(f"  Links: {len(all_images)} images")
    print(f"  Fonts: {font_count}")
    return str(idml_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate IDML by injecting Word content into template.")
    parser.add_argument("--docx", required=True, help="Input Word document")
    parser.add_argument("--output", required=True, help="Output IDML file path")
    parser.add_argument("--layout-rules", default=None)
    parser.add_argument("--shared-rules", default=None)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    generate_idml(args.docx, args.output, args.layout_rules)


if __name__ == "__main__":
    main()
