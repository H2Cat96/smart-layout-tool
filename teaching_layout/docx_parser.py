from __future__ import annotations

import re
from pathlib import Path
from typing import Any

ANSWER_LINE_SENTINEL = "\uE000_ANSWER_LINE"


def clean_text(text: str) -> str:
    """Normalize hard breaks while preserving meaningful spaces inside text."""
    return re.sub(r"[\r\n\t]+", " ", text).strip()


def is_underlined_blank_paragraph(paragraph: Any) -> bool:
    """Return True for Word paragraphs used as answer lines.

    The source documents express answer lines as runs containing only spaces
    with a `w:u` underline property. Treating those as empty text would drop
    printable answer lines, so the parser preserves them as a sentinel token.
    """
    text = "".join(run.text for run in paragraph.runs)
    if text.strip():
        return False
    return any("<w:u" in getattr(run._element, "xml", "") for run in paragraph.runs)


def load_docx_paragraphs(docx_path: Path) -> list[str]:
    from docx import Document

    doc = Document(str(docx_path))
    paragraphs: list[str] = []
    for paragraph in doc.paragraphs:
        if is_underlined_blank_paragraph(paragraph):
            paragraphs.append(ANSWER_LINE_SENTINEL)
            continue
        text = clean_text(paragraph.text)
        if text:
            paragraphs.append(text)
    return paragraphs


def split_practice_and_answers(paragraphs: list[str]) -> tuple[list[str], list[str]]:
    split_idx = next(
        (i for i, text in enumerate(paragraphs) if "参考答案" in text),
        len(paragraphs),
    )
    return paragraphs[:split_idx], paragraphs[split_idx:]


def summarize_docx(docx_path: Path) -> dict[str, int]:
    paragraphs = load_docx_paragraphs(docx_path)
    practice, answers = split_practice_and_answers(paragraphs)
    return {
        "paragraphs": len(paragraphs),
        "practice": len(practice),
        "answers": len(answers),
        "answer_lines": sum(1 for p in paragraphs if p == ANSWER_LINE_SENTINEL),
    }
