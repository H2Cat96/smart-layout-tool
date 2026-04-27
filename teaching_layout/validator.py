from __future__ import annotations

import json
from pathlib import Path

from pypdf import PdfReader

from .layout_rules import FORBIDDEN_LINE_START


def validate_pdf(pdf_path: Path, manifest_path: Path | None = None) -> dict[str, object]:
    reader = PdfReader(str(pdf_path))
    bad_line_starts: list[dict[str, object]] = []
    fonts: set[str] = set()
    fitz_available = True
    try:
        import fitz  # type: ignore[import-not-found]
    except ModuleNotFoundError:
        fitz_available = False
    if fitz_available:
        document = fitz.open(str(pdf_path))
        for page_number, page in enumerate(document, 1):
            for font in page.get_fonts(full=True):
                fonts.add(font[3])
            for block in page.get_text("dict").get("blocks", []):
                for line in block.get("lines", []):
                    text = "".join(span.get("text", "") for span in line.get("spans", [])).strip()
                    if text and text[0] in FORBIDDEN_LINE_START:
                        bad_line_starts.append({"page": page_number, "text": text[:80]})

    result: dict[str, object] = {
        "pages": len(reader.pages),
        "fonts": sorted(fonts),
        "bad_line_starts": bad_line_starts,
        "text_inspection": "fitz" if fitz_available else "unavailable",
    }
    if manifest_path and manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        result["practice"] = manifest.get("practice_result")
        result["answers"] = manifest.get("answer_result")
        result["background_mode"] = manifest.get("background_mode")
    return result
