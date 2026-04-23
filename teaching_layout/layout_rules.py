from __future__ import annotations

from collections.abc import Callable


FORBIDDEN_LINE_START = set("，。！？；：、,.!?;:)]}）】》〉」』”’％‰℃…")
FORBIDDEN_LINE_END = set("([{（【《〈「『“‘")


def wrap_cjk_text(text: str, max_width: float, width_fn: Callable[[str], float]) -> list[str]:
    """Wrap CJK text by width with basic kinsoku shori rules.

    `width_fn` receives a candidate string and returns its rendered width.
    The function intentionally does not know about a PDF backend, making it
    easy to test independently from ReportLab or font registration.
    """
    lines: list[str] = []
    current = ""
    for char in text:
        if not current and lines and char in FORBIDDEN_LINE_START:
            lines[-1] += char
            continue
        candidate = current + char
        if current and width_fn(candidate) > max_width:
            if char in FORBIDDEN_LINE_START:
                lines.append(candidate)
                current = ""
            elif current[-1] in FORBIDDEN_LINE_END and len(current) > 1:
                lines.append(current[:-1])
                current = current[-1] + char
            else:
                lines.append(current)
                current = char
        else:
            current = candidate
    if current:
        lines.append(current)
    return lines or [""]
