from __future__ import annotations

import os
import re
from pathlib import Path

from core.latex_utils import latex_escape


PLACEHOLDER = "{{highlights}}"


def _load_bullets(path: Path) -> list[str]:
    if not path.exists():
        return []
    bullets: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        if cleaned.startswith("#"):
            continue
        if cleaned.casefold() == "highlights of research accomplishments and impact".casefold():
            continue
        bullets.append(cleaned)
    return bullets


def _latex_escape_preserve_spaces(text: str | None) -> str:
    if text is None:
        return ""

    escaped = str(text)
    escaped = escaped.replace("\u2013", "--")
    escaped = escaped.replace("\u2014", "---")
    escaped = escaped.replace("\u2019", "'")
    escaped = escaped.replace("\u2018", "'")
    escaped = escaped.replace("\u201c", '"')
    escaped = escaped.replace("\u201d", '"')
    escaped = escaped.replace("\xa0", " ")

    replacements = {
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }

    for old, new in replacements.items():
        escaped = escaped.replace(old, new)

    return re.sub(r"\s+", " ", escaped)


def _render_inline_markup(text: str) -> str:
    parts = re.split(r"(\*\*.*?\*\*)", text)
    rendered: list[str] = []
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            rendered.append(rf"\textbf{{{_latex_escape_preserve_spaces(part[2:-2])}}}")
        else:
            rendered.append(_latex_escape_preserve_spaces(part))
    return "".join(rendered)


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    del doc
    del document_text
    del mode

    if PLACEHOLDER not in text_content:
        return text_content

    file_env = os.getenv("CV_HIGHLIGHTS_FILE", "").strip()
    if not file_env:
        return text_content.replace(PLACEHOLDER, "")

    bullets = _load_bullets(Path(file_env))
    if not bullets:
        return text_content.replace(PLACEHOLDER, "")

    latex = "\n\\section*{HIGHLIGHTS OF RESEARCH ACCOMPLISHMENTS AND IMPACT}\n"
    latex += "\\begin{itemize}[leftmargin=*, itemsep=0pt]\n"
    for bullet in bullets:
        latex += f"\\item {_render_inline_markup(bullet)}\n"
    latex += "\\end{itemize}\n"

    return text_content.replace(PLACEHOLDER, latex)
