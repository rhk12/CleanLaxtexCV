from __future__ import annotations

import re
from typing import List

from core.io_utils import extract_text_between_markers
from core.latex_utils import spaced_note


def _clean(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace("&", r"\&")
    text = text.replace("$", r"\$")
    text = text.replace("#", r"\#")
    return text


def _extract_year(text: str) -> int:
    years = re.findall(r"\b(19\d{2}|20\d{2})\b", text or "")
    if not years:
        return 0
    return max(int(y) for y in years)


def _extract_entries(block_text: str) -> List[str]:
    if not block_text:
        return []
    return [b.strip() for b in block_text.split("\n\n") if b.strip()]


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("editorial_board.apply requires document_text as a string.")

    editorial_board_text = extract_text_between_markers(
        document_text,
        "Outreach - Editorial Responsibilities",
        "Outreach - Peer Review of Grant Proposals",
    )

    entries = _extract_entries(editorial_board_text)
    entries = [_clean(e) for e in entries if e.strip()]
    entries = sorted(entries, key=_extract_year, reverse=True)

    if mode == "three-page":
        entries = entries[:5]

    latex = "\n\\section*{EDITORIAL BOARD POSITIONS}\n\n"
    for entry in entries:
        latex += spaced_note(entry)
    latex += "\n"

    return text_content.replace("{{editorialboard}}", latex)