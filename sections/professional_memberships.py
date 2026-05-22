from __future__ import annotations

import re
from typing import List

from core.io_utils import extract_text_between_markers
from core.section_helpers import extract_clean_entries
from core.latex_utils import tight_note


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _extract_year(text: str) -> int:
    years = re.findall(r"\b(19\d{2}|20\d{2})\b", text or "")
    if not years:
        return 0
    return max(int(y) for y in years)


def _extract_entries(block_text: str) -> List[str]:
    return extract_clean_entries(block_text, noise_entries={"National"})


def _strip_parenthesized_dates(text: str) -> str:
    cleaned = re.sub(
        r"\s*\(([^()]*(?:19\d{2}|20\d{2})[^()]*)\)\.",
        lambda m: f". {m.group(1)}.",
        text,
    )
    return re.sub(r"\.\.\s+", ". ", cleaned)


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("professional_memberships.apply requires document_text as a string.")

    professional_text = extract_text_between_markers(
        document_text,
        "Record of Membership in Professional and Learned Societies",
        "Description of New Courses and/or Programs Developed, Including Service Learning and Outreach Courses",
    )

    entries = _extract_entries(professional_text)
    entries = [_clean(e) for e in entries if e.strip()]
    entries = sorted(entries, key=_extract_year, reverse=True)

    if mode == "three-page":
        entries = entries[:5]

    latex = "\n\\section*{PROFESSIONAL MEMBERSHIPS}\n\n"
    for entry in entries:
        latex += tight_note(_strip_parenthesized_dates(entry))
    latex += "\n"

    return text_content.replace("{{professionalmembership}}", latex)
