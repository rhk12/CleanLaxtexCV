from __future__ import annotations

import re
from typing import List, Tuple

from core.io_utils import extract_text_between_markers
from core.latex_utils import right_date_line


def _parse_award_entry(block: str) -> tuple[str, str] | None:
    text = re.sub(r"\s+", " ", block).strip()
    if not text:
        return None

    parts = [p.strip() for p in text.split(".") if p.strip()]
    if not parts:
        return None

    title = parts[0]
    year_match = re.search(r"\b(20\d{2}|19\d{2})\b", text)
    year = year_match.group(1) if year_match else ""

    return title, year


def _filter_awards(entries: List[Tuple[str, str]], mode: str) -> List[Tuple[str, str]]:
    if mode == "full":
        return entries
    elif mode == "three-page":
        return entries[:10]
    else:
        return entries

def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("awards_honors.apply requires document_text as a string.")

    markers = [
        # Research/scholarship awards
        (
            "Honors or Awards for Scholarship, Research, or Creative Activities\n\nScholarship/Research",
            "Technology Transferred or Adapted in the Field",
        ),
        # Teaching awards
        (
            "Honors or Awards for Excellence in Teaching and Advising\n\nTeaching",
            "Supervision of, and Membership on, Graduate and Undergraduate Dissertations, Theses, Projects, Monographs, Performances, Productions, and Exhibitions Required for Degrees; Types of Degrees and Years Granted",
        ),
        # Professional/service awards
        (
            "Service, Professional\n\n",
            "EXTERNAL LETTERS OF ASSESSMENT",
        ),
    ]

    blocks: list[str] = []
    for start_marker, end_marker in markers:
        extracted = extract_text_between_markers(document_text, start_marker, end_marker)
        if extracted:
            blocks.extend([b.strip() for b in extracted.split("\n\n") if b.strip()])

    parsed = []
    seen = set()

    for block in blocks:
        parsed_entry = _parse_award_entry(block)
        if not parsed_entry:
            continue
        title, year = parsed_entry
        key = (title.lower(), year)
        if key in seen:
            continue
        seen.add(key)
        parsed.append((title, year))

    parsed.sort(key=lambda x: x[1], reverse=True)
    parsed = _filter_awards(parsed, mode)

    latex = "\n\\section*{AWARDS AND HONORS}\n\n"
    for title, year in parsed:
        latex += right_date_line(title, year)
    latex += "\n"

    return text_content.replace("{{awards_and_honor}}", latex)