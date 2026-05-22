from __future__ import annotations

import re
from typing import List, Dict

from core.io_utils import extract_text_between_markers
from core.latex_utils import latex_escape, latex_quotes


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _extract_entries(section_text: str, split_token: str, label: str) -> List[Dict[str, str]]:
    entries: List[Dict[str, str]] = []

    for block in section_text.split("\n\n"):
        block = block.strip()
        if not block:
            continue

        first_line_match = re.search(r"^(.*?)(?:\n|$)", block, flags=re.DOTALL)
        if first_line_match:
            block = first_line_match.group(1).strip()

        if split_token not in block:
            continue

        parts = block.split(split_token, 1)
        author = _clean(parts[0].strip().rstrip(","))
        rest = parts[1].strip()

        match = re.match(r"^(.*)\s\(([^)]+)\)\.?$", rest)
        if match:
            title = match.group(1).strip()
            dates = match.group(2).strip()
        else:
            title = rest
            dates = ""

        title = re.sub(r"Date Graduated:.*$", "", title).strip()
        title = re.sub(r"\.$", "", title).strip()

        entries.append(
            {
                "author": author,
                "title": _clean(title),
                "dates": _clean(dates),
                "section": label,
                "year": _extract_year(dates),
            }
        )

    return entries


def _extract_postdoc_entries(section_text: str) -> List[Dict[str, str]]:
    entries: List[Dict[str, str]] = []

    for block in section_text.split("\n\n"):
        block = block.strip()
        if not block:
            continue

        match = re.match(r"^(.*)\s\(([^)]+)\)\.?$", block)
        if match:
            text = _clean(match.group(1).strip())
            dates = _clean(match.group(2).strip())
        else:
            text = _clean(block)
            dates = ""

        entries.append(
            {
                "author": text,
                "title": "",
                "dates": dates,
                "section": "Postdoctoral Mentorship",
                "year": _extract_year(dates),
            }
        )

    return entries


def _extract_year(text: str) -> int:
    years = re.findall(r"\b(19\d{2}|20\d{2})\b", text or "")
    if not years:
        return 0
    return max(int(y) for y in years)


def _format_entry(entry: Dict[str, str]) -> str:
    if entry["section"] == "Postdoctoral Mentorship":
        return rf"\item {latex_escape(entry['author'])}, {latex_escape(entry['dates'])}"
    return rf"\item {latex_escape(entry['author'])}, {latex_quotes(entry['title'])}, {latex_escape(entry['dates'])}"


def _limit_entries(entries: List[Dict[str, str]], max_items: int) -> List[Dict[str, str]]:
    entries = sorted(entries, key=lambda x: x["year"], reverse=True)
    return entries[:max_items]


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("directed_student_learning.apply requires document_text as a string.")

    phd_text = extract_text_between_markers(
        document_text,
        "Ph.D. Dissertation Advisor",
        "Ph.D. Dissertation Committee Member",
    )

    masters_text = extract_text_between_markers(
        document_text,
        "Master's Thesis Advisor",
        "Master's Thesis Committee Member",
    )

    postdoc_text = extract_text_between_markers(
        document_text,
        "Postdoctoral Mentorship Advisor",
        "Research Activity Advisor",
    )

    undergrad_text = extract_text_between_markers(
        document_text,
        "Undergraduate Honors Thesis Advisor",
        "THE SCHOLARSHIP OF Research and Creative Accomplishments",
    )

    phd_entries = _extract_entries(phd_text, "Ph.D.", "Ph.D. Dissertation")
    masters_entries = _extract_entries(masters_text, "MS.", "Master’s Thesis")
    postdoc_entries = _extract_postdoc_entries(postdoc_text)
    undergrad_entries = _extract_entries(undergrad_text, "Undergraduate.", "Undergraduate Honors Thesis")

    if mode == "three-page":
        phd_entries = _limit_entries(phd_entries, 4)
        masters_entries = _limit_entries(masters_entries, 3)
        postdoc_entries = _limit_entries(postdoc_entries, 2)
        undergrad_entries = _limit_entries(undergrad_entries, 3)

    latex = "\n\\section*{DIRECTED STUDENT LEARNING}\n"

    if phd_entries:
        latex += "\\subsection*{Ph.D. Dissertation}\n\\begin{enumerate}\n"
        for entry in phd_entries:
            latex += _format_entry(entry) + "\n"
        latex += "\\end{enumerate}\n\n"

    if masters_entries:
        latex += "\\subsection*{Master’s Thesis}\n\\begin{enumerate}\n"
        for entry in masters_entries:
            latex += _format_entry(entry) + "\n"
        latex += "\\end{enumerate}\n\n"

    if postdoc_entries:
        latex += "\\subsection*{Postdoctoral Mentorship}\n\\begin{enumerate}\n"
        for entry in postdoc_entries:
            latex += _format_entry(entry) + "\n"
        latex += "\\end{enumerate}\n\n"

    if undergrad_entries:
        latex += "\\subsection*{Undergraduate Honors Thesis}\n\\begin{enumerate}\n"
        for entry in undergrad_entries:
            latex += _format_entry(entry) + "\n"
        latex += "\\end{enumerate}\n\n"

    return text_content.replace("{{directedstudent}}", latex)
