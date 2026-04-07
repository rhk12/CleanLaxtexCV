from __future__ import annotations

import re
from typing import Dict, List, Tuple

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


def _normalize_quotes(text: str) -> str:
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    return text


def _extract_entries(block_text: str) -> List[str]:
    if not block_text:
        return []
    return [b.strip() for b in block_text.split("\n\n") if b.strip()]


def _categorize_service_entries(entries: List[str]) -> Dict[str, List[Tuple[int, str]]]:
    categories: Dict[str, List[Tuple[int, str]]] = {
        "College": [],
        "Department": [],
        "University": [],
        "Profession": [],
        "Society": [],
    }

    for entry in entries:
        entry = _normalize_quotes(_clean(entry))
        year = _extract_year(entry)
        lowered = entry.lower()

        if any(
            k in lowered
            for k in [
                "department",
                "mechanical engineering",
                "promotion and tenure committee",
                "research advancement committee",
                "teaching load policy committee",
                "strategic plan",
                "facilities committee",
                "faculty search committee",
            ]
        ):
            categories["Department"].append((year, entry))
        elif any(
            k in lowered
            for k in [
                "college of engineering",
                "engineering alumni",
                "college representative",
                "college",
                "engineering laptop",
            ]
        ):
            categories["College"].append((year, entry))
        elif any(
            k in lowered
            for k in [
                "institute for computational",
                "graduate council",
                "university",
                "cyberscience",
                "faculty engagement",
            ]
        ):
            categories["University"].append((year, entry))
        elif any(
            k in lowered
            for k in [
                "asme",
                "imece",
                "conference",
                "symposium",
                "co-organizer",
                "co-chair",
                "technical chair",
                "vice technical chair",
                "organizing conferences",
                "track chair",
                "track co-chair",
                "profession",
            ]
        ):
            categories["Profession"].append((year, entry))
        else:
            categories["Society"].append((year, entry))

    return categories


def _render_category(title: str, entries: List[Tuple[int, str]]) -> str:
    if not entries:
        return ""
    latex = rf"\subsection*{{{title}}}" + "\n\n"
    for _, entry in entries:
        latex += spaced_note(entry)
    latex += "\n"
    return latex


def _limit_entries(entries: List[Tuple[int, str]], max_items: int) -> List[Tuple[int, str]]:
    entries = sorted(entries, key=lambda x: x[0], reverse=True)
    return entries[:max_items]


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("service.apply requires document_text as a string.")

    committee_work_text = extract_text_between_markers(
        document_text,
        "Record of Committee Work at Department, Division, School, Campus, College, and University Levels",
        "Record of Academic Leadership and Support Work (College Representative, Faculty Mentoring, Assessment Activities, etc.)",
    )

    academic_leadership_text = extract_text_between_markers(
        document_text,
        "Record of Academic Leadership and Support Work (College Representative, Faculty Mentoring, Assessment Activities, etc.)",
        "Professional Service",
    )

    professional_service_text = extract_text_between_markers(
        document_text,
        "Professional Service",
        "THE SCHOLARSHIP OF THE PROFESSIONAL",
    )

    combined_entries: List[str] = []
    combined_entries.extend(_extract_entries(committee_work_text))
    combined_entries.extend(_extract_entries(academic_leadership_text))
    combined_entries.extend(_extract_entries(professional_service_text))

    categories = _categorize_service_entries(combined_entries)

    if mode == "three-page":
        categories["College"] = _limit_entries(categories["College"], 2)
        categories["Department"] = _limit_entries(categories["Department"], 4)
        categories["University"] = _limit_entries(categories["University"], 3)
        categories["Profession"] = _limit_entries(categories["Profession"], 6)
        categories["Society"] = _limit_entries(categories["Society"], 2)
    else:
        for key in categories:
            categories[key] = sorted(categories[key], key=lambda x: x[0], reverse=True)

    latex = "\n\\section*{SERVICE}\n"
    latex += _render_category("College", categories["College"])
    latex += _render_category("Department", categories["Department"])
    latex += _render_category("University", categories["University"])
    latex += _render_category("Profession", categories["Profession"])
    latex += _render_category("Society", categories["Society"])

    return text_content.replace("{{service}}", latex)