from __future__ import annotations

import re

from core.latex_utils import right_date_line

EMPLOYER_KEY = "Previous Employers with City/State\nIncluding U.S. Military\n(Most Recent First)"
TITLE_KEY = "Rank or Title"
DATE_KEY = "Dates"


def _table_to_dicts(table):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    rows = []
    for row in table.rows[1:]:
        values = [cell.text.strip() for cell in row.cells]
        if any(v for v in values):
            rows.append(dict(zip(headers, values)))
    return rows


def _format_years(dates: str) -> str:
    if not dates:
        return ""
    years = re.findall(r"\b(\d{4})\b", dates)
    if not years:
        return ""
    if len(years) == 1:
        return years[0]
    start, end = years[0], years[-1]
    if "present" in dates.lower():
        return f"{start} - Present"
    return f"{start} - {end}"


def _normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _current_position_from_text(document_text: str | None):
    """
    Pull the current exact rank/title from the dossier text rather than the
    first table, because the table parsing can misalign cells and return
    'University Park' instead of the title.
    """
    if not isinstance(document_text, str):
        return None

    text = _normalize_space(document_text)

    # First try the exact dossier header structure.
    pattern = (
        r"Exact Rank and Title of Position\s+"
        r"Kraft\s+Reuben H\.\s+"
        r"(.+?)\s+"
        r"College\s+Department/Division/School\s+Location of Residence"
    )
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if match:
        title = _normalize_space(match.group(1))
        if title:
            return (f"{title}, The Pennsylvania State University, University Park, PA", "2024 - Present")

    # Fallback: look for the known current title explicitly.
    fallback_match = re.search(
        r"\b(Professor of [A-Za-z &/\-]+Engineering)\b",
        text,
        flags=re.IGNORECASE,
    )
    if fallback_match:
        title = _normalize_space(fallback_match.group(1))
        return (f"{title}, The Pennsylvania State University, University Park, PA", "2024 - Present")

    return None


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    rows = _table_to_dicts(doc.tables[3])

    categories = {
        "Academic": [],
        "Government": [],
        "Professional": [],
    }

    # Add current position first from the dossier text.
    current_position = _current_position_from_text(document_text)
    if current_position:
        categories["Academic"].append(current_position)

    for row in rows:
        employer = _normalize_space(row.get(EMPLOYER_KEY, "").replace("\n", " "))
        title = _normalize_space(row.get(TITLE_KEY, "")).replace(" (Courtesy)", "").strip()
        dates = _normalize_space(row.get(DATE_KEY, ""))

        if not employer or not title:
            continue

        line_left = f"{title}, {employer}".strip(", ")
        line_right = _format_years(dates)

        if "Pennsylvania State University" in employer:
            categories["Academic"].append((line_left, line_right))
        elif "U.S. Army Research Laboratory" in employer or "Oak Ridge" in employer:
            categories["Government"].append((line_left, line_right))
        else:
            categories["Professional"].append((line_left, line_right))

    # Deduplicate while preserving order.
    for key in categories:
        seen = set()
        deduped = []
        for left, right in categories[key]:
            token = (_normalize_space(left.lower()), right)
            if token in seen:
                continue
            seen.add(token)
            deduped.append((left, right))
        categories[key] = deduped

    latex = "\n\\section*{PROFESSIONAL POSITIONS}\n\n"
    for heading, entries in categories.items():
        if not entries:
            continue
        latex += rf"\subsection*{{{heading}}}" + "\n\n"
        for left, right in entries:
            latex += right_date_line(left, right)
        latex += "\n"

    return text_content.replace("{{positions}}", latex)