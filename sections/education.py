from __future__ import annotations

import datetime
import re

from core.latex_utils import labeled_line, right_date_line


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
    current_year = str(datetime.datetime.now().year)

    if "present" in dates.lower():
        return f"{start} - Present"
    if end == current_year and "current" in dates.lower():
        return f"{start} - Present"
    return f"{start} - {end}"


def _institution_name(raw: str) -> str:
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) >= 2:
        return f"{parts[0]}, {parts[1]}"
    return raw.strip()


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    rows = _table_to_dicts(doc.tables[1])

    latex = "\n\\section*{EDUCATION}\n\n"

    for row in rows:
        institution = _institution_name(row.get("Name and City/State of Institution", ""))
        major = row.get("Major Subjects", "").strip()
        degree_dates = row.get("Degrees - Dates", "").strip()

        degree_parts = [p.strip() for p in degree_dates.split(",") if p.strip()]
        degree = degree_parts[0] if degree_parts else ""
        year_text = _format_years(degree_dates)

        latex += right_date_line(f"{degree}, {institution}", year_text)

        if major:
            label = "Concentration" if "post-doctoral" in degree.lower() else "Major"
            latex += labeled_line(label, major)

        latex += "\n"

    return text_content.replace("{{education}}", latex)