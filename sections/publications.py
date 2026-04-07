from __future__ import annotations

import re
from typing import List, Dict

from core.io_utils import extract_text_between_markers


def _underline_students(text: str) -> str:
    pattern = r'([A-Z][a-zA-Z\-\s\.,]+) \((?:Primary Author|Author|Co-Author|Student Author)(?: -? (Graduate Student|Undergraduate Student|Postdoctoral Student))?\)'
    return re.sub(pattern, r'\\underline{\1}', text)


def _replace_special_chars(text: str) -> str:
    replacements = {
        "á": r"\'a",
        "é": r"\'e",
        "í": r"\'i",
        "ó": r"\'o",
        "ú": r"\'u",
        "ñ": r"\~n",
        "ü": r"\"u",
        "Á": r"\'A",
        "É": r"\'E",
        "Í": r"\'I",
        "Ó": r"\'O",
        "Ú": r"\'U",
        "Ñ": r"\~N",
        "Ü": r"\"U",
        "ç": r"\c{c}",
        "Ç": r"\c{C}",
        "ö": r"\"o",
        "Ö": r"\"O",
    }
    for char, repl in replacements.items():
        text = text.replace(char, repl)
    return text


def _extract_year(text: str) -> int:
    years = re.findall(r"\b(19\d{2}|20\d{2})\b", text)
    if not years:
        return 0
    return max(int(y) for y in years)


def _format_publication_entry(publication: str) -> str:
    publication = re.sub(r"\s+", " ", publication).strip()

    doi_match = re.search(r"DOI:\s*([^\s]+)", publication)
    doi_text = ""
    if doi_match:
        doi = doi_match.group(1).rstrip(".,;")
        doi_text = f" Published. \\url{{https://doi.org/{doi}}}"
        publication = re.sub(r"DOI:\s*[^\s]+", "", publication).strip()

    publication = publication.replace("&", r"\&")
    publication = _replace_special_chars(publication)
    publication = _underline_students(publication)
    publication = re.sub(r"(Kraft,)(\s*R\.\s*H\.)", r"\\textbf{\\textbf{\1}\2}", publication)
    publication = re.sub(r"^(\d+)\.\s*", "", publication)

    return rf"\item {publication}{doi_text}"


def _parse_publication_blocks(publications_text: str, category: str) -> List[Dict]:
    entries: List[Dict] = []
    skip_headers = {"Journal Article", "Refereed Conference Proceedings", "Pre-Print", "Technical Report"}

    for block in publications_text.split("\n\n"):
        block = block.strip()
        if not block or block in skip_headers:
            continue
        entries.append(
            {
                "category": category,
                "year": _extract_year(block),
                "text": block,
            }
        )
    return entries


def _filter_entries(entries: List[Dict], mode: str) -> List[Dict]:
    if mode == "full":
        return entries

    policies = {
        "journal": {"max_items": 8, "years_back": 5},
        "conference": {"max_items": 4, "years_back": 5},
        "other": {"max_items": 0, "years_back": 5},
    }

    current_year = max((e["year"] for e in entries), default=0)
    filtered: List[Dict] = []

    for category, policy in policies.items():
        subset = [e for e in entries if e["category"] == category]
        subset.sort(key=lambda x: x["year"], reverse=True)

        if policy["years_back"] > 0 and current_year > 0:
            cutoff = current_year - policy["years_back"] + 1
            subset = [e for e in subset if e["year"] >= cutoff]

        subset = subset[: policy["max_items"]]
        filtered.extend(subset)

    category_order = {"journal": 0, "conference": 1, "other": 2}
    filtered.sort(key=lambda x: (category_order.get(x["category"], 99), -x["year"]))
    return filtered


def _build_enum_block(title: str, entries: List[Dict]) -> str:
    latex = rf"\subsection*{{{title}}}" + "\n\\begin{enumerate}\n"
    for entry in entries:
        latex += _format_publication_entry(entry["text"]) + "\n"
    latex += "\\end{enumerate}\n\n"
    return latex


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("publications.apply requires document_text as a string.")

    journal_publications = extract_text_between_markers(document_text, "Journal Article", "Parts of Books")
    conference_publications = extract_text_between_markers(document_text, "Refereed Conference Proceedings", "Other Works")
    preprint_publications = extract_text_between_markers(document_text, "Pre-Print", "Technical Report")
    technical_reports = extract_text_between_markers(document_text, "Technical Report", "Projects, Grants, Commissions, and Contracts")

    journal_entries = _parse_publication_blocks(journal_publications, "journal")
    conference_entries = _parse_publication_blocks(conference_publications, "conference")

    merged_other = ""
    if preprint_publications:
        merged_other += preprint_publications.strip() + "\n\n"
    if technical_reports:
        merged_other += technical_reports.strip()

    other_entries = _parse_publication_blocks(merged_other, "other")

    journal_entries = _filter_entries(journal_entries, mode)
    conference_entries = _filter_entries(conference_entries, mode)
    other_entries = _filter_entries(other_entries, mode)

    latex = "\n\\section*{PUBLICATIONS}\n"
    latex += "\\textit{Mentored student and postdoc co-authors are underlined.}\n\n"

    if journal_entries:
        latex += _build_enum_block("Journal Articles", journal_entries)
    if conference_entries:
        latex += _build_enum_block("Conference Proceedings", conference_entries)
    if other_entries:
        latex += _build_enum_block("Preprints and Technical Reports", other_entries)

    return text_content.replace("{{publications}}", latex)