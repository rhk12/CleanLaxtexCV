from __future__ import annotations

import re

from core.io_utils import extract_text_between_markers
from core.latex_utils import tight_note


def _clean_block(text: str) -> str:
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _extract_application_filed(text: str) -> str:
    match = re.search(
        r"\(application:\s*(?P<date>\d{4}|[A-Za-z]+\s+\d{4})\)",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return ""

    year_text = match.group("date").strip()
    if re.fullmatch(r"\d{4}", year_text):
        return f"Application Filed {year_text}"
    return f"Application Filed {year_text}"


def _extract_author_and_title(block: str) -> str:
    author_match = re.match(r"^(Kraft,\s*R\.\s*H\.)", block, flags=re.IGNORECASE)
    author = "Kraft, R. H."
    if author_match:
        block = block[author_match.end():].strip()

    quote_match = re.search(r'"([^"]+)"', block)
    if not quote_match:
        return _clean_block(f"{author} {block}")

    title = quote_match.group(1).strip()
    if not title.endswith(".") and not title.endswith(","):
        title += "."

    app_text = _extract_application_filed(block)
    if app_text:
        return f'{author} "{title}" {app_text}.'

    return f'{author} "{title}".'


def apply(text_content: str, doc, document_text: str | None = None, mode: str = "full") -> str:
    if not isinstance(document_text, str):
        raise TypeError("intellectual_property.apply requires document_text as a string.")

    extracted_text = extract_text_between_markers(
        document_text,
        "Patent Intellectual Property",
        "Impact in Society of Research Scholarship and Creative Accomplishment",
    )

    latex = "\n\\section*{INTELLECTUAL PROPERTY}\n\n"

    if extracted_text:
        blocks = [b.strip() for b in extracted_text.split("\n\n") if b.strip()]
        entries: list[str] = []

        for block in blocks:
            block = _clean_block(block)
            if not block:
                continue

            if not block.lower().startswith("kraft, r. h."):
                continue

            entry = _extract_author_and_title(block)
            if entry:
                entries.append(entry)

        if mode == "three-page":
            entries = entries[:5]

        for entry in entries:
            latex += tight_note(entry)

    latex += "\n"

    return text_content.replace("{{Intellectual}}", latex)
