from __future__ import annotations

import re

from core.io_utils import extract_text_between_markers
from core.latex_utils import spaced_note


def _clean_ip_entry(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace("&", r"\&")
    text = text.replace("$", r"\$")
    text = text.replace("#", r"\#")
    return text


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

        entries = []
        for block in blocks:
            block = _clean_ip_entry(block)
            if not block:
                continue

            quote_match = re.search(r'(.*?"[^"]*")', block)
            date_match = re.search(r'\(application:\s*(?:\d{4}|\w+\s+\d{4})\)', block)

            if quote_match:
                quote_text = quote_match.group(1).strip()
            else:
                quote_text = block

            date_text = date_match.group(0).strip() if date_match else ""
            entry = f"{quote_text} {date_text}".strip()
            entries.append(entry)

        if mode == "three-page":
            entries = entries[:5]

        for entry in entries:
            latex += spaced_note(entry)

    latex += "\n"

    return text_content.replace("{{Intellectual}}", latex)