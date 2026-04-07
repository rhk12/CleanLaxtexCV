from __future__ import annotations

import re

from core.io_utils import extract_text_between_markers


def apply(text_content: str, doc, document_text: str, verbose: bool = False) -> str:
    extracted_text = extract_text_between_markers(document_text, 'Projects, Grants, Commissions, and Contracts', 'Pending')
    latex_content = "\n\\section*{CONTRACT, FELLOWSHIPS, GRANTS AND SPONSORED RESEARCH}\n"

    award_table_index = 34
    for award_text in extracted_text.split("\n\n"):
        if award_text == 'Awarded' or not award_text.strip():
            continue

        result = {item.split(':', 1)[0].strip(): item.split(':', 1)[1].strip() for item in award_text.split("\n") if ':' in item}
        table = doc.tables[award_table_index]
        table_data: dict[str, str] = {}
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if ":" in text:
                    key, value = map(str.strip, text.split(":", 1))
                    table_data[key] = value
                else:
                    table_data["Data"] = text

        result.update(table_data)
        award_amount = re.search(r'\$\d{1,3}(,\d{1,3})*(\.\d{2})?', result.get('Award Amount', ''))
        amount = award_amount.group() if award_amount else "Amount unavailable"
        award_point = (
            f"{result.get('Principal Investigator', '')} (Principal Investigator), "
            f"``{result.get('Project Title', '')}'', Sponsored by {result.get('Agency', '')}, {amount}. "
            f"({result.get('Start Date', '')} - {result.get('End Date', '')})."
        )
        award_point = award_point.replace("$", r"\$").replace("&", r"\&").replace("#", r"\#")
        latex_content += "\n\\noindent " + award_point + r"\vspace{0.25cm}" + "\n"
        award_table_index += 1

    if verbose:
        print("[debug] Built section: grants")
    return text_content.replace("{{awarded}}", latex_content)
