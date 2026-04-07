from __future__ import annotations

import re


def latex_escape(text: str | None) -> str:
    if text is None:
        return ""

    text = str(text)

    # Normalize common Word punctuation safely.
    text = text.replace("\u2013", "--")   # en dash
    text = text.replace("\u2014", "---")  # em dash
    text = text.replace("\u2019", "'")    # right apostrophe
    text = text.replace("\u2018", "'")    # left apostrophe
    text = text.replace("\u201c", '"')    # left quote
    text = text.replace("\u201d", '"')    # right quote
    text = text.replace("\xa0", " ")      # non-breaking space

    replacements = {
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text).strip()
    return text


def latex_quotes(text: str | None) -> str:
    return f"``{latex_escape(text)}''"


def right_date_line(left_text: str, right_text: str) -> str:
    left = latex_escape(left_text)
    right = latex_escape(right_text)
    return (
        rf"\noindent \parbox[t]{{0.8\linewidth}}{{\raggedright {left}}} "
        rf"\hfill \parbox[t]{{0.2\linewidth}}{{\raggedleft {right}}} \\"
        "\n"
    )


def labeled_line(label: str, value: str) -> str:
    label_esc = latex_escape(label)
    value_esc = latex_escape(value)
    return (
        rf"\noindent \parbox[t]{{0.8\linewidth}}{{\raggedright \textbf{{{label_esc}:}} {value_esc}}} \\"
        "\n"
    )


def spaced_note(text: str) -> str:
    return rf"\noindent {latex_escape(text)}\vspace{{0.25cm}}" + "\n"