from __future__ import annotations

import datetime
from pathlib import Path


DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "template_cv.tex"


def create_template_latex(template_path: str | Path | None = None, header_date: str | None = None, header_name: str = "REUBEN H. KRAFT") -> str:
    path = Path(template_path) if template_path else DEFAULT_TEMPLATE_PATH
    text = path.read_text(encoding="utf-8")
    if not header_date:
        header_date = datetime.datetime.now().strftime("%B %Y")
    return text.replace("{header_date}", header_date).replace("{header_name}", header_name)


def add_compact_layout(text_content: str) -> str:
    compact_commands = r"""
\usepackage[a4paper, margin=0.6in]{geometry}
\setlength{\parskip}{3pt}
\setlist[enumerate]{itemsep=2pt,topsep=2pt,leftmargin=*}
\setlist[itemize]{itemsep=2pt,topsep=2pt,leftmargin=*}
\linespread{0.97}
"""
    text_content = text_content.replace(r"\documentclass[a4paper,10pt]{article}", r"\documentclass[a4paper,11pt]{article}")
    insertion_point = text_content.find(r"\begin{document}")
    if insertion_point != -1:
        return text_content[:insertion_point] + compact_commands + text_content[insertion_point:]
    return text_content
