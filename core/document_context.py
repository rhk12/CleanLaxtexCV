from __future__ import annotations

from dataclasses import dataclass

from docx import Document


@dataclass
class DocumentContext:
    doc: Document
    document_text: str
