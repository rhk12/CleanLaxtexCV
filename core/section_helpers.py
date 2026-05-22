from __future__ import annotations

import re
from typing import Iterable


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def extract_clean_entries(block_text: str, noise_entries: Iterable[str] = ()) -> list[str]:
    if not block_text:
        return []

    noise = {re.sub(r"\s+", " ", entry).strip().lower() for entry in noise_entries}
    entries: list[str] = []

    for block in block_text.split("\n\n"):
        block = block.strip()
        if not block:
            continue

        lines = [_clean(line) for line in block.splitlines() if _clean(line)]
        if not lines:
            continue

        if len(lines) == 1 and lines[0].lower() in noise:
            continue

        if all(line.lower() in noise for line in lines):
            continue

        entry = _clean(" ".join(lines))
        if entry.lower() in noise:
            continue

        if len(entry.split()) <= 2 and "." not in entry and "," not in entry:
            continue

        entries.append(entry)

    return entries
