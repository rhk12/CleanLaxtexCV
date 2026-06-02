from __future__ import annotations

import json
import re
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


START_HEADING = "Projects, Grants, Commissions, and Contracts"
END_HEADING = "Patent Intellectual Property"
TABLE_START_INDEX = 34
SECTION_TITLE = r"\section*{CONTRACTS, FELLOWSHIPS, GRANTS, AND SPONSORED RESEARCH}"
USER_NAME = "Kraft, Reuben H."

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
CATALOG_PATH = DATA_DIR / "grants_catalog.json"
RULES_PATH = DATA_DIR / "grants_rules.json"

MONTH_PATTERN = (
    r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|"
    r"Dec(?:ember)?)"
)
DATE_RANGE_PATTERN = re.compile(
    rf"(?P<start>{MONTH_PATTERN}\s+\d{{1,2}},\s+\d{{4}}|{MONTH_PATTERN}\s+\d{{4}})"
    rf"\s*[-–—]\s*"
    rf"(?P<end>{MONTH_PATTERN}\s+\d{{1,2}},\s+\d{{4}}|{MONTH_PATTERN}\s+\d{{4}})",
    re.IGNORECASE,
)


@dataclass
class GrantRecord:
    display_name: str
    role_label: str
    title: str
    sponsor: str
    start: datetime | None = None
    end: datetime | None = None
    awarded: datetime | None = None
    submitted: datetime | None = None
    anticipated_total_usd: float | None = None
    is_awarded: bool = False
    order: int = 0
    matched_catalog: str | None = None


def extract_records(doc) -> list[GrantRecord]:
    return _build_records(doc)


def _load_json(path: Path) -> dict:
    if not path.exists():
        return {}

    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _load_rules() -> dict:
    return _load_json(RULES_PATH)


def _load_catalog() -> list[dict]:
    payload = _load_json(CATALOG_PATH)
    entries = payload.get("entries", [])
    return entries if isinstance(entries, list) else []


def _clean_text(text: str) -> str:
    cleaned = text.replace("[MP]", "")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def _normalize_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", _clean_text(text).lower()).strip()


def _parse_date(text: str) -> datetime | None:
    candidate = _clean_text(text).replace("\u2013", "-").replace("\u2014", "-")
    for fmt in ("%B %d, %Y", "%b %d, %Y", "%B %Y", "%b %Y"):
        try:
            return datetime.strptime(candidate, fmt)
        except ValueError:
            continue
    return None


def _month_year(value: datetime | None) -> str:
    if value is None:
        return ""
    return value.strftime("%B %Y")


def _parse_month_year(text: str) -> datetime | None:
    cleaned = _clean_text(text)
    for fmt in ("%B %Y", "%b %Y"):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue
    return None


def _latex_escape(text: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
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
    escaped = text
    for old, new in replacements.items():
        escaped = escaped.replace(old, new)
    return escaped


def _extract_section_paragraphs(doc) -> list[str]:
    paragraphs = [para.text for para in doc.paragraphs]
    try:
        start_index = next(idx for idx, text in enumerate(paragraphs) if text.strip() == START_HEADING)
    except StopIteration:
        return []

    end_index = next(
        (idx for idx in range(start_index + 1, len(paragraphs)) if END_HEADING in paragraphs[idx]),
        len(paragraphs),
    )
    return paragraphs[start_index + 1 : end_index]


def _split_grant_blocks(paragraphs: list[str]) -> list[tuple[str, str]]:
    blocks: list[tuple[str, str]] = []
    current: list[str] = []
    current_group = "unknown"

    for paragraph in paragraphs:
        cleaned = paragraph.strip()
        if not cleaned:
            continue

        # Penn State dossier uses group markers inside the section.
        if cleaned in {"Awarded", "Pending", "Not Funded"}:
            if current:
                blocks.append((current_group, "\n".join(current)))
                current = []
            current_group = cleaned
            continue

        if cleaned.startswith("OSP Number:"):
            if current:
                blocks.append((current_group, "\n".join(current)))
            current = [cleaned]
            continue

        if current:
            current.append(cleaned)

    if current:
        blocks.append((current_group, "\n".join(current)))

    return blocks


def _table_text(table) -> str:
    lines: list[str] = []
    for row in table.rows:
        for cell in row.cells:
            for line in cell.text.splitlines():
                cleaned = line.strip()
                if cleaned:
                    lines.append(cleaned)
    return "\n".join(lines)


def _parse_key_values(text: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in text.splitlines():
        line = line.strip()
        if ":" not in line:
            continue

        # Some table lines contain multiple key/value pairs, e.g.:
        # "Award Amount: $X Total Anticipated: $Y"
        for match in re.finditer(r"([^:]+):\s*([^:]*?)(?=(?:\s+[^:]+:\s*)|$)", line):
            key = match.group(1).strip()
            value = _clean_text(match.group(2))
            if key:
                values[key] = value
    return values


def _best_field_value(block_text: str, field_name: str) -> str:
    # Some blocks contain multiple occurrences (e.g. amendments). Prefer the most informative value.
    candidates: list[str] = []
    for line in block_text.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        if key.strip() != field_name:
            continue
        cleaned = _clean_text(value)
        if cleaned:
            candidates.append(cleaned)
    if not candidates:
        return ""
    # Prefer longer values (SBIR/STTR short labels lose the real title).
    candidates.sort(key=lambda v: (len(v), v), reverse=True)
    return candidates[0]


def _parse_usd(value: str) -> float | None:
    cleaned = _clean_text(value)
    if not cleaned:
        return None
    cleaned = cleaned.replace("$", "").replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _date_spans_from_text(text: str) -> list[tuple[datetime, datetime]]:
    spans: list[tuple[datetime, datetime]] = []
    for match in DATE_RANGE_PATTERN.finditer(text):
        start = _parse_date(match.group("start"))
        end = _parse_date(match.group("end"))
        if start and end:
            spans.append((start, end))
    return spans


def _collect_record_spans(block_text: str, table_text: str) -> list[tuple[datetime, datetime]]:
    spans: list[tuple[datetime, datetime]] = []

    table_fields = _parse_key_values(table_text)
    start_date = _parse_date(table_fields.get("Start Date", ""))
    end_date = _parse_date(table_fields.get("End Date", ""))
    if start_date and end_date:
        spans.append((start_date, end_date))

    spans.extend(_date_spans_from_text(block_text))
    return spans


def _table_dates(table_text: str) -> tuple[datetime | None, datetime | None]:
    table_fields = _parse_key_values(table_text)
    submitted = _parse_date(table_fields.get("Submitted for Funding", ""))
    awarded = _parse_date(table_fields.get("Awarded Date", ""))
    return submitted, awarded


def _funding_total_from_table(table_text: str) -> float | None:
    table_fields = _parse_key_values(table_text)
    candidates = [
        _parse_usd(table_fields.get("Total Anticipated", "")),
        _parse_usd(table_fields.get("Total Requested", "")),
        _parse_usd(table_fields.get("Award Amount", "")),
    ]
    candidates = [c for c in candidates if c is not None]
    if not candidates:
        return None
    return max(candidates)


def _is_grant_detail_table(table_text: str) -> bool:
    # The detail table we want consistently contains the date fields.
    keys = set(_parse_key_values(table_text).keys())
    return bool({"Start Date", "End Date", "Awarded Date", "Submitted for Funding"} & keys)


def _find_detail_table_start_index(doc) -> int:
    # Dossier exports can shift table indices between downloads. Find the first table that looks like
    # a grant detail table (contains the standard date fields).
    for idx, table in enumerate(getattr(doc, "tables", [])):
        txt = _table_text(table)
        if _is_grant_detail_table(txt):
            return idx
    return TABLE_START_INDEX


def _normalize_candidates(values: list[str]) -> list[str]:
    return [_normalize_key(value) for value in values if _normalize_key(value)]


def _matches_value(text: str, candidate: str) -> bool:
    left = _normalize_key(text)
    right = _normalize_key(candidate)
    if not left or not right:
        return False
    return left == right or left in right or right in left


def _catalog_candidates(entry: dict) -> tuple[list[str], list[str]]:
    title_candidates = [
        _clean_text(entry.get("canonical_title", "")),
        *entry.get("title_aliases", []),
    ]
    sponsor_candidates = [
        _clean_text(entry.get("canonical_sponsor", "")),
        *entry.get("sponsor_aliases", []),
    ]
    return title_candidates, sponsor_candidates


def _match_catalog_entry(title: str, sponsor: str, catalog: list[dict]) -> dict | None:
    title_clean = _clean_text(title)
    sponsor_clean = _clean_text(sponsor)

    for entry in catalog:
        title_candidates, sponsor_candidates = _catalog_candidates(entry)

        title_ok = any(_matches_value(title_clean, candidate) for candidate in title_candidates if candidate)
        sponsor_ok = any(_matches_value(sponsor_clean, candidate) for candidate in sponsor_candidates if candidate)

        if title_ok and sponsor_ok:
            return entry

    return None


def _choose_display_name(fields: dict[str, str]) -> tuple[str, str]:
    pi = _clean_text(fields.get("Principal Investigator", ""))
    cois = _clean_text(fields.get("Co-Investigator(s)", ""))
    role = _clean_text(fields.get("Role", ""))

    user_key = _normalize_key(USER_NAME)
    if pi and user_key in _normalize_key(pi):
        return pi, "Principal Investigator"
    if cois and user_key in _normalize_key(cois):
        return USER_NAME, "Co-Investigator"
    if pi:
        return pi, "Principal Investigator"
    if role:
        return role, "Role"
    if cois:
        return cois, "Co-Investigator"
    return "", ""


def _apply_cleanup_rules(text: str, rules: dict) -> str:
    cleaned = _clean_text(text)
    for pattern in rules.get("title_cleanup_patterns", []):
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)
    return _clean_text(cleaned)


def _should_exclude_from_rules(title: str, sponsor: str, block_text: str, rules: dict) -> bool:
    for rule in rules.get("exclude_entries", []):
        title_match = True
        sponsor_match = True
        block_match = True

        if "match" in rule:
            title_match = _matches_value(title, rule["match"])
        if "match_contains" in rule:
            title_match = _normalize_key(str(rule["match_contains"])) in _normalize_key(title)
        if "sponsor_match" in rule:
            sponsor_match = _matches_value(sponsor, rule["sponsor_match"])
        if "sponsor_match_contains" in rule:
            sponsor_match = _normalize_key(str(rule["sponsor_match_contains"])) in _normalize_key(sponsor)
        if "block_contains" in rule:
            block_match = _normalize_key(str(rule["block_contains"])) in _normalize_key(block_text)

        if title_match and sponsor_match and block_match:
            return True

    return False


def _build_records(doc) -> list[GrantRecord]:
    rules = _load_rules()
    catalog = _load_catalog()
    paragraphs = _extract_section_paragraphs(doc)
    blocks = _split_grant_blocks(paragraphs)

    grouped: OrderedDict[tuple[str, str, str, str], GrantRecord] = OrderedDict()
    # The dossier stores a run of grant detail tables (with Start/End/Submitted/Awarded)
    # but they do not align 1:1 with paragraph blocks once Pending/Not Funded items appear.
    detail_tables: list[str] = []
    table_start = _find_detail_table_start_index(doc)
    for table in doc.tables[table_start:]:
        txt = _table_text(table)
        if _is_grant_detail_table(txt):
            detail_tables.append(txt)
    awarded_table_cursor = 0

    for idx, (group_label, block_text) in enumerate(blocks):
        table_text = ""
        if group_label == "Awarded" and awarded_table_cursor < len(detail_tables):
            table_text = detail_tables[awarded_table_cursor]
            awarded_table_cursor += 1

        fields = _parse_key_values(block_text)
        raw_title = _best_field_value(block_text, "Project Title") or fields.get("Project Title", "")
        raw_sponsor = _best_field_value(block_text, "Agency") or fields.get("Agency", "")
        display_name, role_label = _choose_display_name(fields)

        cleaned_title = _apply_cleanup_rules(raw_title, rules)
        cleaned_sponsor = _clean_text(raw_sponsor)

        catalog_entry = _match_catalog_entry(cleaned_title, cleaned_sponsor, catalog)
        if catalog_entry:
            if catalog_entry.get("exclude"):
                continue
            canonical_title = _clean_text(catalog_entry.get("canonical_title", cleaned_title))
            canonical_sponsor = _clean_text(catalog_entry.get("canonical_sponsor", cleaned_sponsor))
            if catalog_entry.get("canonical_role_label"):
                role_label = _clean_text(catalog_entry.get("canonical_role_label", role_label))
            canonical_start = _parse_month_year(str(catalog_entry.get("canonical_start", "")))
            canonical_end = _parse_month_year(str(catalog_entry.get("canonical_end", "")))
            matched_catalog = canonical_title
        else:
            canonical_title = cleaned_title
            canonical_sponsor = cleaned_sponsor
            canonical_start = None
            canonical_end = None
            matched_catalog = None

        if _should_exclude_from_rules(canonical_title, canonical_sponsor, block_text, rules):
            continue

        spans = _collect_record_spans(block_text, table_text)
        submitted, awarded = _table_dates(table_text)
        is_awarded = group_label == "Awarded"
        anticipated_total_usd = _funding_total_from_table(table_text) if is_awarded else None
        key = (
            _normalize_key(canonical_title),
            _normalize_key(canonical_sponsor),
            _normalize_key(display_name),
            _normalize_key(role_label),
        )

        if key not in grouped:
            grouped[key] = GrantRecord(
                display_name=display_name,
                role_label=role_label,
                title=canonical_title,
                sponsor=canonical_sponsor,
                order=idx,
                matched_catalog=matched_catalog,
                is_awarded=is_awarded,
            )

        record = grouped[key]
        if not record.display_name and display_name:
            record.display_name = display_name
        if not record.role_label and role_label:
            record.role_label = role_label
        if not record.title and canonical_title:
            record.title = canonical_title
        if not record.sponsor and canonical_sponsor:
            record.sponsor = canonical_sponsor
        if matched_catalog and not record.matched_catalog:
            record.matched_catalog = matched_catalog
        if canonical_start and (record.start is None or canonical_start < record.start):
            record.start = canonical_start
        if canonical_end and (record.end is None or canonical_end > record.end):
            record.end = canonical_end
        if anticipated_total_usd is not None and (
            record.anticipated_total_usd is None or anticipated_total_usd > record.anticipated_total_usd
        ):
            record.anticipated_total_usd = anticipated_total_usd
        if is_awarded:
            record.is_awarded = True
        if submitted and (record.submitted is None or submitted > record.submitted):
            record.submitted = submitted
        if awarded and (record.awarded is None or awarded > record.awarded):
            record.awarded = awarded

        for start, end in spans:
            if record.start is None or start < record.start:
                record.start = start
            if record.end is None or end > record.end:
                record.end = end

    records = list(grouped.values())
    records.sort(
        key=lambda record: (
            record.end or record.start or datetime.min,
            record.start or datetime.min,
            -record.order,
        ),
        reverse=True,
    )
    return records


def _format_record(record: GrantRecord) -> str:
    parts: list[str] = []

    if record.display_name:
        if record.role_label:
            parts.append(f"{_latex_escape(record.display_name)} ({_latex_escape(record.role_label)})")
        else:
            parts.append(_latex_escape(record.display_name))

    if record.title:
        parts.append(f"``{_latex_escape(record.title)}''")

    if record.sponsor:
        parts.append(f"Sponsored by {_latex_escape(record.sponsor)}")

    line = ", ".join(parts)

    if record.start and record.end:
        start_text = _month_year(record.start)
        end_text = _month_year(record.end)
        if start_text == end_text:
            line += f". {start_text}."
        else:
            line += f". {start_text} - {end_text}."
    elif record.start:
        line += f". {_month_year(record.start)}."
    elif record.end:
        line += f". {_month_year(record.end)}."
    else:
        line += "."

    return line


def apply(text_content: str, doc, document_text: str, mode: str = "full") -> str:
    del document_text
    del mode

    records = _build_records(doc)
    latex_content = "\n" + SECTION_TITLE + "\n"

    for record in records:
        latex_content += "\n\\noindent " + _format_record(record) + r"\vspace{0.25cm}" + "\n"

    return text_content.replace("{{awarded}}", latex_content)
