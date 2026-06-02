from __future__ import annotations

import html
import json
import re
from difflib import SequenceMatcher
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from core.io_utils import extract_text_between_markers
from sections import grants as grants_section
from sections import publications as publications_section


DATA_DIR = Path(__file__).resolve().parent.parent / "data"
CATALOG_PATH = DATA_DIR / "grants_catalog.json"


@dataclass
class FundingRow:
    title: str
    sponsor: str
    start: datetime | None
    end: datetime | None
    total_usd: float | None
    my_pct: float | None
    cluster_id: int = 0

    @property
    def my_share_usd(self) -> float | None:
        if self.total_usd is None or self.my_pct is None:
            return None
        return self.total_usd * (self.my_pct / 100.0)


def _load_catalog() -> list[dict[str, Any]]:
    if not CATALOG_PATH.exists():
        return []
    payload = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    entries = payload.get("entries", [])
    return entries if isinstance(entries, list) else []


def _normalize_key(text: str) -> str:
    return grants_section._normalize_key(text)  # noqa: SLF001 - intentional reuse


def _find_catalog_entry(title: str, sponsor: str, catalog: list[dict[str, Any]]) -> dict[str, Any] | None:
    tkey = _normalize_key(title)
    skey = _normalize_key(sponsor)
    for entry in catalog:
        if _normalize_key(str(entry.get("canonical_title", ""))) != tkey:
            continue
        if _normalize_key(str(entry.get("canonical_sponsor", ""))) != skey:
            continue
        return entry
    return None


def _fmt_month_year(value: datetime | None) -> str:
    if value is None:
        return ""
    return value.strftime("%b %Y")


def _fmt_usd(value: float | None) -> str:
    if value is None:
        return ""
    return f"${value:,.0f}"


def _fmt_pct(value: float | None) -> str:
    if value is None:
        return ""
    if float(value).is_integer():
        return f"{int(value)}%"
    return f"{value:.1f}%"


def _norm_text(text: str) -> str:
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _similar(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def _cluster_and_sort_rows(rows: list[FundingRow]) -> list[FundingRow]:
    by_sponsor: dict[str, list[FundingRow]] = {}
    for r in rows:
        by_sponsor.setdefault(_norm_text(r.sponsor), []).append(r)

    clustered: list[FundingRow] = []
    next_cluster_id = 1
    cluster_latest: dict[int, datetime] = {}

    for sponsor_rows in by_sponsor.values():
        sponsor_rows.sort(key=lambda r: (r.end or r.start or datetime.min, r.start or datetime.min), reverse=True)
        clusters: list[tuple[int, str]] = []  # (cluster_id, rep_norm_title)

        for row in sponsor_rows:
            norm_title = _norm_text(row.title)

            best_id = None
            best_score = 0.0
            for cid, rep in clusters:
                score = _similar(norm_title, rep)
                if score > best_score:
                    best_score = score
                    best_id = cid

            if best_id is None or best_score < 0.92:
                best_id = next_cluster_id
                next_cluster_id += 1
                clusters.append((best_id, norm_title))

            row.cluster_id = best_id
            clustered.append(row)

            ld = row.end or row.start or datetime.min
            if best_id not in cluster_latest or ld > cluster_latest[best_id]:
                cluster_latest[best_id] = ld

    clustered.sort(
        key=lambda r: (
            cluster_latest.get(r.cluster_id, datetime.min),
            r.end or r.start or datetime.min,
            _norm_text(r.title),
        ),
        reverse=True,
    )
    return clustered


def compute_publication_metrics(document_text: str) -> dict[str, int]:
    journal_publications = extract_text_between_markers(document_text, "Journal Article", "Parts of Books")
    conference_publications = extract_text_between_markers(document_text, "Refereed Conference Proceedings", "Other Works")
    preprint_publications = extract_text_between_markers(document_text, "Pre-Print", "Technical Report")
    technical_reports = extract_text_between_markers(document_text, "Technical Report", "Projects, Grants, Commissions, and Contracts")

    journal_entries = publications_section._parse_publication_blocks(journal_publications or "", "journal")  # noqa: SLF001
    conference_entries = publications_section._parse_publication_blocks(conference_publications or "", "conference")  # noqa: SLF001

    merged_other = ""
    if preprint_publications:
        merged_other += preprint_publications.strip() + "\n\n"
    if technical_reports:
        merged_other += technical_reports.strip()
    other_sections = publications_section._parse_other_publication_blocks(merged_other)  # noqa: SLF001

    published_total = len(journal_entries) + len(conference_entries)
    accepted_total = len(other_sections.get("accepted", []))

    return {
        "journal_published_count": len(journal_entries),
        "conference_published_count": len(conference_entries),
        "published_total_count": published_total,
        "accepted_total_count": accepted_total,
        "accepted_or_published_total_count": published_total + accepted_total,
    }


def _count_graduated(section_text: str) -> int:
    if not section_text:
        return 0
    count = 0
    for block in section_text.split("\n\n"):
        block = block.strip()
        if not block:
            continue
        first_line = block.splitlines()[0].strip()
        lower = first_line.lower()
        if "date graduated" in lower:
            count += 1
            continue
        # Otherwise treat entries with a non-Present end date as graduated.
        m = re.search(r"\(([^)]*)\)\.?\s*$", first_line)
        if not m:
            continue
        dates = m.group(1).strip()
        if "-" not in dates:
            continue
        end = dates.split("-", 1)[1].strip().lower()
        if "present" not in end:
            count += 1
    return count


def compute_student_metrics(document_text: str) -> dict[str, int]:
    phd_text = extract_text_between_markers(
        document_text,
        "Ph.D. Dissertation Advisor",
        "Ph.D. Dissertation Committee Member",
    )
    masters_text = extract_text_between_markers(
        document_text,
        "Master's Thesis Advisor",
        "Master's Thesis Committee Member",
    )

    phd_grad = _count_graduated(phd_text or "")
    ms_grad = _count_graduated(masters_text or "")
    return {
        "phd_graduated_count": phd_grad,
        "masters_graduated_count": ms_grad,
        "grad_total_graduated_count": phd_grad + ms_grad,
    }


def compute_funding_rows(doc) -> list[FundingRow]:
    catalog = _load_catalog()
    records = grants_section.extract_records(doc)

    rows: list[FundingRow] = []
    for record in records:
        if not getattr(record, "is_awarded", False):
            continue
        if not getattr(record, "title", ""):
            continue

        entry = _find_catalog_entry(record.title, record.sponsor, catalog)
        total_usd = record.anticipated_total_usd
        my_pct = None
        if entry:
            total_raw = entry.get("anticipated_total_usd")
            pct_raw = entry.get("my_share_pct")
            try:
                if total_raw is not None and str(total_raw).strip() != "":
                    total_usd = float(total_raw)
            except ValueError:
                pass
            try:
                my_pct = float(pct_raw) if pct_raw is not None and str(pct_raw).strip() != "" else None
            except ValueError:
                my_pct = None

        rows.append(
            FundingRow(
                title=record.title,
                sponsor=record.sponsor,
                start=record.start,
                end=record.end,
                total_usd=total_usd,
                my_pct=my_pct,
            )
        )

    return _cluster_and_sort_rows(rows)


def generate_metrics_html(document_text: str, doc) -> str:
    pub = compute_publication_metrics(document_text)
    students = compute_student_metrics(document_text)
    funding_rows = compute_funding_rows(doc)

    total_anticipated = sum((r.total_usd or 0.0) for r in funding_rows)
    total_my_share = sum((r.my_share_usd or 0.0) for r in funding_rows)

    def esc(value: str) -> str:
        return html.escape(value or "")

    funding_table_rows = []
    for r in funding_rows:
        funding_table_rows.append(
            "<tr>"
            f"<td>{esc(r.title)}</td>"
            f"<td>{esc(r.sponsor)}</td>"
            f"<td>{esc(_fmt_month_year(r.start))}</td>"
            f"<td>{esc(_fmt_month_year(r.end))}</td>"
            f"<td style='text-align:right'>{esc(_fmt_usd(r.total_usd))}</td>"
            f"<td style='text-align:right'>{esc(_fmt_pct(r.my_pct))}</td>"
            f"<td style='text-align:right'>{esc(_fmt_usd(r.my_share_usd))}</td>"
            "</tr>"
        )

    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Research Metrics</title>
    <style>
      :root {{
        --bg: #0b0f14;
        --card: #111823;
        --text: #e8eef7;
        --muted: #a9b7c9;
        --accent: #4fd1c5;
        --border: rgba(255,255,255,0.10);
      }}
      html, body {{ background: var(--bg); color: var(--text); font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial; }}
      .wrap {{ max-width: 1100px; margin: 28px auto; padding: 0 16px; }}
      h1 {{ margin: 0 0 6px; font-size: 24px; }}
      .sub {{ color: var(--muted); margin-bottom: 18px; }}
      .grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin: 14px 0 18px; }}
      .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 14px 14px; }}
      .k {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: 0.06em; }}
      .v {{ font-size: 26px; margin-top: 6px; }}
      .accent {{ color: var(--accent); }}
      table {{ width: 100%; border-collapse: collapse; background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }}
      th, td {{ padding: 10px 10px; border-bottom: 1px solid var(--border); vertical-align: top; }}
      th {{ text-align: left; color: var(--muted); font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.06em; }}
      tr:last-child td {{ border-bottom: none; }}
      .totals {{ display:flex; gap: 14px; flex-wrap: wrap; margin: 12px 0 10px; }}
      .pill {{ background: rgba(79, 209, 197, 0.12); border: 1px solid rgba(79, 209, 197, 0.25); color: var(--text); padding: 8px 10px; border-radius: 999px; }}
      .pill b {{ color: var(--accent); }}
      @media (max-width: 900px) {{ .grid {{ grid-template-columns: 1fr; }} }}
    </style>
  </head>
  <body>
    <div class="wrap">
      <h1>Research Metrics</h1>
      <div class="sub">Generated {datetime.now().strftime("%B %d, %Y")}</div>

      <div class="grid">
        <div class="card">
          <div class="k">Published total</div>
          <div class="v accent">{pub["published_total_count"]}</div>
          <div style="color: var(--muted); margin-top: 6px; font-size: 13px;">
            Journal: {pub["journal_published_count"]}<br/>
            Proceedings: {pub["conference_published_count"]}
          </div>
        </div>
        <div class="card">
          <div class="k">Accepted or published total</div>
          <div class="v accent">{pub["accepted_or_published_total_count"]}</div>
        </div>
        <div class="card">
          <div class="k">Manuscripts accepted</div>
          <div class="v accent">{pub["accepted_total_count"]}</div>
        </div>
        <div class="card">
          <div class="k">Graduate students graduated</div>
          <div class="v accent">{students["grad_total_graduated_count"]}</div>
          <div style="color: var(--muted); margin-top: 6px; font-size: 13px;">
            PhD: {students["phd_graduated_count"]}<br/>
            MS: {students["masters_graduated_count"]}
          </div>
        </div>
      </div>

      <h2 style="margin: 18px 0 8px; font-size: 16px;">Research Dollars</h2>
      <div class="totals">
        <div class="pill">Total anticipated: <b>{_fmt_usd(total_anticipated)}</b></div>
        <div class="pill">Your share (est.): <b>{_fmt_usd(total_my_share)}</b></div>
      </div>

      <table>
        <thead>
          <tr>
            <th>Grant / Project</th>
            <th>Sponsor</th>
            <th>Start</th>
            <th>End</th>
            <th style="text-align:right">Anticipated Total</th>
            <th style="text-align:right">Your %</th>
            <th style="text-align:right">Your Share</th>
          </tr>
        </thead>
        <tbody>
          {"".join(funding_table_rows)}
        </tbody>
      </table>
    </div>
  </body>
</html>
"""
