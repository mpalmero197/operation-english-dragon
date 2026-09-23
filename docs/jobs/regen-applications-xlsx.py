#!/usr/bin/env python3
"""Regenerate applications-living.xlsx and applications-living.csv from applications.json.

This script is intentionally self-contained and writes outputs beside this file.
Requires: openpyxl (pip install openpyxl)
"""
from __future__ import annotations

import csv
import json
import re
import shutil
from collections import Counter
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - helpful CLI failure
    raise SystemExit("openpyxl is required; install it with: python3 -m pip install openpyxl") from exc

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "applications.json"
XLSX = HERE / "applications-living.xlsx"
CSV = HERE / "applications-living.csv"
CT = ZoneInfo("America/Chicago")

HEADERS = [
    "School", "City", "Role / job title", "Status", "Pay / band",
    "Z-sponsor (Yes/No/ask)", "Mandarin required", "Cert required first",
    "Applied date", "Last email date", "Last email subject",
    "Contact / apply URL", "Email watch keywords", "Notes", "Application ID",
]

# Excel-friendly fills: the source status text is retained verbatim.
STATUS_FILLS = {
    "received": PatternFill("solid", fgColor="D9EAD3"),
    "submitted": PatternFill("solid", fgColor="D9EAF7"),
    "skipped": PatternFill("solid", fgColor="E7E6E6"),
    "blocked": PatternFill("solid", fgColor="E7E6E6"),
    "interview": PatternFill("solid", fgColor="FCE4D6"),
}


def parse_dt(value: object) -> datetime | None:
    if not value:
        return None
    text = str(value).strip()
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        return dt.astimezone(CT).replace(tzinfo=None)
    except (TypeError, ValueError):
        return None


def date_value(value: object) -> datetime | str:
    return parse_dt(value) or ""


def date_text(value: object) -> str:
    dt = parse_dt(value)
    return dt.strftime("%Y-%m-%d %H:%M CT") if dt else ""


def sponsor_bucket(value: object) -> str:
    """Reduce the source's explanatory sponsorship text to Yes/No/ask."""
    text = str(value or "").strip().lower()
    if re.search(r"\byes\b", text):
        return "Yes"
    if re.search(r"\bno\b", text):
        return "No"
    return "ask"


def bool_text(value: object) -> str:
    return "Yes" if bool(value) else "No"


def row_for(app: dict) -> list[object]:
    return [
        app.get("school", ""),
        app.get("city", ""),
        app.get("role", ""),
        app.get("status", ""),
        app.get("pay", ""),
        sponsor_bucket(app.get("zSponsor")),
        bool_text(app.get("mandarinRequired")),
        bool_text(app.get("certRequiredFirst")),
        date_value(app.get("appliedAt")),
        date_value(app.get("lastEmailAt")),
        app.get("lastEmailSubject", "") or "",
        app.get("url", "") or "",
        "; ".join(str(x) for x in (app.get("emailWatch") or [])),
        app.get("notes", "") or "",
        app.get("id", ""),
    ]


def csv_row(app: dict) -> list[str]:
    row = row_for(app)
    row[8] = date_text(app.get("appliedAt"))
    row[9] = date_text(app.get("lastEmailAt"))
    return ["" if x is None else str(x) for x in row]


def status_kind(status: object) -> str | None:
    text = str(status or "").lower()
    # Interview is checked first so an interview-related status is orange.
    if "interview" in text:
        return "interview"
    if "received" in text:
        return "received"
    if "submitted" in text:
        return "submitted"
    if "skipped" in text:
        return "skipped"
    if "blocked" in text:
        return "blocked"
    return None


def style_sheet(ws, rows: list[list[object]]) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(rows) + 1}"
    ws.row_dimensions[1].height = 30
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(r_idx, c_idx, value)
            cell.alignment = Alignment(vertical="top", wrap_text=(c_idx in (3, 4, 11, 13, 14)))
            if c_idx in (9, 10) and isinstance(value, datetime):
                cell.number_format = 'yyyy-mm-dd hh:mm "CT"'
            if c_idx == 4:
                fill = STATUS_FILLS.get(status_kind(value))
                if fill:
                    cell.fill = fill
                    cell.font = Font(bold=True)
            if c_idx == 12 and value:
                cell.hyperlink = str(value)
                cell.style = "Hyperlink"
        ws.row_dimensions[r_idx].height = 36 if len(str(row[13])) > 220 else 22
    widths = [34, 22, 42, 38, 34, 18, 18, 18, 22, 22, 48, 48, 38, 72, 34]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def build_summary(ws, apps: list[dict], updated: object) -> None:
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:B1")
    ws["A1"] = "Guangzhou job applications — living list"
    ws["A1"].font = Font(size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 28
    ws["A3"] = "Metric"
    ws["B3"] = "Value"
    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor="5B9BD5")
        cell.font = Font(color="FFFFFF", bold=True)
    exact = Counter(str(a.get("status", "")) for a in apps)
    buckets = Counter(status_kind(a.get("status")) or "Other" for a in apps)
    metrics = [
        ("Total applications", len(apps)),
        ("Submitted", buckets.get("submitted", 0)),
        ("Received", buckets.get("received", 0)),
        ("Skipped/Blocked", buckets.get("skipped", 0) + buckets.get("blocked", 0)),
        ("Interview", buckets.get("interview", 0)),
        ("Other statuses", buckets.get("Other", 0)),
        ("Updated", date_text(updated)),
        ("Note", "English 1 rebook may still be open / applying is paused per Michael."),
    ]
    for r, (label, value) in enumerate(metrics, start=4):
        ws.cell(r, 1, label).font = Font(bold=(r <= 10))
        ws.cell(r, 2, value)
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws["A14"] = "Status"
    ws["B14"] = "Count"
    for cell in ws[14]:
        cell.fill = PatternFill("solid", fgColor="5B9BD5")
        cell.font = Font(color="FFFFFF", bold=True)
    for r, (status, count) in enumerate(sorted(exact.items(), key=lambda item: (-item[1], item[0].lower())), start=15):
        ws.cell(r, 1, status)
        ws.cell(r, 2, count)
        fill = STATUS_FILLS.get(status_kind(status))
        if fill:
            ws.cell(r, 1).fill = fill
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 92
    ws.freeze_panes = "A4"


def main() -> None:
    data = json.loads(SOURCE.read_text(encoding="utf-8"))
    apps = data.get("applications", [])
    all_rows = [row_for(app) for app in apps]
    sorted_apps = sorted(apps, key=lambda app: (str(app.get("status", "")).lower(), str(app.get("school", "")).lower()))
    by_status_rows = [row_for(app) for app in sorted_apps]

    wb = Workbook()
    # Keep both published trees byte-identical across regenerations.
    generated_at = parse_dt(data.get("updated")) or datetime(2000, 1, 1)
    wb.properties.created = generated_at
    wb.properties.modified = generated_at
    ws_all = wb.active
    ws_all.title = "All applications"
    ws_all.append(HEADERS)
    for row in all_rows:
        ws_all.append(row)
    style_sheet(ws_all, all_rows)

    ws_status = wb.create_sheet("By status")
    ws_status.append(HEADERS)
    for row in by_status_rows:
        ws_status.append(row)
    style_sheet(ws_status, by_status_rows)

    ws_summary = wb.create_sheet("Summary")
    build_summary(ws_summary, apps, data.get("updated"))
    wb.save(XLSX)

    with CSV.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(HEADERS)
        writer.writerows(csv_row(app) for app in apps)
    print(f"Wrote {XLSX} and {CSV} ({len(apps)} applications)")


if __name__ == "__main__":
    main()
