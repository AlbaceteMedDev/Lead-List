"""Excel workbook generation with tier tabs, Summary stats, and conditional formatting."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd

log = logging.getLogger(__name__)

TIER_ORDER = [
    "Tier 1 (0-30 min)",
    "Tier 2 (30-60 min)",
    "Tier 3 (60-120 min)",
    "Tier 4 (120-180 min)",
    "Tier 5 (180+ min drivable)",
    "Tier 6 (Requires flight)",
    "Hospital-Based",
]

OUTPUT_COLUMNS = [
    "HCP NPI", "First Name", "Last Name", "Credential", "Specialty",
    "Email", "Email Status", "Verified Phone", "Phone Status",
    "Primary Site of Care", "Practice Type",
    "Address 1", "City", "State", "Postal Code",
    "Tier", "MAC Jurisdiction", "Microlyte Eligible",
    "Joint Repl Vol", "Knee Vol", "Hip Vol", "Shoulder Vol", "Open Ortho Vol",
    "Medical School", "HCP URL",
    "Subject Line", "Draft Email",
]

COLUMN_WIDTHS = {
    "HCP NPI": 14, "First Name": 14, "Last Name": 16, "Credential": 12,
    "Specialty": 30, "Email": 32, "Email Status": 26, "Verified Phone": 14,
    "Phone Status": 22, "Primary Site of Care": 30, "Practice Type": 18,
    "Address 1": 28, "City": 16, "State": 6, "Postal Code": 10,
    "Tier": 24, "MAC Jurisdiction": 14, "Microlyte Eligible": 14,
    "Joint Repl Vol": 12, "Knee Vol": 10, "Hip Vol": 10, "Shoulder Vol": 12,
    "Open Ortho Vol": 14, "Medical School": 28, "HCP URL": 32,
    "Subject Line": 40, "Draft Email": 60,
}


def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[OUTPUT_COLUMNS]


def _sort_key(df: pd.DataFrame) -> pd.Series:
    def _num(v):
        try:
            return float(str(v).replace(",", "")) if str(v).strip() else 0.0
        except ValueError:
            return 0.0
    return df["Joint Repl Vol"].map(_num)


def _summary_rows(df: pd.DataFrame) -> list[tuple[str, object]]:
    total = len(df)
    rows: list[tuple[str, object]] = [("Total leads", total)]
    for tier in TIER_ORDER:
        rows.append((f"  {tier}", int((df["Tier"] == tier).sum())))
    rows.append(("", ""))
    rows.append(("Phone Status", ""))
    for v in ["Verified", "Added from NPPES", "Updated (NPPES differs)", "Missing"]:
        rows.append((f"  {v}", int((df["Phone Status"] == v).sum())))
    rows.append(("", ""))
    rows.append(("Email Status", ""))
    for v in sorted(df["Email Status"].dropna().unique().tolist()):
        rows.append((f"  {v}", int((df["Email Status"] == v).sum())))
    rows.append(("", ""))
    rows.append(("Microlyte Eligible", ""))
    for v in ["Yes", "No", "Unknown"]:
        rows.append((f"  {v}", int((df["Microlyte Eligible"] == v).sum())))
    rows.append(("", ""))
    rows.append(("Email Track", ""))
    if "Email Track" in df.columns:
        for v in sorted(df["Email Track"].dropna().unique().tolist()):
            rows.append((f"  {v}", int((df["Email Track"] == v).sum())))
    return rows


def write_workbook(
    df: pd.DataFrame,
    output_path: Path,
    tiers: Optional[Iterable[str]] = None,
) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    summary_rows = _summary_rows(df)
    df = _ensure_columns(df)
    tiers = list(tiers) if tiers else TIER_ORDER

    wb = Workbook()
    default = wb.active
    wb.remove(default)

    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
    header_fill = PatternFill("solid", fgColor="FF1B4F72")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Arial", size=9)
    alt_fill = PatternFill("solid", fgColor="FFEBF5FB")
    thin = Side(border_style="thin", color="FFD5D8DC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    fills = {
        "email_verified": PatternFill("solid", fgColor="FFD5F5E3"),
        "email_missing": PatternFill("solid", fgColor="FFFADBD8"),
        "phone_added": PatternFill("solid", fgColor="FFFEF9E7"),
        "microlyte_yes": PatternFill("solid", fgColor="FFD4EFDF"),
    }

    # --- Summary tab ---
    summary = wb.create_sheet("Summary")
    summary["A1"] = "Albacete MedDev - Lead List Pipeline Summary"
    summary["A1"].font = Font(name="Arial", size=12, bold=True)
    summary["A3"] = "Metric"
    summary["B3"] = "Value"
    for cell in (summary["A3"], summary["B3"]):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
    for i, (label, value) in enumerate(summary_rows, start=4):
        summary.cell(row=i, column=1, value=label).font = data_font
        summary.cell(row=i, column=2, value=value).font = data_font
    summary.column_dimensions["A"].width = 40
    summary.column_dimensions["B"].width = 16
    summary.freeze_panes = "A4"

    # --- Tier tabs ---
    email_status_col = OUTPUT_COLUMNS.index("Email Status") + 1
    phone_status_col = OUTPUT_COLUMNS.index("Phone Status") + 1
    microlyte_col = OUTPUT_COLUMNS.index("Microlyte Eligible") + 1

    for tier in tiers:
        subset = df[df["Tier"] == tier].copy()
        if subset.empty:
            continue
        subset = subset.assign(_sort=_sort_key(subset)).sort_values("_sort", ascending=False).drop(columns="_sort")
        sheet_name = tier if len(tier) <= 31 else tier[:31]
        ws = wb.create_sheet(sheet_name)

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = border
            ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS.get(col_name, 16)

        for row_idx, (_, row) in enumerate(subset.iterrows(), start=2):
            for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
                value = row[col_name]
                cell = ws.cell(row=row_idx, column=col_idx, value=value if value != "" else None)
                cell.font = data_font
                cell.border = border
                if row_idx % 2 == 0:
                    cell.fill = alt_fill
                if col_name == "Draft Email":
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

            email_status = row["Email Status"]
            if email_status and "Verified" in email_status:
                ws.cell(row=row_idx, column=email_status_col).fill = fills["email_verified"]
            elif email_status == "Missing":
                ws.cell(row=row_idx, column=email_status_col).fill = fills["email_missing"]
            if row["Phone Status"] == "Added from NPPES":
                ws.cell(row=row_idx, column=phone_status_col).fill = fills["phone_added"]
            if row["Microlyte Eligible"] == "Yes":
                ws.cell(row=row_idx, column=microlyte_col).fill = fills["microlyte_yes"]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    log.info("Wrote workbook: %s", output_path)
    return output_path
