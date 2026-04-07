"""Step 8: Excel Output — Generate formatted .xlsx workbook with tier tabs."""

import logging
import os

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Output columns in spec order
OUTPUT_COLUMNS = [
    "HCP NPI", "First Name", "Last Name", "Credential", "Specialty",
    "Email", "Email Status", "Verified Phone", "Phone Status",
    "Primary Site of Care", "Practice Type", "Address 1", "City", "State",
    "Postal Code", "Tier", "MAC Jurisdiction", "Microlyte Eligible",
    "Joint Replacement - Procedure Volume",
    "Knee Joint Replacement - Procedure Volume",
    "Hip Joint Replacement - Procedure Volume",
    "Shoulder Joint Replacement - Procedure Volume",
    "Open Orthopedic Procedures - Procedure Volume",
    "Medical School", "HCP URL", "Subject Line", "Draft Email",
]

# Short column headers for the workbook
COLUMN_DISPLAY_NAMES = {
    "Joint Replacement - Procedure Volume": "Joint Repl Vol",
    "Knee Joint Replacement - Procedure Volume": "Knee Vol",
    "Hip Joint Replacement - Procedure Volume": "Hip Vol",
    "Shoulder Joint Replacement - Procedure Volume": "Shoulder Vol",
    "Open Orthopedic Procedures - Procedure Volume": "Open Ortho Vol",
}

# Column widths
COLUMN_WIDTHS = {
    "HCP NPI": 14, "First Name": 14, "Last Name": 16, "Credential": 10,
    "Specialty": 25, "Email": 32, "Email Status": 22, "Verified Phone": 15,
    "Phone Status": 20, "Primary Site of Care": 30, "Practice Type": 16,
    "Address 1": 25, "City": 16, "State": 8, "Postal Code": 12,
    "Tier": 18, "MAC Jurisdiction": 16, "Microlyte Eligible": 14,
    "Joint Repl Vol": 12, "Knee Vol": 10, "Hip Vol": 10,
    "Shoulder Vol": 12, "Open Ortho Vol": 12,
    "Medical School": 30, "HCP URL": 30, "Subject Line": 35, "Draft Email": 60,
}

# Style constants
HEADER_FILL = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
DATA_FONT = Font(name="Arial", size=9)
ALT_ROW_FILL = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

# Status color fills
GREEN_FILL = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
RED_FILL = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
MICROLYTE_GREEN = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")

# Tab definitions
TIER_TABS = [
    ("Tier 1 (0-30 min)", "Tier 1 (0-30 min)"),
    ("Tier 2 (30-60 min)", "Tier 2 (30-60 min)"),
    ("Tier 3 (60-120 min)", "Tier 3 (60-120 min)"),
    ("Tier 4 (120-180 min)", "Tier 4 (120-180 min)"),
    ("Tier 5 (180+ drivable)", "Tier 5 (180+ drivable)"),
    ("Tier 6 (Requires flight)", "Tier 6 (Requires flight)"),
    ("Hospital-Based", "Hospital-Based"),
]


def _sort_vol_descending(df: pd.DataFrame) -> pd.DataFrame:
    """Sort by Joint Replacement Procedure Volume descending."""
    vol_col = "Joint Replacement - Procedure Volume"
    if vol_col in df.columns:
        df = df.copy()
        df["_sort_vol"] = pd.to_numeric(df[vol_col], errors="coerce").fillna(0)
        df = df.sort_values("_sort_vol", ascending=False).drop(columns=["_sort_vol"])
    return df


def _prepare_output_df(df: pd.DataFrame) -> pd.DataFrame:
    """Select and order output columns, adding missing ones as empty."""
    out = pd.DataFrame()
    for col in OUTPUT_COLUMNS:
        if col in df.columns:
            out[col] = df[col].values
        else:
            out[col] = None
    return out


def _write_sheet(ws, df: pd.DataFrame):
    """Write a DataFrame to a worksheet with full formatting."""
    display_headers = [COLUMN_DISPLAY_NAMES.get(c, c) for c in OUTPUT_COLUMNS]

    # Write headers
    for col_idx, header in enumerate(display_headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Find status column indices (1-based)
    email_status_idx = OUTPUT_COLUMNS.index("Email Status") + 1
    phone_status_idx = OUTPUT_COLUMNS.index("Phone Status") + 1
    microlyte_idx = OUTPUT_COLUMNS.index("Microlyte Eligible") + 1

    # Write data rows
    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        is_alt = (row_idx - 2) % 2 == 1  # alternating rows
        base_fill = ALT_ROW_FILL if is_alt else WHITE_FILL

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
            value = row.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = DATA_FONT
            cell.fill = base_fill
            cell.alignment = Alignment(vertical="top", wrap_text=(col_name == "Draft Email"))

        # Color-coded status cells
        email_status = row.get("Email Status", "")
        phone_status = row.get("Phone Status", "")
        microlyte = row.get("Microlyte Eligible", "")

        if email_status and "Verified" in str(email_status):
            ws.cell(row=row_idx, column=email_status_idx).fill = GREEN_FILL
        elif email_status == "Missing":
            ws.cell(row=row_idx, column=email_status_idx).fill = RED_FILL

        if phone_status == "Added from NPPES":
            ws.cell(row=row_idx, column=phone_status_idx).fill = YELLOW_FILL

        if microlyte == "Yes":
            ws.cell(row=row_idx, column=microlyte_idx).fill = MICROLYTE_GREEN

    # Set column widths
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
        display_name = COLUMN_DISPLAY_NAMES.get(col_name, col_name)
        width = COLUMN_WIDTHS.get(display_name, COLUMN_WIDTHS.get(col_name, 15))
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-filter
    if len(df) > 0:
        last_col = get_column_letter(len(OUTPUT_COLUMNS))
        ws.auto_filter.ref = f"A1:{last_col}{len(df) + 1}"


def _write_summary(ws, df: pd.DataFrame):
    """Write the Summary tab with pipeline statistics."""
    ws.sheet_properties.tabColor = "1B4F72"

    title_font = Font(name="Arial", size=14, bold=True, color="1B4F72")
    header_font = Font(name="Arial", size=11, bold=True)
    data_font = Font(name="Arial", size=10)

    row = 1
    ws.cell(row=row, column=1, value="Lead List Pipeline Summary").font = title_font
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 20

    row += 2
    ws.cell(row=row, column=1, value="Total Leads").font = header_font
    ws.cell(row=row, column=2, value=len(df)).font = data_font

    # Tier counts
    row += 2
    ws.cell(row=row, column=1, value="Tier Distribution").font = header_font
    row += 1
    tier_counts = df["Tier"].value_counts()
    for tier_name, _ in TIER_TABS:
        count = tier_counts.get(tier_name, 0)
        ws.cell(row=row, column=1, value=tier_name).font = data_font
        ws.cell(row=row, column=2, value=count).font = data_font
        row += 1

    # Phone status
    row += 1
    ws.cell(row=row, column=1, value="Phone Status").font = header_font
    row += 1
    if "Phone Status" in df.columns:
        for status, count in df["Phone Status"].value_counts().items():
            ws.cell(row=row, column=1, value=status).font = data_font
            ws.cell(row=row, column=2, value=count).font = data_font
            row += 1

    # Email status
    row += 1
    ws.cell(row=row, column=1, value="Email Status").font = header_font
    row += 1
    if "Email Status" in df.columns:
        for status, count in df["Email Status"].value_counts().items():
            ws.cell(row=row, column=1, value=status).font = data_font
            ws.cell(row=row, column=2, value=count).font = data_font
            row += 1

    # Microlyte eligibility
    row += 1
    ws.cell(row=row, column=1, value="Microlyte Eligibility").font = header_font
    row += 1
    if "Microlyte Eligible" in df.columns:
        for val, count in df["Microlyte Eligible"].value_counts().items():
            ws.cell(row=row, column=1, value=f"Microlyte Eligible = {val}").font = data_font
            ws.cell(row=row, column=2, value=count).font = data_font
            row += 1

    # Template track split
    row += 1
    ws.cell(row=row, column=1, value="Outreach Template Track").font = header_font
    row += 1
    track_a = len(df[df.get("Microlyte Eligible", pd.Series()) == "No"]) if "Microlyte Eligible" in df.columns else 0
    track_b = len(df[df.get("Microlyte Eligible", pd.Series()) == "Yes"]) if "Microlyte Eligible" in df.columns else 0
    ws.cell(row=row, column=1, value="Track A (ProPacks Only)").font = data_font
    ws.cell(row=row, column=2, value=track_a).font = data_font
    row += 1
    ws.cell(row=row, column=1, value="Track B (ProPacks + Microlyte)").font = data_font
    ws.cell(row=row, column=2, value=track_b).font = data_font


def export_workbook(
    df: pd.DataFrame,
    output_path: str = "data/output/lead_list_enriched.xlsx",
    tiers: list[int] | None = None,
) -> str:
    """Export the enriched DataFrame to a formatted Excel workbook."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    wb = Workbook()

    # Summary tab (use the default sheet)
    ws_summary = wb.active
    ws_summary.title = "Summary"
    _write_summary(ws_summary, df)

    # Tier tabs
    for tier_name, tier_value in TIER_TABS:
        # Filter by tiers if specified
        if tiers and tier_name != "Hospital-Based":
            tier_num = int(tier_name.split()[1]) if tier_name.startswith("Tier") else None
            if tier_num and tier_num not in tiers:
                continue

        tier_df = df[df["Tier"] == tier_value].copy()
        if len(tier_df) == 0:
            # Still create the tab, but empty
            ws = wb.create_sheet(title=tier_name)
            _write_sheet(ws, _prepare_output_df(tier_df))
            continue

        tier_df = _sort_vol_descending(tier_df)
        out_df = _prepare_output_df(tier_df)

        ws = wb.create_sheet(title=tier_name)
        _write_sheet(ws, out_df)
        logger.info(f"  Tab '{tier_name}': {len(out_df)} leads")

    wb.save(output_path)
    logger.info(f"Workbook saved to {output_path}")
    return output_path
