"""Rebuild the Dashboard tab with working formulas matching current tab structure."""

import logging
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
logger = logging.getLogger("dashboard")

# Styles
TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

SECTION_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
JR_SECTION_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SN_SECTION_FILL = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
OOS_SECTION_FILL = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
WEEK_SECTION_FILL = PatternFill(start_color="203764", end_color="203764", fill_type="solid")

HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
DATA_FONT = Font(name="Calibri", size=10)
DATA_BOLD = Font(name="Calibri", size=11, bold=True)
PCT_FONT = Font(name="Calibri", size=10, color="0070C0")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")


def write_section_header(ws, row, title, fill, ncols=10):
    """Write a colored section header row spanning ncols."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = SECTION_FONT
    cell.fill = fill
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def write_subheader_row(ws, row, labels):
    """Write a subheader row of column labels."""
    for i, label in enumerate(labels, 1):
        cell = ws.cell(row=row, column=i, value=label)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def write_data_row(ws, row, values, pct_cols=None):
    """Write a data row with formulas/values. pct_cols is a list of 1-indexed column positions to format as percentages."""
    pct_cols = pct_cols or []
    for i, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=i, value=v)
        if i in pct_cols:
            cell.font = PCT_FONT
            cell.number_format = "0.0%"
        else:
            cell.font = DATA_BOLD
            cell.number_format = "#,##0"
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[row].height = 22


def build_portfolio_block(ws, row, title, section_fill, tab, total_range, vol_col_letter, vol_label):
    """Build a portfolio summary block for one dataset."""
    write_section_header(ws, row, title, section_fill)
    row += 1

    write_subheader_row(ws, row, [
        "Total Leads", "Tier 1", "Tier 2", "Tier 3", "Tier 4",
        "Hospital-Based", "Private Practice", "Microlyte Eligible",
        vol_label, "",
    ])
    row += 1

    t = tab  # shorter
    write_data_row(ws, row, [
        f"=COUNTA('{t}'!A2:A{total_range})",
        f'=COUNTIF(\'{t}\'!L2:L{total_range},"Tier 1*")',
        f'=COUNTIF(\'{t}\'!L2:L{total_range},"Tier 2*")',
        f'=COUNTIF(\'{t}\'!L2:L{total_range},"Tier 3*")',
        f'=COUNTIF(\'{t}\'!L2:L{total_range},"Tier 4*")',
        f'=COUNTIF(\'{t}\'!K2:K{total_range},"Hospital*")',
        f'=COUNTIF(\'{t}\'!K2:K{total_range},"Private*")',
        f'=COUNTIF(\'{t}\'!P2:P{total_range},"Yes")',
        f"=IFERROR(ROUND(AVERAGE('{t}'!{vol_col_letter}2:{vol_col_letter}{total_range}),0),0)",
        "",
    ])
    return row + 2


def build_pipeline_block(ws, row, title, section_fill, tab, total_range):
    """Build a Lead Status pipeline block."""
    write_section_header(ws, row, title, section_fill)
    row += 1

    write_subheader_row(ws, row, [
        "New", "Contacted", "Meeting Sched", "Meeting Done",
        "Proposal Sent", "Won", "Lost", "Nurture", "", "",
    ])
    row += 1

    t = tab
    write_data_row(ws, row, [
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"New")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Contacted")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Meeting Scheduled")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Meeting Completed")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Proposal Sent")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Won")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Lost")',
        f'=COUNTIF(\'{t}\'!N2:N{total_range},"Nurture")',
        "", "",
    ])
    return row + 2


def build_call_totals_block(ws, row, title, section_fill, tab, total_range):
    """Build call totals block aggregating across all 5 call attempts."""
    write_section_header(ws, row, title, section_fill)
    row += 1

    write_subheader_row(ws, row, [
        "Total Calls", "Connected", "Voicemail", "No Answer", "Gatekeeper",
        "Meetings Set", "Not Interested", "Bad Numbers", "Connect Rate", "Meeting Rate",
    ])
    row += 1

    t = tab
    tr = total_range
    # Call outcome columns: W, Z, AC, AF, AI (shifted +4 after collagen/incision insertion)
    outcome_cols = ["W", "Z", "AC", "AF", "AI"]
    counta_parts = [f"COUNTA('{t}'!{c}2:{c}{tr})" for c in outcome_cols]
    total_calls_formula = "=" + "+".join(counta_parts)

    def countif(pattern):
        parts = [f'COUNTIF(\'{t}\'!{c}2:{c}{tr},"{pattern}")' for c in outcome_cols]
        return "=" + "+".join(parts)

    # We have total calls in col A, connected in col B
    data_row = row
    write_data_row(ws, data_row, [
        total_calls_formula,
        countif("Connected*"),
        countif("Voicemail"),
        countif("No Answer"),
        countif("Gatekeeper*"),
        countif("Connected - Meeting Set"),
        countif("Connected - Not Interested"),
        countif("Wrong Number"),
        f"=IFERROR(B{data_row}/A{data_row},0)",
        f"=IFERROR(F{data_row}/A{data_row},0)",
    ], pct_cols=[9, 10])
    return row + 2


def build_email_totals_block(ws, row, title, section_fill, tab, total_range):
    """Build email totals block aggregating across all 3 email attempts."""
    write_section_header(ws, row, title, section_fill)
    row += 1

    write_subheader_row(ws, row, [
        "Total Sent", "Opened", "Replied", "Bounced", "No Response",
        "Unsubscribed", "", "", "Open Rate", "Reply Rate",
    ])
    row += 1

    t = tab
    tr = total_range
    # Email 1-3 Outcome columns are at cols T, X, AB in Email Tracker (based on 31-col layout)
    # Email Tracker col map: A-R metadata (17 cols + Email 1 Date col R + Email 1 Subject S + Email 1 Outcome T + Email 1 Notes U)
    # Email 1 Outcome = T, Email 2 Outcome = X, Email 3 Outcome = AB
    outcome_cols = ["T", "X", "AB"]
    counta_parts = [f"COUNTA('{t}'!{c}2:{c}{tr})" for c in outcome_cols]
    total_sent_formula = "=" + "+".join(counta_parts)

    def countif(pattern):
        parts = [f'COUNTIF(\'{t}\'!{c}2:{c}{tr},"{pattern}")' for c in outcome_cols]
        return "=" + "+".join(parts)

    data_row = row
    write_data_row(ws, data_row, [
        total_sent_formula,
        countif("Opened"),
        countif("Replied"),
        countif("Bounced"),
        countif("No Response"),
        countif("Unsubscribed"),
        "", "",
        f"=IFERROR(B{data_row}/A{data_row},0)",
        f"=IFERROR(C{data_row}/A{data_row},0)",
    ], pct_cols=[9, 10])
    return row + 2


def build_collagen_block(ws, row, title, section_fill, tab, total_range):
    """Build collagen volume averages + totals block."""
    write_section_header(ws, row, title, section_fill)
    row += 1
    write_subheader_row(ws, row, [
        "Leads w/ Collagen Use", "Avg Lg Collagen", "Avg Sm/Md Collagen",
        "Avg Collagen Powder", "Total Lg Collagen", "Total Sm/Md Collagen",
        "Total Collagen Powder", "Max Lg Collagen", "", "",
    ])
    row += 1
    t = tab
    tr = total_range
    # Columns: R = Lg Collagen, S = Sm/Md Collagen, T = Collagen Powder
    write_data_row(ws, row, [
        f"=COUNTIF('{t}'!R2:R{tr},\">0\")+COUNTIF('{t}'!S2:S{tr},\">0\")+COUNTIF('{t}'!T2:T{tr},\">0\")",
        f"=IFERROR(ROUND(AVERAGEIF('{t}'!R2:R{tr},\">0\"),0),0)",
        f"=IFERROR(ROUND(AVERAGEIF('{t}'!S2:S{tr},\">0\"),0),0)",
        f"=IFERROR(ROUND(AVERAGEIF('{t}'!T2:T{tr},\">0\"),0),0)",
        f"=IFERROR(SUM('{t}'!R2:R{tr}),0)",
        f"=IFERROR(SUM('{t}'!S2:S{tr}),0)",
        f"=IFERROR(SUM('{t}'!T2:T{tr}),0)",
        f"=IFERROR(MAX('{t}'!R2:R{tr}),0)",
        "", "",
    ])
    return row + 2


def build_incision_block(ws, row, title, section_fill, tab, total_range):
    """Build incision likelihood distribution block."""
    write_section_header(ws, row, title, section_fill)
    row += 1
    write_subheader_row(ws, row, [
        "High", "Medium-High", "Medium", "Low", "Unlikely",
        "% High/Med-High", "High + Med-High Count", "", "", "",
    ])
    row += 1
    t = tab
    tr = total_range
    # Column U = Lg Incision Likelihood
    data_row = row
    write_data_row(ws, data_row, [
        f'=COUNTIF(\'{t}\'!U2:U{tr},"High")',
        f'=COUNTIF(\'{t}\'!U2:U{tr},"Medium-High")',
        f'=COUNTIF(\'{t}\'!U2:U{tr},"Medium")',
        f'=COUNTIF(\'{t}\'!U2:U{tr},"Low")',
        f'=COUNTIF(\'{t}\'!U2:U{tr},"Unlikely")',
        f"=IFERROR((A{data_row}+B{data_row})/COUNTA('{t}'!U2:U{tr}),0)",
        f"=A{data_row}+B{data_row}",
        "", "", "",
    ], pct_cols=[6])
    return row + 2


def build_specialty_block(ws, row, title, section_fill, tab, total_range, specialty_map, vol_col, vol_label):
    """Build a specialty breakdown block."""
    write_section_header(ws, row, title, section_fill)
    row += 1

    labels = list(specialty_map.keys()) + [vol_label]
    # Pad to 10 columns
    labels = labels + [""] * (10 - len(labels))
    write_subheader_row(ws, row, labels[:10])
    row += 1

    t = tab
    values = []
    for display_name, spec_patterns in specialty_map.items():
        if isinstance(spec_patterns, str):
            spec_patterns = [spec_patterns]
        parts = [f'COUNTIF(\'{t}\'!E2:E{total_range},"{p}")' for p in spec_patterns]
        values.append("=" + "+".join(parts))
    values.append(f"=IFERROR(ROUND(AVERAGE('{t}'!{vol_col}2:{vol_col}{total_range}),0),0)")
    values = values + [""] * (10 - len(values))
    write_data_row(ws, row, values[:10])
    return row + 2


def main():
    input_file = "Master_Lead_List_Tracker (3).xlsx"

    logger.info("Loading workbook...")
    wb = openpyxl.load_workbook(input_file)

    # Remove external links that cause the warning
    # openpyxl doesn't track external links directly, but we can clear them
    # by accessing the _external_links
    if hasattr(wb, '_external_links'):
        wb._external_links = []

    # Remove calcPr which can hold external link refs
    # and clear any defined names that reference externals
    names_to_remove = []
    for name in wb.defined_names:
        try:
            val = str(wb.defined_names[name].value)
            if ".xlsx" in val or ".xls]" in val or "['" in val:
                names_to_remove.append(name)
        except Exception:
            pass
    for name in names_to_remove:
        del wb.defined_names[name]

    # Delete and recreate Dashboard
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)  # insert at position 0
    ws.sheet_properties.tabColor = "1F4E78"

    # Column widths
    for i in range(1, 11):
        ws.column_dimensions[get_column_letter(i)].width = 16

    # Title
    ws.merge_cells("A1:J1")
    title_cell = ws.cell(row=1, column=1, value="ALBACETE MEDDEV — OUTREACH DASHBOARD")
    title_cell.font = TITLE_FONT
    title_cell.fill = TITLE_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.cell(row=2, column=1, value="Last updated: auto-refresh on open").font = Font(italic=True, size=9, color="808080")

    row = 4

    # ========== JR ==========
    row = build_portfolio_block(ws, row, "JOINT REPLACEMENT PORTFOLIO",
                                JR_SECTION_FILL, "Call Tracker - JR", "4442", "Q", "Avg Joint Vol")
    row = build_collagen_block(ws, row, "JR — COLLAGEN USAGE",
                               JR_SECTION_FILL, "Call Tracker - JR", "4442")
    row = build_incision_block(ws, row, "JR — INCISION LIKELIHOOD",
                               JR_SECTION_FILL, "Call Tracker - JR", "4442")
    row = build_pipeline_block(ws, row, "JR — PIPELINE STATUS",
                               JR_SECTION_FILL, "Call Tracker - JR", "4442")
    row = build_call_totals_block(ws, row, "JR — CALL TOTALS",
                                  JR_SECTION_FILL, "Call Tracker - JR", "4442")
    row = build_email_totals_block(ws, row, "JR — EMAIL TOTALS",
                                   JR_SECTION_FILL, "Email Tracker - JR", "4442")

    # ========== S&N ==========
    row += 1
    row = build_portfolio_block(ws, row, "SPINE & NEURO PORTFOLIO",
                                SN_SECTION_FILL, "Call Tracker - S&N", "10001", "Q", "Avg Spine Vol")

    row = build_specialty_block(ws, row, "S&N — SPECIALTY BREAKDOWN", SN_SECTION_FILL,
                                "Call Tracker - S&N", "10001",
                                {
                                    "Neurosurgery": "Neurological Surgery",
                                    "Spine Ortho": "Orthopaedic Surgery > Orthopaedic Surgery of the Spine",
                                    "Ortho": "Orthopaedic Surgery",
                                    "Vascular": "Surgery > Vascular Surgery",
                                    "General Surg": "Surgery",
                                    "Pain Med": ["Pain Medicine > Interventional Pain Medicine", "Pain Medicine"],
                                    "Otolaryngology": "Otolaryngology",
                                    "Internal Med": "Internal Medicine",
                                    "Specialist": "Specialist",
                                }, "Q", "Avg Spine Vol")

    row = build_collagen_block(ws, row, "S&N — COLLAGEN USAGE",
                               SN_SECTION_FILL, "Call Tracker - S&N", "10001")
    row = build_incision_block(ws, row, "S&N — INCISION LIKELIHOOD",
                               SN_SECTION_FILL, "Call Tracker - S&N", "10001")

    row = build_pipeline_block(ws, row, "S&N — PIPELINE STATUS",
                               SN_SECTION_FILL, "Call Tracker - S&N", "10001")
    row = build_call_totals_block(ws, row, "S&N — CALL TOTALS",
                                  SN_SECTION_FILL, "Call Tracker - S&N", "10001")
    row = build_email_totals_block(ws, row, "S&N — EMAIL TOTALS",
                                   SN_SECTION_FILL, "Email Tracker - S&N", "10001")

    # ========== OOS ==========
    row += 1
    row = build_portfolio_block(ws, row, "OUTSIDE ORTHO & SPINE PORTFOLIO",
                                OOS_SECTION_FILL, "Call Tracker - OOS", "10001", "Q", "Avg Proc Vol")

    row = build_specialty_block(ws, row, "OOS — SPECIALTY BREAKDOWN", OOS_SECTION_FILL,
                                "Call Tracker - OOS", "10001",
                                {
                                    "OB/GYN": "Obstetrics & Gynecology",
                                    "General Surg": "Surgery",
                                    "Dermatology": "Dermatology",
                                    "Vascular": "Surgery > Vascular Surgery",
                                    "Plastics": "Plastic Surgery",
                                    "Colorectal": "Colon & Rectal Surgery",
                                    "Emerg Med": "Emergency Medicine",
                                    "GynOnc": "Obstetrics & Gynecology > Gynecologic Oncology",
                                    "Mohs": "Dermatology > MOHS-Micrographic Surgery",
                                }, "Q", "Avg Proc Vol")

    row = build_collagen_block(ws, row, "OOS — COLLAGEN USAGE",
                               OOS_SECTION_FILL, "Call Tracker - OOS", "10001")
    row = build_incision_block(ws, row, "OOS — INCISION LIKELIHOOD",
                               OOS_SECTION_FILL, "Call Tracker - OOS", "10001")

    row = build_pipeline_block(ws, row, "OOS — PIPELINE STATUS",
                               OOS_SECTION_FILL, "Call Tracker - OOS", "10001")
    row = build_call_totals_block(ws, row, "OOS — CALL TOTALS",
                                  OOS_SECTION_FILL, "Call Tracker - OOS", "10001")
    row = build_email_totals_block(ws, row, "OOS — EMAIL TOTALS",
                                   OOS_SECTION_FILL, "Email Tracker - OOS", "10001")

    # ========== TOTALS ROLLUP ==========
    row += 1
    write_section_header(ws, row, "ALL PORTFOLIOS — COMBINED TOTALS", WEEK_SECTION_FILL)
    row += 1
    write_subheader_row(ws, row, [
        "Total Leads", "JR Leads", "S&N Leads", "OOS Leads",
        "Private Practice", "Hospital-Based", "Microlyte Eligible",
        "Total Calls", "Total Emails", "",
    ])
    row += 1
    # Direct formulas that don't depend on other dashboard cells
    write_data_row(ws, row, [
        "=COUNTA('Call Tracker - JR'!A2:A4442)+COUNTA('Call Tracker - S&N'!A2:A10001)+COUNTA('Call Tracker - OOS'!A2:A10001)",
        "=COUNTA('Call Tracker - JR'!A2:A4442)",
        "=COUNTA('Call Tracker - S&N'!A2:A10001)",
        "=COUNTA('Call Tracker - OOS'!A2:A10001)",
        "=COUNTIF('Call Tracker - JR'!K2:K4442,\"Private*\")+COUNTIF('Call Tracker - S&N'!K2:K10001,\"Private*\")+COUNTIF('Call Tracker - OOS'!K2:K10001,\"Private*\")",
        "=COUNTIF('Call Tracker - JR'!K2:K4442,\"Hospital*\")+COUNTIF('Call Tracker - S&N'!K2:K10001,\"Hospital*\")+COUNTIF('Call Tracker - OOS'!K2:K10001,\"Hospital*\")",
        "=COUNTIF('Call Tracker - JR'!P2:P4442,\"Yes\")+COUNTIF('Call Tracker - S&N'!P2:P10001,\"Yes\")+COUNTIF('Call Tracker - OOS'!P2:P10001,\"Yes\")",
        "=COUNTA('Call Tracker - JR'!S2:S4442)+COUNTA('Call Tracker - JR'!V2:V4442)+COUNTA('Call Tracker - JR'!Y2:Y4442)+COUNTA('Call Tracker - JR'!AB2:AB4442)+COUNTA('Call Tracker - JR'!AE2:AE4442)+COUNTA('Call Tracker - S&N'!S2:S10001)+COUNTA('Call Tracker - S&N'!V2:V10001)+COUNTA('Call Tracker - S&N'!Y2:Y10001)+COUNTA('Call Tracker - S&N'!AB2:AB10001)+COUNTA('Call Tracker - S&N'!AE2:AE10001)+COUNTA('Call Tracker - OOS'!S2:S10001)+COUNTA('Call Tracker - OOS'!V2:V10001)+COUNTA('Call Tracker - OOS'!Y2:Y10001)+COUNTA('Call Tracker - OOS'!AB2:AB10001)+COUNTA('Call Tracker - OOS'!AE2:AE10001)",
        "=COUNTA('Email Tracker - JR'!T2:T4442)+COUNTA('Email Tracker - JR'!X2:X4442)+COUNTA('Email Tracker - JR'!AB2:AB4442)+COUNTA('Email Tracker - S&N'!T2:T10001)+COUNTA('Email Tracker - S&N'!X2:X10001)+COUNTA('Email Tracker - S&N'!AB2:AB10001)+COUNTA('Email Tracker - OOS'!T2:T10001)+COUNTA('Email Tracker - OOS'!X2:X10001)+COUNTA('Email Tracker - OOS'!AB2:AB10001)",
        "",
    ])

    # Freeze panes
    ws.freeze_panes = "A3"

    wb.save(input_file)
    logger.info(f"Dashboard rebuilt in {input_file}")
    logger.info(f"Total tabs: {wb.sheetnames}")


if __name__ == "__main__":
    main()
