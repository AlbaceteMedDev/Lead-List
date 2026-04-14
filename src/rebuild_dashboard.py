"""Rebuild Dashboard with a clean, easy-to-scan KPI card layout."""

import logging
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
logger = logging.getLogger("dashboard")

# ============================================================
# STYLE SYSTEM
# ============================================================

# Colors
NAVY = "1F4E78"
PURPLE = "7030A0"
RED = "C00000"
TEAL = "006666"
DARK_GRAY = "3B3B3B"
LIGHT_GRAY_BG = "F2F2F2"
MED_GRAY_BG = "E7E6E6"
WHITE = "FFFFFF"

# Portfolio colors
JR_COLOR = NAVY
SN_COLOR = PURPLE
OOS_COLOR = RED

# Fonts
TITLE_FONT = Font(name="Calibri", size=22, bold=True, color=WHITE)
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="808080")
SECTION_BANNER_FONT = Font(name="Calibri", size=14, bold=True, color=WHITE)
SUBSECTION_FONT = Font(name="Calibri", size=11, bold=True, color=DARK_GRAY)

KPI_LABEL_FONT = Font(name="Calibri", size=10, bold=True, color="595959")
KPI_VALUE_FONT_LG = Font(name="Calibri", size=20, bold=True, color=DARK_GRAY)
KPI_VALUE_FONT_MD = Font(name="Calibri", size=14, bold=True, color=DARK_GRAY)
KPI_VALUE_FONT_ACCENT_JR = Font(name="Calibri", size=14, bold=True, color=JR_COLOR)
KPI_VALUE_FONT_ACCENT_SN = Font(name="Calibri", size=14, bold=True, color=SN_COLOR)
KPI_VALUE_FONT_ACCENT_OOS = Font(name="Calibri", size=14, bold=True, color=OOS_COLOR)

TABLE_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color=WHITE)
TABLE_CELL_FONT = Font(name="Calibri", size=11, bold=True, color=DARK_GRAY)

# Fills
TITLE_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
JR_BANNER_FILL = PatternFill(start_color=JR_COLOR, end_color=JR_COLOR, fill_type="solid")
SN_BANNER_FILL = PatternFill(start_color=SN_COLOR, end_color=SN_COLOR, fill_type="solid")
OOS_BANNER_FILL = PatternFill(start_color=OOS_COLOR, end_color=OOS_COLOR, fill_type="solid")
TEAL_BANNER_FILL = PatternFill(start_color=TEAL, end_color=TEAL, fill_type="solid")

CARD_FILL = PatternFill(start_color=LIGHT_GRAY_BG, end_color=LIGHT_GRAY_BG, fill_type="solid")
TABLE_HEADER_FILL = PatternFill(start_color=DARK_GRAY, end_color=DARK_GRAY, fill_type="solid")
TABLE_ROW_FILL = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
TABLE_ALT_FILL = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

# Borders
THIN_LIGHT = Side(style="thin", color="D0D0D0")
THICK_DARK = Side(style="medium", color=DARK_GRAY)
CARD_BORDER = Border(left=THIN_LIGHT, right=THIN_LIGHT, top=THIN_LIGHT, bottom=THIN_LIGHT)
TABLE_BORDER = Border(left=THIN_LIGHT, right=THIN_LIGHT, top=THIN_LIGHT, bottom=THIN_LIGHT)

# Alignment
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_INDENT = Alignment(horizontal="left", vertical="center", indent=1)


# ============================================================
# BUILDING BLOCKS
# ============================================================

def draw_title(ws, row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = TITLE_FONT
    cell.fill = TITLE_FILL
    cell.alignment = CENTER
    ws.row_dimensions[row].height = 42


def draw_subtitle(ws, row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SUBTITLE_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 18


def draw_section_banner(ws, row, text, fill):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_BANNER_FONT
    cell.fill = fill
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 28


def draw_subsection_header(ws, row, text):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SUBSECTION_FONT
    cell.alignment = LEFT_INDENT
    ws.row_dimensions[row].height = 20


def draw_kpi_card(ws, row, col_start, col_end, label, formula, value_font=None):
    """Draw a KPI card: label on top, value on bottom."""
    if value_font is None:
        value_font = KPI_VALUE_FONT_LG

    # Label row
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    label_cell = ws.cell(row=row, column=col_start, value=label)
    label_cell.font = KPI_LABEL_FONT
    label_cell.fill = CARD_FILL
    label_cell.alignment = CENTER
    label_cell.border = CARD_BORDER

    # Apply border to all merged cells
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).border = CARD_BORDER
        ws.cell(row=row, column=c).fill = CARD_FILL

    # Value row
    ws.merge_cells(start_row=row + 1, start_column=col_start, end_row=row + 1, end_column=col_end)
    val_cell = ws.cell(row=row + 1, column=col_start, value=formula)
    val_cell.font = value_font
    val_cell.fill = CARD_FILL
    val_cell.alignment = CENTER
    val_cell.border = CARD_BORDER
    val_cell.number_format = "#,##0"
    for c in range(col_start, col_end + 1):
        ws.cell(row=row + 1, column=c).border = CARD_BORDER
        ws.cell(row=row + 1, column=c).fill = CARD_FILL

    ws.row_dimensions[row].height = 22
    ws.row_dimensions[row + 1].height = 34


def draw_kpi_card_pct(ws, row, col_start, col_end, label, formula, value_font=None):
    """KPI card with percentage formatting."""
    if value_font is None:
        value_font = KPI_VALUE_FONT_LG
    draw_kpi_card(ws, row, col_start, col_end, label, formula, value_font)
    val_cell = ws.cell(row=row + 1, column=col_start)
    val_cell.number_format = "0.0%"


def draw_mini_table(ws, start_row, start_col, headers, values, widths=None):
    """Draw a clean mini-table with header row on top."""
    # Header row
    for i, h in enumerate(headers):
        cell = ws.cell(row=start_row, column=start_col + i, value=h)
        cell.font = TABLE_HEADER_FONT
        cell.fill = TABLE_HEADER_FILL
        cell.alignment = CENTER
        cell.border = TABLE_BORDER
    ws.row_dimensions[start_row].height = 28

    # Value row
    for i, v in enumerate(values):
        cell = ws.cell(row=start_row + 1, column=start_col + i, value=v)
        cell.font = TABLE_CELL_FONT
        cell.fill = TABLE_ROW_FILL
        cell.alignment = CENTER
        cell.border = TABLE_BORDER
        cell.number_format = "#,##0"
    ws.row_dimensions[start_row + 1].height = 26


def draw_mini_table_pct(ws, start_row, start_col, headers, values, pct_indices):
    """Mini-table with some columns formatted as percentage."""
    draw_mini_table(ws, start_row, start_col, headers, values)
    for pi in pct_indices:
        cell = ws.cell(row=start_row + 1, column=start_col + pi)
        cell.number_format = "0.0%"


# ============================================================
# PORTFOLIO SECTION BUILDER
# ============================================================

def build_portfolio_section(ws, start_row, portfolio_name, color_fill,
                            call_tab, email_tab, total_range, vol_col, vol_label,
                            accent_font):
    """Build a complete portfolio section: banner + KPIs + mini-tables."""
    row = start_row

    # Section banner
    draw_section_banner(ws, row, f"  ▍ {portfolio_name}", color_fill)
    row += 2

    # KPI cards row (4 cards, 3 cols each)
    t = call_tab
    tr = total_range

    draw_kpi_card(ws, row, 1, 3, "TOTAL LEADS",
                  f"=COUNTA('{t}'!A2:A{tr})", accent_font)
    draw_kpi_card(ws, row, 4, 6, "PRIVATE PRACTICE",
                  f"=COUNTIF('{t}'!K2:K{tr},\"Private*\")", accent_font)
    draw_kpi_card(ws, row, 7, 9, "HOSPITAL-BASED",
                  f"=COUNTIF('{t}'!K2:K{tr},\"Hospital*\")", accent_font)
    draw_kpi_card(ws, row, 10, 12, "MICROLYTE ELIGIBLE",
                  f"=COUNTIF('{t}'!P2:P{tr},\"Yes\")", accent_font)
    row += 3

    # Second KPI row — drive time tiers
    draw_kpi_card(ws, row, 1, 3, "TIER 1 (0-30 MIN)",
                  f"=COUNTIF('{t}'!L2:L{tr},\"Tier 1*\")", KPI_VALUE_FONT_MD)
    draw_kpi_card(ws, row, 4, 6, "TIER 2 (30-60 MIN)",
                  f"=COUNTIF('{t}'!L2:L{tr},\"Tier 2*\")", KPI_VALUE_FONT_MD)
    draw_kpi_card(ws, row, 7, 9, "TIER 3-4 (1-3 HRS)",
                  f"=COUNTIF('{t}'!L2:L{tr},\"Tier 3*\")+COUNTIF('{t}'!L2:L{tr},\"Tier 4*\")",
                  KPI_VALUE_FONT_MD)
    draw_kpi_card(ws, row, 10, 12, vol_label.upper(),
                  f"=IFERROR(ROUND(AVERAGE('{t}'!{vol_col}2:{vol_col}{tr}),0),0)",
                  KPI_VALUE_FONT_MD)
    row += 3

    # ---- PIPELINE STATUS ----
    draw_subsection_header(ws, row, "  Pipeline Status")
    row += 1
    draw_mini_table(ws, row, 1,
                    ["New", "Contacted", "Mtg Sched", "Mtg Done", "Proposal", "Won", "Lost", "Nurture"],
                    [
                        f"=COUNTIF('{t}'!N2:N{tr},\"New\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Contacted\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Meeting Scheduled\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Meeting Completed\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Proposal Sent\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Won\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Lost\")",
                        f"=COUNTIF('{t}'!N2:N{tr},\"Nurture\")",
                    ])
    row += 3

    # ---- CALL ACTIVITY ----
    draw_subsection_header(ws, row, "  Call Activity")
    row += 1
    outcome_cols = ["W", "Z", "AC", "AF", "AI"]
    total_calls_f = "=" + "+".join(f"COUNTA('{t}'!{c}2:{c}{tr})" for c in outcome_cols)
    connected_f = "=" + "+".join(f"COUNTIF('{t}'!{c}2:{c}{tr},\"Connected*\")" for c in outcome_cols)
    voicemail_f = "=" + "+".join(f"COUNTIF('{t}'!{c}2:{c}{tr},\"Voicemail\")" for c in outcome_cols)
    no_answer_f = "=" + "+".join(f"COUNTIF('{t}'!{c}2:{c}{tr},\"No Answer\")" for c in outcome_cols)
    mtg_set_f = "=" + "+".join(f"COUNTIF('{t}'!{c}2:{c}{tr},\"Connected - Meeting Set\")" for c in outcome_cols)
    connect_rate_f = f"=IFERROR(({connected_f[1:]})/({total_calls_f[1:]}),0)"
    meeting_rate_f = f"=IFERROR(({mtg_set_f[1:]})/({total_calls_f[1:]}),0)"
    draw_mini_table_pct(ws, row, 1,
                        ["Total Calls", "Connected", "Voicemail", "No Answer",
                         "Meetings Set", "Connect Rate", "Meeting Rate", ""],
                        [total_calls_f, connected_f, voicemail_f, no_answer_f,
                         mtg_set_f, connect_rate_f, meeting_rate_f, ""],
                        pct_indices=[5, 6])
    row += 3

    # ---- EMAIL ACTIVITY ----
    draw_subsection_header(ws, row, "  Email Activity")
    row += 1
    et = email_tab
    email_outcome_cols = ["T", "X", "AB"]
    total_emails_f = "=" + "+".join(f"COUNTA('{et}'!{c}2:{c}{tr})" for c in email_outcome_cols)
    opened_f = "=" + "+".join(f"COUNTIF('{et}'!{c}2:{c}{tr},\"Opened\")" for c in email_outcome_cols)
    replied_f = "=" + "+".join(f"COUNTIF('{et}'!{c}2:{c}{tr},\"Replied\")" for c in email_outcome_cols)
    bounced_f = "=" + "+".join(f"COUNTIF('{et}'!{c}2:{c}{tr},\"Bounced\")" for c in email_outcome_cols)
    open_rate_f = f"=IFERROR(({opened_f[1:]})/({total_emails_f[1:]}),0)"
    reply_rate_f = f"=IFERROR(({replied_f[1:]})/({total_emails_f[1:]}),0)"
    draw_mini_table_pct(ws, row, 1,
                        ["Total Sent", "Opened", "Replied", "Bounced",
                         "Open Rate", "Reply Rate", "", ""],
                        [total_emails_f, opened_f, replied_f, bounced_f,
                         open_rate_f, reply_rate_f, "", ""],
                        pct_indices=[4, 5])
    row += 3

    # ---- COLLAGEN USAGE ----
    draw_subsection_header(ws, row, "  Collagen Usage")
    row += 1
    draw_mini_table(ws, row, 1,
                    ["Leads w/ Collagen", "Avg Lg Sheet", "Avg Sm/Md Sheet",
                     "Avg Powder", "Total Lg Sheet", "Total Sm/Md", "Total Powder", "Max Lg Sheet"],
                    [
                        f"=COUNTIF('{t}'!R2:R{tr},\">0\")+COUNTIF('{t}'!S2:S{tr},\">0\")+COUNTIF('{t}'!T2:T{tr},\">0\")",
                        f"=IFERROR(ROUND(AVERAGEIF('{t}'!R2:R{tr},\">0\"),0),0)",
                        f"=IFERROR(ROUND(AVERAGEIF('{t}'!S2:S{tr},\">0\"),0),0)",
                        f"=IFERROR(ROUND(AVERAGEIF('{t}'!T2:T{tr},\">0\"),0),0)",
                        f"=IFERROR(SUM('{t}'!R2:R{tr}),0)",
                        f"=IFERROR(SUM('{t}'!S2:S{tr}),0)",
                        f"=IFERROR(SUM('{t}'!T2:T{tr}),0)",
                        f"=IFERROR(MAX('{t}'!R2:R{tr}),0)",
                    ])
    row += 3

    # ---- INCISION LIKELIHOOD ----
    draw_subsection_header(ws, row, "  Large Incision Likelihood")
    row += 1
    high_f = f"=COUNTIF('{t}'!U2:U{tr},\"High\")"
    medhigh_f = f"=COUNTIF('{t}'!U2:U{tr},\"Medium-High\")"
    medium_f = f"=COUNTIF('{t}'!U2:U{tr},\"Medium\")"
    low_f = f"=COUNTIF('{t}'!U2:U{tr},\"Low\")"
    unlikely_f = f"=COUNTIF('{t}'!U2:U{tr},\"Unlikely\")"
    draw_mini_table_pct(ws, row, 1,
                        ["High 🟢", "Medium-High", "Medium", "Low", "Unlikely",
                         "High+MH", "% High+MH", ""],
                        [high_f, medhigh_f, medium_f, low_f, unlikely_f,
                         f"=({high_f[1:]})+({medhigh_f[1:]})",
                         f"=IFERROR((({high_f[1:]})+({medhigh_f[1:]}))/COUNTA('{t}'!U2:U{tr}),0)",
                         ""],
                        pct_indices=[6])
    row += 4

    return row


# ============================================================
# MAIN
# ============================================================

def main():
    input_file = "Master_Lead_List_Tracker (3).xlsx"

    logger.info("Loading workbook...")
    wb = openpyxl.load_workbook(input_file)

    # Strip external links
    if hasattr(wb, "_external_links"):
        wb._external_links = []

    # Remove old Dashboard
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_properties.tabColor = NAVY

    # Set column widths — 12 cols, each relatively wide for readability
    col_widths = [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Hide gridlines
    ws.sheet_view.showGridLines = False

    # === TITLE ===
    row = 1
    draw_title(ws, row, "ALBACETE MEDDEV — OUTREACH DASHBOARD")
    row += 1
    draw_subtitle(ws, row, "Cold Outreach Performance | Updates automatically on open")
    row += 2

    # === TOP-LEVEL OVERVIEW ===
    draw_section_banner(ws, row, "  📊 OVERVIEW — ALL PORTFOLIOS COMBINED", TEAL_BANNER_FILL)
    row += 2

    # Big combined KPI cards (4 cards, 3 cols each)
    total_leads_f = ("=COUNTA('Call Tracker - JR'!A2:A4442)+"
                    "COUNTA('Call Tracker - S&N'!A2:A10001)+"
                    "COUNTA('Call Tracker - OOS'!A2:A10001)")
    total_pp_f = ("=COUNTIF('Call Tracker - JR'!K2:K4442,\"Private*\")+"
                 "COUNTIF('Call Tracker - S&N'!K2:K10001,\"Private*\")+"
                 "COUNTIF('Call Tracker - OOS'!K2:K10001,\"Private*\")")
    total_hospital_f = ("=COUNTIF('Call Tracker - JR'!K2:K4442,\"Hospital*\")+"
                       "COUNTIF('Call Tracker - S&N'!K2:K10001,\"Hospital*\")+"
                       "COUNTIF('Call Tracker - OOS'!K2:K10001,\"Hospital*\")")
    total_elig_f = ("=COUNTIF('Call Tracker - JR'!P2:P4442,\"Yes\")+"
                   "COUNTIF('Call Tracker - S&N'!P2:P10001,\"Yes\")+"
                   "COUNTIF('Call Tracker - OOS'!P2:P10001,\"Yes\")")

    draw_kpi_card(ws, row, 1, 3, "TOTAL LEADS", total_leads_f, KPI_VALUE_FONT_LG)
    draw_kpi_card(ws, row, 4, 6, "PRIVATE PRACTICES", total_pp_f, KPI_VALUE_FONT_LG)
    draw_kpi_card(ws, row, 7, 9, "HOSPITAL-BASED", total_hospital_f, KPI_VALUE_FONT_LG)
    draw_kpi_card(ws, row, 10, 12, "MICROLYTE ELIGIBLE", total_elig_f, KPI_VALUE_FONT_LG)
    row += 3

    # Portfolio comparison mini-table
    draw_subsection_header(ws, row, "  Portfolio Breakdown")
    row += 1
    draw_mini_table(ws, row, 1,
                    ["Joint Replacement", "Spine & Neuro", "Outside Ortho",
                     "JR Pvt Pract", "S&N Pvt Pract", "OOS Pvt Pract",
                     "JR Microlyte", "S&N Microlyte", "OOS Microlyte",
                     "JR Hospital", "S&N Hospital", "OOS Hospital"],
                    [
                        "=COUNTA('Call Tracker - JR'!A2:A4442)",
                        "=COUNTA('Call Tracker - S&N'!A2:A10001)",
                        "=COUNTA('Call Tracker - OOS'!A2:A10001)",
                        "=COUNTIF('Call Tracker - JR'!K2:K4442,\"Private*\")",
                        "=COUNTIF('Call Tracker - S&N'!K2:K10001,\"Private*\")",
                        "=COUNTIF('Call Tracker - OOS'!K2:K10001,\"Private*\")",
                        "=COUNTIF('Call Tracker - JR'!P2:P4442,\"Yes\")",
                        "=COUNTIF('Call Tracker - S&N'!P2:P10001,\"Yes\")",
                        "=COUNTIF('Call Tracker - OOS'!P2:P10001,\"Yes\")",
                        "=COUNTIF('Call Tracker - JR'!K2:K4442,\"Hospital*\")",
                        "=COUNTIF('Call Tracker - S&N'!K2:K10001,\"Hospital*\")",
                        "=COUNTIF('Call Tracker - OOS'!K2:K10001,\"Hospital*\")",
                    ])
    row += 4

    # === JR PORTFOLIO ===
    row = build_portfolio_section(
        ws, row, "JOINT REPLACEMENT — JR",
        JR_BANNER_FILL, "Call Tracker - JR", "Email Tracker - JR",
        "4442", "Q", "Avg Joint Vol", KPI_VALUE_FONT_ACCENT_JR,
    )

    # === S&N PORTFOLIO ===
    row += 1
    row = build_portfolio_section(
        ws, row, "SPINE & NEURO — S&N",
        SN_BANNER_FILL, "Call Tracker - S&N", "Email Tracker - S&N",
        "10001", "Q", "Avg Spine Vol", KPI_VALUE_FONT_ACCENT_SN,
    )

    # === OOS PORTFOLIO ===
    row += 1
    row = build_portfolio_section(
        ws, row, "OUTSIDE ORTHO & SPINE — OOS",
        OOS_BANNER_FILL, "Call Tracker - OOS", "Email Tracker - OOS",
        "10001", "Q", "Avg Proc Vol", KPI_VALUE_FONT_ACCENT_OOS,
    )

    # === READING GUIDE ===
    row += 1
    draw_section_banner(ws, row, "  ❓ HOW TO READ THIS DASHBOARD", TEAL_BANNER_FILL)
    row += 2
    guide_notes = [
        ("Total Leads", "Count of unique physicians in each portfolio — the pool you can call/email."),
        ("Private Practice vs Hospital-Based", "Private practices are your primary targets — easier to reach decision-makers."),
        ("Microlyte Eligible", "Leads in non-LCD states where you can pitch Microlyte SAM alongside ProPacks. LCD states get ProPacks-only emails."),
        ("Tier 1/2/3/4", "Drive-time tiers from NYC Midtown. Tier 1 = 0–30 min, Tier 2 = 30–60 min. Start with Tier 1-2 Private Practices for in-person visits."),
        ("Pipeline Status", "Updates automatically as you change the 'Lead Status' dropdown in each Call Tracker tab."),
        ("Call Activity", "Aggregated across all 5 call attempts per lead. Populated from 'Call 1-5 Outcome' dropdowns."),
        ("Email Activity", "Aggregated across all 3 email attempts. Populated from 'Email 1-3 Outcome' dropdowns."),
        ("Collagen Usage", "Procedure volume of current collagen sheet/powder products — signals existing wound-product buyers."),
        ("Incision Likelihood", "Predicted surgical incision size based on procedure mix. High/Medium-High = best Microlyte targets."),
    ]
    for label, desc in guide_notes:
        ws.cell(row=row, column=1, value=f"  {label}").font = Font(name="Calibri", size=10, bold=True, color=DARK_GRAY)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=12)
        desc_cell = ws.cell(row=row, column=2, value=desc)
        desc_cell.font = Font(name="Calibri", size=10, color="595959")
        desc_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 22
        row += 1

    # Freeze panes so title stays visible
    ws.freeze_panes = "A4"

    wb.save(input_file)
    logger.info(f"Dashboard rebuilt in {input_file}")


if __name__ == "__main__":
    main()
