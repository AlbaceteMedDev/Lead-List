"""Self-contained HTML dashboard for call/email activity tracking.

Renders a single static .html file (Chart.js via CDN) that shows pipeline
totals, call/email activity KPIs derived from the activity cache, and a
filterable lead table with lead-status / product-line / tier filters.
"""

from __future__ import annotations

import json
import logging
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

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

LEAD_STATUS_ORDER = [
    "New", "Queued", "Attempting Contact", "Connected", "Interested",
    "Meeting Booked", "Nurture", "Not Interested", "Do Not Contact",
    "Closed - Won", "Closed - Lost",
]


def _parse_date(value: object) -> Optional[date]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def _counts(series: pd.Series, keys: Optional[list[str]] = None) -> dict:
    values = series.fillna("").astype(str)
    counted = values.value_counts().to_dict()
    if keys is not None:
        return {k: int(counted.get(k, 0)) for k in keys}
    return {str(k): int(v) for k, v in counted.items() if k != ""}


def _activity_stats(df: pd.DataFrame) -> dict:
    today = date.today()
    week_ago = today - timedelta(days=7)

    call_totals = 0
    call_connected = 0
    call_this_week = 0
    call_today = 0
    email_totals = 0
    email_this_week = 0
    calls_by_date: dict[str, int] = {}
    call_outcome_counts: dict[str, int] = {}
    email_outcome_counts: dict[str, int] = {}

    for i in range(1, 6):
        date_col = f"Call {i} Date"
        outcome_col = f"Call {i} Outcome"
        if date_col not in df.columns:
            continue
        for d_raw, outcome in zip(df[date_col].fillna(""), df.get(outcome_col, pd.Series([""] * len(df))).fillna("")):
            parsed = _parse_date(d_raw)
            if parsed is None:
                continue
            call_totals += 1
            iso = parsed.isoformat()
            calls_by_date[iso] = calls_by_date.get(iso, 0) + 1
            if parsed == today:
                call_today += 1
            if parsed >= week_ago:
                call_this_week += 1
            outcome_s = str(outcome).strip() or "Unspecified"
            call_outcome_counts[outcome_s] = call_outcome_counts.get(outcome_s, 0) + 1
            if outcome_s.lower().startswith("connected") or outcome_s == "Meeting Booked" or outcome_s == "Callback Requested":
                call_connected += 1

    for i in range(1, 4):
        date_col = f"Email {i} Date"
        outcome_col = f"Email {i} Outcome"
        if date_col not in df.columns:
            continue
        for d_raw, outcome in zip(df[date_col].fillna(""), df.get(outcome_col, pd.Series([""] * len(df))).fillna("")):
            parsed = _parse_date(d_raw)
            if parsed is None:
                continue
            email_totals += 1
            if parsed >= week_ago:
                email_this_week += 1
            outcome_s = str(outcome).strip() or "Unspecified"
            email_outcome_counts[outcome_s] = email_outcome_counts.get(outcome_s, 0) + 1

    pickup_rate = (call_connected / call_totals * 100.0) if call_totals else 0.0
    meetings = int((df.get("Lead Status", pd.Series(dtype=str)) == "Meeting Booked").sum()) if len(df) else 0

    return {
        "call_totals": call_totals,
        "call_connected": call_connected,
        "call_today": call_today,
        "call_this_week": call_this_week,
        "pickup_rate": round(pickup_rate, 1),
        "email_totals": email_totals,
        "email_this_week": email_this_week,
        "meetings_booked": meetings,
        "calls_by_date": calls_by_date,
        "call_outcomes": call_outcome_counts,
        "email_outcomes": email_outcome_counts,
    }


def build_stats(df: pd.DataFrame) -> dict:
    total = len(df)
    return {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "total_leads": total,
        "tier_counts": _counts(df.get("Tier", pd.Series(dtype=str)), TIER_ORDER),
        "practice_type_counts": _counts(df.get("Practice Type", pd.Series(dtype=str))),
        "product_line_counts": _counts(df.get("Product Line", pd.Series(dtype=str)), ["JR", "S&N", "OOS"]),
        "lead_status_counts": _counts(df.get("Lead Status", pd.Series(dtype=str)), LEAD_STATUS_ORDER),
        "target_tier_counts": _counts(df.get("Target Tier", pd.Series(dtype=str)), ["A+", "A", "B", "C", "D"]),
        "phone_status_counts": _counts(df.get("Phone Status", pd.Series(dtype=str)),
                                       ["Verified", "Added from NPPES", "Updated (NPPES differs)", "Missing"]),
        "email_status_counts": _counts(df.get("Email Status", pd.Series(dtype=str))),
        "mac_counts": _counts(df.get("MAC Jurisdiction", pd.Series(dtype=str))),
        "microlyte_counts": _counts(df.get("Microlyte Eligible", pd.Series(dtype=str)), ["Yes", "No", "Unknown"]),
        "incision_counts": _counts(df.get("Lg Incision Likelihood", pd.Series(dtype=str)),
                                   ["High", "Medium-High", "Medium", "Low"]),
        "activity": _activity_stats(df),
    }


def _dataset(df: pd.DataFrame) -> list[dict]:
    cols = [
        "HCP NPI", "First Name", "Last Name", "Credential", "Specialty",
        "Email", "Email Status", "Verified Phone", "Phone Status",
        "Primary Site of Care", "Practice Type", "City", "State",
        "Tier", "MAC Jurisdiction", "Microlyte Eligible",
        "Address 1", "Postal Code",
        "Other Locations", "Location Count",
        "Practice Match", "NPPES Practice Address",
        "Alternate Phones",
        "Product Line", "Lead Priority", "Lead Status", "Target Tier", "Target Score",
        "Lg Incision Likelihood", "Next Action", "Next Action Date",
        "Last Touch Date", "Touch Count",
        "Joint Repl Vol", "Knee Vol", "Hip Vol", "Shoulder Vol",
        "Open Spine Vol", "Open Ortho Vol", "Procedure Vol",
        "Lg Collagen Vol", "Sm/Md Collagen Vol", "Collagen Powder Vol",
        "Total Collagen Vol", "Wound Care DME Vol", "All DME Vol",
        "Why Target?", "Target Tier Reason", "Best Approach",
        "Subject Line", "Draft Email", "Email Track",
    ]
    for i in range(1, 6):
        cols += [f"Call {i} Date", f"Call {i} Outcome", f"Call {i} Notes"]
    for i in range(1, 4):
        cols += [f"Email {i} Date", f"Email {i} Subject", f"Email {i} Outcome", f"Email {i} Notes"]
    present = [c for c in cols if c in df.columns]
    out = df[present].fillna("").astype(str).to_dict(orient="records")
    return out


def write_dashboard(df: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    stats = build_stats(df)
    rows = _dataset(df)
    payload = json.dumps({"stats": stats, "rows": rows}, separators=(",", ":"))
    html = _HTML_TEMPLATE.replace("__PAYLOAD__", payload)
    output_path.write_text(html, encoding="utf-8")
    log.info("Wrote dashboard: %s", output_path)
    return output_path


_HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Albacete MedDev - Outreach Tracker</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root { --navy:#1b4f72; --ink:#1c2833; --muted:#566573; --line:#d5d8dc; --bg:#f4f6f7; --card:#fff; --accent:#1b9d7a; --warn:#d68910; }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: -apple-system, Segoe UI, Arial, sans-serif; color: var(--ink); background: var(--bg); }
  header { background: var(--navy); color: #fff; padding: 12px 24px; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
  header h1 { margin: 0; font-size: 17px; font-weight: 600; }
  header .meta { color: #d5dbe1; font-size: 12px; flex: 1; }
  header button, header .hdr-btn { background: rgba(255,255,255,0.14); color: #fff; border: 1px solid rgba(255,255,255,0.3); border-radius: 4px; padding: 6px 12px; font-size: 12px; cursor: pointer; text-decoration: none; display: inline-block; }
  header button:hover, header .hdr-btn:hover { background: rgba(255,255,255,0.25); }
  header .hdr-btn { background: #27ae60; border-color: #27ae60; font-weight: 600; }
  header .hdr-btn:hover { background: #229954; }
  header .pending { background: var(--warn); padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: 600; display: none; }
  header .pending.show { display: inline-block; }
  main { max-width: 1600px; margin: 0 auto; padding: 16px; }
  .tabs { display: flex; gap: 4px; margin-bottom: 14px; border-bottom: 2px solid var(--navy); }
  .tab { padding: 8px 16px; cursor: pointer; background: #eaeded; border: 1px solid var(--line); border-bottom: none; border-radius: 6px 6px 0 0; font-size: 13px; font-weight: 500; color: var(--muted); }
  .tab.active { background: var(--navy); color: #fff; border-color: var(--navy); }
  .panel { display: none; }
  .panel.active { display: block; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 10px; margin-bottom: 16px; }
  .kpi { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 11px 13px; }
  .kpi .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--muted); }
  .kpi .value { font-size: 22px; font-weight: 600; margin-top: 2px; color: var(--navy); }
  .kpi .sub { font-size: 11px; color: var(--muted); margin-top: 2px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap: 12px; margin-bottom: 16px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 14px; }
  .card h2 { margin: 0 0 10px 0; font-size: 13px; color: var(--navy); font-weight: 600; }
  .card canvas { max-height: 240px; }
  .controls { display: flex; flex-wrap: wrap; gap: 10px; align-items: flex-end; background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 10px; margin-bottom: 10px; }
  .controls label { font-size: 11px; color: var(--muted); display: flex; flex-direction: column; gap: 3px; }
  .controls select, .controls input { font: inherit; padding: 5px 8px; border: 1px solid var(--line); border-radius: 4px; min-width: 140px; }
  .table-wrap { background: var(--card); border: 1px solid var(--line); border-radius: 6px; overflow: auto; max-height: 68vh; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th, td { padding: 5px 8px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }
  th { background: var(--navy); color: #fff; font-weight: 600; position: sticky; top: 0; z-index: 1; }
  tr.clickable { cursor: pointer; }
  tr.clickable:hover td { background: #eaf4fb; }
  tr.edited td:first-child { border-left: 3px solid var(--warn); }
  .pill { display: inline-block; padding: 1px 7px; border-radius: 9px; font-size: 10px; font-weight: 600; }
  .pill.green { background: #d5f5e3; color: #186a3b; }
  .pill.red { background: #fadbd8; color: #922b21; }
  .pill.amber { background: #fef9e7; color: #9a7d0a; }
  .pill.navy { background: #d6eaf8; color: var(--navy); }
  .muted { color: var(--muted); font-size: 11px; }

  .drawer-bg { position: fixed; inset: 0; background: rgba(0,0,0,0.3); display: none; z-index: 10; }
  .drawer-bg.show { display: block; }
  .drawer { position: fixed; right: 0; top: 0; bottom: 0; width: min(560px, 100%); background: #fff; box-shadow: -4px 0 20px rgba(0,0,0,0.18); z-index: 11; transform: translateX(100%); transition: transform 0.22s ease; display: flex; flex-direction: column; }
  .drawer.show { transform: translateX(0); }
  .drawer header { position: relative; background: var(--navy); padding: 14px 18px; display: block; }
  .drawer header .close { position: absolute; right: 14px; top: 12px; background: transparent; border: none; color: #fff; font-size: 22px; cursor: pointer; line-height: 1; }
  .drawer h2 { margin: 0; color: #fff; font-size: 16px; font-weight: 600; }
  .drawer .subhead { margin-top: 3px; color: #d5dbe1; font-size: 12px; }
  .drawer-body { overflow-y: auto; padding: 14px 18px; flex: 1; }
  .drawer-body section { margin-bottom: 18px; }
  .drawer-body section h3 { font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); margin: 0 0 8px 0; border-bottom: 1px solid var(--line); padding-bottom: 4px; }
  .field { display: grid; grid-template-columns: 110px 1fr; gap: 8px; margin-bottom: 6px; align-items: center; }
  .field label { font-size: 12px; color: var(--muted); }
  .field input, .field select, .field textarea { font: inherit; font-size: 12px; padding: 5px 8px; border: 1px solid var(--line); border-radius: 4px; width: 100%; }
  .field textarea { min-height: 44px; resize: vertical; }
  .round-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; border: 1px dashed var(--line); padding: 8px; border-radius: 4px; margin-bottom: 8px; }
  .round-grid h4 { grid-column: 1 / -1; margin: 0 0 4px 0; font-size: 11px; color: var(--navy); }
  .round-grid .full { grid-column: 1 / -1; }
  .drawer-footer { border-top: 1px solid var(--line); padding: 10px 18px; display: flex; gap: 8px; }
  .drawer-footer button { font: inherit; font-size: 12px; padding: 7px 14px; border-radius: 4px; border: 1px solid var(--navy); background: var(--navy); color: #fff; cursor: pointer; }
  .drawer-footer button.secondary { background: #fff; color: var(--navy); }
  .drawer-footer .spacer { flex: 1; }
  .drawer-footer .quick { font-size: 11px; color: var(--muted); }

  .btn-draft { font-size: 10px; padding: 3px 8px; background: #fff; border: 1px solid var(--navy); color: var(--navy); border-radius: 3px; cursor: pointer; font-weight: 600; }
  .btn-draft:hover { background: var(--navy); color: #fff; }

  .quick-log { display: flex; gap: 8px; margin-bottom: 10px; }
  .quick-btn { font: inherit; font-size: 13px; font-weight: 600; padding: 10px 14px; border-radius: 5px; border: 1px solid var(--navy); background: #fff; color: var(--navy); cursor: pointer; flex: 1; }
  .quick-btn:hover { background: #e8f3fb; }
  .quick-btn.primary { background: var(--navy); color: #fff; }
  .quick-btn.primary:hover { background: #143b57; }
  .quick-log-form { background: #f7fafc; border: 1px solid var(--line); border-radius: 5px; padding: 10px; margin-top: 6px; }
  .quick-log-form .field { grid-template-columns: 80px 1fr; }
  .date-quick { display: flex; gap: 4px; align-items: center; }
  .date-quick input[type="date"] { flex: 1; }
  .date-btn { font: inherit; font-size: 11px; padding: 4px 8px; border: 1px solid var(--line); background: #fff; border-radius: 3px; cursor: pointer; }
  .date-btn:hover { background: #e8f3fb; border-color: var(--navy); }
  .quick-actions { display: flex; gap: 6px; justify-content: flex-end; margin-top: 6px; }
  .quick-actions button { font: inherit; font-size: 12px; padding: 6px 12px; border-radius: 4px; border: 1px solid var(--navy); background: #fff; color: var(--navy); cursor: pointer; }
  .quick-actions button.primary { background: var(--navy); color: #fff; }

  .draft-bg { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: none; z-index: 20; }
  .draft-bg.show { display: block; }
  .draft-modal { position: fixed; top: 8vh; left: 50%; transform: translateX(-50%); width: min(720px, 92%); max-height: 84vh; background: #fff; border-radius: 8px; box-shadow: 0 10px 40px rgba(0,0,0,0.25); z-index: 21; display: none; flex-direction: column; }
  .draft-modal.show { display: flex; }
  .draft-modal header { background: var(--navy); color: #fff; padding: 12px 18px; border-radius: 8px 8px 0 0; position: relative; }
  .draft-modal header .close { position: absolute; right: 12px; top: 8px; background: transparent; border: none; color: #fff; font-size: 22px; cursor: pointer; }
  .draft-modal h2 { margin: 0; font-size: 14px; font-weight: 600; }
  .draft-modal .subhead { margin-top: 2px; font-size: 12px; color: #d5dbe1; }
  .draft-body { padding: 14px 18px; overflow-y: auto; }
  .draft-body .field-row { display: grid; grid-template-columns: 80px 1fr; gap: 8px; margin-bottom: 8px; align-items: center; }
  .draft-body .field-row label { font-size: 12px; color: var(--muted); }
  .draft-body input, .draft-body textarea { font: inherit; font-size: 12px; padding: 6px 8px; border: 1px solid var(--line); border-radius: 4px; width: 100%; }
  .draft-body textarea { min-height: 320px; font-family: ui-monospace, Menlo, Consolas, monospace; resize: vertical; white-space: pre-wrap; }
  .draft-footer { border-top: 1px solid var(--line); padding: 10px 18px; display: flex; align-items: center; gap: 8px; }
  .draft-footer button { font: inherit; font-size: 12px; padding: 7px 14px; border-radius: 4px; border: 1px solid var(--navy); background: var(--navy); color: #fff; cursor: pointer; }
  .draft-footer button:nth-of-type(1) { background: #fff; color: var(--navy); }
  .draft-footer .spacer { flex: 1; }
</style>
</head>
<body>
<header>
  <h1>Albacete MedDev - Outreach Tracker</h1>
  <span class="meta" id="gen-meta"></span>
  <span class="pending" id="pending-count">0 unsaved edits</span>
  <a href="lead_list.xlsx" download class="hdr-btn">Download Excel Lead List</a>
  <button id="btn-download">Download Updates</button>
  <button id="btn-import">Import Updates</button>
  <button id="btn-discard">Discard</button>
  <input type="file" id="file-import" accept=".json,application/json" style="display:none" />
</header>
<main>
  <div class="tabs">
    <div class="tab active" data-panel="overview">Overview</div>
    <div class="tab" data-panel="activity">Call / Email Activity</div>
    <div class="tab" data-panel="leads">Leads (click to log)</div>
  </div>

  <section class="panel active" id="panel-overview">
    <div class="kpis" id="kpis-overview"></div>
    <div class="grid">
      <div class="card"><h2>Leads by Tier</h2><canvas id="chart-tier"></canvas></div>
      <div class="card"><h2>Lead Status</h2><canvas id="chart-status"></canvas></div>
      <div class="card"><h2>Target Tier</h2><canvas id="chart-target"></canvas></div>
      <div class="card"><h2>Microlyte Eligibility</h2><canvas id="chart-microlyte"></canvas></div>
      <div class="card"><h2>Product Line Split</h2><canvas id="chart-line"></canvas></div>
      <div class="card"><h2>Incision Likelihood</h2><canvas id="chart-incision"></canvas></div>
      <div class="card"><h2>Phone Verification</h2><canvas id="chart-phone"></canvas></div>
      <div class="card"><h2>Email Enrichment</h2><canvas id="chart-email"></canvas></div>
    </div>
  </section>

  <section class="panel" id="panel-activity">
    <div class="kpis" id="kpis-activity"></div>
    <div class="grid">
      <div class="card"><h2>Calls per Day</h2><canvas id="chart-calls-day"></canvas></div>
      <div class="card"><h2>Call Outcomes</h2><canvas id="chart-call-outcomes"></canvas></div>
      <div class="card"><h2>Email Outcomes</h2><canvas id="chart-email-outcomes"></canvas></div>
      <div class="card"><h2>Lead Status Funnel</h2><canvas id="chart-status-funnel"></canvas></div>
    </div>
  </section>

  <section class="panel" id="panel-leads">
    <div class="controls">
      <label>Product Line<select id="filter-line"><option value="">All</option></select></label>
      <label>Tier<select id="filter-tier"><option value="">All</option></select></label>
      <label>Lead Status<select id="filter-status"><option value="">All</option></select></label>
      <label>Target Tier<select id="filter-target"><option value="">All</option></select></label>
      <label>Microlyte<select id="filter-microlyte"><option value="">All</option></select></label>
      <label style="flex:1;">Search<input id="filter-search" placeholder="name, practice, email, NPI..." /></label>
      <span class="muted" id="row-count"></span>
    </div>
    <div class="table-wrap">
      <table id="leads-table">
        <thead>
          <tr>
            <th>NPI</th><th>Name</th><th>Specialty</th><th>Practice</th><th>Type</th><th>City,ST</th>
            <th>Line</th><th>Tier</th><th>Target</th><th>Status</th>
            <th>Phone</th><th>Email</th>
            <th>Proc Vol</th><th>Collagen Vol</th>
            <th>Last Touch</th><th>Touches</th><th>Next Action</th><th>Draft</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>
    <p class="muted" style="margin-top:8px;">Click any row to log a call or email. Edits are saved in your browser. When you are done, click "Download Updates" and commit the file to <code>data/cache/activity.json</code> or hand it off to Gabe.</p>
  </section>
</main>

<div class="draft-bg" id="draft-bg"></div>
<div class="draft-modal" id="draft-modal">
  <header>
    <button class="close" id="draft-close">&times;</button>
    <h2 id="draft-title">Draft Email</h2>
    <div class="subhead" id="draft-subhead"></div>
  </header>
  <div class="draft-body">
    <div class="field-row"><label>To</label><input id="draft-to" readonly /></div>
    <div class="field-row"><label>Subject</label><input id="draft-subject" readonly /></div>
    <textarea id="draft-body-text" readonly></textarea>
  </div>
  <div class="draft-footer">
    <span class="muted" id="draft-track"></span>
    <div class="spacer"></div>
    <button id="draft-copy">Copy to Clipboard</button>
    <button id="draft-mailto">Open in Mail</button>
  </div>
</div>

<div class="drawer-bg" id="drawer-bg"></div>
<aside class="drawer" id="drawer">
  <header>
    <button class="close" id="drawer-close">&times;</button>
    <h2 id="drawer-title">Lead</h2>
    <div class="subhead" id="drawer-subhead"></div>
  </header>
  <div class="drawer-body" id="drawer-body"></div>
  <div class="drawer-footer">
    <span class="quick" id="drawer-quick-info"></span>
    <div class="spacer"></div>
    <button class="secondary" id="drawer-cancel">Close</button>
    <button id="drawer-save">Save to Browser</button>
  </div>
</aside>

<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
const DATA = JSON.parse(document.getElementById('payload').textContent);
const stats = DATA.stats;
const act = stats.activity;
const rows = DATA.rows;

const LS_KEY = 'albacete_activity_edits_v1';
let edits = {};
try { edits = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e) { edits = {}; }

const LEAD_STATUSES = ['','New','Queued','Attempting Contact','Connected','Interested','Meeting Booked','Nurture','Not Interested','Do Not Contact','Closed - Won','Closed - Lost'];
const CALL_OUTCOMES = ['','No Answer','Voicemail','Gatekeeper - Declined','Gatekeeper - Gave Info','Wrong Number','Bad Number','Do Not Call','Connected - Not Interested','Connected - Interested','Meeting Booked','Callback Requested'];
const EMAIL_OUTCOMES = ['','Sent','Bounced','Opened','Replied - Interested','Replied - Not Interested','Meeting Booked','Unsubscribed'];
const DM_OPTIONS = ['','Yes','No','Unknown'];

document.getElementById('gen-meta').textContent = 'Generated ' + stats.generated_at;

document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(x => x.classList.remove('active'));
  t.classList.add('active');
  document.getElementById('panel-' + t.dataset.panel).classList.add('active');
}));

function effective(row) {
  const out = Object.assign({}, row);
  const patch = edits[row['HCP NPI']];
  if (patch) Object.assign(out, patch);
  return out;
}

function updatePendingBadge() {
  const n = Object.keys(edits).length;
  const badge = document.getElementById('pending-count');
  badge.textContent = n + ' unsaved edit' + (n === 1 ? '' : 's');
  badge.classList.toggle('show', n > 0);
}

function kpi(container, label, value, sub) {
  const el = document.createElement('div'); el.className = 'kpi';
  el.innerHTML = '<div class="label">' + label + '</div><div class="value">' + value + '</div><div class="sub">' + (sub || '') + '</div>';
  container.appendChild(el);
}

const verifiedEmail = Object.entries(stats.email_status_counts).filter(([k]) => k.indexOf('Verified') !== -1 || k.indexOf('Inferred') !== -1).reduce((a, [, v]) => a + v, 0);
const missingEmail = stats.email_status_counts['Missing'] || 0;
const verifiedPhone = (stats.phone_status_counts['Verified'] || 0) + (stats.phone_status_counts['Added from NPPES'] || 0);
const t12 = (stats.tier_counts['Tier 1 (0-30 min)'] || 0) + (stats.tier_counts['Tier 2 (30-60 min)'] || 0);

const overviewCt = document.getElementById('kpis-overview');
kpi(overviewCt, 'Total Leads', stats.total_leads.toLocaleString(), stats.generated_at);
kpi(overviewCt, 'Target A+/A', (stats.target_tier_counts['A+'] || 0) + (stats.target_tier_counts['A'] || 0), 'Top-priority leads');
kpi(overviewCt, 'Phone Ready', verifiedPhone.toLocaleString(), (stats.phone_status_counts['Missing'] || 0).toLocaleString() + ' missing');
kpi(overviewCt, 'Email Ready', verifiedEmail.toLocaleString(), missingEmail.toLocaleString() + ' missing');
kpi(overviewCt, 'Microlyte Eligible', (stats.microlyte_counts['Yes'] || 0).toLocaleString(), (stats.microlyte_counts['No'] || 0).toLocaleString() + ' LCD-blocked');
kpi(overviewCt, 'Drive <=1h', t12.toLocaleString(), 'Tier 1+2 private practice');
kpi(overviewCt, 'High Incision', (stats.incision_counts['High'] || 0) + (stats.incision_counts['Medium-High'] || 0), 'Best Microlyte / ProPacks fit');
kpi(overviewCt, 'Hospital-Based', (stats.tier_counts['Hospital-Based'] || 0).toLocaleString(), 'Separate workflow');

const activityCt = document.getElementById('kpis-activity');
kpi(activityCt, 'Calls Today', act.call_today, act.call_this_week + ' this week');
kpi(activityCt, 'Calls (total)', act.call_totals, act.call_connected + ' connected');
kpi(activityCt, 'Pickup Rate', act.pickup_rate + '%', 'Connected / total calls');
kpi(activityCt, 'Meetings Booked', act.meetings_booked, 'Lead Status = Meeting Booked');
kpi(activityCt, 'Emails Sent', act.email_totals, act.email_this_week + ' this week');
kpi(activityCt, 'Call Days', Object.keys(act.calls_by_date).length, 'See chart');

const palette = ['#1b4f72','#2874a6','#2e86c1','#5dade2','#85c1e9','#aed6f1','#d6eaf8','#c0392b','#d68910','#27ae60','#1b9d7a'];
function chart(id, type, labels, data, opts) {
  const ctx = document.getElementById(id); if (!ctx) return;
  new Chart(ctx, {
    type: type,
    data: { labels: labels, datasets: [{ data: data, backgroundColor: labels.map((_, i) => palette[i % palette.length]), borderWidth: 0 }] },
    options: Object.assign({ responsive: true, plugins: { legend: { display: type !== 'bar', position: 'right', labels: { font: { size: 10 } } } } }, opts || {}),
  });
}

chart('chart-tier', 'bar', Object.keys(stats.tier_counts), Object.values(stats.tier_counts), { indexAxis: 'y', plugins: { legend: { display: false } } });
chart('chart-status', 'bar', Object.keys(stats.lead_status_counts), Object.values(stats.lead_status_counts), { indexAxis: 'y', plugins: { legend: { display: false } } });
chart('chart-target', 'bar', Object.keys(stats.target_tier_counts), Object.values(stats.target_tier_counts), { plugins: { legend: { display: false } } });
chart('chart-microlyte', 'doughnut', Object.keys(stats.microlyte_counts), Object.values(stats.microlyte_counts));
chart('chart-line', 'doughnut', Object.keys(stats.product_line_counts), Object.values(stats.product_line_counts));
chart('chart-incision', 'bar', Object.keys(stats.incision_counts), Object.values(stats.incision_counts), { plugins: { legend: { display: false } } });
chart('chart-phone', 'doughnut', Object.keys(stats.phone_status_counts), Object.values(stats.phone_status_counts));
chart('chart-email', 'doughnut', Object.keys(stats.email_status_counts), Object.values(stats.email_status_counts));
const callDates = Object.keys(act.calls_by_date).sort();
chart('chart-calls-day', 'bar', callDates, callDates.map(d => act.calls_by_date[d]), { plugins: { legend: { display: false } } });
chart('chart-call-outcomes', 'bar', Object.keys(act.call_outcomes), Object.values(act.call_outcomes), { indexAxis: 'y', plugins: { legend: { display: false } } });
chart('chart-email-outcomes', 'bar', Object.keys(act.email_outcomes), Object.values(act.email_outcomes), { indexAxis: 'y', plugins: { legend: { display: false } } });
const funnelKeys = ['New','Attempting Contact','Connected','Interested','Meeting Booked','Closed - Won'];
chart('chart-status-funnel', 'bar', funnelKeys, funnelKeys.map(k => stats.lead_status_counts[k] || 0), { plugins: { legend: { display: false } } });

function uniq(col) { const s = new Set(); for (const r of rows) { if (r[col]) s.add(r[col]); } return [...s].sort(); }
function fillSelect(id, values) { const el = document.getElementById(id); for (const v of values) { const o = document.createElement('option'); o.value = v; o.textContent = v; el.appendChild(o); } }
fillSelect('filter-line', uniq('Product Line'));
fillSelect('filter-tier', uniq('Tier'));
fillSelect('filter-status', uniq('Lead Status'));
fillSelect('filter-target', uniq('Target Tier'));
fillSelect('filter-microlyte', uniq('Microlyte Eligible'));

function pill(text, kind) { return '<span class="pill ' + kind + '">' + text + '</span>'; }
function statusPill(s) {
  if (!s) return '';
  if (s === 'Meeting Booked' || s === 'Closed - Won' || s === 'Interested') return pill(s, 'green');
  if (s === 'Not Interested' || s === 'Do Not Contact' || s === 'Closed - Lost') return pill(s, 'red');
  if (s === 'Connected' || s === 'Attempting Contact' || s === 'Callback Requested') return pill(s, 'amber');
  return pill(s, 'navy');
}
function targetPill(s) {
  if (s === 'A+' || s === 'A') return pill(s, 'green');
  if (s === 'B') return pill(s, 'amber');
  if (s === 'C' || s === 'D') return pill(s, 'navy');
  return s || '';
}
function escapeHtml(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]); }

function latestTouch(row) {
  const dates = [];
  for (let i = 1; i <= 5; i++) if (row['Call ' + i + ' Date']) dates.push(row['Call ' + i + ' Date']);
  for (let i = 1; i <= 3; i++) if (row['Email ' + i + ' Date']) dates.push(row['Email ' + i + ' Date']);
  dates.sort();
  return { last: dates[dates.length - 1] || '', count: dates.length };
}

function num(v) { const n = parseFloat(String(v || '0').replace(/,/g, '')); return isFinite(n) ? n : 0; }

function procedureVolume(row) {
  const line = row['Product Line'];
  if (line === 'JR') return num(row['Joint Repl Vol']);
  if (line === 'S&N') return num(row['Open Spine Vol']) || num(row['Joint Repl Vol']);
  if (line === 'OOS') return num(row['Procedure Vol']) || num(row['Open Ortho Vol']);
  return num(row['Joint Repl Vol']) || num(row['Open Spine Vol']) || num(row['Procedure Vol']);
}

function phonePillCell(row) {
  const phone = row['Verified Phone'] || '';
  const status = row['Phone Status'] || '';
  if (!phone) return '<span class="pill red">No phone</span>';
  let kind = 'navy';
  let tag = 'Unverified';
  if (status === 'Verified') { kind = 'green'; tag = 'NPPES-verified'; }
  else if (status === 'Added from NPPES') { kind = 'amber'; tag = 'NPPES-added'; }
  else if (status === 'Updated (NPPES differs)') { kind = 'amber'; tag = 'NPPES-updated'; }
  else if (status === 'Missing') { kind = 'navy'; tag = 'AcuityMD (unverified)'; }
  return escapeHtml(phone) + '<br><span class="pill ' + kind + '">' + tag + '</span>';
}

function emailPillCell(row) {
  const email = row['Email'] || '';
  const status = row['Email Status'] || '';
  if (!email) return '<span class="pill red">Missing</span>';
  let kind = 'navy';
  let tag = 'Review';
  if (status.indexOf('Verified') !== -1) { kind = 'green'; tag = 'High'; }
  else if (status === 'Hospital System Email') { kind = 'green'; tag = 'Hospital'; }
  else if (status.indexOf('Inferred') !== -1) { kind = 'amber'; tag = 'Inferred (med)'; }
  else if (status === 'Personal Email (name match)') { kind = 'amber'; tag = 'Personal'; }
  else if (status === 'Personal Email (no name match)') { kind = 'red'; tag = 'Personal?'; }
  else if (status === 'Generic Office Email') { kind = 'red'; tag = 'Generic'; }
  else if (status === 'Practice Email (review recommended)') { kind = 'amber'; tag = 'Review'; }
  else if (status === 'Missing') { kind = 'red'; tag = 'Missing'; }
  return escapeHtml(email) + '<br><span class="pill ' + kind + '">' + tag + '</span>';
}

function collagenVolume(row) {
  if (row['Total Collagen Vol']) return num(row['Total Collagen Vol']);
  return num(row['Lg Collagen Vol']) + num(row['Sm/Md Collagen Vol']) + num(row['Collagen Powder Vol']);
}

function render() {
  const line = document.getElementById('filter-line').value;
  const tier = document.getElementById('filter-tier').value;
  const status = document.getElementById('filter-status').value;
  const target = document.getElementById('filter-target').value;
  const micro = document.getElementById('filter-microlyte').value;
  const q = document.getElementById('filter-search').value.trim().toLowerCase();

  const filtered = rows.map(r => ({ raw: r, eff: effective(r) })).filter(({ eff }) => {
    if (line && eff['Product Line'] !== line) return false;
    if (tier && eff['Tier'] !== tier) return false;
    if (status && eff['Lead Status'] !== status) return false;
    if (target && eff['Target Tier'] !== target) return false;
    if (micro && eff['Microlyte Eligible'] !== micro) return false;
    if (q) {
      const hay = [eff['First Name'], eff['Last Name'], eff['Primary Site of Care'], eff['Email'], eff['HCP NPI']].join(' ').toLowerCase();
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  });
  document.getElementById('row-count').textContent = filtered.length + ' of ' + rows.length;

  const tbody = document.querySelector('#leads-table tbody');
  tbody.innerHTML = '';
  const limit = Math.min(filtered.length, 500);
  for (let i = 0; i < limit; i++) {
    const { raw, eff } = filtered[i];
    const tr = document.createElement('tr');
    tr.className = 'clickable' + (edits[raw['HCP NPI']] ? ' edited' : '');
    tr.dataset.npi = raw['HCP NPI'];
    const t = latestTouch(eff);
    const practiceTypeBadge = eff['Practice Type'] === 'Hospital-Based'
      ? '<span class="pill amber">Hospital</span>'
      : (eff['Practice Type'] === 'Private Practice' ? '<span class="pill green">Private</span>' : escapeHtml(eff['Practice Type'] || ''));
    const cells = [
      escapeHtml(eff['HCP NPI'] || ''),
      escapeHtml(((eff['First Name'] || '') + ' ' + (eff['Last Name'] || '')).trim()),
      escapeHtml(eff['Specialty'] || ''),
      escapeHtml(eff['Primary Site of Care'] || ''),
      practiceTypeBadge,
      escapeHtml((eff['City'] || '') + (eff['State'] ? ', ' + eff['State'] : '')),
      escapeHtml(eff['Product Line'] || ''),
      escapeHtml((eff['Tier'] || '').replace(' (', '\n(')),
      targetPill(eff['Target Tier']),
      statusPill(eff['Lead Status']),
      phonePillCell(eff),
      emailPillCell(eff),
      procedureVolume(eff).toLocaleString(),
      collagenVolume(eff).toLocaleString(),
      escapeHtml(t.last),
      t.count,
      escapeHtml(eff['Next Action'] || ''),
      eff['Draft Email'] ? '<button class="btn-draft" data-npi="' + escapeHtml(raw['HCP NPI']) + '">Draft Email</button>' : '',
    ];
    for (let k = 0; k < cells.length; k++) {
      const td = document.createElement('td');
      td.innerHTML = cells[k];
      tr.appendChild(td);
    }
    tr.addEventListener('click', (ev) => {
      if (ev.target.classList && ev.target.classList.contains('btn-draft')) return;
      openDrawer(raw['HCP NPI']);
    });
    tbody.appendChild(tr);
  }
  if (filtered.length > limit) {
    const tr = document.createElement('tr');
    const td = document.createElement('td'); td.colSpan = 18; td.className = 'muted'; td.style.textAlign = 'center';
    td.textContent = 'Showing first ' + limit + ' rows; refine filters to see more.';
    tr.appendChild(td); tbody.appendChild(tr);
  }
  updatePendingBadge();
}
for (const id of ['filter-line','filter-tier','filter-status','filter-target','filter-microlyte','filter-search']) {
  document.getElementById(id).addEventListener('input', render);
}

document.querySelector('#leads-table tbody').addEventListener('click', (ev) => {
  const btn = ev.target.closest('.btn-draft');
  if (!btn) return;
  ev.stopPropagation();
  openDraft(btn.dataset.npi);
});

function openDraft(npi) {
  const row = ROW_BY_NPI[npi];
  if (!row) return;
  document.getElementById('draft-title').textContent = 'Draft for ' + (row['First Name'] || '') + ' ' + (row['Last Name'] || '');
  document.getElementById('draft-subhead').innerHTML = escapeHtml(row['Primary Site of Care'] || '') + ' | ' + escapeHtml(row['City'] || '') + ', ' + escapeHtml(row['State'] || '') + ' | ' + escapeHtml(row['Product Line'] || '');
  document.getElementById('draft-to').value = row['Email'] || '';
  document.getElementById('draft-subject').value = row['Subject Line'] || '';
  document.getElementById('draft-body-text').value = row['Draft Email'] || '';
  document.getElementById('draft-track').textContent = row['Email Track'] || '';
  document.getElementById('draft-bg').classList.add('show');
  document.getElementById('draft-modal').classList.add('show');
}
function closeDraft() {
  document.getElementById('draft-bg').classList.remove('show');
  document.getElementById('draft-modal').classList.remove('show');
}
document.getElementById('draft-bg').addEventListener('click', closeDraft);
document.getElementById('draft-close').addEventListener('click', closeDraft);
document.getElementById('draft-copy').addEventListener('click', async () => {
  const text = 'Subject: ' + document.getElementById('draft-subject').value + '\n\n' + document.getElementById('draft-body-text').value;
  try { await navigator.clipboard.writeText(text); document.getElementById('draft-copy').textContent = 'Copied!'; setTimeout(() => document.getElementById('draft-copy').textContent = 'Copy to Clipboard', 1400); }
  catch (e) { alert('Could not copy: ' + e.message); }
});
document.getElementById('draft-mailto').addEventListener('click', () => {
  const to = encodeURIComponent(document.getElementById('draft-to').value);
  const subj = encodeURIComponent(document.getElementById('draft-subject').value);
  const body = encodeURIComponent(document.getElementById('draft-body-text').value);
  window.location.href = 'mailto:' + to + '?subject=' + subj + '&body=' + body;
});

const ROW_BY_NPI = {};
for (const r of rows) ROW_BY_NPI[r['HCP NPI']] = r;
let drawerNpi = null;

function selectField(name, value, options) {
  const opts = options.map(o => '<option value="' + escapeHtml(o) + '"' + (o === (value || '') ? ' selected' : '') + '>' + escapeHtml(o || '(blank)') + '</option>').join('');
  return '<select data-field="' + name + '">' + opts + '</select>';
}
function inputField(name, value, type) {
  if (type === 'date') {
    return '<div class="date-quick">'
      + '<input type="date" data-field="' + name + '" value="' + escapeHtml(value || '') + '" />'
      + '<button type="button" class="date-btn quick-set" data-for="' + escapeHtml(name) + '" data-offset="0">Today</button>'
      + '</div>';
  }
  return '<input type="' + (type || 'text') + '" data-field="' + name + '" value="' + escapeHtml(value || '') + '" />';
}
function textareaField(name, value) {
  return '<textarea data-field="' + name + '">' + escapeHtml(value || '') + '</textarea>';
}

function buildDrawerBody(row) {
  const eff = effective(row);
  const sections = [];
  const volRows = [
    ['Joint Replacement', eff['Joint Repl Vol']],
    ['Knee', eff['Knee Vol']],
    ['Hip', eff['Hip Vol']],
    ['Shoulder', eff['Shoulder Vol']],
    ['Open Ortho', eff['Open Ortho Vol']],
    ['Open Spine', eff['Open Spine Vol']],
    ['Other Procedure', eff['Procedure Vol']],
    ['Lg Collagen Sheet', eff['Lg Collagen Vol']],
    ['Sm/Md Collagen Sheet', eff['Sm/Md Collagen Vol']],
    ['Collagen Powder', eff['Collagen Powder Vol']],
    ['Total Collagen', eff['Total Collagen Vol']],
    ['Wound Care DME', eff['Wound Care DME Vol']],
    ['All DME', eff['All DME Vol']],
  ].filter(([, v]) => v && num(v) > 0);
  if (volRows.length) {
    sections.push('<section><h3>Volumes</h3>' + volRows.map(([k, v]) =>
      '<div class="field"><label>' + k + '</label><span>' + num(v).toLocaleString() + '</span></div>'
    ).join('') + '</section>');
  }
  sections.push('<section><h3>Verification</h3>'
    + '<div class="field"><label>Practice Type</label><span>' + escapeHtml(eff['Practice Type'] || '') + '</span></div>'
    + '<div class="field"><label>Primary Site</label><span>' + escapeHtml(eff['Primary Site of Care'] || '') + '</span></div>'
    + '<div class="field"><label>Address</label><span>' + escapeHtml([eff['Address 1'] || '', eff['City'] || '', eff['State'] || '', eff['Postal Code'] || ''].filter(x => x).join(', ')) + '</span></div>'
    + '<div class="field"><label>Drive Tier (NYC)</label><span>' + escapeHtml(eff['Tier'] || '') + '</span></div>'
    + '<div class="field"><label>Phone</label><span>' + escapeHtml(eff['Verified Phone'] || '(none)') + ' - ' + escapeHtml(eff['Phone Status'] || '') + '</span></div>'
    + (eff['Alternate Phones'] ? '<div class="field"><label>Alt. Phones</label><span>' + escapeHtml(eff['Alternate Phones']) + '</span></div>' : '')
    + '<div class="field"><label>Email</label><span>' + escapeHtml(eff['Email'] || '(none)') + ' - ' + escapeHtml(eff['Email Status'] || '') + '</span></div>'
    + '<div class="field"><label>MAC / Microlyte</label><span>' + escapeHtml(eff['MAC Jurisdiction'] || '') + ' / ' + escapeHtml(eff['Microlyte Eligible'] || '') + '</span></div>'
    + '<div class="field"><label>Incision Likelihood</label><span>' + escapeHtml(eff['Lg Incision Likelihood'] || '') + '</span></div>'
    + (eff['Practice Match'] ? '<div class="field"><label>Practice Match</label><span>' + escapeHtml(eff['Practice Match']) + '</span></div>' : '')
    + (eff['NPPES Practice Address'] ? '<div class="field"><label>NPPES Practice</label><span>' + escapeHtml(eff['NPPES Practice Address']) + '</span></div>' : '')
    + '</section>');

  if (eff['Practice Match'] && eff['Practice Match'].indexOf('Different') !== -1) {
    sections.push('<section><h3 style="color:#d68910;">⚠️ Multi-Affiliation Flag</h3>'
      + '<div style="background:#fef9e7;border:1px solid #f9e79f;padding:10px;border-radius:4px;font-size:12px;">'
      + '<strong>This doctor likely has multiple practice locations.</strong><br>'
      + 'AcuityMD lists them at <em>' + escapeHtml(eff['Primary Site of Care'] || '') + '</em> in ' + escapeHtml(eff['City'] || '') + ', ' + escapeHtml(eff['State'] || '') + '.<br>'
      + 'NPPES has them at <em>' + escapeHtml(eff['NPPES Practice Address'] || '') + '</em>.<br>'
      + 'Worth asking which one they spend most of their OR time at when you call.'
      + '</div></section>');
  }

  if (eff['Other Locations']) {
    try {
      const others = JSON.parse(eff['Other Locations']);
      if (others.length > 0) {
        const rows = others.map(l =>
          '<div style="padding:6px 0;border-bottom:1px dashed var(--line);font-size:12px;">'
          + '<div style="font-weight:600;">' + escapeHtml(l.site || '(no practice name)') + '</div>'
          + '<div class="muted">' + escapeHtml([l.address, l.city, l.state, l.zip].filter(x => x).join(', ')) + '</div>'
          + (l.phone ? '<div class="muted">📞 ' + escapeHtml(l.phone) + '</div>' : '')
          + (l.email ? '<div class="muted">✉️ ' + escapeHtml(l.email) + '</div>' : '')
          + '</div>'
        ).join('');
        sections.push('<section><h3>Secondary Locations (' + others.length + ')</h3>' + rows + '</section>');
      }
    } catch(e) {}
  }
  if (eff['Target Tier Reason']) {
    sections.push('<section><h3>Why this grade (' + escapeHtml(eff['Target Tier'] || '') + ')</h3>'
      + '<div style="font-size:12px;color:var(--ink);background:#f7f9fa;padding:10px;border-radius:4px;white-space:pre-wrap;">'
      + escapeHtml(eff['Target Tier Reason']).replace(/\|/g, '<br>')
      + '</div></section>');
  }
  sections.push('<section><h3>Quick Log</h3>'
    + '<div class="quick-log">'
    + '<button type="button" class="quick-btn primary" id="qlog-call" data-npi="' + escapeHtml(row['HCP NPI']) + '">📞 Log Call Now</button>'
    + '<button type="button" class="quick-btn" id="qlog-email" data-npi="' + escapeHtml(row['HCP NPI']) + '">✉️ Log Email Now</button>'
    + '</div>'
    + '<div class="quick-log-form" id="quick-log-form" style="display:none;">'
    + '<div class="field"><label>Date</label>'
    + '<div class="date-quick">'
    + '<input type="date" id="quick-date" />'
    + '<button type="button" class="date-btn" data-offset="0">Today</button>'
    + '<button type="button" class="date-btn" data-offset="-1">Yesterday</button>'
    + '</div></div>'
    + '<div class="field"><label>Outcome</label><select id="quick-outcome"></select></div>'
    + '<div class="field"><label id="quick-subject-lbl" style="display:none;">Subject</label><input id="quick-subject" style="display:none;" /></div>'
    + '<div class="field"><label>Notes</label><textarea id="quick-notes" placeholder="What happened?"></textarea></div>'
    + '<div class="field"><label>Update Status?</label>'
    + '<select id="quick-status">'
    + LEAD_STATUSES.map(s => '<option value="' + escapeHtml(s) + '">' + (s ? escapeHtml(s) : '(leave unchanged)') + '</option>').join('')
    + '</select></div>'
    + '<div class="quick-actions"><button type="button" class="primary" id="quick-save">Save to Browser</button> <button type="button" id="quick-cancel">Cancel</button></div>'
    + '</div>'
    + '</section>');

  sections.push('<section><h3>Lead</h3>'
    + '<div class="field"><label>Lead Status</label>' + selectField('Lead Status', eff['Lead Status'], LEAD_STATUSES) + '</div>'
    + '<div class="field"><label>Lead Priority</label><span>' + escapeHtml(eff['Lead Priority'] || '') + ' (' + escapeHtml(eff['Target Tier'] || '') + ')</span></div>'
    + '<div class="field"><label>Decision Maker?</label>' + selectField('Decision Maker?', eff['Decision Maker?'], DM_OPTIONS) + '</div>'
    + '<div class="field"><label>Next Action</label>' + inputField('Next Action', eff['Next Action']) + '</div>'
    + '<div class="field"><label>Next Action Date</label>' + inputField('Next Action Date', eff['Next Action Date'], 'date') + '</div>'
    + '</section>');

  sections.push('<section><h3>Calls</h3>');
  for (let i = 1; i <= 5; i++) {
    sections.push('<div class="round-grid"><h4>Call ' + i + '</h4>'
      + '<div><label class="muted">Date</label>' + inputField('Call ' + i + ' Date', eff['Call ' + i + ' Date'], 'date') + '</div>'
      + '<div><label class="muted">Outcome</label>' + selectField('Call ' + i + ' Outcome', eff['Call ' + i + ' Outcome'], CALL_OUTCOMES) + '</div>'
      + '<div class="full"><label class="muted">Notes</label>' + textareaField('Call ' + i + ' Notes', eff['Call ' + i + ' Notes']) + '</div>'
      + '</div>');
  }
  sections.push('</section>');

  sections.push('<section><h3>Emails</h3>');
  for (let i = 1; i <= 3; i++) {
    sections.push('<div class="round-grid"><h4>Email ' + i + '</h4>'
      + '<div><label class="muted">Date</label>' + inputField('Email ' + i + ' Date', eff['Email ' + i + ' Date'], 'date') + '</div>'
      + '<div><label class="muted">Outcome</label>' + selectField('Email ' + i + ' Outcome', eff['Email ' + i + ' Outcome'], EMAIL_OUTCOMES) + '</div>'
      + '<div class="full"><label class="muted">Subject</label>' + inputField('Email ' + i + ' Subject', eff['Email ' + i + ' Subject']) + '</div>'
      + '<div class="full"><label class="muted">Notes</label>' + textareaField('Email ' + i + ' Notes', eff['Email ' + i + ' Notes']) + '</div>'
      + '</div>');
  }
  sections.push('</section>');

  return sections.join('');
}

function openDrawer(npi) {
  drawerNpi = npi;
  const row = ROW_BY_NPI[npi];
  if (!row) return;
  const eff = effective(row);
  document.getElementById('drawer-title').textContent = (eff['First Name'] || '') + ' ' + (eff['Last Name'] || '') + ' - ' + (eff['Primary Site of Care'] || '');
  document.getElementById('drawer-subhead').innerHTML = 'NPI ' + escapeHtml(npi) + ' | ' + escapeHtml(eff['Product Line'] || '') + ' | ' + escapeHtml(eff['Tier'] || '') + ' | ' + escapeHtml(eff['MAC Jurisdiction'] || '') + ' | Microlyte ' + escapeHtml(eff['Microlyte Eligible'] || '');
  document.getElementById('drawer-quick-info').textContent = 'Phone: ' + (eff['Verified Phone'] || '-') + ' | Email: ' + (eff['Email'] || '-');
  document.getElementById('drawer-body').innerHTML = buildDrawerBody(row);
  document.getElementById('drawer-bg').classList.add('show');
  document.getElementById('drawer').classList.add('show');
}

function closeDrawer() {
  drawerNpi = null;
  document.getElementById('drawer-bg').classList.remove('show');
  document.getElementById('drawer').classList.remove('show');
}

document.getElementById('drawer-bg').addEventListener('click', closeDrawer);
document.getElementById('drawer-close').addEventListener('click', closeDrawer);
document.getElementById('drawer-cancel').addEventListener('click', closeDrawer);

let quickMode = null;
function todayIso() { const d = new Date(); d.setHours(12); return d.toISOString().slice(0,10); }
function dateWithOffset(days) { const d = new Date(); d.setDate(d.getDate() + days); d.setHours(12); return d.toISOString().slice(0,10); }

document.getElementById('drawer-body').addEventListener('click', (ev) => {
  const quickSet = ev.target.closest('.date-btn.quick-set');
  if (quickSet) {
    const target = document.querySelector('#drawer-body [data-field="' + quickSet.dataset.for + '"]');
    if (target) target.value = dateWithOffset(parseInt(quickSet.dataset.offset, 10) || 0);
    return;
  }
  const call = ev.target.closest('#qlog-call');
  const email = ev.target.closest('#qlog-email');
  if (!call && !email) {
    const db = ev.target.closest('.date-btn');
    if (db) { document.getElementById('quick-date').value = dateWithOffset(parseInt(db.dataset.offset, 10) || 0); }
    return;
  }
  quickMode = call ? 'call' : 'email';
  const form = document.getElementById('quick-log-form');
  form.style.display = 'block';
  document.getElementById('quick-date').value = todayIso();
  document.getElementById('quick-notes').value = '';
  const subjLbl = document.getElementById('quick-subject-lbl');
  const subjIn = document.getElementById('quick-subject');
  const outcomeSel = document.getElementById('quick-outcome');
  outcomeSel.innerHTML = '';
  const opts = quickMode === 'call' ? CALL_OUTCOMES : EMAIL_OUTCOMES;
  for (const o of opts) { const op = document.createElement('option'); op.value = o; op.textContent = o || '(pick outcome)'; outcomeSel.appendChild(op); }
  outcomeSel.value = '';
  if (quickMode === 'email') { subjLbl.style.display = ''; subjIn.style.display = ''; subjIn.value = ''; }
  else { subjLbl.style.display = 'none'; subjIn.style.display = 'none'; }
  document.getElementById('quick-status').value = '';
  outcomeSel.focus();
});

document.getElementById('drawer-body').addEventListener('click', (ev) => {
  if (!ev.target.matches('#quick-cancel')) return;
  document.getElementById('quick-log-form').style.display = 'none';
  quickMode = null;
});

document.getElementById('drawer-body').addEventListener('click', (ev) => {
  if (!ev.target.matches('#quick-save')) return;
  if (!drawerNpi || !quickMode) return;
  const outcome = document.getElementById('quick-outcome').value;
  if (!outcome) { alert('Pick an outcome before saving.'); return; }
  const date = document.getElementById('quick-date').value || todayIso();
  const notes = document.getElementById('quick-notes').value.trim();
  const subject = document.getElementById('quick-subject').value.trim();
  const newStatus = document.getElementById('quick-status').value;
  const row = ROW_BY_NPI[drawerNpi] || {};
  const patch = edits[drawerNpi] ? Object.assign({}, edits[drawerNpi]) : {};

  const rounds = quickMode === 'call' ? 5 : 3;
  const label = quickMode === 'call' ? 'Call' : 'Email';
  let slot = 0;
  for (let i = 1; i <= rounds; i++) {
    const existing = (patch[label + ' ' + i + ' Date'] !== undefined ? patch[label + ' ' + i + ' Date'] : row[label + ' ' + i + ' Date']) || '';
    if (!existing.trim()) { slot = i; break; }
  }
  if (!slot) { slot = rounds; }

  patch[label + ' ' + slot + ' Date'] = date;
  patch[label + ' ' + slot + ' Outcome'] = outcome;
  if (notes) patch[label + ' ' + slot + ' Notes'] = notes;
  if (quickMode === 'email' && subject) patch[label + ' ' + slot + ' Subject'] = subject;
  if (newStatus) patch['Lead Status'] = newStatus;

  edits[drawerNpi] = patch;
  localStorage.setItem(LS_KEY, JSON.stringify(edits));
  const savedNpi = drawerNpi;
  closeDrawer();
  render();
  const m = quickMode === 'call' ? 'Call' : 'Email';
  const badge = document.createElement('div');
  badge.textContent = m + ' #' + slot + ' logged for ' + (row['First Name'] || '') + ' ' + (row['Last Name'] || '') + '. Remember to Download Updates.';
  badge.style.cssText = 'position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:#27ae60;color:#fff;padding:10px 18px;border-radius:4px;box-shadow:0 4px 12px rgba(0,0,0,0.2);z-index:50;font-size:13px;';
  document.body.appendChild(badge);
  setTimeout(() => badge.remove(), 3500);
  quickMode = null;
});

document.getElementById('drawer-save').addEventListener('click', () => {
  if (!drawerNpi) return;
  const row = ROW_BY_NPI[drawerNpi] || {};
  const patch = {};
  document.querySelectorAll('#drawer-body [data-field]').forEach(el => {
    const name = el.dataset.field;
    const value = (el.value || '').trim();
    if (value !== (row[name] || '').trim()) patch[name] = value;
  });
  if (Object.keys(patch).length === 0) {
    delete edits[drawerNpi];
  } else {
    edits[drawerNpi] = patch;
  }
  localStorage.setItem(LS_KEY, JSON.stringify(edits));
  closeDrawer();
  render();
});

document.getElementById('btn-download').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify(edits, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const ts = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
  a.download = 'activity_edits_' + ts + '.json';
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
});
document.getElementById('btn-import').addEventListener('click', () => document.getElementById('file-import').click());
document.getElementById('file-import').addEventListener('change', async (e) => {
  const f = e.target.files && e.target.files[0];
  if (!f) return;
  try {
    const text = await f.text();
    const incoming = JSON.parse(text);
    for (const npi of Object.keys(incoming)) {
      edits[npi] = Object.assign({}, edits[npi] || {}, incoming[npi]);
    }
    localStorage.setItem(LS_KEY, JSON.stringify(edits));
    render();
    alert('Imported ' + Object.keys(incoming).length + ' edit entries.');
  } catch (err) {
    alert('Could not read the file: ' + err.message);
  } finally {
    e.target.value = '';
  }
});
document.getElementById('btn-discard').addEventListener('click', () => {
  if (Object.keys(edits).length === 0) return;
  if (!confirm('Discard all ' + Object.keys(edits).length + ' unsaved edits? This clears your browser copy, but anything you already downloaded is safe.')) return;
  edits = {};
  localStorage.removeItem(LS_KEY);
  render();
});

render();
</script>
</body>
</html>
"""
