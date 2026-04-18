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
        "Product Line", "Lead Priority", "Lead Status", "Target Tier", "Target Score",
        "Lg Incision Likelihood", "Next Action", "Next Action Date",
        "Last Touch Date", "Touch Count",
        "Joint Repl Vol", "Open Spine Vol", "Why Target?", "Best Approach",
    ]
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
  :root { --navy:#1b4f72; --ink:#1c2833; --muted:#566573; --line:#d5d8dc; --bg:#f4f6f7; --card:#fff; --accent:#1b9d7a; }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: -apple-system, Segoe UI, Arial, sans-serif; color: var(--ink); background: var(--bg); }
  header { background: var(--navy); color: #fff; padding: 14px 24px; display: flex; align-items: baseline; gap: 16px; }
  header h1 { margin: 0; font-size: 17px; font-weight: 600; }
  header .meta { color: #d5dbe1; font-size: 12px; }
  main { max-width: 1500px; margin: 0 auto; padding: 18px; }
  .tabs { display: flex; gap: 4px; margin-bottom: 16px; border-bottom: 2px solid var(--navy); }
  .tab { padding: 8px 16px; cursor: pointer; background: #eaeded; border: 1px solid var(--line); border-bottom: none; border-radius: 6px 6px 0 0; font-size: 13px; font-weight: 500; color: var(--muted); }
  .tab.active { background: var(--navy); color: #fff; border-color: var(--navy); }
  .panel { display: none; }
  .panel.active { display: block; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin-bottom: 16px; }
  .kpi { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 12px 14px; }
  .kpi .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--muted); }
  .kpi .value { font-size: 22px; font-weight: 600; margin-top: 3px; color: var(--navy); }
  .kpi .sub { font-size: 11px; color: var(--muted); margin-top: 3px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap: 14px; margin-bottom: 16px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 14px; }
  .card h2 { margin: 0 0 10px 0; font-size: 13px; color: var(--navy); font-weight: 600; }
  .card canvas { max-height: 240px; }
  .controls { display: flex; flex-wrap: wrap; gap: 10px; align-items: flex-end; background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 10px; margin-bottom: 10px; }
  .controls label { font-size: 11px; color: var(--muted); display: flex; flex-direction: column; gap: 3px; }
  .controls select, .controls input { font: inherit; padding: 5px 8px; border: 1px solid var(--line); border-radius: 4px; min-width: 150px; }
  .table-wrap { background: var(--card); border: 1px solid var(--line); border-radius: 6px; overflow: auto; max-height: 65vh; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th, td { padding: 5px 8px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }
  th { background: var(--navy); color: #fff; font-weight: 600; position: sticky; top: 0; z-index: 1; }
  .pill { display: inline-block; padding: 1px 7px; border-radius: 9px; font-size: 10px; font-weight: 600; }
  .pill.green { background: #d5f5e3; color: #186a3b; }
  .pill.red { background: #fadbd8; color: #922b21; }
  .pill.amber { background: #fef9e7; color: #9a7d0a; }
  .pill.navy { background: #d6eaf8; color: var(--navy); }
  .muted { color: var(--muted); font-size: 11px; }
</style>
</head>
<body>
<header>
  <h1>Albacete MedDev - Outreach Tracker</h1>
  <span class="meta" id="gen-meta"></span>
</header>
<main>
  <div class="tabs">
    <div class="tab active" data-panel="overview">Overview</div>
    <div class="tab" data-panel="activity">Call / Email Activity</div>
    <div class="tab" data-panel="leads">Leads</div>
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
      <div class="card"><h2>Lead Status (Funnel)</h2><canvas id="chart-status-funnel"></canvas></div>
    </div>
  </section>

  <section class="panel" id="panel-leads">
    <div class="controls">
      <label>Product Line<select id="filter-line"><option value="">All</option></select></label>
      <label>Tier<select id="filter-tier"><option value="">All</option></select></label>
      <label>Lead Status<select id="filter-status"><option value="">All</option></select></label>
      <label>Target Tier<select id="filter-target"><option value="">All</option></select></label>
      <label>Microlyte<select id="filter-microlyte"><option value="">All</option></select></label>
      <label style="flex: 1;">Search<input id="filter-search" placeholder="name, practice, email, NPI..." /></label>
      <span class="muted" id="row-count"></span>
    </div>
    <div class="table-wrap">
      <table id="leads-table">
        <thead>
          <tr>
            <th>NPI</th><th>Name</th><th>Practice</th><th>City,ST</th>
            <th>Line</th><th>Tier</th><th>Target</th><th>Status</th>
            <th>Phone</th><th>Email</th>
            <th>Last Touch</th><th>Touches</th><th>Next Action</th><th>Why Target?</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>
  </section>
</main>

<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
const DATA = JSON.parse(document.getElementById('payload').textContent);
const stats = DATA.stats;
const act = stats.activity;
const rows = DATA.rows;

document.getElementById('gen-meta').textContent = 'Generated ' + stats.generated_at;

document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(x => x.classList.remove('active'));
  t.classList.add('active');
  document.getElementById('panel-' + t.dataset.panel).classList.add('active');
}));

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
kpi(activityCt, 'Leads Touched', Object.keys(act.calls_by_date).length ? 'see chart' : '0', 'See daily activity chart');

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

function uniq(col) {
  const s = new Set();
  for (const r of rows) { if (r[col]) s.add(r[col]); }
  return [...s].sort();
}
function fillSelect(id, values) {
  const el = document.getElementById(id);
  for (const v of values) { const opt = document.createElement('option'); opt.value = v; opt.textContent = v; el.appendChild(opt); }
}
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

function render() {
  const line = document.getElementById('filter-line').value;
  const tier = document.getElementById('filter-tier').value;
  const status = document.getElementById('filter-status').value;
  const target = document.getElementById('filter-target').value;
  const micro = document.getElementById('filter-microlyte').value;
  const q = document.getElementById('filter-search').value.trim().toLowerCase();

  const filtered = rows.filter(r => {
    if (line && r['Product Line'] !== line) return false;
    if (tier && r['Tier'] !== tier) return false;
    if (status && r['Lead Status'] !== status) return false;
    if (target && r['Target Tier'] !== target) return false;
    if (micro && r['Microlyte Eligible'] !== micro) return false;
    if (q) {
      const hay = [r['First Name'], r['Last Name'], r['Primary Site of Care'], r['Email'], r['HCP NPI']].join(' ').toLowerCase();
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  });
  document.getElementById('row-count').textContent = filtered.length + ' of ' + rows.length;

  const tbody = document.querySelector('#leads-table tbody');
  tbody.innerHTML = '';
  const limit = Math.min(filtered.length, 500);
  for (let i = 0; i < limit; i++) {
    const r = filtered[i];
    const tr = document.createElement('tr');
    const cells = [
      r['HCP NPI'] || '',
      ((r['First Name'] || '') + ' ' + (r['Last Name'] || '')).trim(),
      r['Primary Site of Care'] || '',
      ((r['City'] || '') + (r['State'] ? ', ' + r['State'] : '')),
      r['Product Line'] || '',
      (r['Tier'] || '').replace(' (', '\n('),
      targetPill(r['Target Tier']),
      statusPill(r['Lead Status']),
      r['Verified Phone'] || '',
      r['Email'] || '',
      r['Last Touch Date'] || '',
      r['Touch Count'] || '0',
      r['Next Action'] || '',
      r['Why Target?'] || '',
    ];
    for (const c of cells) {
      const td = document.createElement('td');
      td.innerHTML = c;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  if (filtered.length > limit) {
    const tr = document.createElement('tr');
    const td = document.createElement('td'); td.colSpan = 14; td.className = 'muted'; td.style.textAlign = 'center';
    td.textContent = 'Showing first ' + limit + ' rows; refine filters to see more.';
    tr.appendChild(td); tbody.appendChild(tr);
  }
}
for (const id of ['filter-line','filter-tier','filter-status','filter-target','filter-microlyte','filter-search']) {
  document.getElementById(id).addEventListener('input', render);
}
render();
</script>
</body>
</html>
"""
