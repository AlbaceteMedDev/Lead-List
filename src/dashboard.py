"""Generate a self-contained HTML dashboard for tracking and monitoring the lead pipeline.

The dashboard is a single static file — no server required. It uses Chart.js
from CDN for the charts and embeds the enriched dataset inline so sales ops can
open it in any browser and filter/drill-down by tier, MAC, email status, and
Microlyte eligibility.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
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


def _counts(series: pd.Series, keys: Optional[list[str]] = None) -> dict:
    values = series.fillna("").astype(str)
    counted = values.value_counts().to_dict()
    if keys is not None:
        return {k: int(counted.get(k, 0)) for k in keys}
    return {str(k): int(v) for k, v in counted.items()}


def build_stats(df: pd.DataFrame) -> dict:
    total = len(df)
    return {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "total_leads": total,
        "tier_counts": _counts(df.get("Tier", pd.Series(dtype=str)), TIER_ORDER),
        "practice_type_counts": _counts(df.get("Practice Type", pd.Series(dtype=str))),
        "phone_status_counts": _counts(
            df.get("Phone Status", pd.Series(dtype=str)),
            ["Verified", "Added from NPPES", "Updated (NPPES differs)", "Missing"],
        ),
        "email_status_counts": _counts(df.get("Email Status", pd.Series(dtype=str))),
        "mac_counts": _counts(df.get("MAC Jurisdiction", pd.Series(dtype=str))),
        "microlyte_counts": _counts(
            df.get("Microlyte Eligible", pd.Series(dtype=str)), ["Yes", "No", "Unknown"]
        ),
        "track_counts": _counts(df.get("Email Track", pd.Series(dtype=str))),
        "state_counts": _counts(df.get("State", pd.Series(dtype=str))),
    }


def _dataset(df: pd.DataFrame) -> list[dict]:
    cols = [
        "HCP NPI", "First Name", "Last Name", "Credential", "Specialty",
        "Email", "Email Status", "Verified Phone", "Phone Status",
        "Primary Site of Care", "Practice Type", "City", "State",
        "Tier", "MAC Jurisdiction", "Microlyte Eligible",
        "Joint Repl Vol", "Subject Line", "Email Track",
    ]
    present = [c for c in cols if c in df.columns]
    return df[present].fillna("").astype(str).to_dict(orient="records")


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
<title>Albacete MedDev - Lead List Dashboard</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {
    --navy: #1b4f72;
    --ink: #1c2833;
    --muted: #566573;
    --line: #d5d8dc;
    --bg: #f4f6f7;
    --green: #27ae60;
    --red: #c0392b;
    --amber: #d68910;
    --card: #ffffff;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: -apple-system, Segoe UI, Arial, sans-serif; color: var(--ink); background: var(--bg); }
  header { background: var(--navy); color: #fff; padding: 16px 24px; display: flex; align-items: baseline; gap: 16px; }
  header h1 { margin: 0; font-size: 18px; font-weight: 600; }
  header .meta { color: #d5dbe1; font-size: 12px; }
  main { max-width: 1400px; margin: 0 auto; padding: 20px; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 20px; }
  .kpi { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 14px 16px; }
  .kpi .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--muted); }
  .kpi .value { font-size: 24px; font-weight: 600; margin-top: 4px; color: var(--navy); }
  .kpi .sub { font-size: 12px; color: var(--muted); margin-top: 4px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap: 16px; margin-bottom: 20px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 16px; }
  .card h2 { margin: 0 0 12px 0; font-size: 14px; color: var(--navy); font-weight: 600; }
  .card canvas { max-height: 260px; }
  .controls { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 12px; margin-bottom: 12px; }
  .controls label { font-size: 12px; color: var(--muted); display: flex; flex-direction: column; gap: 4px; }
  .controls select, .controls input { font: inherit; padding: 6px 8px; border: 1px solid var(--line); border-radius: 4px; min-width: 160px; }
  table { width: 100%; border-collapse: collapse; background: var(--card); font-size: 12px; }
  th, td { padding: 6px 8px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: top; }
  th { background: var(--navy); color: #fff; font-weight: 600; position: sticky; top: 0; }
  .pill { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
  .pill.green { background: #d5f5e3; color: #186a3b; }
  .pill.red { background: #fadbd8; color: #922b21; }
  .pill.amber { background: #fef9e7; color: #9a7d0a; }
  .pill.navy { background: #d6eaf8; color: var(--navy); }
  .table-wrap { background: var(--card); border: 1px solid var(--line); border-radius: 6px; overflow: auto; max-height: 60vh; }
  .muted { color: var(--muted); font-size: 12px; }
</style>
</head>
<body>
<header>
  <h1>Albacete MedDev - Lead List Tracker</h1>
  <span class="meta" id="gen-meta"></span>
</header>
<main>
  <section class="kpis" id="kpis"></section>

  <section class="grid">
    <div class="card"><h2>Leads by Tier</h2><canvas id="chart-tier"></canvas></div>
    <div class="card"><h2>Phone Verification Status</h2><canvas id="chart-phone"></canvas></div>
    <div class="card"><h2>Email Enrichment Status</h2><canvas id="chart-email"></canvas></div>
    <div class="card"><h2>Microlyte Eligibility</h2><canvas id="chart-microlyte"></canvas></div>
    <div class="card"><h2>MAC Jurisdictions</h2><canvas id="chart-mac"></canvas></div>
    <div class="card"><h2>Email Track Split</h2><canvas id="chart-track"></canvas></div>
  </section>

  <div class="controls">
    <label>Tier
      <select id="filter-tier"><option value="">All</option></select>
    </label>
    <label>MAC
      <select id="filter-mac"><option value="">All</option></select>
    </label>
    <label>Email Status
      <select id="filter-email"><option value="">All</option></select>
    </label>
    <label>Microlyte
      <select id="filter-microlyte"><option value="">All</option></select>
    </label>
    <label style="flex: 1;">Search (name, practice, email, NPI)
      <input id="filter-search" placeholder="type to filter..." />
    </label>
    <span class="muted" id="row-count"></span>
  </div>

  <div class="table-wrap">
    <table id="leads-table">
      <thead>
        <tr>
          <th>NPI</th><th>Name</th><th>Credential</th><th>Practice</th>
          <th>City, State</th><th>Tier</th><th>MAC</th>
          <th>Microlyte</th><th>Phone</th><th>Phone Status</th>
          <th>Email</th><th>Email Status</th><th>Joint Vol</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  </div>
</main>

<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
const DATA = JSON.parse(document.getElementById('payload').textContent);
const stats = DATA.stats;
const rows = DATA.rows;

document.getElementById('gen-meta').textContent = 'Generated ' + stats.generated_at;

function sum(obj) { return Object.values(obj).reduce((a, b) => a + (b || 0), 0); }

const verifiedEmail = Object.entries(stats.email_status_counts)
  .filter(([k]) => k.indexOf('Verified') !== -1 || k.indexOf('Inferred') !== -1)
  .reduce((a, [, v]) => a + v, 0);
const missingEmail = stats.email_status_counts['Missing'] || 0;
const verifiedPhone = (stats.phone_status_counts['Verified'] || 0) + (stats.phone_status_counts['Added from NPPES'] || 0);
const microlyteYes = stats.microlyte_counts['Yes'] || 0;

const kpiDefs = [
  { label: 'Total Leads', value: stats.total_leads, sub: stats.generated_at },
  { label: 'Phone Ready', value: verifiedPhone, sub: (stats.phone_status_counts['Missing'] || 0) + ' missing' },
  { label: 'Email Ready', value: verifiedEmail, sub: missingEmail + ' missing' },
  { label: 'Microlyte Eligible', value: microlyteYes, sub: (stats.microlyte_counts['No'] || 0) + ' LCD-blocked' },
  { label: 'Tier 1-2 (drive ≤1h)', value: (stats.tier_counts['Tier 1 (0-30 min)'] || 0) + (stats.tier_counts['Tier 2 (30-60 min)'] || 0), sub: 'Priority outreach' },
  { label: 'Hospital-Based', value: stats.tier_counts['Hospital-Based'] || 0, sub: 'Separate workflow' },
];
const kpiEl = document.getElementById('kpis');
for (const k of kpiDefs) {
  const el = document.createElement('div');
  el.className = 'kpi';
  el.innerHTML = '<div class="label">' + k.label + '</div><div class="value">' + k.value + '</div><div class="sub">' + k.sub + '</div>';
  kpiEl.appendChild(el);
}

const palette = ['#1b4f72', '#2874a6', '#2e86c1', '#5dade2', '#85c1e9', '#aed6f1', '#d6eaf8', '#c0392b', '#d68910', '#27ae60'];
function makeChart(id, type, labels, data, options) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  new Chart(ctx, {
    type: type,
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: labels.map((_, i) => palette[i % palette.length]),
        borderWidth: 0,
      }],
    },
    options: Object.assign({
      responsive: true,
      plugins: { legend: { display: type !== 'bar', position: 'right', labels: { font: { size: 11 } } } },
    }, options || {}),
  });
}

makeChart('chart-tier', 'bar', Object.keys(stats.tier_counts), Object.values(stats.tier_counts), { indexAxis: 'y', plugins: { legend: { display: false } } });
makeChart('chart-phone', 'doughnut', Object.keys(stats.phone_status_counts), Object.values(stats.phone_status_counts));
makeChart('chart-email', 'doughnut', Object.keys(stats.email_status_counts), Object.values(stats.email_status_counts));
makeChart('chart-microlyte', 'doughnut', Object.keys(stats.microlyte_counts), Object.values(stats.microlyte_counts));
makeChart('chart-mac', 'bar', Object.keys(stats.mac_counts), Object.values(stats.mac_counts), { plugins: { legend: { display: false } } });
makeChart('chart-track', 'doughnut', Object.keys(stats.track_counts), Object.values(stats.track_counts));

function uniq(col) {
  const s = new Set();
  for (const r of rows) { if (r[col]) s.add(r[col]); }
  return [...s].sort();
}
function fillSelect(id, values) {
  const el = document.getElementById(id);
  for (const v of values) {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    el.appendChild(opt);
  }
}
fillSelect('filter-tier', uniq('Tier'));
fillSelect('filter-mac', uniq('MAC Jurisdiction'));
fillSelect('filter-email', uniq('Email Status'));
fillSelect('filter-microlyte', uniq('Microlyte Eligible'));

function pill(text, kind) { return '<span class="pill ' + kind + '">' + text + '</span>'; }
function emailPill(s) {
  if (!s) return '';
  if (s.indexOf('Verified') !== -1) return pill(s, 'green');
  if (s === 'Missing') return pill(s, 'red');
  if (s.indexOf('Inferred') !== -1) return pill(s, 'amber');
  return pill(s, 'navy');
}
function phonePill(s) {
  if (!s) return '';
  if (s === 'Verified') return pill(s, 'green');
  if (s === 'Added from NPPES') return pill(s, 'amber');
  if (s === 'Missing') return pill(s, 'red');
  return pill(s, 'navy');
}
function microlytePill(s) {
  if (s === 'Yes') return pill('Yes', 'green');
  if (s === 'No') return pill('No', 'red');
  return pill(s || '-', 'navy');
}

function render() {
  const tier = document.getElementById('filter-tier').value;
  const mac = document.getElementById('filter-mac').value;
  const email = document.getElementById('filter-email').value;
  const micro = document.getElementById('filter-microlyte').value;
  const q = document.getElementById('filter-search').value.trim().toLowerCase();

  const filtered = rows.filter(r => {
    if (tier && r['Tier'] !== tier) return false;
    if (mac && r['MAC Jurisdiction'] !== mac) return false;
    if (email && r['Email Status'] !== email) return false;
    if (micro && r['Microlyte Eligible'] !== micro) return false;
    if (q) {
      const hay = [r['First Name'], r['Last Name'], r['Primary Site of Care'], r['Email'], r['HCP NPI']].join(' ').toLowerCase();
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  });

  document.getElementById('row-count').textContent = filtered.length + ' of ' + rows.length + ' leads';
  const tbody = document.querySelector('#leads-table tbody');
  tbody.innerHTML = '';
  const limit = Math.min(filtered.length, 500);
  for (let i = 0; i < limit; i++) {
    const r = filtered[i];
    const tr = document.createElement('tr');
    const cells = [
      r['HCP NPI'] || '',
      ((r['First Name'] || '') + ' ' + (r['Last Name'] || '')).trim(),
      r['Credential'] || '',
      r['Primary Site of Care'] || '',
      ((r['City'] || '') + (r['State'] ? ', ' + r['State'] : '')),
      r['Tier'] || '',
      r['MAC Jurisdiction'] || '',
      microlytePill(r['Microlyte Eligible']),
      r['Verified Phone'] || '',
      phonePill(r['Phone Status']),
      r['Email'] || '',
      emailPill(r['Email Status']),
      r['Joint Repl Vol'] || '',
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
    const td = document.createElement('td');
    td.colSpan = 13;
    td.className = 'muted';
    td.style.textAlign = 'center';
    td.textContent = 'Showing first ' + limit + ' rows; refine filters to see more.';
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
}
for (const id of ['filter-tier', 'filter-mac', 'filter-email', 'filter-microlyte', 'filter-search']) {
  document.getElementById(id).addEventListener('input', render);
}
render();
</script>
</body>
</html>
"""
