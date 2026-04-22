"""Mobile-first cold-caller dashboard.

Emits ``mobile.html`` — a single-column, tap-to-call, thumb-optimized UI
for working leads on a phone. Shares the same data schema as the
desktop dashboard but renders a vertical card feed instead of a table.

Key mobile features:
- Tap a card → full-screen detail view
- ``tel:`` link on phone numbers (one-tap dial)
- Big "Log Call" button that opens a bottom sheet with outcome dropdown
- Top priority (A+/A) filter by default
- Sticky search + filter chips at top
- localStorage edits sync with desktop dashboard (same key)
"""

from __future__ import annotations

import json
import logging
from pathlib import Path

import pandas as pd

from src import dashboard as desktop  # reuse build_stats + dataset

log = logging.getLogger(__name__)


def write_mobile(df: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    stats = desktop.build_stats(df)
    rows = desktop._dataset(df)
    payload = json.dumps({"stats": stats, "rows": rows}, separators=(",", ":"))
    html = _HTML_TEMPLATE.replace("__PAYLOAD__", payload)
    output_path.write_text(html, encoding="utf-8")
    log.info("Wrote mobile dashboard: %s", output_path)
    return output_path


_HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Albacete MedDev - Leads Mobile</title>
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
<meta name="theme-color" content="#1b4f72" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-title" content="AMD Leads" />
<style>
  :root { --navy:#1b4f72; --ink:#1c2833; --muted:#566573; --line:#d5d8dc; --bg:#f4f6f7; --card:#fff; --green:#27ae60; --red:#c0392b; --amber:#d68910; --accent:#1b9d7a; }
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  html, body { margin: 0; padding: 0; height: 100%; overscroll-behavior-y: contain; }
  body { font-family: -apple-system, Segoe UI, Arial, sans-serif; color: var(--ink); background: var(--bg); font-size: 15px; line-height: 1.35; }
  header { position: sticky; top: 0; z-index: 5; background: var(--navy); color: #fff; padding: 12px 14px calc(12px + env(safe-area-inset-top)) 14px; padding-top: calc(12px + env(safe-area-inset-top)); }
  header h1 { margin: 0; font-size: 16px; font-weight: 600; }
  header .meta { font-size: 11px; color: #d5dbe1; margin-top: 2px; }
  header .actions { display: flex; gap: 6px; margin-top: 8px; }
  header button, header .ghost { flex: 1; background: rgba(255,255,255,0.16); color: #fff; border: 1px solid rgba(255,255,255,0.3); border-radius: 6px; padding: 8px 10px; font-size: 12px; cursor: pointer; text-align: center; text-decoration: none; }
  header button:active { background: rgba(255,255,255,0.32); }

  .chips { display: flex; gap: 6px; padding: 8px 14px; overflow-x: auto; background: #fff; border-bottom: 1px solid var(--line); -webkit-overflow-scrolling: touch; }
  .chip { flex-shrink: 0; font-size: 12px; padding: 6px 12px; border-radius: 18px; border: 1px solid var(--line); background: #fff; color: var(--muted); white-space: nowrap; cursor: pointer; }
  .chip.active { background: var(--navy); color: #fff; border-color: var(--navy); }

  .search { padding: 8px 14px; background: #fff; border-bottom: 1px solid var(--line); }
  .search input { width: 100%; padding: 9px 12px; border: 1px solid var(--line); border-radius: 7px; font-size: 14px; background: var(--bg); }

  .feed { padding: 10px 10px 120px 10px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 12px; margin-bottom: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }
  .card.edited { border-left: 4px solid var(--amber); }
  .card-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; }
  .card-name { font-weight: 600; font-size: 15px; color: var(--ink); }
  .card-sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .card-badges { display: flex; gap: 4px; flex-shrink: 0; flex-direction: column; align-items: flex-end; }
  .pill { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 700; white-space: nowrap; }
  .pill.green { background: #d5f5e3; color: #186a3b; }
  .pill.red { background: #fadbd8; color: #922b21; }
  .pill.amber { background: #fef9e7; color: #9a7d0a; }
  .pill.navy { background: #d6eaf8; color: var(--navy); }
  .card-phone { margin-top: 10px; display: flex; flex-direction: column; gap: 4px; }
  .card-phone a { font-size: 14px; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; gap: 8px; padding: 8px 10px; background: var(--bg); border-radius: 6px; color: var(--navy); border: 1px solid var(--line); }
  .card-phone a .phone-src { font-size: 10px; font-weight: 600; padding: 2px 7px; border-radius: 9px; background: #d5f5e3; color: #186a3b; text-transform: uppercase; }
  .card-phone a .phone-src.acuity { background: #d6eaf8; color: var(--navy); }
  .card-phone a .phone-src.web { background: #d5f5e3; color: #186a3b; }
  .card-phone a.missing { color: var(--red); background: #fadbd8; border-color: #f5b7b1; }
  .card-meta { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; font-size: 11px; color: var(--muted); }
  .card-meta span { display: inline-flex; align-items: center; gap: 3px; }
  .card-actions { display: flex; gap: 6px; margin-top: 10px; }
  .card-actions button { flex: 1; font: inherit; font-size: 12px; font-weight: 600; padding: 10px 8px; border-radius: 7px; border: 1px solid var(--line); background: #fff; color: var(--ink); cursor: pointer; }
  .card-actions button.primary { background: var(--navy); color: #fff; border-color: var(--navy); }
  .card-actions button:active { opacity: 0.7; }

  .footer { position: fixed; bottom: 0; left: 0; right: 0; background: #fff; border-top: 1px solid var(--line); padding: 10px 14px calc(10px + env(safe-area-inset-bottom)) 14px; display: flex; justify-content: space-between; align-items: center; font-size: 12px; color: var(--muted); z-index: 4; }
  .footer .pending { background: var(--amber); color: #fff; padding: 4px 10px; border-radius: 10px; font-weight: 700; font-size: 11px; display: none; }
  .footer .pending.show { display: inline-block; }
  .footer button { font: inherit; font-size: 12px; padding: 7px 12px; border-radius: 6px; background: var(--navy); color: #fff; border: none; font-weight: 600; }

  /* Sheet (bottom) for lead detail + log call */
  .sheet-bg { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: none; z-index: 8; }
  .sheet-bg.show { display: block; }
  .sheet { position: fixed; left: 0; right: 0; bottom: 0; max-height: 92vh; background: #fff; border-radius: 14px 14px 0 0; transform: translateY(100%); transition: transform 0.22s ease; z-index: 9; display: flex; flex-direction: column; }
  .sheet.show { transform: translateY(0); }
  .sheet-handle { width: 40px; height: 4px; background: var(--line); border-radius: 2px; margin: 8px auto 0 auto; }
  .sheet-head { padding: 12px 16px 8px 16px; border-bottom: 1px solid var(--line); }
  .sheet-head h2 { margin: 0; font-size: 17px; font-weight: 600; }
  .sheet-head .sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .sheet-body { overflow-y: auto; padding: 14px 16px; flex: 1; padding-bottom: calc(80px + env(safe-area-inset-bottom)); }
  .sheet-body section { margin-bottom: 14px; }
  .sheet-body h3 { font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted); margin: 0 0 6px 0; }
  .field { display: flex; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid var(--line); font-size: 13px; gap: 10px; }
  .field .lbl { color: var(--muted); font-size: 12px; }
  .field .val { color: var(--ink); text-align: right; flex: 1; word-break: break-word; }
  .log-form label { display: block; font-size: 11px; color: var(--muted); margin-top: 8px; margin-bottom: 3px; text-transform: uppercase; letter-spacing: 0.03em; }
  .log-form input, .log-form select, .log-form textarea { width: 100%; font: inherit; font-size: 14px; padding: 10px 12px; border: 1px solid var(--line); border-radius: 8px; background: #fff; }
  .log-form textarea { min-height: 60px; resize: vertical; }
  .date-row { display: flex; gap: 6px; align-items: center; }
  .date-row input { flex: 1; }
  .date-row button { font: inherit; font-size: 11px; padding: 7px 10px; background: #fff; border: 1px solid var(--line); border-radius: 6px; color: var(--muted); font-weight: 600; }
  .sheet-foot { position: absolute; bottom: 0; left: 0; right: 0; padding: 10px 14px calc(10px + env(safe-area-inset-bottom)) 14px; background: #fff; border-top: 1px solid var(--line); display: flex; gap: 8px; }
  .sheet-foot button { flex: 1; font: inherit; font-size: 14px; font-weight: 600; padding: 12px; border-radius: 8px; border: 1px solid var(--line); background: #fff; color: var(--ink); }
  .sheet-foot button.primary { background: var(--navy); color: #fff; border-color: var(--navy); }

  .action-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .act-btn { display: flex; align-items: center; justify-content: center; gap: 6px; font: inherit; font-size: 13px; font-weight: 600; padding: 12px 10px; border-radius: 8px; border: 1px solid var(--navy); background: var(--navy); color: #fff; cursor: pointer; text-decoration: none; text-align: center; }
  .act-btn.ghost { background: #fff; color: var(--navy); border-color: var(--line); }
  .act-btn:active { opacity: 0.7; }

  .history-item { padding: 8px 10px; background: #f7f9fa; border-radius: 8px; margin-bottom: 6px; }
  .history-row { display: flex; justify-content: space-between; align-items: center; gap: 8px; font-size: 12px; }
  .history-row .lbl { color: var(--muted); font-weight: 600; }
  .history-note { font-size: 12px; color: var(--ink); margin-top: 4px; white-space: pre-wrap; }

  .toast { position: fixed; top: calc(env(safe-area-inset-top) + 60px); left: 16px; right: 16px; background: var(--green); color: #fff; padding: 12px 16px; border-radius: 10px; z-index: 20; transform: translateY(-120%); transition: transform 0.25s ease; font-size: 14px; box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
  .toast.show { transform: translateY(0); }

  .empty { text-align: center; color: var(--muted); padding: 40px 20px; font-size: 13px; }
</style>
</head>
<body>
<header>
  <h1>Albacete MedDev — Leads</h1>
  <div class="meta" id="meta"></div>
  <div class="actions">
    <a href="dashboard.html" class="ghost">Desktop View</a>
    <button id="btn-download">Download Updates</button>
  </div>
</header>

<div class="chips" id="chips"></div>

<div class="search">
  <input id="search" placeholder="Search name, practice, city, NPI..." />
</div>

<main class="feed" id="feed"></main>

<div class="footer">
  <span id="feed-count">0 leads</span>
  <span class="pending" id="pending-count">0 unsaved</span>
  <button id="btn-toggle-all">Show All</button>
</div>

<div class="sheet-bg" id="sheet-bg"></div>
<section class="sheet" id="sheet">
  <div class="sheet-handle"></div>
  <div class="sheet-head">
    <h2 id="sh-name"></h2>
    <div class="sub" id="sh-sub"></div>
  </div>
  <div class="sheet-body" id="sh-body"></div>
  <div class="sheet-foot">
    <button id="sh-close">Close</button>
    <button class="primary" id="sh-save">Save</button>
  </div>
</section>

<div class="toast" id="toast"></div>

<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
const DATA = JSON.parse(document.getElementById('payload').textContent);
const rows = DATA.rows;
const stats = DATA.stats;

const LS_KEY = 'albacete_activity_edits_v1'; // same key as desktop so they sync
let edits = {};
try { edits = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e) { edits = {}; }

const LEAD_STATUSES = ['','New','Queued','Attempting Contact','Connected','Interested','Meeting Booked','Nurture','Not Interested','Do Not Contact','Closed - Won','Closed - Lost'];
const CALL_OUTCOMES = ['','No Answer','Voicemail','Gatekeeper - Declined','Gatekeeper - Gave Info','Wrong Number','Bad Number','Do Not Call','Connected - Not Interested','Connected - Interested','Meeting Booked','Callback Requested'];

document.getElementById('meta').textContent = stats.generated_at + ' · ' + stats.total_leads.toLocaleString() + ' leads · ' + (stats.target_tier_counts['A+'] || 0) + ' A+ · ' + (stats.target_tier_counts['A'] || 0) + ' A';

// Filter chips
const chipDefs = [
  { id: 'priority', label: 'Priority (A+/A)', test: r => effective(r)['Target Tier'] === 'A+' || effective(r)['Target Tier'] === 'A' },
  { id: 'jr', label: 'JR', test: r => effective(r)['Product Line'] === 'JR' },
  { id: 'sn', label: 'S&N', test: r => effective(r)['Product Line'] === 'S&N' },
  { id: 'oos', label: 'OOS', test: r => effective(r)['Product Line'] === 'OOS' },
  { id: 'private', label: 'Private Practice', test: r => effective(r)['Practice Type'] === 'Private Practice' },
  { id: 'hospital', label: 'Hospital', test: r => effective(r)['Practice Type'] === 'Hospital-Based' },
  { id: 'microlyte', label: 'Microlyte OK', test: r => effective(r)['Microlyte Eligible'] === 'Yes' },
  { id: 'queued', label: 'Queued', test: r => effective(r)['Lead Status'] === 'Queued' },
  { id: 'interested', label: 'Interested', test: r => effective(r)['Lead Status'] === 'Interested' },
  { id: 'booked', label: 'Meeting Booked', test: r => effective(r)['Lead Status'] === 'Meeting Booked' },
  { id: 'callback', label: 'Callback', test: r => effective(r)['Lead Status'] === 'Callback Requested' },
];
let activeChip = 'priority';
let showAll = false;

const chipsEl = document.getElementById('chips');
for (const c of chipDefs) {
  const el = document.createElement('div');
  el.className = 'chip' + (c.id === activeChip ? ' active' : '');
  el.textContent = c.label;
  el.dataset.id = c.id;
  el.addEventListener('click', () => {
    activeChip = c.id;
    document.querySelectorAll('.chip').forEach(x => x.classList.toggle('active', x.dataset.id === activeChip));
    render();
  });
  chipsEl.appendChild(el);
}

function effective(row) {
  const patch = edits[row['HCP NPI']];
  return patch ? Object.assign({}, row, patch) : row;
}

function updatePending() {
  const n = Object.keys(edits).length;
  const badge = document.getElementById('pending-count');
  badge.textContent = n + ' unsaved';
  badge.classList.toggle('show', n > 0);
}

function phonePill(row) {
  const web = (row['Web Phone'] || '').replace(/\D/g, '');
  const nppes = (row['NPPES Phone'] || '').replace(/\D/g, '');
  const acuity = (row['Phone Number'] || '').replace(/\D/g, '');
  const isPrivate = row['Practice Type'] === 'Private Practice';
  const pieces = [];
  const seen = new Set();
  function add(num, label, klass) {
    if (!num || seen.has(num)) return;
    seen.add(num);
    const fmt = num.length === 10 ? '(' + num.slice(0,3) + ') ' + num.slice(3,6) + '-' + num.slice(6) : num;
    pieces.push('<a href="tel:' + num + '"><span class="phone-src ' + klass + '">' + label + '</span>' + fmt + '</a>');
  }
  if (isPrivate && web) add(web, 'Web', 'web');
  if (nppes) add(nppes, 'NPPES', 'web');
  if (acuity) add(acuity, 'Acuity', 'acuity');
  if (!isPrivate && web) add(web, 'Web', 'web');
  if (!pieces.length) return '<a class="missing">No phone on file</a>';
  return pieces.join('');
}

function targetBadge(t) {
  if (t === 'A+' || t === 'A') return '<span class="pill green">' + t + '</span>';
  if (t === 'B') return '<span class="pill amber">B</span>';
  if (t === 'C' || t === 'D') return '<span class="pill navy">' + t + '</span>';
  if (t === 'F') return '<span class="pill red">F</span>';
  return '';
}
function statusBadge(s) {
  if (!s || s === 'New') return '';
  if (s === 'Meeting Booked' || s === 'Closed - Won' || s === 'Interested') return '<span class="pill green">' + s + '</span>';
  if (s === 'Not Interested' || s === 'Do Not Contact' || s === 'Closed - Lost') return '<span class="pill red">' + s + '</span>';
  return '<span class="pill amber">' + s + '</span>';
}
function escHtml(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]); }

function render() {
  const q = document.getElementById('search').value.trim().toLowerCase();
  const chip = chipDefs.find(c => c.id === activeChip);
  let filtered = rows;
  if (!showAll && chip) filtered = filtered.filter(r => chip.test(r));
  if (q) {
    filtered = filtered.filter(r => {
      const e = effective(r);
      const hay = [e['First Name'], e['Last Name'], e['Primary Site of Care'], e['Email'], e['HCP NPI'], e['City']].join(' ').toLowerCase();
      return hay.indexOf(q) !== -1;
    });
  }
  // Sort: priority (A+/A first), then by Target Score desc
  filtered = filtered.slice().sort((a, b) => {
    const ea = effective(a), eb = effective(b);
    const rank = t => ({ 'A+': 0, 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'F': 5 })[t] ?? 6;
    const rd = rank(ea['Target Tier']) - rank(eb['Target Tier']);
    if (rd !== 0) return rd;
    return (parseInt(eb['Target Score'] || '0', 10) - parseInt(ea['Target Score'] || '0', 10));
  });

  document.getElementById('feed-count').textContent = filtered.length.toLocaleString() + ' leads';
  const feed = document.getElementById('feed');
  feed.innerHTML = '';
  const limit = Math.min(filtered.length, 200);
  if (!limit) {
    feed.innerHTML = '<div class="empty">No leads match. Tap another chip or clear the search.</div>';
    updatePending();
    return;
  }
  for (let i = 0; i < limit; i++) {
    const raw = filtered[i];
    const e = effective(raw);
    const card = document.createElement('div');
    card.className = 'card' + (edits[raw['HCP NPI']] ? ' edited' : '');
    card.innerHTML =
      '<div class="card-head">'
      + '<div>'
      + '<div class="card-name">' + escHtml(((e['First Name'] || '') + ' ' + (e['Last Name'] || '')).trim()) + ' <span class="pill navy">' + escHtml(e['Credential'] || '') + '</span></div>'
      + '<div class="card-sub">' + escHtml(e['Primary Site of Care'] || '') + '</div>'
      + '<div class="card-sub">' + escHtml(e['City'] || '') + (e['State'] ? ', ' + escHtml(e['State']) : '') + ' · ' + escHtml(e['Specialty'] || '') + '</div>'
      + '</div>'
      + '<div class="card-badges">' + targetBadge(e['Target Tier']) + statusBadge(e['Lead Status']) + '</div>'
      + '</div>'
      + '<div class="card-phone">' + phonePill(e) + '</div>'
      + '<div class="card-meta">'
      + '<span>' + escHtml(e['Product Line'] || '') + '</span>'
      + (e['Practice Type'] ? '<span>· ' + escHtml(e['Practice Type']) + '</span>' : '')
      + (e['Tier'] ? '<span>· ' + escHtml((e['Tier'] || '').replace(' (0', '').replace(' min)', '').split(' (')[0]) + '</span>' : '')
      + (e['Microlyte Eligible'] === 'Yes' ? '<span>· Microlyte ✓</span>' : '')
      + (e['Joint Repl Vol'] ? '<span>· JR Vol ' + escHtml(e['Joint Repl Vol']) + '</span>' : '')
      + '</div>'
      + '<div class="card-actions">'
      + '<button class="card-log">📞 Log Call</button>'
      + '<button class="card-open">View</button>'
      + '</div>';
    card.querySelector('.card-log').addEventListener('click', ev => { ev.stopPropagation(); openSheet(raw['HCP NPI'], 'logcall'); });
    card.querySelector('.card-open').addEventListener('click', () => openSheet(raw['HCP NPI'], 'view'));
    feed.appendChild(card);
  }
  if (filtered.length > limit) {
    const el = document.createElement('div');
    el.className = 'empty';
    el.textContent = 'Showing top ' + limit + ' of ' + filtered.length + '. Narrow the filter for more.';
    feed.appendChild(el);
  }
  updatePending();
}

document.getElementById('search').addEventListener('input', render);
document.getElementById('btn-toggle-all').addEventListener('click', () => {
  showAll = !showAll;
  document.getElementById('btn-toggle-all').textContent = showAll ? 'Apply Filter' : 'Show All';
  render();
});

const ROW_BY_NPI = {};
for (const r of rows) ROW_BY_NPI[r['HCP NPI']] = r;
let currentNpi = null;
let currentMode = 'view';

function todayIso() { const d = new Date(); d.setHours(12); return d.toISOString().slice(0,10); }

const EMAIL_OUTCOMES = ['','Sent','Bounced','Opened','Replied - Interested','Replied - Not Interested','Meeting Booked','Unsubscribed'];

function openSheet(npi, mode) {
  currentNpi = npi; currentMode = mode || 'view';
  const raw = ROW_BY_NPI[npi];
  if (!raw) return;
  const e = effective(raw);
  document.getElementById('sh-name').textContent = ((e['First Name'] || '') + ' ' + (e['Last Name'] || '')).trim();
  document.getElementById('sh-sub').innerHTML = escHtml(e['Primary Site of Care'] || '') + ' · ' + escHtml(e['City'] || '') + ', ' + escHtml(e['State'] || '') + ' · ' + escHtml(e['Specialty'] || '');

  let html = '';
  if (currentMode === 'logcall') {
    html += '<section class="log-form">'
      + '<h3>Log a call</h3>'
      + '<label>Date</label><div class="date-row"><input type="date" id="f-date" value="' + todayIso() + '" /><button data-d="0">Today</button><button data-d="-1">Yesterday</button></div>'
      + '<label>Outcome</label><select id="f-outcome">'
      + CALL_OUTCOMES.map(o => '<option value="' + escHtml(o) + '">' + (o || '(pick)') + '</option>').join('')
      + '</select>'
      + '<label>Notes</label><textarea id="f-notes" placeholder="What happened on the call?"></textarea>'
      + '<label>Update Lead Status</label><select id="f-status">'
      + LEAD_STATUSES.map(s => '<option value="' + escHtml(s) + '">' + (s || '(leave unchanged)') + '</option>').join('')
      + '</select>'
      + '</section>';
  } else if (currentMode === 'logemail') {
    html += '<section class="log-form">'
      + '<h3>Log an email</h3>'
      + '<label>Date</label><div class="date-row"><input type="date" id="f-date" value="' + todayIso() + '" /><button data-d="0">Today</button><button data-d="-1">Yesterday</button></div>'
      + '<label>Outcome</label><select id="f-outcome">'
      + EMAIL_OUTCOMES.map(o => '<option value="' + escHtml(o) + '">' + (o || '(pick)') + '</option>').join('')
      + '</select>'
      + '<label>Subject</label><input id="f-subject" value="' + escHtml(e['Subject Line'] || '') + '" />'
      + '<label>Notes</label><textarea id="f-notes" placeholder="Response, follow-up, etc."></textarea>'
      + '<label>Update Lead Status</label><select id="f-status">'
      + LEAD_STATUSES.map(s => '<option value="' + escHtml(s) + '">' + (s || '(leave unchanged)') + '</option>').join('')
      + '</select>'
      + '</section>';
  } else if (currentMode === 'appointment') {
    html += '<section class="log-form">'
      + '<h3>Schedule appointment</h3>'
      + '<label>Appointment Date</label><div class="date-row"><input type="date" id="f-appt-date" value="' + todayIso() + '" /></div>'
      + '<label>Next Action Description</label><input id="f-action" placeholder="e.g. Lunch & learn at their office" />'
      + '<label>Set Lead Status</label><select id="f-status">'
      + LEAD_STATUSES.map(s => '<option value="' + escHtml(s) + '"' + (s === 'Meeting Booked' ? ' selected' : '') + '>' + (s || '(leave unchanged)') + '</option>').join('')
      + '</select>'
      + '</section>';
  }

  // Quick actions toolbar
  const emailTo = (e['Email'] || '').trim();
  const subj = encodeURIComponent(e['Subject Line'] || '');
  const body = encodeURIComponent(e['Draft Email'] || '');
  const gmailUrl = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(emailTo) + '&su=' + subj + '&body=' + body;
  const mailto = 'mailto:' + encodeURIComponent(emailTo) + '?subject=' + subj + '&body=' + body;
  html += '<section><h3>Actions</h3>'
    + '<div class="action-grid">'
    + '<button class="act-btn" data-act="logcall">📞 Log Call</button>'
    + '<button class="act-btn" data-act="logemail">✉️ Log Email</button>'
    + '<button class="act-btn" data-act="appointment">📅 Appointment</button>'
    + (emailTo ? '<a class="act-btn" target="_blank" rel="noopener" href="' + gmailUrl + '">✉️ Gmail Compose</a>' : '')
    + (emailTo ? '<a class="act-btn" target="_blank" rel="noopener" href="' + mailto + '">✉️ Mail App</a>' : '')
    + (e['Draft Email'] ? '<button class="act-btn" data-act="showdraft">📝 View Draft</button>' : '')
    + '</div></section>';

  // Web lookup links
  const nameForSearch = ((e['First Name'] || '') + ' ' + (e['Last Name'] || '')).trim();
  const locForSearch = ((e['City'] || '') + ' ' + (e['State'] || '')).trim();
  const spec = (e['Specialty'] || '').split(',')[0].split('>')[0].trim();
  const googleQ = encodeURIComponent(nameForSearch + ' ' + spec + ' ' + locForSearch);
  const npiQ = encodeURIComponent(e['HCP NPI'] || '');
  html += '<section><h3>Look up online</h3>'
    + '<div class="action-grid">'
    + '<a class="act-btn ghost" target="_blank" rel="noopener" href="https://www.google.com/search?q=' + googleQ + '">🔍 Google</a>'
    + '<a class="act-btn ghost" target="_blank" rel="noopener" href="https://www.healthgrades.com/usearch?what=' + encodeURIComponent(nameForSearch) + '&where=' + encodeURIComponent(locForSearch) + '">Healthgrades</a>'
    + '<a class="act-btn ghost" target="_blank" rel="noopener" href="https://doctor.webmd.com/results?q=' + encodeURIComponent(nameForSearch) + '&location=' + encodeURIComponent(locForSearch) + '">WebMD</a>'
    + '<a class="act-btn ghost" target="_blank" rel="noopener" href="https://opennpi.com/provider/' + npiQ + '">OpenNPI</a>'
    + '</div></section>';

  html += '<section><h3>Phone</h3>' + phonePill(e).replace(/<a/g, '<a style="display:block;margin-bottom:6px;"') + '</section>';

  // Call history (rounds that already have dates)
  const callHistory = [];
  for (let i = 1; i <= 5; i++) {
    const d = e['Call ' + i + ' Date'];
    if (d && String(d).trim()) {
      callHistory.push('<div class="history-item"><div class="history-row"><span class="lbl">Call ' + i + ' · ' + escHtml(d) + '</span><span class="val pill navy">' + escHtml(e['Call ' + i + ' Outcome'] || '—') + '</span></div>'
        + (e['Call ' + i + ' Notes'] ? '<div class="history-note">' + escHtml(e['Call ' + i + ' Notes']) + '</div>' : '')
        + '</div>');
    }
  }
  if (callHistory.length) {
    html += '<section><h3>Call history (' + callHistory.length + ')</h3>' + callHistory.join('') + '</section>';
  }

  // Email history
  const emailHistory = [];
  for (let i = 1; i <= 3; i++) {
    const d = e['Email ' + i + ' Date'];
    if (d && String(d).trim()) {
      emailHistory.push('<div class="history-item"><div class="history-row"><span class="lbl">Email ' + i + ' · ' + escHtml(d) + '</span><span class="val pill navy">' + escHtml(e['Email ' + i + ' Outcome'] || '—') + '</span></div>'
        + (e['Email ' + i + ' Subject'] ? '<div class="history-note"><strong>' + escHtml(e['Email ' + i + ' Subject']) + '</strong></div>' : '')
        + (e['Email ' + i + ' Notes'] ? '<div class="history-note">' + escHtml(e['Email ' + i + ' Notes']) + '</div>' : '')
        + '</div>');
    }
  }
  if (emailHistory.length) {
    html += '<section><h3>Email history (' + emailHistory.length + ')</h3>' + emailHistory.join('') + '</section>';
  }

  // Next action / appointment
  if (e['Next Action'] || e['Next Action Date']) {
    html += '<section><h3>Next step</h3>'
      + '<div class="field"><span class="lbl">Action</span><span class="val">' + escHtml(e['Next Action'] || '-') + '</span></div>'
      + '<div class="field"><span class="lbl">Date</span><span class="val">' + escHtml(e['Next Action Date'] || '-') + '</span></div>'
      + '</section>';
  }

  html += '<section><h3>Lead</h3>'
    + '<div class="field"><span class="lbl">Target Tier</span><span class="val">' + escHtml(e['Target Tier'] || '') + '</span></div>'
    + '<div class="field"><span class="lbl">Practice Type</span><span class="val">' + escHtml(e['Practice Type'] || '') + '</span></div>'
    + '<div class="field"><span class="lbl">Status</span><span class="val">' + escHtml(e['Lead Status'] || 'New') + '</span></div>'
    + '<div class="field"><span class="lbl">MAC / Microlyte</span><span class="val">' + escHtml(e['MAC Jurisdiction'] || '') + ' / ' + escHtml(e['Microlyte Eligible'] || '') + '</span></div>'
    + '<div class="field"><span class="lbl">Email</span><span class="val">' + escHtml(e['Email'] || '(none)') + '</span></div>'
    + '</section>';

  if (e['Target Tier Reason']) {
    html += '<section><h3>Why ' + escHtml(e['Target Tier'] || '') + '</h3><div style="font-size:12px;background:#f7f9fa;padding:10px;border-radius:8px;white-space:pre-wrap;">' + escHtml(e['Target Tier Reason']).replace(/\|/g, '<br>') + '</div></section>';
  }

  document.getElementById('sh-body').innerHTML = html;
  const isLog = currentMode === 'logcall' || currentMode === 'logemail' || currentMode === 'appointment';
  document.getElementById('sh-save').style.display = isLog ? '' : 'none';
  document.getElementById('sh-close').textContent = isLog ? 'Cancel' : 'Close';
  document.getElementById('sheet').classList.add('show');
  document.getElementById('sheet-bg').classList.add('show');
  document.getElementById('sh-body').scrollTop = 0;

  // Wire up action toolbar buttons
  document.querySelectorAll('.act-btn[data-act]').forEach(b => {
    b.addEventListener('click', () => {
      const act = b.dataset.act;
      if (act === 'showdraft') {
        const body = (raw['Draft Email'] || '');
        const subj = (raw['Subject Line'] || '');
        if (navigator.clipboard) {
          navigator.clipboard.writeText('Subject: ' + subj + '\n\n' + body).then(() => toast('Draft copied to clipboard'));
        } else {
          alert('Subject: ' + subj + '\n\n' + body);
        }
        return;
      }
      openSheet(npi, act);
    });
  });

  document.querySelectorAll('.date-row button').forEach(b => {
    b.addEventListener('click', () => {
      const d = new Date(); d.setDate(d.getDate() + (parseInt(b.dataset.d, 10) || 0)); d.setHours(12);
      const target = document.getElementById('f-date') || document.getElementById('f-appt-date');
      if (target) target.value = d.toISOString().slice(0,10);
    });
  });
}

function closeSheet() {
  document.getElementById('sheet').classList.remove('show');
  document.getElementById('sheet-bg').classList.remove('show');
  currentNpi = null;
}
document.getElementById('sh-close').addEventListener('click', closeSheet);
document.getElementById('sheet-bg').addEventListener('click', closeSheet);

document.getElementById('sh-save').addEventListener('click', () => {
  if (!currentNpi) return;
  const raw = ROW_BY_NPI[currentNpi] || {};
  const patch = edits[currentNpi] ? Object.assign({}, edits[currentNpi]) : {};
  const fullName = ((raw['First Name'] || '') + ' ' + (raw['Last Name'] || '')).trim();

  if (currentMode === 'logcall') {
    const outcome = document.getElementById('f-outcome').value;
    if (!outcome) { toast('Pick an outcome first', true); return; }
    const date = document.getElementById('f-date').value || todayIso();
    const notes = document.getElementById('f-notes').value.trim();
    const newStatus = document.getElementById('f-status').value;
    let slot = 0;
    for (let i = 1; i <= 5; i++) {
      const existing = (patch['Call ' + i + ' Date'] !== undefined ? patch['Call ' + i + ' Date'] : raw['Call ' + i + ' Date']) || '';
      if (!existing.trim()) { slot = i; break; }
    }
    if (!slot) slot = 5;
    patch['Call ' + slot + ' Date'] = date;
    patch['Call ' + slot + ' Outcome'] = outcome;
    if (notes) patch['Call ' + slot + ' Notes'] = notes;
    if (newStatus) patch['Lead Status'] = newStatus;
    edits[currentNpi] = patch;
    localStorage.setItem(LS_KEY, JSON.stringify(edits));
    toast('Call ' + slot + ' logged for ' + fullName);
  } else if (currentMode === 'logemail') {
    const outcome = document.getElementById('f-outcome').value;
    if (!outcome) { toast('Pick an outcome first', true); return; }
    const date = document.getElementById('f-date').value || todayIso();
    const subject = document.getElementById('f-subject').value.trim();
    const notes = document.getElementById('f-notes').value.trim();
    const newStatus = document.getElementById('f-status').value;
    let slot = 0;
    for (let i = 1; i <= 3; i++) {
      const existing = (patch['Email ' + i + ' Date'] !== undefined ? patch['Email ' + i + ' Date'] : raw['Email ' + i + ' Date']) || '';
      if (!existing.trim()) { slot = i; break; }
    }
    if (!slot) slot = 3;
    patch['Email ' + slot + ' Date'] = date;
    patch['Email ' + slot + ' Outcome'] = outcome;
    if (subject) patch['Email ' + slot + ' Subject'] = subject;
    if (notes) patch['Email ' + slot + ' Notes'] = notes;
    if (newStatus) patch['Lead Status'] = newStatus;
    edits[currentNpi] = patch;
    localStorage.setItem(LS_KEY, JSON.stringify(edits));
    toast('Email ' + slot + ' logged for ' + fullName);
  } else if (currentMode === 'appointment') {
    const apptDate = document.getElementById('f-appt-date').value || todayIso();
    const action = document.getElementById('f-action').value.trim() || 'Meeting/Lunch & Learn';
    const newStatus = document.getElementById('f-status').value || 'Meeting Booked';
    patch['Next Action'] = action;
    patch['Next Action Date'] = apptDate;
    patch['Lead Status'] = newStatus;
    edits[currentNpi] = patch;
    localStorage.setItem(LS_KEY, JSON.stringify(edits));
    toast('Appointment set for ' + fullName);
  }
  closeSheet();
  render();
});

function toast(msg, isErr) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.background = isErr ? 'var(--red)' : 'var(--green)';
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2800);
}

document.getElementById('btn-download').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify(edits, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const ts = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
  a.download = 'activity_edits_' + ts + '.json';
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  toast('Downloaded ' + Object.keys(edits).length + ' edits');
});

render();
</script>
</body>
</html>
"""
