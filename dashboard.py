"""
dashboard.py  –  CHECKSETGO
Run: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
import re
from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

st.set_page_config(
    page_title="CHECKSETGO",
    page_icon="🦴",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --ivory-buff: #f4ecd8;
  --ivory-surface: #fbf6ea;
  --ivory-highlight: #fffaf1;
  --glaucous-light: #bfd1d2;
  --glaucous-soft: #e4efef;
  --glaucous-mid: #6e8f95;
  --green-blue: #2f6f73;
  --green-blue-deep: #244f56;
  --green-blue-soft: #dbeaea;
  --ink-main: #22383d;
  --ink-muted: #61777c;
  --line-soft: #d8ccb4;
  --buff-deep: #8c6b3f;
  --danger-soft: #f8dfdb;
  --danger-ink: #a13b35;
  --font-body: 16px;
  --font-label: 13px;
  --font-caption: 12px;
  --font-h1: clamp(28px, 4.6vw, 40px);
}

html, body {
  font-family: 'Inter', sans-serif;
  background-color: var(--ivory-buff);
  color: var(--ink-main);
  font-size: var(--font-body);
  line-height: 1.5;
}

body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
  background:
    radial-gradient(circle at top right, rgba(191, 209, 210, 0.35), transparent 26%),
    linear-gradient(180deg, var(--ivory-highlight) 0%, var(--ivory-buff) 38%, #f1e8d3 100%);
  color: var(--ink-main);
}

[data-testid="stHeader"] {
  background: rgba(244, 236, 216, 0.9);
  border-bottom: 1px solid var(--line-soft);
  backdrop-filter: blur(8px);
}

.main .block-container {
  padding-top: 1.8rem;
}

section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #ecf3f2 0%, #f7f0e0 100%);
  border-right: 1px solid var(--line-soft);
}

section[data-testid="stSidebar"] * {
  color: var(--ink-main) !important;
}

.stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown h5, .stMarkdown h6 {
  color: var(--green-blue-deep) !important;
  letter-spacing: -0.02em;
}

.stMarkdown h1 {
  font-size: var(--font-h1);
}

.stMarkdown p, .stMarkdown li {
  color: var(--ink-main);
  font-size: var(--font-body);
}

.stCaption, label, small {
  color: var(--ink-main);
  font-size: var(--font-label);
}

.app-title {
  font-family: 'JetBrains Mono', monospace;
  font-size: var(--font-h1);
  font-weight: 800;
  letter-spacing: -0.05em;
  color: var(--green-blue-deep);
  line-height: 1;
}

.page-timestamp {
  text-align: right;
  color: var(--ink-muted);
  font-size: var(--font-label);
  font-weight: 600;
  padding-top: 14px;
}

.app-footer {
  font-size: var(--font-caption);
  color: var(--ink-muted);
  text-align: center;
}

hr {
  border-color: var(--line-soft);
}

button[data-baseweb="tab"], button[role="tab"] {
  background: var(--ivory-surface) !important;
  border: 1px solid var(--line-soft) !important;
  border-radius: 12px 12px 0 0 !important;
  color: var(--ink-main) !important;
  padding: 0.5rem 0.95rem !important;
  transition: background-color .15s ease, color .15s ease, border-color .15s ease;
}

button[data-baseweb="tab"] *, button[role="tab"] * {
  color: inherit !important;
  font-weight: 700 !important;
  font-size: 16px !important;
}

button[data-baseweb="tab"]:hover, button[role="tab"]:hover {
  background: var(--glaucous-soft) !important;
  color: var(--green-blue-deep) !important;
}

button[data-baseweb="tab"][aria-selected="true"], button[role="tab"][aria-selected="true"] {
  background: var(--green-blue-soft) !important;
  border-color: var(--green-blue) !important;
  color: var(--green-blue-deep) !important;
  box-shadow: inset 0 -2px 0 var(--green-blue);
}

div[data-baseweb="tab-list"] {
  gap: 0.45rem;
  border-bottom: 1px solid var(--line-soft);
}

div[data-baseweb="tab-highlight"] {
  background: transparent !important;
}

.stButton > button {
  background: var(--green-blue);
  color: var(--ivory-highlight);
  border: 1px solid var(--green-blue-deep);
  border-radius: 10px;
  font-weight: 700;
}

.stButton > button:hover {
  background: var(--green-blue-deep);
  color: var(--ivory-highlight);
  border-color: var(--green-blue-deep);
}

div[data-baseweb="input"] > div {
  background: var(--ivory-highlight);
  border-color: var(--glaucous-light);
}

div[data-baseweb="input"] input {
  color: var(--ink-main);
  font-size: 16px !important;
}

textarea, input, select {
  font-size: 16px !important;
}

div[data-baseweb="input"]:focus-within > div {
  border-color: var(--green-blue);
  box-shadow: 0 0 0 1px rgba(47, 111, 115, 0.18);
}

details[data-testid="stExpander"] {
  background: var(--ivory-surface);
  border: 1px solid var(--line-soft);
  border-radius: 12px;
}

details[data-testid="stExpander"] summary,
details[data-testid="stExpander"] summary * {
  color: var(--ink-main) !important;
}

[data-testid="stDataFrame"] {
  background: var(--ivory-highlight);
  border: 1px solid var(--line-soft);
  border-radius: 12px;
  overflow: hidden;
}

.stAlert {
  background: var(--ivory-surface);
  border: 1px solid var(--line-soft);
  color: var(--ink-main);
}

.kpi-card { background: var(--ivory-surface); border: 1px solid var(--line-soft); border-radius: 10px; padding: 14px 18px; text-align: center; box-shadow: 0 1px 4px rgba(36,79,86,.08); }
.kpi-label { font-size: 12px; letter-spacing: .09em; color: var(--ink-muted); text-transform: uppercase; margin-bottom: 4px; font-weight: 600; }
.kpi-value { font-size: 26px; font-weight: 600; color: var(--green-blue-deep); font-family: 'JetBrains Mono', monospace; }
.kpi-value.warn  { color: var(--buff-deep); }
.kpi-value.alert { color: var(--danger-ink); }
.kpi-value.ok    { color: var(--green-blue); }

.avail-badge { display: inline-block; font-family: 'JetBrains Mono', monospace; font-size: 22px; font-weight: 700; padding: 4px 16px; border-radius: 8px; }
.avail-ok   { background: var(--green-blue-soft); color: var(--green-blue-deep); }
.avail-low  { background: #f5ead2; color: var(--buff-deep); }
.avail-zero { background: var(--danger-soft); color: var(--danger-ink); }

.inv-row { padding: 16px 0; border-bottom: 1px solid var(--line-soft); }
.inv-name { font-family: 'JetBrains Mono', monospace; font-size: 28px; font-weight: 700; color: var(--green-blue-deep); line-height: 1.2; }
.inv-sub  { font-size: 13px; color: var(--ink-muted); margin-top: 4px; }
.office-set-ids { font-family: 'JetBrains Mono', monospace; font-size: 24px; font-weight: 700; letter-spacing: -.04em; line-height: 1.15; color: var(--green-blue); }
.office-set-ids.is-standby { color: var(--buff-deep); }
.office-set-empty { color: var(--ink-muted); font-size: 12px; font-style: italic; }
.booking-next { margin-top: 8px; font-size: 12px; font-weight: 700; color: var(--green-blue); letter-spacing: .01em; }
.home-set-wrap { display: flex; flex-direction: column; gap: 8px; margin-top: 12px; }
.home-set-title { font-size: 12px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; color: var(--ink-muted); }
.home-set-list { display: flex; flex-wrap: wrap; gap: 8px; }
.home-set-chip { display: inline-flex; flex-direction: column; align-items: flex-start; gap: 3px; padding: 7px 12px; border-radius: 14px; background: var(--ivory-highlight); border: 1px solid var(--glaucous-light); box-shadow: 0 1px 2px rgba(36,79,86,.06); font-family: 'JetBrains Mono', monospace; font-size: 13px; font-weight: 800; line-height: 1.2; color: var(--green-blue-deep); }
.home-set-main { display: inline-flex; align-items: center; gap: 8px; }
.home-set-home { color: var(--buff-deep); font-size: 11px; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; }
.home-set-date { color: var(--green-blue); font-size: 11px; font-weight: 700; }
.service-set-wrap { display: flex; flex-direction: column; gap: 8px; margin-top: 12px; }
.service-set-title { font-size: 12px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; color: var(--ink-muted); }
.service-set-list { display: flex; flex-wrap: wrap; gap: 8px; }
.service-set-chip { display: inline-flex; flex-direction: column; align-items: flex-start; gap: 3px; padding: 7px 12px; border-radius: 14px; background: #f7e7df; border: 1px solid #d8b7aa; box-shadow: 0 1px 2px rgba(161,59,53,.05); font-family: 'JetBrains Mono', monospace; font-size: 13px; font-weight: 800; line-height: 1.2; color: var(--danger-ink); }
.service-set-main { display: inline-flex; align-items: center; gap: 8px; }
.service-set-date { color: #8e4741; font-size: 11px; font-weight: 700; }


.out-line { padding: 5px 0; margin-top: 6px; line-height: 1.6; border-left: 3px solid var(--line-soft); padding-left: 14px; margin-left: 4px; }
.out-order { font-size: 10px; font-weight: 700; color: var(--ink-muted); margin-right: 6px; letter-spacing: .02em; }
.out-tag { display: inline-block; font-size: 10px; font-weight: 700; letter-spacing: .06em; background: #f4e7cb; color: var(--buff-deep); border-radius: 3px; padding: 2px 7px; margin-right: 8px; vertical-align: middle; }
.out-tag.out-tag-booked { background: var(--glaucous-soft); color: var(--green-blue-deep); }
.out-sep { color: #8d9fa3; }
.out-set { font-family: 'JetBrains Mono', monospace; color: var(--green-blue); font-weight: 700; font-size: 15px; margin-right: 2px; }
.out-hosp-wrap { display: inline-flex; align-items: center; gap: 0; font-weight: 700; font-size: 15px; vertical-align: middle; }
.out-hosp-led { display: none; }
.out-hosp-wrap .out-hosp-name { display: inline-block; padding: 2px 9px 2px 8px; border-radius: 5px; border-left: 3px solid var(--glaucous-light); background: var(--ivory-highlight); color: var(--ink-main); font-size: 14px; font-weight: 600; line-height: 1.5; letter-spacing: -.01em; }
.out-hosp-wrap.is-delivered .out-hosp-name  { border-left-color: #0ea5a4; background: #f0fdfa; color: #0f766e; }
.out-hosp-wrap.is-surgery .out-hosp-name    { border-left-color: #f59e0b; background: #fffbeb; color: #92400e; }
.out-hosp-wrap.is-cancelled .out-hosp-name  { border-left-color: #ef4444; background: #fef2f2; color: #b91c1c; }
.out-hosp-wrap.is-postponed .out-hosp-name  { border-left-color: #f97316; background: #fff7ed; color: #c2410c; }
.out-hosp-wrap.is-sales-posted .out-hosp-name,
.out-hosp-wrap.is-sales .out-hosp-name      { border-left-color: #3b82f6; background: #eff6ff; color: #1d4ed8; }
.out-hosp-wrap.is-in-transit .out-hosp-name { border-left-color: #0ea5e9; background: #ecfeff; color: #0369a1; }
.out-hosp-wrap.is-in-transit.is-its .out-hosp-name { border-left-color: #0ea5e9; }
.out-hosp-wrap.is-in-transit.is-itd .out-hosp-name { border-left-color: #06b6d4; }
.out-hosp-wrap.is-checking .out-hosp-name   { border-left-color: #8b5cf6; background: #f5f3ff; color: #6d28d9; }
.out-hosp-wrap.is-shelf .out-hosp-name,
.out-hosp-wrap.is-collect .out-hosp-name,
.out-hosp-wrap.is-future .out-hosp-name,
.out-hosp-wrap.is-today .out-hosp-name,
.out-hosp-wrap.is-past .out-hosp-name       { border-left-color: #22c55e; background: #f0fdf4; color: #166534; }
.out-hosp-wrap.is-plate .out-hosp-name      { font-size: 15px; padding: 3px 10px 3px 9px; }
.out-surg  { color: var(--ink-muted); font-size: 14px; }

.sec-header { font-size: 12px; letter-spacing: .12em; text-transform: uppercase; color: var(--ink-muted); font-weight: 700; border-bottom: 1px solid var(--line-soft); padding-bottom: 6px; margin: 24px 0 14px 0; }

/* ── Meeple track ── */
.meeple-track { display: flex; align-items: flex-start; width: 100%; margin: 8px 0 4px 0; overflow-x: auto; }
.meeple-step { display: flex; flex-direction: column; align-items: center; flex: 0 0 auto; min-width: 58px; gap: 2px; }
.meeple-dot-row { height: 8px; display: flex; align-items: center; justify-content: center; margin-bottom: 1px; }
.meeple-current-dot { width: 8px; height: 8px; border-radius: 50%; }
.meeple-pill { min-width: 46px; height: 28px; padding: 0 8px; border-radius: 999px; border: 2px solid; display: inline-flex; align-items: center; justify-content: center; font-family: 'Inter', sans-serif; font-size: 10px; font-weight: 700; line-height: 1; letter-spacing: .03em; white-space: nowrap; }
.meeple-pill.mp-done    { color: #fff; }
.meeple-pill.mp-current { }
.meeple-pill.mp-future  { background: var(--ivory-highlight); border-color: var(--line-soft); color: var(--buff-deep); }
.meeple-step-label { font-size: 9px; font-weight: 700; text-align: center; line-height: 1.2; max-width: 58px; }
.meeple-step-date  { font-size: 16px; font-weight: 800; color: var(--ink-main); text-align: center; min-height: 20px; }
.meeple-connector  { flex: 1 1 0; min-width: 8px; height: 3px; border-radius: 999px; margin-top: 19px; align-self: flex-start; background: var(--line-soft); }
.meeple-terminal { display: inline-flex; align-items: center; gap: 6px; padding: 4px 12px; border-radius: 999px; font-size: 11px; font-weight: 700; letter-spacing: .05em; margin: 8px 0 2px 0; background: var(--glaucous-soft); color: var(--green-blue-deep); border: 1px solid var(--glaucous-light); }
.meeple-terminal.mp-pp  { background:#f5ead2; color:var(--buff-deep); border:2px solid #c8aa79; }

/* ── Upcoming chips ── */
.upcoming-chip-wrap { display: flex; flex-direction: column; gap: 4px; margin-top: 8px; }
.upcoming-chip { display: inline-flex; align-items: center; gap: 6px; background: var(--green-blue-soft); border: 1px solid var(--glaucous-light); border-radius: 7px; padding: 4px 9px; font-size: 11px; font-family: 'JetBrains Mono', monospace; color: var(--green-blue-deep); font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.upcoming-chip.is-booked { background: var(--glaucous-soft); border-color: var(--glaucous-light); color: var(--green-blue-deep); }
.upcoming-chip .uc-date { font-weight: 700; min-width: 44px; }
.upcoming-chip .uc-hosp { color: var(--ink-main); }
.upcoming-chip .uc-set { color: var(--green-blue-deep); font-weight: 700; }
.upcoming-chip .uc-set.is-tentative { color: var(--buff-deep); }
.upcoming-more { font-size: 10px; color: var(--ink-muted); font-style: italic; padding: 1px 4px; }
.out-next-wrap { display: flex; flex-wrap: wrap; gap: 6px; margin: 6px 0 0 26px; }
.out-next-chip { display: inline-flex; align-items: center; gap: 6px; padding: 3px 8px; border-radius: 999px; background: var(--glaucous-soft); color: var(--green-blue-deep); font-size: 10px; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; }
.out-next-chip.is-tentative { background: #f5ead2; color: var(--buff-deep); }

/* ── Plate drawers ── */
.dr-block { margin-top: 8px; border-radius: 8px; overflow: hidden; border: 1.5px solid var(--line-soft); }
.dr-hdr { display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; flex-wrap: wrap; padding: 4px 10px; background: var(--glaucous-soft); font-family: 'JetBrains Mono', monospace; font-size: 11px; font-weight: 700; color: var(--ink-main); letter-spacing: .06em; }
.dr-hdr.dh-out { background:#fff1f2; color:#be123c; }
.dh-out-list { display: flex; flex: 1; flex-wrap: wrap; justify-content: flex-end; gap: 6px; }
.dh-out-tag { display: inline-flex; align-items: center; flex-wrap: wrap; gap: 6px; font-size: 11px; font-weight: 700; background: #fecdd3; color: #9f1239; border-radius: 6px; padding: 4px 8px; }
.dh-out-tag.dht-stk { background:#fde68a; color:#92400e; }
.dh-out-sr    { font-size: 10px; letter-spacing: .08em; text-transform: uppercase; }
.dh-out-surg  { color: #7f1d1d; font-size: 13px; font-weight: 700; }
.dh-out-tag.dht-stk .dh-out-surg { color:#92400e; }
.dh-out-stock { font-size: 11px; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; }
.dr-chips { display: flex; flex-wrap: wrap; gap: 5px; padding: 7px 10px; background: var(--ivory-highlight); border-top: 1px solid var(--line-soft); }
.dr-chips.drc-out { background: #fff8eb; }
.sc { font-family: 'JetBrains Mono', monospace; font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; white-space: nowrap; border: 1.5px solid transparent; }
.sc-std  { background:#dbeafe; color:#1e40af; border-color:#93c5fd; }
.sc-sht  { background:#fce7f3; color:#9d174d; border-color:#f9a8d4; }
.sc-lng  { background:#d1fae5; color:#065f46; border-color:#6ee7b7; }
.sc-xl   { background:#ede9fe; color:#4c1d95; border-color:#c4b5fd; }
.sc-case { background:#fef3c7; color:#92400e; border-color:#fcd34d; }
.sc-none { background:#fee2e2; color:#9f1239; border-color:#fca5a5; text-decoration: line-through; opacity:.8; }
.sr-legend { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 6px; margin-bottom: 2px; }
.sr-legend-item { display: flex; align-items: center; gap: 4px; font-size: 10px; color: var(--ink-muted); }
.sr-legend-dot  { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }

/* ── Triage ── */
.triage-note { margin: 8px 0 14px 0; padding: 11px 14px; border-radius: 12px; border: 1px solid var(--line-soft); background: var(--ivory-surface); color: var(--ink-main); font-size: 13px; line-height: 1.45; }
.triage-note strong { color: var(--green-blue-deep); }
.triage-note.is-warn { background: #f5ead2; border-color: #e2c99b; }
.triage-note.is-warn strong { color: var(--buff-deep); }
.triage-note.is-alert { background: var(--danger-soft); border-color: #e8b7b2; }
.triage-note.is-alert strong { color: var(--danger-ink); }
.triage-empty { margin: 8px 0; padding: 12px 14px; border-radius: 12px; border: 1px dashed var(--glaucous-light); background: var(--ivory-highlight); color: var(--ink-muted); font-size: 13px; font-style: italic; }
.triage-section-title { font-size: 12px; letter-spacing: .1em; text-transform: uppercase; font-weight: 800; color: var(--green-blue-deep); margin: 18px 0 4px 0; }
.triage-table-wrap { width: 100%; overflow-x: auto; margin-top: 8px; border-radius: 12px; border: 1px solid var(--line-soft); background: var(--ivory-highlight); box-shadow: 0 1px 4px rgba(36,79,86,.06); }
.triage-table { width: 100%; border-collapse: collapse; min-width: 760px; color: var(--ink-main); }
.triage-table thead th { position: sticky; top: 0; background: var(--green-blue-soft); color: var(--green-blue-deep); font-size: 13px; font-weight: 800; letter-spacing: .04em; text-align: left; padding: 12px 14px; border-bottom: 1px solid var(--glaucous-light); white-space: nowrap; }
.triage-table tbody td { color: var(--ink-main); font-size: 16px; padding: 12px 14px; border-bottom: 1px solid rgba(216, 204, 180, 0.7); vertical-align: top; line-height: 1.45; }
.triage-table tbody tr:nth-child(even) td { background: rgba(219, 234, 234, 0.22); }
.triage-table tbody tr:hover td { background: rgba(219, 234, 234, 0.38); }
.triage-table td:first-child, .triage-table th:first-child { font-family: 'JetBrains Mono', monospace; }

@media (max-width: 768px) {
  :root {
    --font-body: 16px;
    --font-label: 12px;
    --font-caption: 12px;
    --font-h1: clamp(28px, 8vw, 34px);
  }

  .main .block-container {
    padding-top: 1rem;
    padding-left: 0.9rem;
    padding-right: 0.9rem;
  }

  .page-timestamp {
    text-align: left;
    padding-top: 8px;
  }

  .kpi-card {
    padding: 12px 14px;
  }

  .kpi-value {
    font-size: 28px;
  }

  .triage-note,
  .triage-empty,
  .inv-sub,
  .out-surg {
    font-size: 16px;
  }

  .out-tag,
  .home-set-home,
  .home-set-date,
  .service-set-date,
  .meeple-pill,
  .meeple-step-label,
  .meeple-terminal,
  .upcoming-chip,
  .upcoming-more,
  .out-next-chip,
  .dr-hdr,
  .dh-out-tag,
  .dh-out-sr,
  .dh-out-stock,
  .sc,
  .sr-legend-item,
  .triage-section-title,
  .sec-header,
  .kpi-label {
    font-size: 12px !important;
  }

  .triage-table {
    min-width: 680px;
  }

  .triage-table thead th {
    font-size: 12px;
  }

  .triage-table tbody td {
    font-size: 16px;
  }
}
</style>
""", unsafe_allow_html=True)

KL_TZ = ZoneInfo("Asia/Kuala_Lumpur")
DEFAULT_MASTER_DATA_PATH = str(Path(__file__).resolve().with_name("master_data.py"))
_POWERTOOL_CATS = {"P5503", "P5400", "P8400"}
THEME_GREEN_BLUE = "#2f6f73"
THEME_GLAUCOUS = "#6e8f95"
THEME_BUFF = "#8c6b3f"
THEME_CANCELLED = "#8a5f56"
THEME_LINE_SOFT = "#d8ccb4"

OFFICE_VIEW_ORDER = [
    "STD CANNA 6.5/7.3","STD CANNA 4.0","STD CANNA 2.4","STD CANNA 3.0",
    "RFN","PFN II 170-240","PFN II 340-420 SYSTEM","PFN II 340-420 IMPLANT",
    "ANKLE ARTHRODESIS NAIL","COATLMON CABLE SYSTEM","FOOT SET","FOOT INSTRUMENT",
    "PFN II NAIL REMOVAL","P5503","P5400","P8400","FIBULAR NAIL","FNS","ROI",
    "2.7-4.0","3.5-6.5","PFN","LONG PFN","REAMER","ILN TIBIA SUPSUB","ILN HUMERUS",
    "2.4-2.7","1.5-2.0","ILN RADIUS ULNA","2.0-2.4","2.0 ONLY","TENS",
    "CANNA 2.5","CANNA 3.5","CANNA 4.0","CANNA 5.2","ILN FEMUR","ILN TIBIA",
]
PLATE_UID_ORDER = [
    "PHILOS","OLEI","OLEII","DSC","MSC","CHOOK","DIA","URS","RECON","METAI","METAII",
    "TUBULAR","DFIBII","DFIBIII","DMH","DPLH","DLHI","DLHII","DMT","DLT","CMESH",
    "CCOMBO","PPMT","DPT","PLTI","PLTII","PMT","DLF","TSP","FSP","PFP","APP",
    "FIBHOOK","DLTII","ADT","DRVL",
]
PLATE_UID_RANK = {uid: idx for idx, uid in enumerate(PLATE_UID_ORDER)}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Sidebar
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with st.sidebar:
    st.markdown("### ⚙️ Data Sources")
    master_path   = st.text_input("master_data.py path", value=DEFAULT_MASTER_DATA_PATH)
    cases_source  = st.text_input("Cases CSV (path or URL)", value=(
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IY"
        "nirCfsnGHoMyn5xPoG5c/pub?gid=0&single=true&output=csv"))
    archive_source = st.text_input("Archive CSV (path or URL)", value=(
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IY"
        "nirCfsnGHoMyn5xPoG5c/pub?gid=1320419668&single=true&output=csv"))
    refresh      = st.button("🔄 Refresh Data", use_container_width=True)
    st.divider()
    search_query = st.text_input("🔍 Search", placeholder="cases, sets, plates, hospitals…")
    st.divider()
    st.caption("TZ: Asia/Kuala_Lumpur")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Load / cache
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
@st.cache_data(show_spinner="Loading data…", ttl=60)
def load_report(master: str, cases: str, archive: str, sig: str) -> dict:
    from ops_engine import build_operations_report
    return build_operations_report(master_data_path=master, cases_source=cases or None, archive_source=archive or None)

@st.cache_data(show_spinner=False, ttl=60)
def load_master_sets(master: str, sig: str) -> list[dict]:
    from ops_engine import load_master_data
    raw_sets = load_master_data(master).get("SETS", [])
    return list(raw_sets) if isinstance(raw_sets, list) else []

def _source_sig(value: str) -> str:
    text = str(value or "").strip()
    if not text: return ""
    p = Path(text).expanduser()
    if p.exists():
        s = p.stat()
        return f"{p.resolve()}:{s.st_mtime_ns}:{s.st_size}"
    return text

if refresh:
    load_report.clear()

sig = "|".join([_source_sig(master_path), _source_sig(cases_source), _source_sig(archive_source)])
if "report" not in st.session_state or refresh or st.session_state.get("report_signature") != sig:
    with st.spinner("Fetching data…"):
        try:
            st.session_state["report"] = load_report(master_path, cases_source, archive_source, sig)
            st.session_state["report_signature"] = sig
            st.session_state["load_error"] = None
        except Exception as exc:
            st.session_state["load_error"] = str(exc)

if st.session_state.get("load_error"):
    st.error(f"❌ Failed to load data: {st.session_state['load_error']}")
    st.stop()

master_sets = load_master_sets(master_path, _source_sig(master_path))
report       = st.session_state["report"]
meta         = report["meta"]
case_rows_all = report.get("cases_all", [])
cases_all_df  = pd.DataFrame(case_rows_all)

now_kl = datetime.now(KL_TZ)
try:
    report_today = datetime.strptime(str(meta.get("today_kl","")).strip(), "%Y-%m-%d").date()
except ValueError:
    report_today = now_kl.date()

# ── Per-case lookups ──────────────────────────────────────────────────────────
case_sales_lookup:      dict[str,str]  = {}
case_delivery_lookup:   dict[str,str]  = {}
case_surgery_lookup:    dict[str,str]  = {}
case_return_lookup:     dict[str,str]  = {}
case_prefix_lookup:     dict[str,str]  = {}
case_status_lookup:     dict[str,str]  = {}
case_is_booking_lookup: dict[str,bool] = {}
upcoming_case_rows: list[dict] = report.get("case_buckets", {}).get("to_deliver", [])
if not isinstance(upcoming_case_rows, list):
    upcoming_case_rows = []
# Fallback path for older reports that do not include case_buckets.
if not upcoming_case_rows and not cases_all_df.empty:
    for _, r in cases_all_df.iterrows():
        if not bool(r.get("has_shorthand_only", False)):
            continue
        if bool(r.get("is_cancelled_case", False)):
            continue
        del_str  = str(r.get("delivery_date","") or "").strip()
        del_date = None
        for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y"):
            try:
                del_date = datetime.strptime(del_str, fmt).date()
                break
            except:
                pass
        if del_date is None or del_date < report_today:
            continue
        upcoming_case_rows.append(r.to_dict())
upcoming_case_ids = {
    str(r.get("case_id","")).strip()
    for r in upcoming_case_rows
    if str(r.get("case_id","")).strip()
}

# category_norm → list of upcoming case entries (delivery_date >= today, not cancelled)
category_upcoming:      dict[str,list] = {}
suggested_case_queue_by_set_key: dict[str, list[dict]] = {}

if not cases_all_df.empty and "case_id" in cases_all_df.columns:
    for _, r in cases_all_df.iterrows():
        cid = str(r.get("case_id","")).strip()
        if not cid: continue
        case_sales_lookup[cid]      = str(r.get("sales_code","") or "").strip()
        case_delivery_lookup[cid]   = str(r.get("delivery_date","") or "").strip()
        case_surgery_lookup[cid]    = str(r.get("surgery_date","") or "").strip()
        case_return_lookup[cid]     = str(r.get("return_date","") or "").strip()
        case_prefix_lookup[cid]     = str(r.get("prefix","") or "").strip().upper()
        case_status_lookup[cid]     = str(r.get("status","") or "").strip().upper()
        case_is_booking_lookup[cid] = bool(r.get("is_booking_case", False))
        suggested_sets = r.get("suggested_sets", [])
        if not isinstance(suggested_sets, list):
            suggested_sets = []
        if cid in upcoming_case_ids:
            for item in suggested_sets:
                if not isinstance(item, dict):
                    continue
                set_key = str(item.get("set_key", "")).strip()
                if not set_key:
                    continue
                suggestion_date = None
                for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y"):
                    try:
                        suggestion_date = datetime.strptime(str(r.get("delivery_date", "") or "").strip(), fmt).date()
                        break
                    except:
                        pass
                suggested_case_queue_by_set_key.setdefault(set_key, []).append({
                    "case_id": cid,
                    "hospital": str(r.get("hospital", "") or "").strip(),
                    "date": str(r.get("delivery_date", "") or "").strip(),
                    "date_value": suggestion_date,
                    "confirmed": bool(item.get("confirmed", False)),
                    "category": str(item.get("category", "")).strip(),
                })

    for case in upcoming_case_rows:
        case_id = str(case.get("case_id","")).strip()
        if not case_id:
            continue
        if case.get("has_uid_set") or case.get("has_active_uid_set"):
            continue

        del_str = str(case.get("delivery_date","") or "").strip()
        del_date = None
        for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y"):
            try:
                del_date = datetime.strptime(del_str, fmt).date()
                break
            except:
                pass
        if del_date is None:
            continue

        suggested_sets = case.get("suggested_sets", [])
        if not isinstance(suggested_sets, list):
            suggested_sets = []
        suggested_by_category: dict[str, dict] = {}
        for item in suggested_sets:
            if not isinstance(item, dict):
                continue
            cat_norm = str(item.get("category", "")).upper().strip()
            if cat_norm:
                suggested_by_category[cat_norm] = item

        set_cats = case.get("set_categories", [])
        if not isinstance(set_cats, list):
            set_cats = []
        for cat in set_cats:
            cn = str(cat).upper().strip()
            if not cn:
                continue
            suggestion = suggested_by_category.get(cn, {})
            category_upcoming.setdefault(cn, []).append({
                "case_id":    case_id,
                "hospital":   str(case.get("hospital","") or "").strip(),
                "date":       del_str,
                "date_value": del_date,
                "kind":       "PENDING",
                "suggested_set": str(suggestion.get("set_display", "")).strip(),
                "suggested_confirmed": bool(suggestion.get("confirmed", False)),
            })

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Header
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
c1, c2 = st.columns([3,1])
with c1: st.markdown("<div class='app-title'>CHECKSETGO</div>", unsafe_allow_html=True)
with c2: st.markdown(f"<div class='page-timestamp'>{now_kl.strftime('%A, %d %b %Y  %H:%M')} KL</div>", unsafe_allow_html=True)
st.divider()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Helper functions
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def avail_badge(available: int, total: int) -> str:
    if total == 0: return "<span class='avail-badge avail-ok'>—</span>"
    if available == 0: return f"<span class='avail-badge avail-zero'>{available}/{total}</span>"
    if available/total <= 0.25: return f"<span class='avail-badge avail-low'>{available}/{total}</span>"
    return f"<span class='avail-badge avail-ok'>{available}/{total}</span>"

def _parse_ui_date(value: str) -> Optional[date]:
    text = str(value or "").strip()
    if not text: return None
    for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y"):
        try: return datetime.strptime(text, fmt).date()
        except: pass
    return None

def _safe_int(val) -> int:
    n = pd.to_numeric(val, errors="coerce")
    return int(n) if pd.notna(n) else 0

def _compact_set_id(value: str, fallback: str = "") -> str:
    t = str(value or "").strip() or str(fallback or "").strip()
    if not t: return ""
    return str(int(t)) if re.fullmatch(r"\d+", t) else t

def _format_last_maintained(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    d = _parse_ui_date(text)
    if d is None:
        return text
    return f"{d.day} {d.strftime('%b %Y')}"


def _is_past_or_today(value) -> bool:
    if isinstance(value, date):
        return value <= report_today
    d = _parse_ui_date(str(value or "").strip())
    if d is None:
        return False
    return d <= report_today

def _kpi_card(label: str, value: str | int, tone: str = "ok") -> str:
    return (
        f"<div class='kpi-card'>"
        f"<div class='kpi-label'>{escape(str(label))}</div>"
        f"<div class='kpi-value {tone}'>{escape(str(value))}</div>"
        f"</div>"
    )

def _filter_df_by_search(df: pd.DataFrame, query: str) -> pd.DataFrame:
    text = str(query or "").strip().lower()
    if not text or df.empty:
        return df
    haystack = df.fillna("").astype(str).agg(" ".join, axis=1).str.lower()
    return df[haystack.str.contains(text, regex=False)]

def _triage_note(text: str, tone: str = "") -> str:
    cls = "triage-note"
    if tone:
        cls += f" is-{tone}"
    return f"<div class='{cls}'>{text}</div>"

def _triage_empty(text: str) -> str:
    return f"<div class='triage-empty'>{escape(text)}</div>"

def _case_item_labels(items, fallback: str = "") -> str:
    labels: list[str] = []
    if isinstance(items, list):
        for item in items:
            if isinstance(item, dict):
                for key in ("label", "set_display", "proper_name", "name"):
                    value = str(item.get(key, "") or "").strip()
                    if value:
                        labels.append(value)
                        break
            else:
                value = str(item or "").strip()
                if value:
                    labels.append(value)
    if labels:
        return "; ".join(labels)
    text = str(fallback or "").strip()
    return text or "—"

def _triage_table(df: pd.DataFrame) -> str:
    if df.empty:
        return ""
    headers = "".join(f"<th>{escape(str(col))}</th>" for col in df.columns)
    body_rows: list[str] = []
    for _, row in df.iterrows():
        cells: list[str] = []
        for value in row.tolist():
            text = "—" if pd.isna(value) or str(value).strip() == "" else str(value)
            cells.append(f"<td>{escape(text)}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")
    return (
        "<div class='triage-table-wrap'>"
        "<table class='triage-table'>"
        f"<thead><tr>{headers}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table>"
        "</div>"
    )

# ── Meeple track ──────────────────────────────────────────────────────────────
def _render_meeple_steps(steps: list, accent: str) -> str:
    last_done = -1
    for i, (_, _, done, _) in enumerate(steps):
        if done: last_done = i
    current_idx = min(last_done + 1, len(steps) - 1)

    parts = ['<div class="meeple-track">']
    for i, (label, short, done, date_str) in enumerate(steps):
        is_current = (i == current_idx)
        is_done    = done and not is_current
        if is_done:
            pill_style = f"background:{accent};border-color:{accent};color:#fff"
            pill_cls   = "meeple-pill mp-done"
        elif is_current:
            pill_style = f"background:{accent}18;border-color:{accent};color:{accent}"
            pill_cls   = "meeple-pill mp-current"
        else:
            pill_style = ""
            pill_cls   = "meeple-pill mp-future"
        label_color = accent if (is_done or is_current) else THEME_BUFF
        dot_html = f'<span class="meeple-current-dot" style="background:{accent};box-shadow:0 0 0 3px {accent}33"></span>' if is_current else ""
        d_obj = _parse_ui_date(date_str)
        date_display = d_obj.strftime("%-d %b") if d_obj else ""
        parts.append(
            f'<div class="meeple-step">'
            f'<div class="meeple-dot-row">{dot_html}</div>'
            f'<div class="{pill_cls}" style="{pill_style}">{escape(short)}</div>'
            f'<div class="meeple-step-label" style="color:{label_color}">{escape(label)}</div>'
            f'<div class="meeple-step-date">{escape(date_display)}</div>'
            f'</div>'
        )
        if i < len(steps) - 1:
            parts.append(f'<div class="meeple-connector" style="background:{accent if is_done else THEME_LINE_SOFT}"></div>')
    parts.append('</div>')
    return "".join(parts)


def _build_meeple_steps(*, prefix, delivery, surgery, sales_code, status, is_booking, return_date=""):
    """
    Returns (steps, accent).
    Cancelled: Delivered → Cancelled → In Transit → Checking
    Parking P:  Surgery → Sales Posted → Top Up Prepared → Top Up Delivered
    Booked BC:  Booked → Delivered → Surgery → Sales Posted → In Transit → Checking
    Standard:   Delivered → Surgery → Sales Posted → In Transit → Checking
    Missing surgery → delivery + 1 day.
    """
    is_parking  = prefix.startswith("P") and not is_booking
    is_transit  = status in {"ITS","ITD"}
    is_checking = status in {"ITO","COMPLETED"}
    is_cnx      = status == "CNX"
    tl = "With Saiful" if status=="ITS" else "With Dylan" if status=="ITD" else "In Transit"

    # Infer surgery date if missing
    if not surgery and delivery:
        d = _parse_ui_date(delivery)
        if d: surgery = (d + timedelta(days=1)).strftime("%d/%m/%Y")

    if is_cnx:
        return [
            ("Delivered", "DEL",  _is_past_or_today(delivery),  delivery),
            ("Cancelled", "CNX",  True,                           ""),
            (tl,          "TRNST",is_transit,                     ""),
            ("Checking",  "CHK",  is_checking,                    ""),
        ], THEME_CANCELLED

    if is_parking:
        has_surg      = _is_past_or_today(surgery)
        has_sales     = bool(sales_code)
        has_prep      = _is_past_or_today(return_date)
        has_del_topup = is_checking or (has_sales and has_prep)
        return [
            ("Surgery",         "SURG",  has_surg,      surgery),
            ("Sales Posted",    "SALES", has_sales,     ""),
            ("Top Up Prepared", "PREP",  has_prep,      return_date),
            ("Top Up Delivered","DEL",   has_del_topup, ""),
        ], THEME_BUFF

    if is_booking:
        return [
            ("Booked",      "BKD",   True,                          ""),
            ("Delivered",   "DEL",   _is_past_or_today(delivery),   delivery),
            ("Surgery",     "SURG",  _is_past_or_today(surgery),    surgery),
            ("Sales Posted","SALES", bool(sales_code),               ""),
            (tl,            "TRNST", is_transit,                     ""),
            ("Checking",    "CHK",   is_checking,                    ""),
        ], THEME_GLAUCOUS

    return [
        ("Delivered",   "DEL",   _is_past_or_today(delivery),   delivery),
        ("Surgery",     "SURG",  _is_past_or_today(surgery),    surgery),
        ("Sales Posted","SALES", bool(sales_code),               ""),
        (tl,            "TRNST", is_transit,                     ""),
        ("Checking",    "CHK",   is_checking,                    ""),
    ], THEME_GREEN_BLUE

def _meeple_track_for_case_id(case_id: str, *, surgery="", delivery="", case_status="") -> str:
    cid    = str(case_id or "").strip()
    status = case_status.strip().upper() or case_status_lookup.get(cid, "")
    if status == "PP": return "<div class='meeple-terminal mp-pp'>⏸ Postponed</div>"
    steps, accent = _build_meeple_steps(
        prefix      = case_prefix_lookup.get(cid,""),
        delivery    = delivery or case_delivery_lookup.get(cid,""),
        surgery     = surgery  or case_surgery_lookup.get(cid,""),
        sales_code  = case_sales_lookup.get(cid,""),
        status      = status,
        is_booking  = case_is_booking_lookup.get(cid, False),
        return_date = case_return_lookup.get(cid,""),
    )
    return _render_meeple_steps(steps, accent)


def _hospital_status_class(sv, *, sales_code="", case_status="", is_booked=False, delivery_value="") -> str:
    s   = str(case_status or "").strip().upper()
    bdv = str(delivery_value or "").strip() or (str(sv or "").strip() if is_booked else "")
    if s == "CNX": return "is-cancelled"
    if s == "PP":  return "is-postponed"
    if is_booked and bdv: return "is-delivered"
    if _parse_ui_date(sv) is not None: return "is-surgery"
    if str(sales_code or "").strip(): return "is-sales-posted"
    if s in {"ITS","ITD"}: return f"is-in-transit is-{s.lower()}"
    if s in {"ITO","COMPLETED"}: return "is-checking"
    return "is-collect"

def _hospital_status_label(sv, *, sales_code="", case_status="", is_booked=False, delivery_value="") -> str:
    s   = str(case_status or "").strip().upper()
    bdv = str(delivery_value or "").strip() or (str(sv or "").strip() if is_booked else "")
    if s == "CNX": return "Cancelled"
    if s == "PP":  return "Postponed"
    if is_booked and bdv: return "Delivered"
    if _parse_ui_date(sv) is not None: return "Surgery"
    if str(sales_code or "").strip(): return "Sales posted"
    if s in {"ITS","ITD"}: return "In transit with Saiful" if s=="ITS" else "In transit with Dylan"
    if s in {"ITO","COMPLETED"}: return "Checking"
    return "Collect"

def _hospital_with_led(hospital, sv, *, variant="", sales_code="", case_status="", is_booked=False, delivery_value="") -> str:
    sc  = _hospital_status_class(sv, sales_code=sales_code, case_status=case_status, is_booked=is_booked, delivery_value=delivery_value)
    sl  = _hospital_status_label(sv, sales_code=sales_code, case_status=case_status, is_booked=is_booked, delivery_value=delivery_value)
    vc  = f"is-{variant}" if variant else ""
    ca  = " ".join(p for p in ("out-hosp-wrap", vc, sc) if p)
    return (
        f"<span class='{ca}' title='{escape(sl)}'>"
        f"<span class='out-hosp-led'></span>"
        f"<span class='out-hosp-name'>{escape(str(hospital or '—'))}</span>"
        f"</span>"
    )

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TABS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("<div class='sec-header'>Operational Detail — Inventory Snapshot</div>", unsafe_allow_html=True)
inv_tabs = st.tabs(["🧰 Sets", "🦾 Plates", "🔌 Powertools"])

# ══════════════════════════════════════════════════════════════════════════════
# SETS TAB
# ══════════════════════════════════════════════════════════════════════════════
with inv_tabs[0]:
    set_avail      = pd.DataFrame(report.get("set_category_availability", []))
    set_status_all = pd.DataFrame(report.get("set_office_status", []))
    pt_avail_df    = pd.DataFrame(report.get("powertool_category_availability", []))

    if set_avail.empty and set_status_all.empty:
        st.info("No set data.")
    else:
        for col in ("category","set_display","id","location_now","surgery_date","patient_doctor",
                    "case_id","set_status","home","assignment_kind","delivery_date","case_status","set_key"):
            if col not in set_status_all.columns: set_status_all[col] = ""
        set_status_all = set_status_all.copy()
        set_status_all["category_norm"]       = set_status_all["category"].astype(str).str.upper().str.strip()
        set_status_all["location_norm"]       = set_status_all["location_now"].astype(str).str.upper().str.strip()
        set_status_all["home_norm"]           = set_status_all["home"].astype(str).str.upper().str.strip()
        set_status_all["set_status_norm"]     = set_status_all["set_status"].astype(str).str.upper().str.strip()
        set_status_all["assignment_kind_norm"]= set_status_all["assignment_kind"].astype(str).str.upper().str.strip()
        set_status_all["is_na"]      = set_status_all["set_status_norm"].str.contains("NA", na=False)
        set_status_all["is_standby"] = (
            set_status_all["home_norm"].eq("STANDBY")
            | set_status_all["set_status_norm"].str.contains("STANDBY", na=False)
        )
        set_status_all["set_name"] = set_status_all["set_display"].astype(str).where(
            set_status_all["set_display"].astype(str).str.strip().ne(""),
            set_status_all["id"].astype(str),
        )
        set_status_all = set_status_all[~set_status_all["category_norm"].isin(_POWERTOOL_CATS)]

        if "category" not in set_avail.columns: set_avail["category"] = ""
        set_avail = set_avail.copy()
        set_avail["category_norm"] = set_avail["category"].astype(str).str.upper().str.strip()
        set_avail = set_avail[~set_avail["category_norm"].isin(_POWERTOOL_CATS)]

        avail_lookup:        dict[str,int]  = {}
        total_lookup:        dict[str,int]  = {}
        next_booking_lookup: dict[str,dict] = {}
        for _, r in set_avail.iterrows():
            k = str(r.get("category_norm","")).strip()
            if k:
                avail_lookup[k] = _safe_int(r.get("available",0))
                total_lookup[k] = _safe_int(r.get("total_office",0))
                next_booking_lookup[k] = {
                    "date":     str(r.get("next_booking_date","")).strip(),
                    "hospital": str(r.get("next_booking_hospital","")).strip(),
                    "case_id":  str(r.get("next_booking_case_id","")).strip(),
                    "set":      str(r.get("next_booking_set","")).strip(),
                }

        display_by_norm: dict[str,str] = {c.upper(): c for c in OFFICE_VIEW_ORDER}
        for src_df in (set_status_all, set_avail):
            for cat in src_df.get("category", pd.Series(dtype=str)).astype(str).tolist():
                cn = str(cat).upper().strip()
                if cn and cn not in display_by_norm: display_by_norm[cn] = str(cat).strip()

        ordered_norm = [c.upper() for c in OFFICE_VIEW_ORDER]
        ordered_norm.extend(sorted(n for n in display_by_norm if n not in ordered_norm))
        ordered_core = {c.upper() for c in OFFICE_VIEW_ORDER}
        home_items_by_category: dict[str, list[dict[str, str]]] = {}
        service_items_by_category: dict[str, list[dict[str, str]]] = {}
        for item in master_sets:
            cat_norm = str(item.get("category", "")).upper().strip()
            home_norm = str(item.get("home", "")).upper().strip()
            status_norm = str(item.get("status", "")).upper().strip()
            last_maintained = str(item.get("last_maintained", "")).strip()
            if not cat_norm or cat_norm in _POWERTOOL_CATS:
                continue
            label = _compact_set_id(item.get("id", ""), item.get("uid", ""))
            if "NA" in status_norm:
                if label:
                    service_items_by_category.setdefault(cat_norm, []).append({
                        "label": label,
                        "last_maintained": last_maintained,
                    })
                continue
            if home_norm in {"", "OFFICE", "STANDBY"}:
                continue
            if "STANDBY" in status_norm:
                continue
            home_items_by_category.setdefault(cat_norm, []).append({
                "label": label,
                "home": home_norm,
                "last_maintained": last_maintained,
            })
        for items in home_items_by_category.values():
            items.sort(key=lambda item: (
                0 if str(item.get("label", "")).isdigit() else 1,
                int(item["label"]) if str(item.get("label", "")).isdigit() else 0,
                str(item.get("label", "")),
                str(item.get("home", "")),
            ))
        for items in service_items_by_category.values():
            items.sort(key=lambda item: (
                0 if str(item.get("label", "")).isdigit() else 1,
                int(item["label"]) if str(item.get("label", "")).isdigit() else 0,
                str(item.get("label", "")),
            ))

        summary_rows: list[dict] = []
        out_rows:     list[dict] = []

        for cat_norm in ordered_norm:
            label = display_by_norm.get(cat_norm, cat_norm)
            cr    = set_status_all[set_status_all["category_norm"] == cat_norm]
            office_rows  = cr[(cr["location_norm"]=="OFFICE") & ~cr["is_na"] & ~cr["is_standby"]]
            standby_rows = cr[(cr["location_norm"].eq("STANDBY") | (cr["location_norm"].eq("OFFICE") & cr["is_standby"])) & ~cr["is_na"]]
            booked_rows  = cr[cr["assignment_kind_norm"].eq("BOOKED") & ~cr["is_na"]]
            out_case_rows= cr[~cr["location_norm"].isin(["OFFICE","STANDBY"]) & (cr["location_norm"]!="") & ~cr["is_na"] & ~cr["assignment_kind_norm"].eq("BOOKED")]

            in_count  = len(office_rows)
            sb_count  = len(standby_rows)
            bk_count  = len(booked_rows)
            out_count = len(out_case_rows)
            available = avail_lookup.get(cat_norm, in_count)
            total     = total_lookup.get(cat_norm, in_count+out_count+bk_count)
            total     = total if total > 0 else (in_count+out_count+bk_count)

            if cat_norm not in ordered_core and total==0 and available==0 and out_count==0 and sb_count==0 and bk_count==0:
                continue

            office_items  = [{"name":str(r["set_name"]),"label":_compact_set_id(r.get("id",""),r.get("set_name","")),"standby":False} for _,r in office_rows.iterrows()]
            standby_items = [{"name":str(r["set_name"]),"label":_compact_set_id(r.get("id",""),r.get("set_name","")),"standby":True}  for _,r in standby_rows.iterrows()]

            out_case_items: list[dict] = []
            for _, r in out_case_rows.iterrows():
                out_case_items.append({
                    "kind":"OUT","set_name":str(r.get("set_name","")).strip(),
                    "hospital":str(r.get("location_now","")).strip() or "OUT",
                    "surgery_date":str(r.get("surgery_date","")).strip(),
                    "delivery_date":str(r.get("delivery_date","")).strip(),
                    "date_value":_parse_ui_date(str(r.get("delivery_date",""))) or date.max,
                    "case_id":str(r.get("case_id","")).strip(),
                    "case_status":str(r.get("case_status","")).strip(),
                    "sales_code":case_sales_lookup.get(str(r.get("case_id","")).strip(),""),
                    "set_key":str(r.get("set_key","")).strip(),
                })
            booked_case_items: list[dict] = []
            for _, r in booked_rows.iterrows():
                booked_case_items.append({
                    "kind":"BOOKED","set_name":str(r.get("set_name","")).strip(),
                    "hospital":str(r.get("location_now","")).strip() or "BOOKED",
                    "delivery_date":str(r.get("delivery_date","")).strip(),
                    "date_value":_parse_ui_date(str(r.get("delivery_date",""))) or date.max,
                    "case_id":str(r.get("case_id","")).strip(),
                    "case_status":str(r.get("case_status","")).strip(),
                    "sales_code":case_sales_lookup.get(str(r.get("case_id","")).strip(),""),
                    "set_key":str(r.get("set_key","")).strip(),
                })

            # Upcoming cases — from cases_all index, delivery_date >= today
            raw_up = sorted(category_upcoming.get(cat_norm, []), key=lambda x: (x["date_value"], x["case_id"]))
            seen_up: set[str] = set()
            upcoming_cases: list[dict] = []
            for item in raw_up:
                if item["case_id"] not in seen_up:
                    seen_up.add(item["case_id"])
                    upcoming_cases.append(item)

            out_list = "; ".join(
                f"{x['set_name']} @ {x['hospital']} ({x.get('delivery_date') or x.get('surgery_date') or '-'})"
                for x in sorted(out_case_items+booked_case_items, key=lambda x: x["date_value"])
            )
            upcoming_text = " ".join(f"{x['case_id']} {x['hospital']} {x['date']}" for x in upcoming_cases)

            summary_rows.append({
                "Category":cat_norm,"DisplayLabel":label,"CategoryNorm":cat_norm,
                "In Office":in_count,"Booked":bk_count,"Out":out_count,"Available":available,
                "OfficeItems":office_items,"StandbyItems":standby_items,
                "HomeItems":list(home_items_by_category.get(cat_norm, [])),
                "ServiceItems":list(service_items_by_category.get(cat_norm, [])),
                "NextBooking":next_booking_lookup.get(cat_norm,{}),
                "In Office Sets":", ".join(sorted(i["name"] for i in office_items+standby_items)),
                "Home Sets":", ".join(
                    " ".join(filter(None, [i["label"], i["home"], i.get("last_maintained", "")]))
                    for i in home_items_by_category.get(cat_norm, [])
                ),
                "Service Sets":", ".join(
                    " ".join(filter(None, [i["label"], i.get("last_maintained", "")]))
                    for i in service_items_by_category.get(cat_norm, [])
                ),
                "Out (Hospital•Surgery)":out_list,
                "UpcomingCases":upcoming_cases,"UpcomingCasesText":upcoming_text,
            })
            for item in out_case_items:
                out_rows.append({"Category":cat_norm,"Set":item["set_name"],"Hospital":item["hospital"],
                    "Surgery Date":item["surgery_date"],"Delivery Date":item["delivery_date"],
                    "Case":item["case_id"],"Case Status":item["case_status"],"Sales Code":item["sales_code"],
                    "DateValue":item["date_value"],"set_key":item.get("set_key","")})
            for item in booked_case_items:
                out_rows.append({"Category":cat_norm,"Set":item["set_name"],"Hospital":item["hospital"],
                    "Surgery Date":"","Delivery Date":item["delivery_date"],
                    "Case":item["case_id"],"Case Status":item["case_status"],"Sales Code":item["sales_code"],
                    "Assignment":"BOOKED","DateValue":item["date_value"],"set_key":item.get("set_key","")})

        _set_out: dict[str,list] = {}
        for r in out_rows: _set_out.setdefault(r["Category"].upper().strip(), []).append(r)

        filtered_summary = summary_rows
        if search_query:
            sq = search_query.lower()
            filtered_summary = [
                r for r in summary_rows
                if sq in r["DisplayLabel"].lower()
                or sq in r["In Office Sets"].lower()
                or sq in r.get("Home Sets","").lower()
                or sq in r.get("Service Sets","").lower()
                or sq in r["Out (Hospital•Surgery)"].lower()
                or sq in r.get("UpcomingCasesText","").lower()
            ]

        def _set_row_html(r: dict) -> str:
            cn        = r["CategoryNorm"]
            avail_val = int(r.get("Available",0))
            total_val = int(total_lookup.get(cn, r["In Office"]+r["Out"]+r.get("Booked",0)))
            total_val = total_val if total_val > 0 else (r["In Office"]+r["Out"]+r.get("Booked",0))
            badge = avail_badge(avail_val, total_val)

            oi = list(r.get("OfficeItems",[])); si = list(r.get("StandbyItems",[]))
            in_html = ""
            if oi or si:
                ol = [escape(str(i.get("label",""))) for i in oi if str(i.get("label","")).strip()]
                sl = [escape(str(i.get("label",""))) for i in si if str(i.get("label","")).strip()]
                if ol: in_html += f"<div class='office-set-ids'>{', '.join(ol)}</div>"
                if sl: in_html += f"<div class='office-set-ids is-standby'>{', '.join(sl)} [standby]</div>"
                if not in_html: in_html = "<span class='office-set-empty'>none available</span>"
            else:
                in_html = "<span class='office-set-empty'>none available</span>"

            nb = dict(r.get("NextBooking",{}))
            nb_parts = [p for p in [escape(str(nb.get("date","")).strip()), escape(str(nb.get("hospital","")).strip())] if p]
            nb_html = ""
            if nb_parts:
                nb_lbl = " · ".join(nb_parts)
                if nb.get("set"): nb_lbl += f" · {escape(str(nb['set']))}"
                if nb.get("case_id"): nb_lbl += f" · {escape(str(nb['case_id']))}"
                nb_html = f"<div class='booking-next'>next booking {nb_lbl}</div>"

            upcoming_cases = list(r.get("UpcomingCases",[]) or [])
            up_parts = []
            for item in upcoming_cases[:5]:
                is_booked = str(item.get("kind","")).upper() == "BOOKED"
                chip_cls  = "upcoming-chip is-booked" if is_booked else "upcoming-chip"
                d_obj     = _parse_ui_date(str(item.get("date","")).strip())
                date_lbl  = d_obj.strftime("%-d %b") if d_obj else escape(str(item.get("date","")) or "—")
                da        = (d_obj - report_today).days if d_obj else None
                hint      = " today" if da==0 else " tmrw" if da==1 else (f" {da}d" if da is not None and da<=7 else "")
                suggested_set = str(item.get("suggested_set","")).strip()
                suggested_cls = "uc-set" if bool(item.get("suggested_confirmed", False)) else "uc-set is-tentative"
                suggested_html = (
                    f"<span class='{suggested_cls}'>→ {escape(suggested_set)}</span>"
                    if suggested_set else
                    ""
                )
                up_parts.append(
                    f"<span class='{chip_cls}'>"
                    f"<span class='uc-date'>{date_lbl}{escape(hint)}</span>"
                    f"<span class='uc-hosp'>{escape(str(item.get('hospital','—')))}</span>"
                    f"{suggested_html}"
                    f"<span style='color:#9ca3af;font-size:10px'>{escape(str(item.get('case_id','')))}</span>"
                    f"</span>"
                )
            up_html = ""
            if up_parts:
                more = max(len(upcoming_cases)-5, 0)
                up_html = "<div class='upcoming-chip-wrap'>" + "".join(up_parts) + (f"<span class='upcoming-more'>+{more} more</span>" if more else "") + "</div>"

            home_parts = []
            for item in list(r.get("HomeItems", []) or []):
                label = escape(str(item.get("label", "")).strip())
                home = escape(str(item.get("home", "")).strip())
                maintained = escape(_format_last_maintained(str(item.get("last_maintained", "")).strip()))
                if not label and not home:
                    continue
                home_label_html = f"<span>{label or '—'}</span>"
                home_home_html = f"<span class='home-set-home'>{home}</span>" if home else ""
                maintained_html = f"<span class='home-set-date'>{maintained}</span>" if maintained else ""
                home_parts.append(
                    f"<span class='home-set-chip'>"
                    f"<span class='home-set-main'>{home_label_html}{home_home_html}</span>"
                    f"{maintained_html}"
                    f"</span>"
                )
            home_html = ""
            if home_parts:
                home_html = (
                    "<div class='home-set-wrap'>"
                    "<div class='home-set-title'>🏠 home</div>"
                    f"<div class='home-set-list'>{''.join(home_parts)}</div>"
                    "</div>"
                )

            service_items = [
                {
                    "label": escape(str(item.get("label", "")).strip()),
                    "maintained": escape(_format_last_maintained(str(item.get("last_maintained", "")).strip())),
                }
                for item in list(r.get("ServiceItems", []) or [])
                if str(item.get("label", "")).strip()
            ]
            service_html = ""
            if service_items:
                service_parts = []
                for item in service_items:
                    service_main_html = f"<span class='service-set-main'><span>{item['label']}</span></span>"
                    maintained_html = f"<span class='service-set-date'>{item['maintained']}</span>" if item["maintained"] else ""
                    service_parts.append(
                        f"<span class='service-set-chip'>"
                        f"{service_main_html}"
                        f"{maintained_html}"
                        f"</span>"
                    )
                service_html = (
                    "<div class='service-set-wrap'>"
                    "<div class='service-set-title'>🛠️ maintenance/service</div>"
                    f"<div class='service-set-list'>{''.join(service_parts)}</div>"
                    "</div>"
                )

            left_col = (
                f"<div style='flex:0 0 40%;max-width:40%;padding-right:20px'>"
                f"<div style='display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:8px'>"
                f"<span class='inv-name' style='font-size:20px'>{r['DisplayLabel']}</span>{badge}</div>"
                f"{in_html}{nb_html}{up_html}{home_html}{service_html}</div>"
            )

            out_items = sorted(_set_out.get(cn,[]), key=lambda o: (o.get("DateValue",date.max), str(o.get("Case",""))))
            if out_items:
                lines = []
                for idx, o in enumerate(out_items, 1):
                    assignment = str(o.get("Assignment","")).upper()
                    cid        = str(o.get("Case","")).strip()
                    set_key    = str(o.get("set_key","")).strip()
                    meeple     = _meeple_track_for_case_id(cid, surgery=o.get("Surgery Date",""), delivery=o.get("Delivery Date",""), case_status=o.get("Case Status",""))
                    next_cases = sorted(
                        suggested_case_queue_by_set_key.get(set_key, []),
                        key=lambda item: (item.get("date_value") or date.max, str(item.get("case_id",""))),
                    )
                    next_parts = []
                    for next_item in next_cases[:3]:
                        nx_cls = "out-next-chip" if bool(next_item.get("confirmed", False)) else "out-next-chip is-tentative"
                        next_parts.append(
                            f"<span class='{nx_cls}'>"
                            f"next {escape(str(next_item.get('date','')) or '—')} · "
                            f"{escape(str(next_item.get('hospital','')) or '—')} · "
                            f"{escape(str(next_item.get('case_id','')))}</span>"
                        )
                    if len(next_cases) > 3:
                        next_parts.append(f"<span class='out-next-chip is-tentative'>+{len(next_cases)-3} more</span>")
                    next_html = f"<div class='out-next-wrap'>{''.join(next_parts)}</div>" if next_parts else ""
                    if assignment == "BOOKED":
                        lines.append(
                            f"<div class='out-line'>"
                            f"<span class='out-order'>#{idx}</span>"
                            f"<span class='out-tag out-tag-booked'>BOOKED</span> "
                            f"<span class='out-set'>{o['Set']}</span><span class='out-sep'> → </span>"
                            f"{_hospital_with_led(o['Hospital'], o.get('Delivery Date',''), is_booked=True, sales_code=o.get('Sales Code',''), case_status=o.get('Case Status',''))}"
                            f"{meeple}{next_html}</div>"
                        )
                    else:
                        lines.append(
                            f"<div class='out-line'>"
                            f"<span class='out-order'>#{idx}</span>"
                            f"<span class='out-tag'>OUT</span> "
                            f"<span class='out-set'>{o['Set']}</span><span class='out-sep'> → </span>"
                            f"{_hospital_with_led(o['Hospital'], o['Surgery Date'], sales_code=o.get('Sales Code',''), case_status=o.get('Case Status',''))}"
                            f"{meeple}{next_html}</div>"
                        )
                out_lines = "".join(lines)
            else:
                out_lines = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>all in office</span>"

            return (
                f"<div class='inv-row' style='display:flex;align-items:flex-start'>"
                f"{left_col}<div style='flex:0 0 60%;max-width:60%'>{out_lines}</div></div>"
            )

        if filtered_summary:
            st.markdown("".join(_set_row_html(r) for r in filtered_summary), unsafe_allow_html=True)
        else:
            st.info("No matching sets.")

        # Copy block
        summary_lookup: dict[str,int] = {str(r.get("CategoryNorm","")).upper(): _safe_int(r.get("Available",0)) for r in summary_rows}
        pt_avail_lookup: dict[str,int] = {}
        if not pt_avail_df.empty and "category" in pt_avail_df.columns:
            for _, r in pt_avail_df.iterrows():
                k = str(r.get("category","")).upper().strip()
                if k: pt_avail_lookup[k] = _safe_int(r.get("available",0))

        def _cc(keys, mode="sum"):
            vals = [pt_avail_lookup.get(k.upper()) if k.upper() in _POWERTOOL_CATS else summary_lookup.get(k.upper(),0) for k in keys]
            vals = [v for v in vals if v is not None]
            if not vals: return 0
            return int(min(vals) if mode=="min" else max(vals) if mode=="max" else sum(vals))

        copy_map = [
            ("1.5-2.0",["1.5-2.0"],"sum"),("2.0-2.4",["2.0-2.4"],"sum"),("2.4-2.7",["2.4-2.7"],"sum"),
            ("2.7-4.0",["2.7-4.0"],"sum"),("3.5 - 6.5",["3.5-6.5"],"sum"),("Canna 2.5",["CANNA 2.5"],"sum"),
            ("Canna 3.5",["CANNA 3.5"],"sum"),("Canna 4.0",["CANNA 4.0"],"sum"),("Canna 5.2",["CANNA 5.2"],"sum"),
            ("Std canna 2.4",["STD CANNA 2.4"],"sum"),("Std canna 3.0",["STD CANNA 3.0"],"sum"),
            ("Std canna 4.0",["STD CANNA 4.0"],"sum"),("Std canna 6.5",["STD CANNA 6.5/7.3"],"sum"),
            ("PFN",["PFN"],"sum"),("Reamer set",["REAMER"],"sum"),("ILN Femur",["ILN FEMUR"],"sum"),
            ("ILN Tibia",["ILN TIBIA"],"sum"),("ILN Humerus",["ILN HUMERUS"],"sum"),
            ("ILN Radius & Ulna",["ILN RADIUS ULNA"],"sum"),("TENS",["TENS"],"sum"),
            ("Fibular Nail",["FIBULAR NAIL"],"sum"),("FNS",["FNS"],"sum"),
            ("Foot set",["FOOT SET"],"sum"),("Distal Femoral (RFN)",["RFN"],"sum"),
            ("PFN ll 170-240",["PFN II 170-240"],"sum"),
            ("PFN ll 340-420",["PFN II 340-420 SYSTEM","PFN II 340-420 IMPLANT"],"min"),
            ("Ankle Nail",["ANKLE ARTHRODESIS NAIL"],"sum"),("Coatlmon Cable",["COATLMON CABLE SYSTEM"],"sum"),
            ("ROI",["ROI"],"sum"),
        ]
        copy_lines = ["*Office Sets availability*",""] + [f"{l} - {_cc(k,m)}" for l,k,m in copy_map] + [
            "","Power",f"5503B (normal) - {_cc(['P5503'])}",f"5400 (kwire) - {_cc(['P5400'])}",f"8400 (handpiece) - {_cc(['P8400'])}",
        ]
        st.markdown("##### Copy Block")
        st.code("\n".join(copy_lines), language="text")


# ══════════════════════════════════════════════════════════════════════════════
# PLATES TAB
# ══════════════════════════════════════════════════════════════════════════════
with inv_tabs[1]:
    plate_sum = pd.DataFrame(report.get("plate_uid_summary", []))
    if plate_sum.empty:
        st.info("No plate data.")
    else:
        if "proper_name" not in plate_sum.columns: plate_sum["proper_name"] = plate_sum.get("plate_name","")
        if "screw_sizes" not in plate_sum.columns: plate_sum["screw_sizes"] = ""
        if "status_note" not in plate_sum.columns: plate_sum["status_note"] = ""

        psr = pd.DataFrame(report.get("plate_size_range_availability",[]))
        if not psr.empty and "screw_sizes" in psr.columns and "plate_uid" in psr.columns:
            uid_screw = (
                psr.assign(un=psr["plate_uid"].astype(str).str.upper().str.strip(), ss=psr["screw_sizes"].astype(str).str.strip())
                .groupby("un")["ss"].apply(lambda s: ", ".join(sorted({x for x in s if x}))).to_dict()
            )
            plate_sum["uid_norm"]    = plate_sum["plate_uid"].astype(str).str.upper().str.strip()
            plate_sum["screw_sizes"] = plate_sum.apply(lambda r: r["screw_sizes"] if str(r["screw_sizes"]).strip() else uid_screw.get(str(r["uid_norm"]),""), axis=1)
        else:
            plate_sum["uid_norm"] = plate_sum["plate_uid"].astype(str).str.upper().str.strip()

        plate_sum["order_rank"] = plate_sum["uid_norm"].map(PLATE_UID_RANK).fillna(10_000).astype(int)
        plate_sum = plate_sum.sort_values(["order_rank","uid_norm"])
        if search_query:
            plate_sum = plate_sum[
                plate_sum["plate_uid"].str.contains(search_query,case=False,na=False)
                | plate_sum["proper_name"].str.contains(search_query,case=False,na=False)
                | plate_sum["screw_sizes"].str.contains(search_query,case=False,na=False)
            ]

        _SR_ORDER = ["SHORT","STANDARD","LONG","EXTRA LONG"]
        _REVERSED_LR_UIDS = {"DSC","MSC","DIA","DPLH"}
        _SR_CHIP = {"SHORT":"sc-sht","STANDARD":"sc-std","LONG":"sc-lng","EXTRA LONG":"sc-xl"}
        _SR_DOT  = {"SHORT":"#f9a8d4","STANDARD":"#93c5fd","LONG":"#6ee7b7","EXTRA LONG":"#c4b5fd"}

        def _dsort(v):
            t = str(v or "").strip().upper()
            return (int(t[1:]),t) if t.startswith("D") and t[1:].isdigit() else (10_000,t)

        def _plsort(v, *, rlr=False):
            t = str(v or "").strip().upper().replace(" ","")
            sr, hs = 1, False
            if t.endswith("L"): sr, hs = 0, True
            elif t.endswith("R"): sr, hs = 2, True
            m = re.search(r"(\d+)",t)
            nv = int(m.group(1)) if m else 10_000
            if not hs: nr = nv
            elif rlr: nr = -nv if sr==0 else nv
            else: nr = nv if sr==0 else -nv
            return (sr, nr, t)

        def _hsn(v):
            t = str(v or "").strip().upper().replace(" ","")
            return bool(re.search(r"\d+",t) and t.endswith(("L","R")))

        # Build drawer lookup
        _dl: dict[str,dict[str,dict[str,dict]]] = {}
        _pdd = pd.DataFrame(report.get("plate_drawer_detail",[]))
        if not _pdd.empty:
            for _, drow in _pdd.iterrows():
                uk  = str(drow["plate_uid"])
                dr  = str(drow.get("drawer","?"))
                srk = str(drow.get("size_range","STANDARD"))
                det = drow.get("drawer_size_detail") or []
                if not det:
                    raw = str(drow.get("drawer_sizes",""))
                    det = [{"label":s.strip(),"size_range":srk,"no_stock":False} for s in raw.split(",") if s.strip()]
                ocd = drow.get("drawer_out_case_details", drow.get("out_case_details",[])) or []
                _dl.setdefault(uk,{}).setdefault(dr,{})[srk] = {"sizes":det,"out_case_details":ocd}

        def _plate_row_html(row: pd.Series) -> str:
            uid   = row["plate_uid"]
            badge = avail_badge(int(row["available_units"]), int(row["total_units"]))
            udl   = _dl.get(uid, {})

            # Legend
            all_srs: set[str] = set()
            for sm in udl.values(): all_srs.update(sm.keys())
            sorted_srs = sorted(all_srs, key=lambda x: _SR_ORDER.index(x) if x in _SR_ORDER else 99)
            legend_html = ""
            if len(sorted_srs) > 1:
                _fallback_dot = "#e5e7eb"
                legend_html = "<div class='sr-legend'>" + "".join(
                    f"<span class='sr-legend-item'><span class='sr-legend-dot' style='background:{_SR_DOT.get(sr, _fallback_dot)}'></span>{sr}</span>"
                    for sr in sorted_srs
                ) + "</div>"

            # Drawers
            drawer_blocks = ""
            for drawer in sorted(udl.keys(), key=_dsort):
                sr_map = udl[drawer]

                # Collect unique out cases
                seen: set[str] = set()
                unique_out: list[dict] = []
                for sr, srd in sr_map.items():
                    for cd in (srd.get("out_case_details") or []):
                        k = "|".join([sr,str(cd.get("case_id","")),str(cd.get("hospital","")),
                                      str(cd.get("surgery_date","")),"1" if cd.get("from_stock") else "0",str(cd.get("case_status",""))])
                        if k in seen: continue
                        seen.add(k)
                        unique_out.append({
                            "size_range":sr,"case_id":cd.get("case_id",""),"hospital":cd.get("hospital",""),
                            "surgery_date":cd.get("surgery_date",""),"case_status":cd.get("case_status",""),
                            "from_stock":bool(cd.get("from_stock",False)),
                            "date_value":_parse_ui_date(str(cd.get("surgery_date",""))) or date.max,
                        })

                hdr_cls = "dh-out" if unique_out else ""
                ordered_uo = sorted(unique_out, key=lambda x: (x["date_value"],x.get("case_id","")))

                out_tags = ""
                out_meeples = ""
                meeple_seen_cases: set[str] = set()  # one meeple per case_id
                for cd in ordered_uo:
                    hosp = cd["hospital"] or "—"
                    surg = cd["surgery_date"] or "—"
                    tc   = "dht-stk" if cd["from_stock"] else ""
                    stks = "<span class='dh-out-stock'>[stk]</span>" if cd["from_stock"] else ""
                    out_tags += (
                        f"<span class='dh-out-tag {tc}'>"
                        f"<span class='dh-out-sr'>{escape(str(cd['size_range']))} out</span>"
                        f"<span class='out-sep'>→</span>"
                        f"{_hospital_with_led(hosp, surg, variant='plate', sales_code=case_sales_lookup.get(str(cd.get('case_id','')).strip(),''), case_status=cd.get('case_status',''))}"
                        f"{stks}</span>"
                    )
                    cid = str(cd.get("case_id","")).strip()
                    if cid and cid not in meeple_seen_cases:
                        meeple_seen_cases.add(cid)
                        mp = _meeple_track_for_case_id(cid, surgery=str(cd.get("surgery_date","")), case_status=str(cd.get("case_status","")))
                        out_meeples += (
                            f"<div style='padding:4px 10px 6px 10px;border-top:1px solid #f3f4f6'>"
                            f"<span style='font-size:9px;font-weight:700;color:#9ca3af;letter-spacing:.08em;text-transform:uppercase'>"
                            f"{escape(cid)} · {escape(hosp)}</span>{mp}</div>"
                        )

                # Size chips
                flat: list[dict] = []
                for sr, srd in sr_map.items():
                    sio = bool(srd.get("out_case_details"))
                    for sz in srd["sizes"]:
                        flat.append({"label":sz.get("label",""),"no_stock":bool(sz.get("no_stock",False)),"size_range":sr,"sr_is_out":sio})

                use_sided = bool(flat) and all(_hsn(sz["label"]) for sz in flat)
                rlr = uid in _REVERSED_LR_UIDS
                ordered_sz = sorted(
                    flat,
                    key=lambda sz: (
                        _plsort(sz["label"],rlr=rlr)
                        if use_sided
                        else (_SR_ORDER.index(sz["size_range"]) if sz["size_range"] in _SR_ORDER else 99, *_plsort(sz["label"],rlr=rlr))
                    ),
                )
                chips_html = "".join(
                    f"<span class='sc {'sc-none' if sz['no_stock'] else 'sc-case' if sz['sr_is_out'] else _SR_CHIP.get(sz['size_range'],'sc-std')}'>{sz['label']}</span>"
                    for sz in ordered_sz
                )

                drc_cls = "drc-out" if unique_out else ""
                drawer_blocks += (
                    f"<div class='dr-block'>"
                    f"<div class='dr-hdr {hdr_cls}'><span>{drawer}</span><span class='dh-out-list'>{out_tags}</span></div>"
                    f"{out_meeples}"
                    f"<div class='dr-chips {drc_cls}'>{chips_html}</div>"
                    f"</div>"
                )
            # end for drawer

            return (
                f"<div class='inv-row'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                f"<span class='inv-name'>{row['proper_name']}</span><span>{badge}</span></div>"
                f"<div class='inv-sub'>{uid} &nbsp;·&nbsp; {row['screw_sizes'] or ''}</div>"
                f"<div class='inv-sub'>{row.get('size_ranges','') or ''} &nbsp;·&nbsp; {row.get('status_note','READY') or 'READY'}</div>"
                f"{legend_html}{drawer_blocks}</div>"
            )

        st.markdown("".join(_plate_row_html(r) for _,r in plate_sum.iterrows()), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# POWERTOOLS TAB
# ══════════════════════════════════════════════════════════════════════════════
with inv_tabs[2]:
    pt_avail = pd.DataFrame(report.get("powertool_category_availability",[]))
    pt_del   = pd.DataFrame(report.get("powertool_delivered",[]))
    pt_uid   = pd.DataFrame(report.get("powertool_uid_availability",[]))
    pt_u30   = pd.DataFrame(report.get("powertool_usage_30d",[]))

    if pt_avail.empty and pt_uid.empty:
        st.info("No powertool data.")
    else:
        for col in ("powertool_uid","category","availability","hospital","surgery_date","patient_doctor"):
            if col not in pt_uid.columns: pt_uid[col] = ""
        pt_uid = pt_uid.copy()
        pt_uid["category_norm"]     = pt_uid["category"].astype(str).str.upper().str.strip()
        pt_uid["availability_norm"] = pt_uid["availability"].astype(str).str.upper().str.strip()
        pt_uid["uid_name"]          = pt_uid["powertool_uid"].astype(str).str.strip()
        pt_avail = pt_avail.copy()
        pt_avail["category_norm"]   = pt_avail["category"].astype(str).str.upper().str.strip()

        usage_by_uid: dict[str,int] = {}
        if not pt_u30.empty:
            for _,ur in pt_u30.iterrows():
                k = str(ur.get("powertool_uid","")).strip()
                if k: usage_by_uid[k] = _safe_int(ur.get("usage_30d",0))

        _pt_uid_out: dict[str,list] = {}
        if not pt_del.empty:
            for _,pr in pt_del.iterrows():
                k = str(pr.get("powertool_uid","") or "").strip()
                if k: _pt_uid_out.setdefault(k,[]).append({"uid":k,"hospital":str(pr.get("hospital","") or "").strip(),"surgery":str(pr.get("surgery_date","") or "").strip(),"case_id":str(pr.get("case_id","") or "").strip()})

        av_lk:  dict[str,int] = {}
        us_lk:  dict[str,int] = {}
        ho_lk:  dict[str,int] = {}
        sb_lk:  dict[str,int] = {}
        for _,r in pt_avail.iterrows():
            k = str(r.get("category_norm","")).strip()
            if k:
                av_lk[k] = _safe_int(r.get("available",0))
                us_lk[k] = _safe_int(r.get("usable_total",0))
                ho_lk[k] = _safe_int(r.get("na_hold",0))
                sb_lk[k] = _safe_int(r.get("standby_hold",0))

        ord_cats = ["P5503","P5400","P8400"]
        disc = sorted({*pt_uid["category_norm"].astype(str).tolist(), *pt_avail["category_norm"].astype(str).tolist()})
        ord_cats.extend(c for c in disc if c and c not in ord_cats)

        def _uhtml(uid):
            return f"<span style='display:block;color:#4b5563;font-size:14px;font-weight:700;line-height:1.15;margin-top:2px'>30d use {usage_by_uid.get(uid,0)}</span>"

        def _uchip(uid, hold=False):
            bg,fg,bd = ("#f3f4f6","#6b7280","#d1d5db") if hold else ("#eff6ff","#1d4ed8","#bfdbfe")
            sf = " [hold]" if hold else ""
            return (
                f"<span style='display:inline-flex;flex-direction:column;align-items:flex-start;"
                f"background:{bg};color:{fg};border:1px solid {bd};font-family:\"JetBrains Mono\",monospace;"
                f"border-radius:8px;padding:7px 10px;margin:3px 6px 3px 0;min-width:140px'>"
                f"<span style='font-size:14px;font-weight:700;line-height:1.1'>{uid}{sf}</span>{_uhtml(uid)}</span>"
            )

        pt_sum_rows = []
        for cat in ord_cats:
            cr  = pt_uid[pt_uid["category_norm"]==cat]
            avr = cr[cr["availability_norm"]=="AVAILABLE"]
            hor = cr[cr["availability_norm"]=="NA_HOLD"]
            sbr = cr[cr["availability_norm"]=="STANDBY_HOLD"]
            otr = cr[cr["availability_norm"].str.startswith("OUT")]
            av  = av_lk.get(cat, len(avr))
            ut  = us_lk.get(cat, len(avr)+len(otr))
            ut  = ut if ut > 0 else len(avr)+len(otr)
            if not cat or (ut==0 and len(hor)==0 and len(otr)==0 and len(avr)==0): continue
            out_items = []
            for _,r in otr.sort_values(["surgery_date","uid_name"]).iterrows():
                uid = str(r.get("uid_name","")).strip()
                dets = _pt_uid_out.get(uid,[])
                m = next((d for d in dets if d.get("hospital","")==str(r.get("hospital","")).strip() and d.get("surgery","")==str(r.get("surgery_date","")).strip()), dets[0] if dets else None)
                if m is None: m = {"uid":uid,"hospital":str(r.get("hospital","")).strip(),"surgery":str(r.get("surgery_date","")).strip(),"case_id":str(r.get("case_id","")).strip()}
                out_items.append({"uid":uid,"hospital":m.get("hospital","") or str(r.get("hospital","")).strip(),"surgery":m.get("surgery","") or str(r.get("surgery_date","")).strip(),"case_id":m.get("case_id","") or str(r.get("case_id","")).strip(),"standby":"STANDBY" in str(r.get("availability_norm","")).upper()})
            sb = " ".join([cat," ".join(i["uid"] for i in [{"uid":str(r["uid_name"]).strip()} for _,r in avr.iterrows()]+[{"uid":str(r["uid_name"]).strip()} for _,r in hor.iterrows()])," ".join(f"{i['uid']} {i['hospital']} {i['surgery']}" for i in out_items)]).lower()
            if search_query and search_query.lower() not in sb: continue
            pt_sum_rows.append({"category":cat,"available":av,"usable_total":ut,"na_hold":ho_lk.get(cat,len(hor)),
                "available_units":[{"uid":str(r["uid_name"]).strip(),"hold":False} for _,r in avr.sort_values(["uid_name"]).iterrows()],
                "hold_units":[{"uid":str(r["uid_name"]).strip(),"hold":True} for _,r in hor.sort_values(["uid_name"]).iterrows()],
                "standby_units":[{"uid":str(r["uid_name"]).strip()} for _,r in sbr.sort_values(["uid_name"]).iterrows()],
                "out_items":out_items})

        def _pt_html(row):
            badge = avail_badge(int(row["available"]),int(row["usable_total"]))
            hn = f"<span style='color:#9ca3af;font-size:12px'>[{int(row['na_hold'])} on hold]</span>" if row["na_hold"]>0 else ""
            sn = f"<span style='color:#b45309;font-size:12px'>[{int(sb_lk.get(row['category'],0))} standby]</span>" if sb_lk.get(row['category'],0)>0 else ""
            li = "".join(_uchip(i["uid"],i["hold"]) for i in row["available_units"]+row["hold_units"])
            li += "".join(
                f"<span style='display:inline-flex;flex-direction:column;align-items:flex-start;"
                f"background:#fff7ed;color:#b45309;border:1px solid #fdba74;font-family:\"JetBrains Mono\",monospace;"
                f"border-radius:8px;padding:7px 10px;margin:3px 6px 3px 0;min-width:140px'>"
                f"<span style='font-size:14px;font-weight:700;line-height:1.1'>{i['uid']} [standby]</span>{_uhtml(i['uid'])}</span>"
                for i in row["standby_units"]
            )
            if not li: li = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>none available</span>"
            lc = (
                f"<div style='flex:0 0 39%;padding-right:20px'>"
                f"<div style='display:flex;align-items:center;gap:12px;flex-wrap:wrap'>"
                f"<span class='inv-name' style='font-size:20px'>{row['category']}</span>{badge}{hn}{sn}</div>"
                f"<div style='margin-top:8px'>{li}</div></div>"
            )
            if row["out_items"]:
                ol = "".join(
                    f"<div class='out-line'><span class='out-tag'>OUT</span> "
                    f"<span class='out-set'>{o['uid']}</span><span class='out-sep'> → </span>"
                    f"{_hospital_with_led(o['hospital'],o['surgery'],sales_code=case_sales_lookup.get(str(o.get('case_id','')).strip(),''))}"
                    f"<span class='out-sep'> · </span><span class='out-surg'>surg {o['surgery'] or '—'}</span>"
                    f"<span style='display:inline-block;margin-left:8px;color:#4b5563;font-size:14px;font-weight:700'>30d use {usage_by_uid.get(o['uid'],0)}</span>"
                    + (f"<span style='margin-left:8px;color:#b45309;font-size:12px;font-weight:700'>[standby]</span>" if o.get('standby') else "")
                    + f"</div>"
                    for o in row["out_items"]
                )
            else:
                ol = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>all in office</span>"
            return f"<div class='inv-row' style='display:flex;align-items:flex-start'>{lc}<div style='flex:1'>{ol}</div></div>"

        if pt_sum_rows: st.markdown("".join(_pt_html(r) for r in pt_sum_rows), unsafe_allow_html=True)
        else: st.info("No matching powertools.")

# ══════════════════════════════════════════════════════════════════════════════
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DATA QUALITY
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
total_unknowns = sum(len(report.get("unknown",{}).get(k,[])) for k in ("set_tokens","plate_tokens","powertool_tokens","hospitals_for_routes"))
if total_unknowns > 0:
    with st.expander(f"⚠️ Data quality — {total_unknowns} unrecognised tokens"):
        for tab,key in zip(st.tabs(["Sets","Plates","Powertools","Hospitals"]),["set_tokens","plate_tokens","powertool_tokens","hospitals_for_routes"]):
            with tab:
                df = pd.DataFrame(report.get("unknown",{}).get(key,[]))
                (st.success("None.") if df.empty else st.dataframe(df,use_container_width=True,hide_index=True))

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HOSPITAL DIRECTORY
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if search_query:
    hdf = pd.DataFrame(report.get("hospital_directory",[]))
    if not hdf.empty:
        filt = hdf[hdf["hosp_code"].str.contains(search_query,case=False,na=False)|hdf["name"].str.contains(search_query,case=False,na=False)|hdf["region"].str.contains(search_query,case=False,na=False)]
        if not filt.empty:
            st.markdown("<div class='sec-header'>Hospital Directory — Search Results</div>", unsafe_allow_html=True)
            st.dataframe(filt[["hosp_code","name","region","office_to_hospital_km","tbs_to_hospital_km"]].rename(columns={"hosp_code":"Code","name":"Name","region":"Region","office_to_hospital_km":"Office→Hosp (km)","tbs_to_hospital_km":"TBS→Hosp (km)"}),use_container_width=True,hide_index=True)

# Footer
st.divider()
st.markdown(
    f"<div class='app-footer'>"
    f"Cases: {meta['counts']['cases_rows']} &nbsp;·&nbsp; Archive: {meta['counts']['archive_rows']} &nbsp;·&nbsp; "
    f"Sets: {meta['counts']['master_sets']} &nbsp;·&nbsp; Plates: {meta['counts']['master_plates']} &nbsp;·&nbsp; "
    f"Hospitals: {meta['counts']['master_hospitals']}</div>",
    unsafe_allow_html=True,
)
