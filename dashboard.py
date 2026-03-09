"""
dashboard.py  –  CHECKSETGO
Run: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
import re
from datetime import date, datetime
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

# ── Light theme CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #f8f9fb;
    color: #111827;
}

.stApp { background-color: #f8f9fb; }
section[data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e5e7eb; }

/* KPI tiles */
.kpi-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 14px 18px;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
}
.kpi-label {
    font-size: 10px;
    letter-spacing: .09em;
    color: #6b7280;
    text-transform: uppercase;
    margin-bottom: 4px;
    font-weight: 500;
}
.kpi-value {
    font-size: 26px;
    font-weight: 600;
    color: #111827;
    font-family: 'JetBrains Mono', monospace;
}
.kpi-value.warn  { color: #d97706; }
.kpi-value.alert { color: #dc2626; }
.kpi-value.ok    { color: #16a34a; }

/* Availability badge */
.avail-badge {
    display: inline-block;
    font-family: 'JetBrains Mono', monospace;
    font-size: 22px;
    font-weight: 700;
    padding: 4px 16px;
    border-radius: 8px;
}
.avail-ok   { background: #dcfce7; color: #15803d; }
.avail-low  { background: #fef9c3; color: #a16207; }
.avail-zero { background: #fee2e2; color: #b91c1c; }

/* Inventory rows */
.inv-row {
    padding: 16px 0 16px 0;
    border-bottom: 1px solid #e5e7eb;
}
.inv-name {
    font-family: 'JetBrains Mono', monospace;
    font-size: 28px;
    font-weight: 700;
    color: #111827;
    line-height: 1.2;
}
.inv-sub {
    font-size: 13px;
    color: #6b7280;
    margin-top: 4px;
}
.office-set-ids {
    font-family: 'JetBrains Mono', monospace;
    font-size: 24px;
    font-weight: 700;
    letter-spacing: -.04em;
    line-height: 1.15;
    color: #166534;
}
.office-set-ids.is-standby {
    color: #b45309;
}
.office-set-empty {
    color: #9ca3af;
    font-size: 12px;
    font-style: italic;
}
.out-line {
    font-size: 15px;
    color: #374151;
    padding: 5px 0 5px 0;
    margin-top: 6px;
    line-height: 1.6;
    border-left: 3px solid #e5e7eb;
    padding-left: 14px;
    margin-left: 4px;
}
.out-tag  {
    display: inline-block;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: .06em;
    background: #fef3c7;
    color: #92400e;
    border-radius: 3px;
    padding: 2px 7px;
    margin-right: 8px;
    vertical-align: middle;
}
.out-sep   { color: #9ca3af; }
.out-set   { font-family: 'JetBrains Mono', monospace; color: #1d4ed8; font-weight: 600; font-size: 15px; margin-right: 2px; }
.out-hosp-wrap {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    font-weight: 700;
    font-size: 15px;
    vertical-align: middle;
}
.out-hosp-name { color: #111827; }
.out-hosp-led {
    width: 11px;
    height: 11px;
    border-radius: 999px;
    display: inline-block;
    position: relative;
    flex: 0 0 auto;
}
.out-hosp-wrap.led-retro .out-hosp-led {
    background: #9ca3af;
    border: 1px solid rgba(255,255,255,.85);
    box-shadow:
        inset 0 1px 2px rgba(255,255,255,.65),
        0 0 0 2px rgba(156,163,175,.16),
        0 0 10px rgba(107,114,128,.25);
}
.out-hosp-wrap.led-retro.is-future .out-hosp-led {
    background: #f59e0b;
    box-shadow:
        inset 0 1px 2px rgba(255,255,255,.65),
        0 0 0 2px rgba(245,158,11,.18),
        0 0 12px rgba(245,158,11,.55);
}
.out-hosp-wrap.is-future .out-hosp-name { color: #92400e; }
.out-hosp-wrap.led-retro.is-today .out-hosp-led {
    background: #ef4444;
    box-shadow:
        inset 0 1px 2px rgba(255,255,255,.65),
        0 0 0 2px rgba(239,68,68,.18),
        0 0 12px rgba(239,68,68,.62);
}
.out-hosp-wrap.is-today .out-hosp-name { color: #991b1b; }
.out-hosp-wrap.led-retro.is-past .out-hosp-led {
    background: #22c55e;
    box-shadow:
        inset 0 1px 2px rgba(255,255,255,.65),
        0 0 0 2px rgba(34,197,94,.18),
        0 0 12px rgba(34,197,94,.52);
}
.out-hosp-wrap.led-retro.is-sales .out-hosp-led {
    background: #2563eb;
    box-shadow:
        inset 0 1px 2px rgba(255,255,255,.65),
        0 0 0 2px rgba(37,99,235,.18),
        0 0 12px rgba(37,99,235,.55);
}
.out-hosp-wrap.led-lens .out-hosp-led {
    width: 12px;
    height: 12px;
    border: 1px solid #cbd5e1;
    background: radial-gradient(circle at 32% 28%, rgba(255,255,255,.95) 0 18%, rgba(255,255,255,.45) 19% 30%, #cbd5e1 31% 100%);
    box-shadow:
        inset 0 1px 1px rgba(255,255,255,.88),
        inset 0 -2px 3px rgba(15,23,42,.24),
        0 1px 2px rgba(15,23,42,.18),
        0 0 0 1px rgba(255,255,255,.4);
}
.out-hosp-wrap.led-lens .out-hosp-led::after {
    content: "";
    position: absolute;
    top: 1px;
    left: 2px;
    width: 5px;
    height: 3px;
    border-radius: 999px;
    background: rgba(255,255,255,.55);
    transform: rotate(-18deg);
}
.out-hosp-wrap.led-lens.is-future .out-hosp-led {
    border-color: #d97706;
    background: radial-gradient(circle at 32% 28%, #fff8e1 0 18%, #fde68a 19% 30%, #f59e0b 31% 64%, #b45309 100%);
    box-shadow:
        inset 0 1px 1px rgba(255,255,255,.88),
        inset 0 -2px 3px rgba(120,53,15,.28),
        0 1px 2px rgba(120,53,15,.18),
        0 0 10px rgba(245,158,11,.35);
}
.out-hosp-wrap.led-lens.is-today .out-hosp-led {
    border-color: #b91c1c;
    background: radial-gradient(circle at 32% 28%, #ffe4e6 0 18%, #fda4af 19% 30%, #ef4444 31% 64%, #991b1b 100%);
    box-shadow:
        inset 0 1px 1px rgba(255,255,255,.88),
        inset 0 -2px 3px rgba(127,29,29,.3),
        0 1px 2px rgba(127,29,29,.18),
        0 0 10px rgba(239,68,68,.38);
}
.out-hosp-wrap.led-lens.is-past .out-hosp-led {
    border-color: #15803d;
    background: radial-gradient(circle at 32% 28%, #dcfce7 0 18%, #86efac 19% 30%, #22c55e 31% 64%, #166534 100%);
    box-shadow:
        inset 0 1px 1px rgba(255,255,255,.88),
        inset 0 -2px 3px rgba(20,83,45,.28),
        0 1px 2px rgba(20,83,45,.18),
        0 0 10px rgba(34,197,94,.32);
}
.out-hosp-wrap.led-lens.is-sales .out-hosp-led {
    border-color: #1d4ed8;
    background: radial-gradient(circle at 32% 28%, #dbeafe 0 18%, #93c5fd 19% 30%, #2563eb 31% 64%, #1e3a8a 100%);
    box-shadow:
        inset 0 1px 1px rgba(255,255,255,.88),
        inset 0 -2px 3px rgba(30,58,138,.28),
        0 1px 2px rgba(30,58,138,.18),
        0 0 10px rgba(37,99,235,.36);
}
.out-hosp-wrap.led-panel .out-hosp-led {
    width: 14px;
    height: 14px;
    border: 1px solid #475569;
    background:
        radial-gradient(circle at 36% 32%, rgba(255,255,255,.95) 0 12%, rgba(255,255,255,.28) 13% 18%, transparent 19% 100%),
        radial-gradient(circle at 50% 52%, #64748b 0 52%, #1f2937 72%, #0f172a 100%);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.35),
        inset 0 -2px 3px rgba(0,0,0,.45),
        0 0 0 2px #cbd5e1,
        0 1px 2px rgba(15,23,42,.35);
}
.out-hosp-wrap.led-panel .out-hosp-led::before {
    content: "";
    position: absolute;
    inset: 2px;
    border-radius: 999px;
    background: radial-gradient(circle at 34% 30%, rgba(255,255,255,.88) 0 16%, rgba(255,255,255,.15) 17% 26%, transparent 27% 100%);
    opacity: .95;
}
.out-hosp-wrap.led-panel .out-hosp-led::after {
    content: "";
    position: absolute;
    inset: -4px;
    border-radius: 999px;
    opacity: 0;
    transition: opacity .15s ease;
}
.out-hosp-wrap.led-panel.is-future .out-hosp-led {
    border-color: #78350f;
    background:
        radial-gradient(circle at 36% 32%, rgba(255,255,255,.95) 0 12%, rgba(255,248,220,.35) 13% 18%, transparent 19% 100%),
        radial-gradient(circle at 50% 52%, #fbbf24 0 52%, #d97706 72%, #78350f 100%);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.4),
        inset 0 -2px 3px rgba(120,53,15,.5),
        0 0 0 2px #cbd5e1,
        0 0 14px rgba(245,158,11,.35);
}
.out-hosp-wrap.led-panel.is-future .out-hosp-led::after {
    opacity: 1;
    background: radial-gradient(circle, rgba(251,191,36,.42) 0, rgba(245,158,11,.16) 45%, rgba(245,158,11,0) 72%);
}
.out-hosp-wrap.led-panel.is-today .out-hosp-led {
    border-color: #7f1d1d;
    background:
        radial-gradient(circle at 36% 32%, rgba(255,255,255,.95) 0 12%, rgba(255,230,230,.35) 13% 18%, transparent 19% 100%),
        radial-gradient(circle at 50% 52%, #fb7185 0 52%, #dc2626 72%, #7f1d1d 100%);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.4),
        inset 0 -2px 3px rgba(127,29,29,.55),
        0 0 0 2px #cbd5e1,
        0 0 14px rgba(220,38,38,.42);
}
.out-hosp-wrap.led-panel.is-today .out-hosp-led::after {
    opacity: 1;
    background: radial-gradient(circle, rgba(251,113,133,.42) 0, rgba(220,38,38,.16) 45%, rgba(220,38,38,0) 72%);
}
.out-hosp-wrap.led-panel.is-past .out-hosp-led {
    border-color: #14532d;
    background:
        radial-gradient(circle at 36% 32%, rgba(255,255,255,.95) 0 12%, rgba(220,252,231,.35) 13% 18%, transparent 19% 100%),
        radial-gradient(circle at 50% 52%, #4ade80 0 52%, #16a34a 72%, #14532d 100%);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.4),
        inset 0 -2px 3px rgba(20,83,45,.55),
        0 0 0 2px #cbd5e1,
        0 0 14px rgba(34,197,94,.34);
}
.out-hosp-wrap.led-panel.is-past .out-hosp-led::after {
    opacity: 1;
    background: radial-gradient(circle, rgba(74,222,128,.36) 0, rgba(34,197,94,.14) 45%, rgba(34,197,94,0) 72%);
}
.out-hosp-wrap.led-panel.is-sales .out-hosp-led {
    border-color: #1e3a8a;
    background:
        radial-gradient(circle at 36% 32%, rgba(255,255,255,.95) 0 12%, rgba(219,234,254,.35) 13% 18%, transparent 19% 100%),
        radial-gradient(circle at 50% 52%, #60a5fa 0 52%, #2563eb 72%, #1e3a8a 100%);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.4),
        inset 0 -2px 3px rgba(30,58,138,.52),
        0 0 0 2px #cbd5e1,
        0 0 14px rgba(37,99,235,.40);
}
.out-hosp-wrap.led-panel.is-sales .out-hosp-led::after {
    opacity: 1;
    background: radial-gradient(circle, rgba(96,165,250,.40) 0, rgba(37,99,235,.16) 45%, rgba(37,99,235,0) 72%);
}
.out-hosp-wrap.is-past .out-hosp-name { color: #166534; }
.out-hosp-wrap.is-sales .out-hosp-name { color: #1d4ed8; }
.out-hosp-wrap.is-plate {
    gap: 9px;
}
.out-hosp-wrap.is-plate .out-hosp-led {
    width: 15px;
    height: 15px;
}
.out-hosp-wrap.is-plate .out-hosp-name {
    font-size: 18px;
    line-height: 1;
}
.out-surg  { color: #4b5563; font-size: 14px; }
.out-days  { color: #9ca3af; font-size: 12px; }
.out-stock { color: #9ca3af; font-style: italic; font-size: 13px; }

/* Section header */
.sec-header {
    font-size: 10px;
    letter-spacing: .12em;
    text-transform: uppercase;
    color: #9ca3af;
    font-weight: 600;
    border-bottom: 1px solid #e5e7eb;
    padding-bottom: 6px;
    margin: 24px 0 14px 0;
}
</style>
""", unsafe_allow_html=True)

KL_TZ = ZoneInfo("Asia/Kuala_Lumpur")
APP_BUILD_TAG = "CHECKSETGO-v4-2026-03-05"
APP_FILE = str(Path(__file__).resolve())
DEFAULT_MASTER_DATA_PATH = str(Path(__file__).resolve().with_name("master_data.py"))

# Powertool categories — excluded from the Sets tab
_POWERTOOL_CATS = {"P5503", "P5400", "P8400"}
HOSPITAL_LED_STYLE = "panel"

# Priority order for "In Office" control-tower view
OFFICE_VIEW_ORDER = [
    "STD CANNA 6.5/7.3",
    "STD CANNA 4.0",
    "STD CANNA 2.4",
    "STD CANNA 3.0",
    "RFN",
    "PFN II 170-240",
    "PFN II 340-420 SYSTEM",
    "PFN II 340-420 IMPLANT",
    "ANKLE ARTHRODESIS NAIL",
    "COATLMON CABLE SYSTEM",
    "FOOT SET",
    "FOOT INSTRUMENT",
    "PFN II NAIL REMOVAL",
    "P5503",
    "P5400",
    "P8400",
    "FIBULAR NAIL",
    "FNS",
    "ROI",
    "2.7-4.0",
    "3.5-6.5",
    "PFN",
    "LONG PFN",
    "REAMER",
    "ILN TIBIA SUPSUB",
    "ILN HUMERUS",
    "2.4-2.7",
    "1.5-2.0",
    "ILN RADIUS ULNA",
    "2.0-2.4",
    "2.0 ONLY",
    "TENS",
    "CANNA 2.5",
    "CANNA 3.5",
    "CANNA 4.0",
    "CANNA 5.2",
    "ILN FEMUR",
    "ILN TIBIA",
]

PLATE_UID_ORDER = [
    "PHILOS",
    "OLEI",
    "OLEII",
    "DSC",
    "MSC",
    "CHOOK",
    "DIA",
    "URS",
    "RECON",
    "METAI",
    "METAII",
    "TUBULAR",
    "DFIBII",
    "DFIBIII",
    "DMH",
    "DPLH",
    "DLHI",
    "DLHII",
    "DMT",
    "DLT",
    "CMESH",
    "CCOMBO",
    "PPMT",
    "DPT",
    "PLTI",
    "PLTII",
    "PMT",
    "DLF",
    "TSP",
    "FSP",
    "PFP",
    "APP",
    "FIBHOOK",
    "DLTII",
    "ADT",
    "DRVL",
]
PLATE_UID_RANK = {uid: idx for idx, uid in enumerate(PLATE_UID_ORDER)}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Sidebar
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with st.sidebar:
    st.markdown("### ⚙️ Data Sources")
    master_path = st.text_input(
        "master_data.py path",
        value=DEFAULT_MASTER_DATA_PATH,
        help="Absolute or relative path to master_data.py",
    )
    cases_source = st.text_input(
        "Cases CSV (path or URL)",
        value=(
            "https://docs.google.com/spreadsheets/d/e/"
            "2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IY"
            "nirCfsnGHoMyn5xPoG5c/pub?gid=0&single=true&output=csv"
        ),
    )
    archive_source = st.text_input(
        "Archive CSV (path or URL)",
        value=(
            "https://docs.google.com/spreadsheets/d/e/"
            "2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IY"
            "nirCfsnGHoMyn5xPoG5c/pub?gid=1320419668&single=true&output=csv"
        ),
    )
    refresh = st.button("🔄 Refresh Data", use_container_width=True)
    st.divider()
    search_query = st.text_input("🔍 Search", placeholder="sets, plates, hospitals…")
    st.divider()
    st.caption("TZ: Asia/Kuala_Lumpur")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Load / cache report
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
@st.cache_data(show_spinner="Loading data…", ttl=60)
def load_report(master: str, cases: str, archive: str, source_signature: str) -> dict:
    from ops_engine import build_operations_report
    _ = source_signature  # cache key only
    return build_operations_report(
        master_data_path=master,
        cases_source=cases or None,
        archive_source=archive or None,
    )


def _source_signature(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    path = Path(text).expanduser()
    if path.exists():
        stat = path.stat()
        return f"{path.resolve()}:{stat.st_mtime_ns}:{stat.st_size}"
    return text


if refresh:
    load_report.clear()

report_signature = "|".join([
    _source_signature(master_path),
    _source_signature(cases_source),
    _source_signature(archive_source),
])

if (
    "report" not in st.session_state
    or refresh
    or st.session_state.get("report_signature") != report_signature
):
    with st.spinner("Fetching data…"):
        try:
            st.session_state["report"] = load_report(
                master_path, cases_source, archive_source, report_signature
            )
            st.session_state["report_signature"] = report_signature
            st.session_state["load_error"] = None
        except Exception as exc:
            st.session_state["load_error"] = str(exc)

if st.session_state.get("load_error"):
    st.error(f"❌ Failed to load data: {st.session_state['load_error']}")
    st.stop()

report = st.session_state["report"]
meta   = report["meta"]
cases_all_df = pd.DataFrame(report.get("cases_all", []))
case_sales_lookup: dict[str, str] = {}
if not cases_all_df.empty and "case_id" in cases_all_df.columns:
    if "sales_code" not in cases_all_df.columns:
        cases_all_df["sales_code"] = ""
    case_sales_lookup = {
        str(r.get("case_id", "")).strip(): str(r.get("sales_code", "") or "").strip()
        for _, r in cases_all_df.iterrows()
        if str(r.get("case_id", "")).strip()
    }
now_kl = datetime.now(KL_TZ)
try:
    report_today = datetime.strptime(str(meta.get("today_kl", "")).strip(), "%Y-%m-%d").date()
except ValueError:
    report_today = now_kl.date()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Header
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
col_title, col_time = st.columns([3, 1])
with col_title:
    st.markdown("## CHECKSETGO")
with col_time:
    st.markdown(
        f"<div style='text-align:right;color:#6b7280;font-size:13px;padding-top:14px'>"
        f"{now_kl.strftime('%A, %d %b %Y  %H:%M')} KL</div>",
        unsafe_allow_html=True,
    )
st.divider()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 4 — INVENTORY SNAPSHOT (DETAIL)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def avail_badge(available: int, total: int) -> str:
    if total == 0:
        return "<span class='avail-badge avail-ok'>—</span>"
    if available == 0:
        return f"<span class='avail-badge avail-zero'>{available}/{total}</span>"
    if available / total <= 0.25:
        return f"<span class='avail-badge avail-low'>{available}/{total}</span>"
    return f"<span class='avail-badge avail-ok'>{available}/{total}</span>"


def _parse_ui_date(value: str) -> Optional[date]:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _hospital_status_class(value: str, sales_code: str = "") -> str:
    if str(sales_code or "").strip():
        return "is-sales"
    surgery_date = _parse_ui_date(value)
    if surgery_date is None:
        return ""
    if surgery_date > report_today:
        return "is-future"
    if surgery_date == report_today:
        return "is-today"
    return "is-past"


def _hospital_with_led(
    hospital: str,
    surgery_value: str,
    *,
    variant: str = "",
    sales_code: str = "",
) -> str:
    hosp_text = escape(str(hospital or "—"))
    status_class = _hospital_status_class(surgery_value, sales_code=sales_code)
    led_style = HOSPITAL_LED_STYLE if HOSPITAL_LED_STYLE in {"retro", "lens", "panel"} else "panel"
    variant_class = f"is-{variant}" if variant else ""
    class_attr = " ".join(
        part for part in ("out-hosp-wrap", f"led-{led_style}", variant_class, status_class) if part
    )
    return (
        f"<span class='{class_attr}'>"
        f"<span class='out-hosp-led'></span>"
        f"<span class='out-hosp-name'>{hosp_text}</span>"
        f"</span>"
    )


def _compact_set_id(value: str, fallback: str = "") -> str:
    text = str(value or "").strip()
    if not text:
        text = str(fallback or "").strip()
    if not text:
        return ""
    if re.fullmatch(r"\d+", text):
        return str(int(text))
    return text


st.markdown("<div class='sec-header'>Operational Detail — Inventory Snapshot</div>", unsafe_allow_html=True)
inv_tabs = st.tabs(["🔩 Sets", "🦿 Plates", "⚡ Powertools"])


# ── Sets ──────────────────────────────────────────────────────────────────────
with inv_tabs[0]:
    set_avail = pd.DataFrame(report.get("set_category_availability", []))
    set_status_all = pd.DataFrame(report.get("set_office_status", []))

    if set_avail.empty and set_status_all.empty:
        st.info("No set data.")
    else:
        def _safe_int(val) -> int:
            num = pd.to_numeric(val, errors="coerce")
            return int(num) if pd.notna(num) else 0

        for col in ("category", "set_display", "id", "location_now", "surgery_date", "patient_doctor", "case_id", "set_status", "home"):
            if col not in set_status_all.columns:
                set_status_all[col] = ""
        set_status_all = set_status_all.copy()
        set_status_all["category_norm"] = set_status_all["category"].astype(str).str.upper().str.strip()
        set_status_all["location_norm"] = set_status_all["location_now"].astype(str).str.upper().str.strip()
        set_status_all["home_norm"] = set_status_all["home"].astype(str).str.upper().str.strip()
        set_status_all["set_status_norm"] = set_status_all["set_status"].astype(str).str.upper().str.strip()
        set_status_all["is_na"] = set_status_all["set_status_norm"].str.contains("NA", na=False)
        set_status_all["is_standby"] = (
            set_status_all["home_norm"].eq("STANDBY")
            | set_status_all["set_status_norm"].str.contains("STANDBY", na=False)
        )
        set_status_all["set_name"] = set_status_all["set_display"].astype(str).where(
            set_status_all["set_display"].astype(str).str.strip().ne(""),
            set_status_all["id"].astype(str),
        )
        set_status_all = set_status_all[~set_status_all["category_norm"].isin(_POWERTOOL_CATS)]

        if "category" not in set_avail.columns:
            set_avail["category"] = ""
        set_avail = set_avail.copy()
        set_avail["category_norm"] = set_avail["category"].astype(str).str.upper().str.strip()
        set_avail = set_avail[~set_avail["category_norm"].isin(_POWERTOOL_CATS)]

        avail_lookup: dict[str, int] = {}
        total_lookup: dict[str, int] = {}
        for _, r in set_avail.iterrows():
            key = str(r.get("category_norm", "")).strip()
            if key:
                avail_lookup[key] = _safe_int(r.get("available", 0))
                total_lookup[key] = _safe_int(r.get("total_office", 0))

        display_by_norm: dict[str, str] = {c.upper(): c for c in OFFICE_VIEW_ORDER}
        for src_df in (set_status_all, set_avail):
            for cat in src_df.get("category", pd.Series(dtype=str)).astype(str).tolist():
                cat_norm = str(cat).upper().strip()
                if cat_norm and cat_norm not in display_by_norm:
                    display_by_norm[cat_norm] = str(cat).strip()

        ordered_norm = [c.upper() for c in OFFICE_VIEW_ORDER]
        ordered_norm.extend(sorted([n for n in display_by_norm if n not in ordered_norm]))
        ordered_core = {c.upper() for c in OFFICE_VIEW_ORDER}

        summary_rows = []
        out_rows = []
        for cat_norm in ordered_norm:
            label = display_by_norm.get(cat_norm, cat_norm)
            cat_rows = set_status_all[set_status_all["category_norm"] == cat_norm]
            office_rows = cat_rows[
                (cat_rows["location_norm"] == "OFFICE")
                & (~cat_rows["is_na"])
                & (~cat_rows["is_standby"])
            ]
            standby_rows = cat_rows[
                (
                    cat_rows["location_norm"].eq("STANDBY")
                    | (
                        cat_rows["location_norm"].eq("OFFICE")
                        & cat_rows["is_standby"]
                    )
                )
                & (~cat_rows["is_na"])
            ]
            out_case_rows = cat_rows[
                (~cat_rows["location_norm"].isin(["OFFICE", "STANDBY"]))
                & (cat_rows["location_norm"] != "")
                & (~cat_rows["is_na"])
            ]

            in_count = int(len(office_rows))
            standby_count = int(len(standby_rows))
            out_count = int(len(out_case_rows))
            available = avail_lookup.get(cat_norm, in_count)
            total = total_lookup.get(cat_norm, int(len(office_rows) + len(out_case_rows)))
            total = total if total > 0 else (in_count + out_count)

            if cat_norm not in ordered_core and total == 0 and available == 0 and out_count == 0 and standby_count == 0:
                continue

            office_items = [
                {
                    "name": str(r["set_name"]),
                    "label": _compact_set_id(r.get("id", ""), r.get("set_name", "")),
                    "standby": False,
                }
                for _, r in office_rows.iterrows()
            ]
            standby_items = [
                {
                    "name": str(r["set_name"]),
                    "label": _compact_set_id(r.get("id", ""), r.get("set_name", "")),
                    "standby": True,
                }
                for _, r in standby_rows.iterrows()
            ]
            in_list = "; ".join(sorted([item["name"] for item in office_items + standby_items]))
            out_list = "; ".join(
                sorted([
                    f"{str(r['set_name'])} @ {str(r['location_now']).strip() or 'OUT'} ({str(r['surgery_date']).strip() or '-'})"
                    for _, r in out_case_rows.iterrows()
                ])
            )

            summary_rows.append({
                "Category": label,
                "CategoryNorm": cat_norm,
                "In Office": in_count,
                "Out": out_count,
                "Available": available,
                "Available/Total": f"{available}/{total}",
                "OfficeItems": office_items,
                "StandbyItems": standby_items,
                "In Office Sets": in_list,
                "Out (Hospital • Surgery)": out_list,
            })

            for _, r in out_case_rows.iterrows():
                out_rows.append({
                    "Category": label,
                    "Set": str(r.get("set_name", "")),
                    "Hospital": str(r.get("location_now", "")),
                    "Surgery Date": str(r.get("surgery_date", "")),
                    "Case": str(r.get("case_id", "")),
                    "Sales Code": case_sales_lookup.get(str(r.get("case_id", "")).strip(), ""),
                })

        # Build out-details lookup: category_norm → list of dicts
        _set_out: dict[str, list[dict]] = {}
        for r in out_rows:
            key = r["Category"].upper().strip()
            _set_out.setdefault(key, []).append(r)

        # Filter summary_rows by search
        filtered_summary = summary_rows
        if search_query:
            sq = search_query.lower()
            filtered_summary = [
                r for r in summary_rows
                if sq in r["Category"].lower()
                or sq in r["In Office Sets"].lower()
                or sq in r["Out (Hospital • Surgery)"].lower()
            ]

        def _set_row_html(r: dict) -> str:
            cat_norm = r["CategoryNorm"]
            avail_val = int(r.get("Available", 0))
            total_val = int(total_lookup.get(cat_norm, r["In Office"] + r["Out"]))
            total_val = total_val if total_val > 0 else (r["In Office"] + r["Out"])
            badge = avail_badge(avail_val, total_val)

            # Left: availability badge + prominent in-office set names
            in_sets_html = ""
            office_items = list(r.get("OfficeItems", []))
            standby_items = list(r.get("StandbyItems", []))
            if office_items or standby_items:
                office_labels = [escape(str(item.get("label", ""))) for item in office_items if str(item.get("label", "")).strip()]
                standby_labels = [escape(str(item.get("label", ""))) for item in standby_items if str(item.get("label", "")).strip()]
                if office_labels:
                    in_sets_html += f"<div class='office-set-ids'>{', '.join(office_labels)}</div>"
                if standby_labels:
                    in_sets_html += f"<div class='office-set-ids is-standby'>{', '.join(standby_labels)} [standby]</div>"
                if not in_sets_html:
                    in_sets_html = "<span class='office-set-empty'>none available</span>"
            else:
                in_sets_html = "<span class='office-set-empty'>none available</span>"

            left_col = (
                f"<div style='flex:0 0 39%;padding-right:20px'>"
                f"<div style='display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:8px'>"
                f"<span class='inv-name' style='font-size:20px'>{r['Category']}</span>"
                f"{badge}"
                f"</div>"
                f"{in_sets_html}"
                f"</div>"
            )

            # Right: OUT lines
            out_items = _set_out.get(cat_norm, [])
            if out_items:
                out_lines = "".join(
                    (
                        f"<div class='out-line'>"
                        f"<span class='out-tag'>OUT</span> "
                        f"<span class='out-set'>{o['Set']}</span>"
                        f"<span class='out-sep'> → </span>"
                        f"{_hospital_with_led(o['Hospital'], o['Surgery Date'], sales_code=o.get('Sales Code', ''))}"
                        f"<span class='out-sep'> · </span>"
                        f"<span class='out-surg'>surg {o['Surgery Date'] or '—'}</span>"
                        f"</div>"
                    )
                    for o in out_items
                )
            else:
                out_lines = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>all in office</span>"

            right_col = f"<div style='flex:1'>{out_lines}</div>"

            return (
                f"<div class='inv-row' style='display:flex;align-items:flex-start'>"
                f"{left_col}{right_col}"
                f"</div>"
            )

        if filtered_summary:
            st.markdown(
                "".join(_set_row_html(r) for r in filtered_summary),
                unsafe_allow_html=True,
            )
        else:
            st.info("No matching sets.")

        summary_lookup: dict[str, int] = {}
        for row in summary_rows:
            row_norm = str(row.get("CategoryNorm", "")).strip().upper()
            if row_norm:
                summary_lookup[row_norm] = _safe_int(row.get("Available", 0))

        def _copy_count(category_keys: list[str], mode: str = "sum") -> int:
            values: list[int] = []
            for key in category_keys:
                key_norm = str(key).upper().strip()
                values.append(summary_lookup.get(key_norm, 0))
            if not values:
                return 0
            if mode == "min":
                return int(min(values))
            if mode == "max":
                return int(max(values))
            return int(sum(values))

        copy_lines = ["*Office Sets availability*", ""]
        copy_map: list[tuple[str, list[str], str]] = [
            ("Comus mini 1.5 (1.5-2.0)", ["1.5-2.0"], "sum"),
            ("Comus mini 2.0 (2.0-2.4)", ["2.0-2.4"], "sum"),
            ("2.4 (2.4-2.7)", ["2.4-2.7"], "sum"),
            ("2.7 (2.7-4.0)", ["2.7-4.0"], "sum"),
            ("3.5 (3.5-6.5)", ["3.5-6.5"], "sum"),
            ("Canna 2.5", ["CANNA 2.5"], "sum"),
            ("Canna 3.5", ["CANNA 3.5"], "sum"),
            ("Canna 4.0", ["CANNA 4.0"], "sum"),
            ("Canna 5.2", ["CANNA 5.2"], "sum"),
            ("Std canna 2.4", ["STD CANNA 2.4"], "sum"),
            ("Std canna 3.0", ["STD CANNA 3.0"], "sum"),
            ("Std canna 4.0", ["STD CANNA 4.0"], "sum"),
            ("Std canna 6.5 (STD CANNA 6.5/7.3)", ["STD CANNA 6.5/7.3"], "sum"),
            ("PFN", ["PFN"], "sum"),
            ("Reamer set", ["REAMER"], "sum"),
            ("ILN Femur", ["ILN FEMUR"], "sum"),
            ("ILN Tibia", ["ILN TIBIA"], "sum"),
            ("ILN Humerus", ["ILN HUMERUS"], "sum"),
            ("ILN Radius & Ulna", ["ILN RADIUS ULNA"], "sum"),
            ("TENS", ["TENS"], "sum"),
            ("Fibular Nail", ["FIBULAR NAIL"], "sum"),
            ("FNS", ["FNS"], "sum"),
            ("Foot set", ["FOOT SET"], "sum"),
            ("Distal Femoral (RFN)", ["RFN"], "sum"),
            ("PFN ll 170-240", ["PFN II 170-240"], "sum"),
            ("PFN ll 340-420", ["PFN II 340-420 SYSTEM", "PFN II 340-420 IMPLANT"], "min"),
            ("Ankle Nail", ["ANKLE ARTHRODESIS NAIL"], "sum"),
            ("Coatlmon Cable", ["COATLMON CABLE SYSTEM"], "sum"),
            ("ROI", ["ROI"], "sum"),
        ]
        for label, keys, mode in copy_map:
            copy_lines.append(f"{label} - {_copy_count(keys, mode=mode)}")

        st.markdown("##### Copy Block")
        st.code("\n".join(copy_lines), language="text")


# ── Plates ────────────────────────────────────────────────────────────────────
with inv_tabs[1]:
    plate_sum    = pd.DataFrame(report["plate_uid_summary"])

    if plate_sum.empty:
        st.info("No plate data.")
    else:
        if "proper_name" not in plate_sum.columns:
            plate_sum["proper_name"] = plate_sum.get("plate_name", "")
        if "screw_sizes" not in plate_sum.columns:
            plate_sum["screw_sizes"] = ""
        if "status_note" not in plate_sum.columns:
            plate_sum["status_note"] = ""
        psr_fill = pd.DataFrame(report.get("plate_size_range_availability", []))
        if not psr_fill.empty and "screw_sizes" in psr_fill.columns and "plate_uid" in psr_fill.columns:
            uid_screw_map = (
                psr_fill.assign(
                    uid_norm=psr_fill["plate_uid"].astype(str).str.upper().str.strip(),
                    screw_sizes=psr_fill["screw_sizes"].astype(str).str.strip(),
                )
                .groupby("uid_norm")["screw_sizes"]
                .apply(lambda s: ", ".join(sorted({x for x in s if x})))
                .to_dict()
            )
            plate_sum["uid_norm"] = plate_sum["plate_uid"].astype(str).str.upper().str.strip()
            plate_sum["screw_sizes"] = plate_sum.apply(
                lambda r: r["screw_sizes"] if str(r["screw_sizes"]).strip() else uid_screw_map.get(str(r["uid_norm"]), ""),
                axis=1,
            )
        else:
            plate_sum["uid_norm"] = plate_sum["plate_uid"].astype(str).str.upper().str.strip()
        plate_sum["order_rank"] = plate_sum["uid_norm"].map(PLATE_UID_RANK).fillna(10_000).astype(int)
        plate_sum = plate_sum.sort_values(["order_rank", "uid_norm"])

        if search_query:
            plate_sum = plate_sum[
                plate_sum["plate_uid"].str.contains(search_query, case=False, na=False)
                | plate_sum["proper_name"].str.contains(search_query, case=False, na=False)
                | plate_sum["screw_sizes"].str.contains(search_query, case=False, na=False)
            ]

        # ── CSS ───────────────────────────────────────────────────────────────
        st.markdown("""
<style>
/* Drawer block */
.dr-block {
    margin-top: 8px;
    border-radius: 8px;
    overflow: hidden;
    border: 1.5px solid #e5e7eb;
}
/* Drawer header */
.dr-hdr {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 8px;
    flex-wrap: wrap;
    padding: 4px 10px;
    background: #f3f4f6;
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px;
    font-weight: 700;
    color: #374151;
    letter-spacing: .06em;
}
.dr-hdr.dh-out { background:#fff1f2; color:#be123c; }
.dh-out-list {
    display: flex;
    flex: 1;
    flex-wrap: wrap;
    justify-content: flex-end;
    gap: 6px;
}
/* OUT → hospital tag in header */
.dh-out-tag {
    display: inline-flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 6px;
    font-size: 11px;
    font-weight: 700;
    background: #fecdd3;
    color: #9f1239;
    border-radius: 6px;
    padding: 4px 8px;
}
.dh-out-tag.dht-stk { background:#fde68a; color:#92400e; }
.dh-out-sr {
    font-size: 10px;
    letter-spacing: .08em;
    text-transform: uppercase;
}
.dh-out-surg {
    color: #7f1d1d;
    font-size: 13px;
    font-weight: 700;
}
.dh-out-tag.dht-stk .dh-out-surg { color:#92400e; }
.dh-out-stock {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .05em;
    text-transform: uppercase;
}
/* Chip row */
.dr-chips {
    display: flex; flex-wrap: wrap; gap: 5px;
    padding: 7px 10px;
    background: #ffffff;
    border-top: 1px solid #f3f4f6;
}
.dr-chips.drc-out { background: #fffbf0; }
/* Size chip */
.sc {
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px; font-weight: 600;
    padding: 3px 9px; border-radius: 20px;
    white-space: nowrap; border: 1.5px solid transparent;
}
/* size_range colours — in stock */
.sc-std  { background:#dbeafe; color:#1e40af; border-color:#93c5fd; }
.sc-sht  { background:#fce7f3; color:#9d174d; border-color:#f9a8d4; }
.sc-lng  { background:#d1fae5; color:#065f46; border-color:#6ee7b7; }
.sc-xl   { background:#ede9fe; color:#4c1d95; border-color:#c4b5fd; }
/* out for a case — amber override */
.sc-case { background:#fef3c7; color:#92400e; border-color:#fcd34d; }
/* no stock — red strikethrough */
.sc-none { background:#fee2e2; color:#9f1239; border-color:#fca5a5;
           text-decoration: line-through; opacity:.8; }
/* legend row */
.sr-legend {
    display: flex; flex-wrap: wrap; gap: 8px;
    margin-top: 6px; margin-bottom: 2px;
}
.sr-legend-item {
    display: flex; align-items: center; gap: 4px;
    font-size: 10px; color: #6b7280;
}
.sr-legend-dot {
    width: 10px; height: 10px; border-radius: 50%;
    flex-shrink: 0;
}
</style>
""", unsafe_allow_html=True)

        _SR_ORDER  = ["SHORT", "STANDARD", "LONG", "EXTRA LONG"]
        _REVERSED_LR_UIDS = {"DSC", "MSC", "DIA", "DPLH"}
        # chip class per size_range when in stock
        _SR_CHIP   = {"SHORT": "sc-sht", "STANDARD": "sc-std",
                      "LONG": "sc-lng", "EXTRA LONG": "sc-xl"}
        # legend colour dot
        _SR_DOT    = {"SHORT": "#f9a8d4", "STANDARD": "#93c5fd",
                      "LONG": "#6ee7b7", "EXTRA LONG": "#c4b5fd"}

        def _drawer_sort_key(value: str) -> tuple[int, str]:
            text = str(value or "").strip().upper()
            if text.startswith("D") and text[1:].isdigit():
                return (int(text[1:]), text)
            return (10_000, text)

        def _plate_label_sort_key(value: str, *, reverse_lr: bool = False) -> tuple[int, int, str]:
            text = str(value or "").strip().upper()
            token = text.replace(" ", "")
            side_rank = 1
            numeric_rank = 10_000
            has_side = False
            if token.endswith("L"):
                side_rank = 0
                has_side = True
            elif token.endswith("R"):
                side_rank = 2
                has_side = True
            match = re.search(r"(\d+)", token)
            if match:
                numeric_value = int(match.group(1))
                if not has_side:
                    numeric_rank = numeric_value
                elif reverse_lr:
                    numeric_rank = -numeric_value if side_rank == 0 else numeric_value
                else:
                    numeric_rank = numeric_value if side_rank == 0 else -numeric_value
            return (side_rank, numeric_rank, text)

        def _has_sided_numeric_label(value: str) -> bool:
            text = str(value or "").strip().upper().replace(" ", "")
            return bool(re.search(r"\d+", text) and text.endswith(("L", "R")))

        # ── Build drawer-first lookup from plate_drawer_detail ────────────────
        # uid → { drawer → { size_range → {sizes:[{label,no_stock}], out_case_details} } }
        _drawer_lookup: dict[str, dict[str, dict[str, dict]]] = {}
        _pdd_raw = pd.DataFrame(report.get("plate_drawer_detail", []))
        if not _pdd_raw.empty:
            for _, drow in _pdd_raw.iterrows():
                uid_key = str(drow["plate_uid"])
                drawer  = str(drow.get("drawer", "?"))
                sr_key  = str(drow.get("size_range", "STANDARD"))
                # drawer_size_detail is [{label, size_range, no_stock}]
                detail  = drow.get("drawer_size_detail") or []
                if not detail:
                    raw = str(drow.get("drawer_sizes", ""))
                    detail = [{"label": s.strip(), "size_range": sr_key, "no_stock": False}
                              for s in raw.split(",") if s.strip()]
                out_cds = drow.get("drawer_out_case_details", drow.get("out_case_details", [])) or []
                (
                    _drawer_lookup
                    .setdefault(uid_key, {})
                    .setdefault(drawer, {})
                )[sr_key] = {"sizes": detail, "out_case_details": out_cds}

        def _plate_row_html(row: pd.Series) -> str:
            uid   = row["plate_uid"]
            badge = avail_badge(int(row["available_units"]), int(row["total_units"]))

            uid_drawers = _drawer_lookup.get(uid, {})

            # ── Legend: which size ranges exist for this plate ────────────────
            all_srs_present: set[str] = set()
            for sr_map in uid_drawers.values():
                all_srs_present.update(sr_map.keys())
            sorted_present = sorted(all_srs_present,
                key=lambda x: _SR_ORDER.index(x) if x in _SR_ORDER else 99)
            legend_html = ""
            if len(sorted_present) > 1:
                _fallback_dot = "#e5e7eb"
                items = "".join(
                    "<span class='sr-legend-item'>"
                    "<span class='sr-legend-dot' style='background:"
                    + _SR_DOT.get(sr, _fallback_dot)
                    + "'></span>"
                    + sr + "</span>"
                    for sr in sorted_present
                )
                legend_html = f"<div class='sr-legend'>{items}</div>"

            # ── Drawer blocks ─────────────────────────────────────────────────
            drawer_blocks = ""
            for drawer in sorted(uid_drawers.keys(), key=_drawer_sort_key):
                sr_map = uid_drawers[drawer]  # {size_range: {sizes, out_case_details}}

                seen: set[str] = set()
                unique_out: list[dict] = []
                for sr, sr_data in sr_map.items():
                    for cd in (sr_data.get("out_case_details") or []):
                        key = "|".join([
                            sr,
                            str(cd.get("case_id", "")),
                            str(cd.get("hospital", "")),
                            str(cd.get("surgery_date", "")),
                            "1" if cd.get("from_stock") else "0",
                        ])
                        if key in seen:
                            continue
                        seen.add(key)
                        unique_out.append({
                            "size_range": sr,
                            "case_id": cd.get("case_id", ""),
                            "hospital": cd.get("hospital", ""),
                            "surgery_date": cd.get("surgery_date", ""),
                            "from_stock": bool(cd.get("from_stock", False)),
                        })

                hdr_cls = "dh-out" if unique_out else ""
                out_tags = ""
                for cd in unique_out:
                    hosp = cd["hospital"] or "—"
                    surg = cd["surgery_date"] or "—"
                    tc = "dht-stk" if cd["from_stock"] else ""
                    stk_s = "<span class='dh-out-stock'>[stk]</span>" if cd["from_stock"] else ""
                    out_tags += (
                        f"<span class='dh-out-tag {tc}'>"
                        f"<span class='dh-out-sr'>{escape(str(cd['size_range']))} out</span>"
                        f"<span class='out-sep'>→</span>"
                        f"{_hospital_with_led(hosp, surg, variant='plate', sales_code=case_sales_lookup.get(str(cd.get('case_id', '')).strip(), ''))}"
                        f"<span class='out-sep'>·</span>"
                        f"<span class='dh-out-surg'>surg {escape(str(surg))}</span>"
                        f"{stk_s}"
                        f"</span>"
                    )

                # All chips merged into one row, colour = size_range (or override)
                chips_html = ""
                flat_sizes: list[dict] = []
                for sr, sr_data in sr_map.items():
                    sr_is_out = bool(sr_data.get("out_case_details"))
                    for sz in sr_data["sizes"]:
                        flat_sizes.append({
                            "label": sz.get("label", ""),
                            "no_stock": bool(sz.get("no_stock", False)),
                            "size_range": sr,
                            "sr_is_out": sr_is_out,
                        })

                use_sided_numeric_order = bool(flat_sizes) and all(
                    _has_sided_numeric_label(sz["label"]) for sz in flat_sizes
                )
                reverse_lr = uid in _REVERSED_LR_UIDS
                ordered_sizes = sorted(
                    flat_sizes,
                    key=lambda sz: (
                        _plate_label_sort_key(sz["label"], reverse_lr=reverse_lr)
                        if use_sided_numeric_order
                        else (
                            _SR_ORDER.index(sz["size_range"]) if sz["size_range"] in _SR_ORDER else 99,
                            *_plate_label_sort_key(sz["label"], reverse_lr=reverse_lr),
                        )
                    ),
                )

                for sz in ordered_sizes:
                    lbl      = sz["label"]
                    no_stock = sz["no_stock"]
                    if no_stock:
                        chip_cls = "sc-none"
                    elif sz["sr_is_out"]:
                        chip_cls = "sc-case"
                    else:
                        chip_cls = _SR_CHIP.get(sz["size_range"], "sc-std")
                    chips_html += f"<span class='sc {chip_cls}'>{lbl}</span>"

                drc_cls = "drc-out" if unique_out else ""
                drawer_blocks += (
                    f"<div class='dr-block'>"
                    f"<div class='dr-hdr {hdr_cls}'>"
                    f"<span>{drawer}</span><span class='dh-out-list'>{out_tags}</span>"
                    f"</div>"
                    f"<div class='dr-chips {drc_cls}'>{chips_html}</div>"
                    f"</div>"
                )

            return (
                f"<div class='inv-row'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                f"<span class='inv-name'>{row['proper_name']}</span>"
                f"<span>{badge}</span>"
                f"</div>"
                f"<div class='inv-sub'>{uid} &nbsp;·&nbsp; {row['screw_sizes'] or ''}</div>"
                f"<div class='inv-sub'>{row.get('size_ranges', '')} &nbsp;·&nbsp; {row.get('status_note', 'READY')}</div>"
                f"{legend_html}"
                f"{drawer_blocks}"
                f"</div>"
            )

        st.markdown(
            "".join(_plate_row_html(r) for _, r in plate_sum.iterrows()),
            unsafe_allow_html=True,
        )


# ── Powertools ────────────────────────────────────────────────────────────────
with inv_tabs[2]:
    pt_avail = pd.DataFrame(report["powertool_category_availability"])
    pt_del   = pd.DataFrame(report["powertool_delivered"])
    pt_uid   = pd.DataFrame(report.get("powertool_uid_availability", []))
    pt_u30   = pd.DataFrame(report.get("powertool_usage_30d", []))

    if pt_avail.empty and pt_uid.empty:
        st.info("No powertool data.")
    else:
        def _safe_int(val) -> int:
            num = pd.to_numeric(val, errors="coerce")
            return int(num) if pd.notna(num) else 0

        for col in ("powertool_uid", "category", "availability", "hospital", "surgery_date", "patient_doctor"):
            if col not in pt_uid.columns:
                pt_uid[col] = ""
        pt_uid = pt_uid.copy()
        pt_uid["category_norm"] = pt_uid["category"].astype(str).str.upper().str.strip()
        pt_uid["availability_norm"] = pt_uid["availability"].astype(str).str.upper().str.strip()
        pt_uid["uid_name"] = pt_uid["powertool_uid"].astype(str).str.strip()

        pt_avail = pt_avail.copy()
        pt_avail["category_norm"] = pt_avail["category"].astype(str).str.upper().str.strip()

        usage_by_uid: dict[str, int] = {}
        if not pt_u30.empty:
            for _, ur in pt_u30.iterrows():
                key = str(ur.get("powertool_uid", "")).strip()
                if not key:
                    continue
                usage_by_uid[key] = _safe_int(ur.get("usage_30d", 0))

        # uid -> out details
        _pt_uid_out: dict[str, list[dict]] = {}
        if not pt_del.empty:
            for _, pr in pt_del.iterrows():
                key = str(pr.get("powertool_uid", "")).strip()
                if not key:
                    continue
                _pt_uid_out.setdefault(key, []).append({
                    "uid": str(pr.get("powertool_uid", "") or "").strip(),
                    "hospital": str(pr.get("hospital", "") or "").strip(),
                    "surgery": str(pr.get("surgery_date", "") or "").strip(),
                    "case_id": str(pr.get("case_id", "") or "").strip(),
                })

        avail_lookup: dict[str, int] = {}
        usable_lookup: dict[str, int] = {}
        hold_lookup: dict[str, int] = {}
        standby_lookup: dict[str, int] = {}
        for _, r in pt_avail.iterrows():
            key = str(r.get("category_norm", "")).strip()
            if key:
                avail_lookup[key] = _safe_int(r.get("available", 0))
                usable_lookup[key] = _safe_int(r.get("usable_total", 0))
                hold_lookup[key] = _safe_int(r.get("na_hold", 0))
                standby_lookup[key] = _safe_int(r.get("standby_hold", 0))

        ordered_cats = ["P5503", "P5400", "P8400"]
        discovered = sorted({
            *pt_uid["category_norm"].astype(str).tolist(),
            *pt_avail["category_norm"].astype(str).tolist(),
        })
        ordered_cats.extend([c for c in discovered if c and c not in ordered_cats])

        def _usage_html(uid: str) -> str:
            use_30d = usage_by_uid.get(uid, 0)
            return (
                f"<span style='display:block;color:#4b5563;font-size:14px;"
                f"font-weight:700;line-height:1.15;margin-top:2px'>30d use {use_30d}</span>"
            )

        def _unit_chip(uid: str, hold: bool = False) -> str:
            bg = "#f3f4f6" if hold else "#eff6ff"
            fg = "#6b7280" if hold else "#1d4ed8"
            border = "#d1d5db" if hold else "#bfdbfe"
            suffix = " [hold]" if hold else ""
            return (
                f"<span style='display:inline-flex;flex-direction:column;align-items:flex-start;"
                f"background:{bg};color:{fg};border:1px solid {border};"
                f"font-family:\"JetBrains Mono\",monospace;border-radius:8px;"
                f"padding:7px 10px;margin:3px 6px 3px 0;min-width:140px'>"
                f"<span style='font-size:14px;font-weight:700;line-height:1.1'>{uid}{suffix}</span>"
                f"{_usage_html(uid)}"
                f"</span>"
            )

        summary_rows = []
        for cat_norm in ordered_cats:
            cat_rows = pt_uid[pt_uid["category_norm"] == cat_norm]
            avail_rows = cat_rows[cat_rows["availability_norm"] == "AVAILABLE"]
            hold_rows = cat_rows[cat_rows["availability_norm"] == "NA_HOLD"]
            standby_rows = cat_rows[cat_rows["availability_norm"] == "STANDBY_HOLD"]
            out_rows = cat_rows[cat_rows["availability_norm"].str.startswith("OUT")]

            available = avail_lookup.get(cat_norm, int(len(avail_rows)))
            usable_total = usable_lookup.get(cat_norm, int(len(avail_rows) + len(out_rows)))
            usable_total = usable_total if usable_total > 0 else int(len(avail_rows) + len(out_rows))
            if not cat_norm or (usable_total == 0 and len(hold_rows) == 0 and len(out_rows) == 0 and len(avail_rows) == 0):
                continue

            available_units = [
                {"uid": str(r["uid_name"]).strip(), "hold": False}
                for _, r in avail_rows.sort_values(["uid_name"]).iterrows()
            ]
            hold_units = [
                {"uid": str(r["uid_name"]).strip(), "hold": True}
                for _, r in hold_rows.sort_values(["uid_name"]).iterrows()
            ]
            standby_units = [
                {"uid": str(r["uid_name"]).strip(), "standby": True}
                for _, r in standby_rows.sort_values(["uid_name"]).iterrows()
            ]

            out_items = []
            for _, r in out_rows.sort_values(["surgery_date", "uid_name"]).iterrows():
                uid = str(r.get("uid_name", "")).strip()
                details = _pt_uid_out.get(uid, [])
                matched = None
                for d in details:
                    if (
                        str(d.get("hospital", "")).strip() == str(r.get("hospital", "")).strip()
                        and str(d.get("surgery", "")).strip() == str(r.get("surgery_date", "")).strip()
                    ):
                        matched = d
                        break
                if matched is None:
                    matched = details[0] if details else {
                        "uid": uid,
                        "hospital": str(r.get("hospital", "")).strip(),
                        "surgery": str(r.get("surgery_date", "")).strip(),
                        "case_id": str(r.get("case_id", "")).strip(),
                    }
                out_items.append({
                    "uid": uid,
                    "hospital": matched.get("hospital", "") or str(r.get("hospital", "")).strip(),
                    "surgery": matched.get("surgery", "") or str(r.get("surgery_date", "")).strip(),
                    "case_id": matched.get("case_id", "") or str(r.get("case_id", "")).strip(),
                    "standby": "STANDBY" in str(r.get("availability_norm", "")).upper(),
                })

            search_blob = " ".join(
                [
                    cat_norm,
                    " ".join(item["uid"] for item in available_units),
                    " ".join(item["uid"] for item in hold_units),
                    " ".join(item["uid"] for item in standby_units),
                    " ".join(
                        f"{item['uid']} {item['hospital']} {item['surgery']}"
                        for item in out_items
                    ),
                ]
            ).lower()
            if search_query and search_query.lower() not in search_blob:
                continue

            summary_rows.append({
                "category": cat_norm,
                "available": available,
                "usable_total": usable_total,
                "na_hold": hold_lookup.get(cat_norm, int(len(hold_rows))),
                "available_units": available_units,
                "hold_units": hold_units,
                "standby_units": standby_units,
                "out_items": out_items,
            })

        def _pt_group_html(row: dict) -> str:
            badge = avail_badge(int(row["available"]), int(row["usable_total"]))
            hold_note = (
                f"<span style='color:#9ca3af;font-size:12px'>[{int(row['na_hold'])} on hold]</span>"
                if int(row["na_hold"]) > 0 else ""
            )
            standby_note = (
                f"<span style='color:#b45309;font-size:12px'>[{int(standby_lookup.get(row['category'], 0))} standby]</span>"
                if int(standby_lookup.get(row["category"], 0)) > 0 else ""
            )
            left_items = "".join(_unit_chip(item["uid"], hold=item["hold"]) for item in row["available_units"] + row["hold_units"])
            left_items += "".join(
                f"<span style='display:inline-flex;flex-direction:column;align-items:flex-start;"
                f"background:#fff7ed;color:#b45309;border:1px solid #fdba74;"
                f"font-family:\"JetBrains Mono\",monospace;border-radius:8px;"
                f"padding:7px 10px;margin:3px 6px 3px 0;min-width:140px'>"
                f"<span style='font-size:14px;font-weight:700;line-height:1.1'>{item['uid']} [standby]</span>"
                f"{_usage_html(item['uid'])}"
                f"</span>"
                for item in row["standby_units"]
            )
            if not left_items:
                left_items = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>none available</span>"

            left_col = (
                f"<div style='flex:0 0 39%;padding-right:20px'>"
                f"<div style='display:flex;align-items:center;gap:12px;flex-wrap:wrap'>"
                f"<span class='inv-name' style='font-size:20px'>{row['category']}</span>"
                f"{badge}{hold_note}{standby_note}"
                f"</div>"
                f"<div style='margin-top:8px'>{left_items}</div>"
                f"</div>"
            )

            if row["out_items"]:
                out_lines = "".join(
                    (
                        f"<div class='out-line'>"
                        f"<span class='out-tag'>OUT</span> "
                        f"<span class='out-set'>{o['uid']}</span>"
                        f"<span class='out-sep'> → </span>"
                        f"{_hospital_with_led(o['hospital'], o['surgery'], sales_code=case_sales_lookup.get(str(o.get('case_id', '')).strip(), ''))}"
                        f"<span class='out-sep'> · </span>"
                        f"<span class='out-surg'>surg {o['surgery'] or '—'}</span>"
                        f"<span style='display:inline-block;margin-left:8px;color:#4b5563;font-size:14px;font-weight:700'>"
                        f"30d use {usage_by_uid.get(o['uid'], 0)}</span>"
                        + ("<span style='margin-left:8px;color:#b45309;font-size:12px;font-weight:700'>[standby]</span>" if o.get('standby') else "")
                        + f"</div>"
                    )
                    for o in row["out_items"]
                )
            else:
                out_lines = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>all in office</span>"

            right_col = f"<div style='flex:1'>{out_lines}</div>"
            return (
                f"<div class='inv-row' style='display:flex;align-items:flex-start'>"
                f"{left_col}{right_col}"
                f"</div>"
            )

        if summary_rows:
            st.markdown(
                "".join(_pt_group_html(row) for row in summary_rows),
                unsafe_allow_html=True,
            )
        else:
            st.info("No matching powertools.")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 3 — DATA QUALITY
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
total_unknowns = sum(
    len(report["unknown"][k])
    for k in ("set_tokens", "plate_tokens", "powertool_tokens", "hospitals_for_routes")
)
if total_unknowns > 0:
    with st.expander(f"⚠️ Data quality — {total_unknowns} unrecognised tokens"):
        unk_tabs = st.tabs(["Sets", "Plates", "Powertools", "Hospitals"])
        for tab, key in zip(
            unk_tabs,
            ["set_tokens", "plate_tokens", "powertool_tokens", "hospitals_for_routes"],
        ):
            with tab:
                df = pd.DataFrame(report["unknown"][key])
                if df.empty:
                    st.success("None.")
                else:
                    st.dataframe(df, use_container_width=True, hide_index=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 4 — HOSPITAL DIRECTORY (search-gated)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if search_query:
    hosp_df  = pd.DataFrame(report["hospital_directory"])
    filtered = hosp_df[
        hosp_df["hosp_code"].str.contains(search_query, case=False, na=False)
        | hosp_df["name"].str.contains(search_query, case=False, na=False)
        | hosp_df["region"].str.contains(search_query, case=False, na=False)
    ]
    if not filtered.empty:
        st.markdown(
            "<div class='sec-header'>Hospital Directory — Search Results</div>",
            unsafe_allow_html=True,
        )
        st.dataframe(
            filtered[[
                "hosp_code", "name", "region",
                "office_to_hospital_km", "tbs_to_hospital_km",
            ]].rename(columns={
                "hosp_code": "Code", "name": "Name", "region": "Region",
                "office_to_hospital_km": "Office→Hosp (km)",
                "tbs_to_hospital_km": "TBS→Hosp (km)",
            }),
            use_container_width=True, hide_index=True,
        )

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Footer
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.divider()
st.markdown(
    f"<div style='font-size:11px;color:#9ca3af;text-align:center'>"
    f"Cases: {meta['counts']['cases_rows']} &nbsp;·&nbsp; "
    f"Archive: {meta['counts']['archive_rows']} &nbsp;·&nbsp; "
    f"Sets: {meta['counts']['master_sets']} &nbsp;·&nbsp; "
    f"Plates: {meta['counts']['master_plates']} &nbsp;·&nbsp; "
    f"Hospitals: {meta['counts']['master_hospitals']}"
    f"</div>",
    unsafe_allow_html=True,
)
