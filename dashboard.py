"""
dashboard.py  –  Osteo Ops Commander Dashboard
Run: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

st.set_page_config(
    page_title="Osteo Ops",
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
.out-hosp  { font-weight: 700; color: #111827; font-size: 15px; }
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
APP_BUILD_TAG = "ATC-INOFFICE-v1-2026-03-04"
APP_FILE = str(Path(__file__).resolve())

# Powertool categories — excluded from the Sets tab
_POWERTOOL_CATS = {"P5503", "P5400", "P8400"}

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


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Sidebar
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with st.sidebar:
    st.markdown("### ⚙️ Data Sources")
    master_path = st.text_input(
        "master_data.py path",
        value="master_data.py",
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
@st.cache_data(show_spinner="Loading data…", ttl=300)
def load_report(master: str, cases: str, archive: str) -> dict:
    from ops_engine import build_operations_report
    return build_operations_report(
        master_data_path=master,
        cases_source=cases or None,
        archive_source=archive or None,
    )


if "report" not in st.session_state or refresh:
    with st.spinner("Fetching data…"):
        try:
            st.session_state["report"] = load_report(
                master_path, cases_source, archive_source
            )
            st.session_state["load_error"] = None
        except Exception as exc:
            st.session_state["load_error"] = str(exc)

if st.session_state.get("load_error"):
    st.error(f"❌ Failed to load data: {st.session_state['load_error']}")
    st.stop()

report = st.session_state["report"]
kpis   = report["kpis"]
meta   = report["meta"]
now_kl = datetime.now(KL_TZ)

# Pre-computed frames for ATC tiers
set_status_df = pd.DataFrame(report.get("set_office_status", []))
set_avail_df  = pd.DataFrame(report.get("set_category_availability", []))
tomorrow_df   = pd.DataFrame(report.get("case_buckets", {}).get("to_deliver_tomorrow", []))

if not set_status_df.empty:
    set_status_df = set_status_df.copy()
    for col in (
        "set_status", "location_now", "home", "category", "set_display",
        "case_id", "patient_doctor", "surgery_date", "days_since_surgery",
    ):
        if col not in set_status_df.columns:
            set_status_df[col] = ""
    set_status_df["set_status_norm"] = set_status_df["set_status"].astype(str).str.upper()
    set_status_df["location_norm"] = set_status_df["location_now"].astype(str).str.upper().str.strip()
    set_status_df["is_na"] = set_status_df["set_status_norm"].str.contains("NA", na=False)
    set_status_df["is_out"] = set_status_df["location_norm"].ne("OFFICE") & set_status_df["location_norm"].ne("")
    set_status_df["is_ready"] = (~set_status_df["is_na"]) & (~set_status_df["is_out"]) & set_status_df["location_norm"].eq("OFFICE")
    set_status_df["readiness"] = set_status_df.apply(
        lambda r: "Critical" if r["is_na"] else ("Dirty/Out" if r["is_out"] else ("Ready" if r["is_ready"] else "Unknown")),
        axis=1,
    )
    set_status_df["cue"] = set_status_df["readiness"].map({
        "Critical": "🔴",
        "Dirty/Out": "🟡",
        "Ready": "🟢",
        "Unknown": "⚪",
    }).fillna("⚪")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Header
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
col_title, col_time = st.columns([3, 1])
with col_title:
    st.markdown("## 🦴 Osteo Ops — Commander Dashboard")
with col_time:
    st.markdown(
        f"<div style='text-align:right;color:#6b7280;font-size:13px;padding-top:14px'>"
        f"{now_kl.strftime('%A, %d %b %Y  %H:%M')} KL</div>",
        unsafe_allow_html=True,
    )
st.caption(f"Build: {APP_BUILD_TAG} | Running file: {APP_FILE}")
st.divider()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TIER 1 — PULSE (FLIGHT READINESS)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _kpi(label: str, value: str, style: str = "") -> str:
    cls = f"kpi-value {style}" if style else "kpi-value"
    return (
        f'<div class="kpi-card">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="{cls}">{value}</div>'
        f'</div>'
    )

deployable_fleet = int(set_status_df["is_ready"].sum()) if not set_status_df.empty else 0
dirty_queue = int(set_status_df["is_out"].sum()) if not set_status_df.empty else 0
consignment_mask = set_status_df["home"].astype(str).str.upper().ne("OFFICE") if not set_status_df.empty else pd.Series(dtype=bool)
consignment_total = int(consignment_mask.sum()) if not set_status_df.empty else 0
consignment_ready = int((set_status_df["is_ready"] & consignment_mask).sum()) if not set_status_df.empty else 0
parked_health_pct = (consignment_ready / consignment_total * 100.0) if consignment_total else 0.0
tomorrow_gap = deployable_fleet - (len(tomorrow_df) if not tomorrow_df.empty else 0)

kpi_cols = st.columns(4)
kpi_defs = [
    ("Deployable Fleet", f"{deployable_fleet}", "ok" if deployable_fleet else ""),
    ("Dirty Queue", f"{dirty_queue}", "warn" if dirty_queue else "ok"),
    ("Parked Health", f"{parked_health_pct:.0f}%", "ok" if parked_health_pct >= 70 else "warn"),
    ("Tomorrow's Gap", f"{tomorrow_gap:+d}", "ok" if tomorrow_gap >= 0 else "alert"),
]
for col, (label, val, style) in zip(kpi_cols, kpi_defs):
    col.markdown(_kpi(label, val, style), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TIER 2 — FLEET MAP (WHERE + STATE)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("<div class='sec-header'>Fleet Map</div>", unsafe_allow_html=True)

if set_status_df.empty:
    st.info("No set fleet data.")
else:
    fm1, fm2, fm3 = st.columns(3)
    with fm1:
        set_type_opt = ["All"] + sorted(set_status_df["category"].astype(str).unique().tolist())
        selected_set_type = st.selectbox("Set Type", set_type_opt, index=0)
    with fm2:
        hospital_opt = ["All"] + sorted(set_status_df["location_now"].fillna("").astype(str).replace("", "OFFICE").unique().tolist())
        selected_hospital = st.selectbox("Hospital / Location", hospital_opt, index=0)
    with fm3:
        readiness_opt = ["All", "Ready", "Dirty/Out", "Critical", "Unknown"]
        selected_readiness = st.selectbox("Readiness State", readiness_opt, index=0)

    fleet = set_status_df.copy()
    if selected_set_type != "All":
        fleet = fleet[fleet["category"] == selected_set_type]
    if selected_hospital != "All":
        fleet = fleet[fleet["location_now"].fillna("").replace("", "OFFICE") == selected_hospital]
    if selected_readiness != "All":
        fleet = fleet[fleet["readiness"] == selected_readiness]
    if search_query:
        mask = fleet.apply(lambda r: r.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
        fleet = fleet[mask]

    st.dataframe(
        fleet[[
            "cue", "set_display", "category", "home", "location_now", "readiness",
            "case_id", "patient_doctor", "surgery_date", "days_since_surgery",
        ]].rename(columns={
            "cue": "", "set_display": "Set", "category": "Type", "home": "Home",
            "location_now": "Current", "readiness": "Flight State",
            "case_id": "Case", "patient_doctor": "Patient/Doctor",
            "surgery_date": "Surgery", "days_since_surgery": "Days Out",
        }),
        use_container_width=True,
        hide_index=True,
    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TIER 3 — DEEP DIVE (SET COMPOSITION PROXY)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("<div class='sec-header'>Set Deep Dive</div>", unsafe_allow_html=True)

if set_status_df.empty:
    st.info("No set to inspect.")
else:
    set_choices = set_status_df["set_display"].astype(str).tolist()
    selected_set = st.selectbox("Select Set", set_choices, index=0)
    selected_rows = set_status_df[set_status_df["set_display"] == selected_set]
    srow = selected_rows.iloc[0]

    category = str(srow.get("category", ""))
    availability_row = set_avail_df[set_avail_df["category"] == category] if not set_avail_df.empty else pd.DataFrame()
    available = int(availability_row.iloc[0]["available"]) if not availability_row.empty else 0
    total_office = int(availability_row.iloc[0]["total_office"]) if not availability_row.empty else 0
    score = (available / total_office * 100.0) if total_office else 0.0

    if str(srow.get("readiness", "")) in {"Critical", "Dirty/Out"}:
        recommendation = "Not Recommended for Surgery"
        rec_style = "alert"
    elif score < 60:
        recommendation = "Use With Caution"
        rec_style = "warn"
    else:
        recommendation = "Recommended"
        rec_style = "ok"

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Set", selected_set)
    d2.metric("Current", str(srow.get("location_now", "OFFICE") or "OFFICE"))
    d3.metric("Availability Score", f"{score:.0f}%")
    d4.metric("State", str(srow.get("readiness", "Unknown")))
    st.markdown(_kpi("Recommendation", recommendation, rec_style), unsafe_allow_html=True)

    case_id = str(srow.get("case_id", "") or "")
    all_cases = []
    for _, rows in report.get("case_buckets", {}).items():
        all_cases.extend(rows)
    all_cases_df = pd.DataFrame(all_cases).drop_duplicates(subset=["case_id"], keep="first") if all_cases else pd.DataFrame()
    linked_case = all_cases_df[all_cases_df["case_id"].astype(str) == case_id] if case_id and not all_cases_df.empty else pd.DataFrame()

    if not linked_case.empty:
        crow = linked_case.iloc[0]
        st.markdown(
            f"**Linked Case:** `{case_id}` | Hospital `{crow.get('hospital','')}` | "
            f"Patient/Doctor `{crow.get('patient_doctor','')}`"
        )
        st.code(
            f"Sets: {crow.get('sets','')}\n"
            f"Plates: {crow.get('plates','')}\n"
            f"Powertools: {crow.get('powertools','')}",
            language="text",
        )
    else:
        st.caption("No active linked case details for this set.")


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


st.markdown("<div class='sec-header'>Operational Detail — Inventory Snapshot</div>", unsafe_allow_html=True)
inv_tabs = st.tabs(["🔩 Sets", "🦿 Plates", "⚡ Powertools", "🏢 In Office"])


# ── Sets ──────────────────────────────────────────────────────────────────────
with inv_tabs[0]:
    set_avail      = pd.DataFrame(report["set_category_availability"])
    set_status_all = pd.DataFrame(report["set_office_status"])

    if set_avail.empty:
        st.info("No set data.")
    else:
        # Exclude powertool categories
        set_avail = set_avail[
            ~set_avail["category"].str.upper().isin(_POWERTOOL_CATS)
        ]

        if search_query:
            set_avail = set_avail[
                set_avail["category"].str.contains(search_query, case=False, na=False)
            ]

        # Build lookup: category → list of out-for-case details
        # Key on lowercased category for safe matching
        _set_out: dict[str, list[dict]] = {}
        if not set_status_all.empty:
            for _, sr in set_status_all.iterrows():
                loc = str(sr.get("location_now", "OFFICE")).strip().upper()
                cat_raw = str(sr.get("category", "")).strip()
                if loc != "OFFICE" and cat_raw.upper() not in _POWERTOOL_CATS:
                    _set_out.setdefault(cat_raw, []).append({
                        "set":     str(sr.get("set_display") or sr.get("id") or ""),
                        "hospital":str(sr.get("location_now", "") or "").strip(),
                        "surgery": str(sr.get("surgery_date", "") or "").strip(),
                        "days":    str(sr.get("days_since_surgery", "") or "").strip(),
                        "patient": str(sr.get("patient_doctor", "") or "").strip(),
                    })

        def _set_row_html(row: pd.Series) -> str:
            badge     = avail_badge(int(row["available"]), int(row["total_office"]))
            out_items = _set_out.get(row["category"], [])
            out_lines = ""
            for o in out_items:
                hosp = o["hospital"] or "—"
                surg = o["surgery"]  or "—"
                days = (f"<span class='out-days'> · {o['days']}d out</span>"
                        if o["days"] not in ("", "None", "0") else "")
                out_lines += (
                    f"<div class='out-line'>"
                    f"<span class='out-tag'>OUT</span> "
                    f"<span class='out-set'>{o['set']}</span>"
                    f"<span class='out-sep'> → </span>"
                    f"<span class='out-hosp'>{hosp}</span>"
                    f"<span class='out-sep'> · </span>"
                    f"<span class='out-surg'>surg {surg}</span>"
                    f"{days}"
                    f"</div>"
                )
            return (
                f"<div class='inv-row'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                f"<span class='inv-name'>{row['category']}</span>"
                f"<span>{badge}</span>"
                f"</div>"
                f"{out_lines}"
                f"</div>"
            )

        st.markdown(
            "".join(_set_row_html(r) for _, r in set_avail.iterrows()),
            unsafe_allow_html=True,
        )

        with st.expander("📋 Full set status list"):
            disp = set_status_all[
                ~set_status_all["category"].str.upper().isin(_POWERTOOL_CATS)
            ]
            if search_query:
                disp = disp[
                    disp["category"].str.contains(search_query, case=False, na=False)
                    | disp["location_now"].str.contains(search_query, case=False, na=False)
                ]
            st.dataframe(
                disp[[
                    "set_display", "category", "location_now",
                    "delivery_date", "surgery_date", "days_since_surgery",
                    "patient_doctor", "case_status",
                ]].rename(columns={
                    "set_display": "Set", "category": "Category",
                    "location_now": "Location", "delivery_date": "Delivery",
                    "surgery_date": "Surgery", "days_since_surgery": "Days Out",
                    "patient_doctor": "Patient/Doctor", "case_status": "Status",
                }),
                use_container_width=True, hide_index=True,
            )


# ── Plates ────────────────────────────────────────────────────────────────────
with inv_tabs[1]:
    plate_sum    = pd.DataFrame(report["plate_uid_summary"])
    plate_out_df = pd.DataFrame(report["plate_out_cases"])

    if plate_sum.empty:
        st.info("No plate data.")
    else:
        if search_query:
            plate_sum = plate_sum[
                plate_sum["plate_uid"].str.contains(search_query, case=False, na=False)
                | plate_sum["plate_name"].str.contains(search_query, case=False, na=False)
            ]

        # Build lookup: plate_uid → list of out details
        _plate_out: dict[str, list[dict]] = {}
        if not plate_out_df.empty:
            for _, pr in plate_out_df.iterrows():
                _plate_out.setdefault(str(pr["plate_uid"]), []).append({
                    "size":      str(pr.get("size_range", "") or "").strip(),
                    "hospital":  str(pr.get("hospital", "") or "").strip(),
                    "surgery":   str(pr.get("surgery_date", "") or "").strip(),
                    "from_stock":bool(pr.get("from_stock", False)),
                })

        def _plate_row_html(row: pd.Series) -> str:
            badge     = avail_badge(int(row["available_units"]), int(row["total_units"]))
            out_items = _plate_out.get(row["plate_uid"], [])
            out_lines = ""
            for o in out_items:
                hosp  = o["hospital"] or "—"
                surg  = o["surgery"]  or "—"
                size  = (f"<span style='color:#6b7280;font-size:10px;margin-right:4px'>{o['size']}</span>"
                         if o["size"] else "")
                stock = (" <span class='out-stock'>[stock]</span>" if o["from_stock"] else "")
                out_lines += (
                    f"<div class='out-line'>"
                    f"<span class='out-tag'>OUT</span> "
                    f"{size}"
                    f"<span class='out-sep'>→ </span>"
                    f"<span class='out-hosp'>{hosp}</span>"
                    f"<span class='out-sep'> · </span>"
                    f"<span class='out-surg'>surg {surg}</span>"
                    f"{stock}"
                    f"</div>"
                )
            return (
                f"<div class='inv-row'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                f"<span class='inv-name'>{row['plate_uid']}</span>"
                f"<span>{badge}</span>"
                f"</div>"
                f"<div class='inv-sub'>{row['plate_name']} &nbsp;·&nbsp; {row['size_ranges']}</div>"
                f"{out_lines}"
                f"</div>"
            )

        st.markdown(
            "".join(_plate_row_html(r) for _, r in plate_sum.iterrows()),
            unsafe_allow_html=True,
        )

        with st.expander("📋 Plate size range detail"):
            psr = pd.DataFrame(report["plate_size_range_availability"])
            if search_query:
                psr = psr[
                    psr["plate_uid"].str.contains(search_query, case=False, na=False)
                ]
            st.dataframe(
                psr[[
                    "plate_uid", "plate_name", "size_range", "drawer_locations",
                    "available_units", "out_units", "total_units", "availability",
                ]].rename(columns={
                    "plate_uid": "UID", "plate_name": "Name", "size_range": "Size",
                    "drawer_locations": "Drawers", "available_units": "Avail",
                    "out_units": "Out", "total_units": "Total", "availability": "Avail/Total",
                }),
                use_container_width=True, hide_index=True,
            )


# ── Powertools ────────────────────────────────────────────────────────────────
with inv_tabs[2]:
    pt_avail = pd.DataFrame(report["powertool_category_availability"])
    pt_del   = pd.DataFrame(report["powertool_delivered"])

    if pt_avail.empty:
        st.info("No powertool data.")
    else:
        if search_query:
            pt_avail = pt_avail[
                pt_avail["category"].str.contains(search_query, case=False, na=False)
            ]

        # Build lookup: category → list of out details
        _pt_out: dict[str, list[dict]] = {}
        if not pt_del.empty:
            for _, pr in pt_del.iterrows():
                _pt_out.setdefault(str(pr["category"]), []).append({
                    "uid":     str(pr.get("powertool_uid", "") or "").strip(),
                    "hospital":str(pr.get("hospital", "") or "").strip(),
                    "surgery": str(pr.get("surgery_date", "") or "").strip(),
                    "patient": str(pr.get("patient_doctor", "") or "").strip(),
                })

        def _pt_row_html(row: pd.Series) -> str:
            badge     = avail_badge(int(row["available"]), int(row["usable_total"]))
            out_items = _pt_out.get(row["category"], [])
            na_note   = (
                f" <span style='color:#9ca3af;font-size:11px'>[{row['na_hold']} on hold]</span>"
                if int(row["na_hold"]) > 0 else ""
            )
            out_lines = ""
            for o in out_items:
                hosp = o["hospital"] or "—"
                surg = o["surgery"]  or "—"
                uid  = (f"<span class='out-set'>{o['uid']}</span>"
                        if o["uid"] else "")
                out_lines += (
                    f"<div class='out-line'>"
                    f"<span class='out-tag'>OUT</span> "
                    f"{uid}"
                    f"<span class='out-sep'> → </span>"
                    f"<span class='out-hosp'>{hosp}</span>"
                    f"<span class='out-sep'> · </span>"
                    f"<span class='out-surg'>surg {surg}</span>"
                    f"</div>"
                )
            return (
                f"<div class='inv-row'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                f"<span class='inv-name'>{row['category']}</span>"
                f"<span>{badge}{na_note}</span>"
                f"</div>"
                f"{out_lines}"
                f"</div>"
            )

        st.markdown(
            "".join(_pt_row_html(r) for _, r in pt_avail.iterrows()),
            unsafe_allow_html=True,
        )

        with st.expander("📋 30-day usage"):
            pt_u30 = pd.DataFrame(report["powertool_usage_30d"])
            if not pt_u30.empty:
                used = pt_u30[pt_u30["usage_30d"] > 0].sort_values(
                    "usage_30d", ascending=False
                )
                st.dataframe(
                    used[["powertool_uid", "category", "usage_30d",
                           "window_start", "window_end"]],
                    use_container_width=True, hide_index=True,
                )


# ── In Office (ordered control list) ─────────────────────────────────────────
with inv_tabs[3]:
    office_df = pd.DataFrame(report.get("set_office_status", []))
    if office_df.empty:
        st.info("No set data.")
    else:
        office_df = office_df.copy()
        if "location_now" not in office_df.columns:
            office_df["location_now"] = ""
        if "category" not in office_df.columns:
            office_df["category"] = ""
        if "set_display" not in office_df.columns:
            office_df["set_display"] = office_df.get("id", "")
        if "id" not in office_df.columns:
            office_df["id"] = ""

        office_df["location_norm"] = office_df["location_now"].astype(str).str.upper().str.strip()
        office_df["set_name"] = office_df["set_display"].astype(str).where(
            office_df["set_display"].astype(str).str.strip().ne(""),
            office_df["id"].astype(str),
        )

        categories_in_data = office_df["category"].astype(str).dropna().unique().tolist()
        ordered_categories = OFFICE_VIEW_ORDER + [
            c for c in categories_in_data if c not in OFFICE_VIEW_ORDER
        ]

        rows = []
        for category in ordered_categories:
            cat_rows = office_df[office_df["category"] == category]
            total = int(len(cat_rows))
            office_rows = cat_rows[cat_rows["location_norm"] == "OFFICE"]
            office_count = int(len(office_rows))
            office_list = "; ".join(sorted(office_rows["set_name"].astype(str).tolist()))
            rows.append({
                "Category": category,
                "Total": total,
                "In Office": office_count,
                "In Office List": office_list,
            })

        out = pd.DataFrame(rows)
        if search_query:
            out = out[
                out["Category"].str.contains(search_query, case=False, na=False)
                | out["In Office List"].str.contains(search_query, case=False, na=False)
            ]

        top1, top2 = st.columns(2)
        with top1:
            st.metric("Categories in View", int(len(out)))
        with top2:
            st.metric("Total Sets In Office", int(out["In Office"].sum()) if not out.empty else 0)

        st.dataframe(out, use_container_width=True, hide_index=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 3 — SCHEDULE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("<div class='sec-header'>Schedule</div>", unsafe_allow_html=True)

sched_tabs = st.tabs([
    "✅ Delivered Today",
    "🚚 To Deliver",
    "📅 Tomorrow",
    "↩️ To Collect",
    "🔎 To Follow Up",
    "⚠️ To Check",
    "🔋 To Top Up",
])

_CASE_COLS = [
    "case_id", "hospital", "patient_doctor",
    "delivery_date", "surgery_date", "sales_code",
    "sets", "plates", "powertools", "status", "smart_status",
]
_CASE_RENAME = {
    "case_id": "ID", "hospital": "Hospital", "patient_doctor": "Patient/Doctor",
    "delivery_date": "Delivery", "surgery_date": "Surgery", "sales_code": "Sales Code",
    "sets": "Sets", "plates": "Plates", "powertools": "Powertools",
    "status": "Status", "smart_status": "Smart Status",
}

def _render_bucket(tab_obj, bucket_key: str):
    with tab_obj:
        rows = pd.DataFrame(report["case_buckets"][bucket_key])
        if rows.empty:
            st.success("Nothing here. ✓")
            return
        if search_query:
            mask = rows.apply(
                lambda r: r.astype(str).str.contains(search_query, case=False).any(), axis=1
            )
            rows = rows[mask]
        cols = [c for c in _CASE_COLS if c in rows.columns]
        st.dataframe(
            rows[cols].rename(columns=_CASE_RENAME),
            use_container_width=True, hide_index=True,
        )

_render_bucket(sched_tabs[0], "delivered_today")
_render_bucket(sched_tabs[1], "to_deliver")
_render_bucket(sched_tabs[2], "to_deliver_tomorrow")
_render_bucket(sched_tabs[3], "to_collect")
_render_bucket(sched_tabs[4], "to_follow_up")
_render_bucket(sched_tabs[5], "to_check")
_render_bucket(sched_tabs[6], "to_top_up")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 4 — ROUTES
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("<div class='sec-header'>Routes</div>", unsafe_allow_html=True)
route_tabs = st.tabs(["📅 Tomorrow – Delivery", "✅ Today – Delivered"])

def _render_route(tab_obj, route_key: str):
    with tab_obj:
        rows = pd.DataFrame(report["distance_routes"][route_key])
        if rows.empty:
            st.info("No route data.")
            return
        if search_query:
            mask = rows.apply(
                lambda r: r.astype(str).str.contains(search_query, case=False).any(), axis=1
            )
            rows = rows[mask]
        cols = [
            "hospital", "hospital_name", "delivery_date", "surgery_date",
            "office_est_drive_km", "office_est_drive_min", "sets", "plates", "sales_code",
        ]
        cols = [c for c in cols if c in rows.columns]
        st.dataframe(
            rows[cols].rename(columns={
                "hospital": "Code", "hospital_name": "Hospital",
                "delivery_date": "Delivery", "surgery_date": "Surgery",
                "office_est_drive_km": "~Drive km", "office_est_drive_min": "~Drive min",
                "sets": "Sets", "plates": "Plates", "sales_code": "Sales Code",
            }),
            use_container_width=True, hide_index=True,
        )

_render_route(route_tabs[0], "to_deliver_tomorrow")
_render_route(route_tabs[1], "delivered_today")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 5 — DATA QUALITY
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
# SECTION 6 — HOSPITAL DIRECTORY (search-gated)
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
