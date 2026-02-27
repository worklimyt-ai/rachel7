"""
dashboard.py  –  Osteo Ops Commander Dashboard
Run: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
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

# Powertool categories — excluded from the Sets tab
_POWERTOOL_CATS = {"P5503", "P5400", "P8400"}


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
st.divider()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 1 — KPI BAR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _kpi(label: str, value: int, style: str = "") -> str:
    cls = f"kpi-value {style}" if style else "kpi-value"
    return (
        f'<div class="kpi-card">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="{cls}">{value}</div>'
        f'</div>'
    )

kpi_cols = st.columns(7)
kpi_defs = [
    ("Delivered Today",   kpis["delivered_today"],    "ok"    if kpis["delivered_today"]     else ""),
    ("To Deliver",        kpis["to_deliver"],          "warn"  if kpis["to_deliver"]          else "ok"),
    ("Tomorrow",          kpis["to_deliver_tomorrow"], "warn"  if kpis["to_deliver_tomorrow"] else ""),
    ("To Collect",        kpis["to_collect"],          "warn"  if kpis["to_collect"]          else "ok"),
    ("To Follow Up",      kpis["to_follow_up"],        "alert" if kpis["to_follow_up"]        else "ok"),
    ("To Check (ITO)",    kpis["to_check"],            "alert" if kpis["to_check"]            else "ok"),
    ("To Top Up",         kpis["to_top_up"],           "warn"  if kpis["to_top_up"]           else "ok"),
]
for col, (label, val, style) in zip(kpi_cols, kpi_defs):
    col.markdown(_kpi(label, val, style), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SECTION 2 — INVENTORY SNAPSHOT
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def avail_badge(available: int, total: int) -> str:
    if total == 0:
        return "<span class='avail-badge avail-ok'>—</span>"
    if available == 0:
        return f"<span class='avail-badge avail-zero'>{available}/{total}</span>"
    if available / total <= 0.25:
        return f"<span class='avail-badge avail-low'>{available}/{total}</span>"
    return f"<span class='avail-badge avail-ok'>{available}/{total}</span>"


st.markdown("<div class='sec-header'>Inventory Snapshot</div>", unsafe_allow_html=True)
inv_tabs = st.tabs(["🔩 Sets", "🦿 Plates", "⚡ Powertools"])


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
