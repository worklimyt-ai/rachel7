"""
dashboard.py  –  CHECKSETGO
Run: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
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
APP_BUILD_TAG = "CHECKSETGO-v4-2026-03-05"
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
@st.cache_data(show_spinner="Loading data…", ttl=60)
def load_report(master: str, cases: str, archive: str) -> dict:
    from ops_engine import build_operations_report
    return build_operations_report(
        master_data_path=master,
        cases_source=cases or None,
        archive_source=archive or None,
    )


if refresh:
    load_report.clear()

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
meta   = report["meta"]
now_kl = datetime.now(KL_TZ)


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


st.markdown("<div class='sec-header'>Operational Detail — Inventory Snapshot</div>", unsafe_allow_html=True)
inv_tabs = st.tabs(["🔩 Sets", "🦿 Plates", "⚡ Powertools"])


# ── Sets ──────────────────────────────────────────────────────────────────────
with inv_tabs[0]:
    set_avail = pd.DataFrame(report.get("set_category_availability", []))
    set_status_all = pd.DataFrame(report.get("set_office_status", []))
    pt_avail_df = pd.DataFrame(report.get("powertool_category_availability", []))

    if set_avail.empty and set_status_all.empty:
        st.info("No set data.")
    else:
        def _safe_int(val) -> int:
            num = pd.to_numeric(val, errors="coerce")
            return int(num) if pd.notna(num) else 0

        for col in ("category", "set_display", "id", "location_now", "surgery_date", "patient_doctor", "case_id", "set_status"):
            if col not in set_status_all.columns:
                set_status_all[col] = ""
        set_status_all = set_status_all.copy()
        set_status_all["category_norm"] = set_status_all["category"].astype(str).str.upper().str.strip()
        set_status_all["location_norm"] = set_status_all["location_now"].astype(str).str.upper().str.strip()
        set_status_all["set_status_norm"] = set_status_all["set_status"].astype(str).str.upper().str.strip()
        set_status_all["is_na"] = set_status_all["set_status_norm"].str.contains("NA", na=False)
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
            in_rows = cat_rows[(cat_rows["location_norm"] == "OFFICE") & (~cat_rows["is_na"])]
            out_case_rows = cat_rows[
                (cat_rows["location_norm"] != "OFFICE")
                & (cat_rows["location_norm"] != "")
                & (~cat_rows["is_na"])
            ]

            in_count = int(len(in_rows))
            out_count = int(len(out_case_rows))
            available = avail_lookup.get(cat_norm, in_count)
            total = total_lookup.get(cat_norm, int(len(cat_rows)))
            total = total if total > 0 else (in_count + out_count)

            if cat_norm not in ordered_core and total == 0 and available == 0 and out_count == 0:
                continue

            in_list = "; ".join(sorted(in_rows["set_name"].astype(str).tolist()))
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
                "In Office Sets": in_list,
                "Out (Hospital • Surgery)": out_list,
            })

            for _, r in out_case_rows.iterrows():
                out_rows.append({
                    "Category": label,
                    "Set": str(r.get("set_name", "")),
                    "Hospital": str(r.get("location_now", "")),
                    "Surgery Date": str(r.get("surgery_date", "")),
                    "Patient/Doctor": str(r.get("patient_doctor", "")),
                    "Case": str(r.get("case_id", "")),
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

            # Left: category name + badge + in-office set names
            in_sets_html = ""
            if r["In Office Sets"]:
                names = [s.strip() for s in r["In Office Sets"].split(";") if s.strip()]
                in_sets_html = "".join(
                    f"<span style='display:inline-block;background:#f0fdf4;color:#166534;"
                    f"font-family:\"JetBrains Mono\",monospace;font-size:11px;font-weight:600;"
                    f"border-radius:4px;padding:1px 7px;margin:2px 4px 2px 0'>{n}</span>"
                    for n in names
                )

            left_col = (
                f"<div style='flex:0 0 39%;padding-right:20px'>"
                f"<div style='display:flex;align-items:center;gap:12px;flex-wrap:wrap'>"
                f"<span class='inv-name' style='font-size:20px'>{r['Category']}</span>"
                f"{badge}"
                f"</div>"
                f"<div style='margin-top:6px'>{in_sets_html}</div>"
                f"</div>"
            )

            # Right: OUT lines
            out_items = _set_out.get(cat_norm, [])
            if out_items:
                out_lines = "".join(
                    f"<div class='out-line'>"
                    f"<span class='out-tag'>OUT</span> "
                    f"<span class='out-set'>{o['Set']}</span>"
                    f"<span class='out-sep'> → </span>"
                    f"<span class='out-hosp'>{o['Hospital'] or '—'}</span>"
                    f"<span class='out-sep'> · </span>"
                    f"<span class='out-surg'>surg {o['Surgery Date'] or '—'}</span>"
                    + (f"<span class='out-days' style='margin-left:8px'>{o['Patient/Doctor']}</span>" if o['Patient/Doctor'] else "")
                    + f"</div>"
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

        pt_avail_lookup: dict[str, int] = {}
        if not pt_avail_df.empty and "category" in pt_avail_df.columns:
            for _, r in pt_avail_df.iterrows():
                key = str(r.get("category", "")).upper().strip()
                if key:
                    pt_avail_lookup[key] = _safe_int(r.get("available", 0))

        summary_lookup: dict[str, int] = {}
        for row in summary_rows:
            row_norm = str(row.get("CategoryNorm", "")).strip().upper()
            if row_norm:
                summary_lookup[row_norm] = _safe_int(row.get("Available", 0))

        def _copy_count(category_keys: list[str], mode: str = "sum") -> int:
            total_val = 0
            values: list[int] = []
            for key in category_keys:
                key_norm = str(key).upper().strip()
                if key_norm in _POWERTOOL_CATS:
                    values.append(pt_avail_lookup.get(key_norm, 0))
                else:
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
        copy_lines.append("Power")
        copy_lines.append(f"5503B (normal) - {_copy_count(['P5503'])}")
        copy_lines.append(f"5400 (kwire) - {_copy_count(['P5400'])}")
        copy_lines.append(f"8400 (handpiece) - {_copy_count(['P8400'])}")

        st.markdown("##### Copy Block")
        st.code("\n".join(copy_lines), language="text")


# ── Plates ────────────────────────────────────────────────────────────────────
with inv_tabs[1]:
    plate_sum    = pd.DataFrame(report["plate_uid_summary"])
    plate_out_df = pd.DataFrame(report["plate_out_cases"])

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
                f"<span class='inv-name'>{row['proper_name']}</span>"
                f"<span>{badge}</span>"
                f"</div>"
                f"<div class='inv-sub'>{row['plate_uid']} &nbsp;·&nbsp; {row['screw_sizes'] or ''}{' &nbsp;·&nbsp; ' if row['screw_sizes'] else ''}{row['size_ranges']}</div>"
                f"<div class='inv-sub'>status: {row['status_note'] or 'READY'}</div>"
                f"{out_lines}"
                f"</div>"
            )

        st.markdown(
            "".join(_plate_row_html(r) for _, r in plate_sum.iterrows()),
            unsafe_allow_html=True,
        )

        with st.expander("📋 Plate size range detail"):
            psr = pd.DataFrame(report["plate_size_range_availability"])
            pdd = pd.DataFrame(report.get("plate_drawer_detail", []))
            if "proper_name" not in psr.columns:
                psr["proper_name"] = psr.get("plate_name", "")
            if "screw_sizes" not in psr.columns:
                psr["screw_sizes"] = ""
            psr["uid_norm"] = psr["plate_uid"].astype(str).str.upper().str.strip()
            psr["order_rank"] = psr["uid_norm"].map(PLATE_UID_RANK).fillna(10_000).astype(int)
            psr = psr.sort_values(["order_rank", "uid_norm", "size_range"])
            if search_query:
                psr = psr[
                    psr["plate_uid"].str.contains(search_query, case=False, na=False)
                    | psr["proper_name"].str.contains(search_query, case=False, na=False)
                    | psr["screw_sizes"].str.contains(search_query, case=False, na=False)
                ]
                if not pdd.empty:
                    pdd = pdd[
                        pdd["plate_uid"].str.contains(search_query, case=False, na=False)
                        | pdd["proper_name"].str.contains(search_query, case=False, na=False)
                        | pdd["screw_sizes"].str.contains(search_query, case=False, na=False)
                        | pdd["drawer"].str.contains(search_query, case=False, na=False)
                        | pdd["drawer_sizes"].str.contains(search_query, case=False, na=False)
                    ]
            st.dataframe(
                psr[[
                    "plate_uid", "proper_name", "screw_sizes", "size_range", "drawer_locations",
                    "available_units", "out_units", "total_units", "availability", "range_status",
                ]].rename(columns={
                    "plate_uid": "UID", "proper_name": "Proper Name", "screw_sizes": "Screw Sizes", "size_range": "Size",
                    "drawer_locations": "Drawers", "available_units": "Avail",
                    "out_units": "Out", "total_units": "Total", "availability": "Avail/Total", "range_status": "Status",
                }),
                use_container_width=True, hide_index=True,
            )
            if not pdd.empty:
                if "proper_name" not in pdd.columns:
                    pdd["proper_name"] = ""
                if "screw_sizes" not in pdd.columns:
                    pdd["screw_sizes"] = ""
                pdd["uid_norm"] = pdd["plate_uid"].astype(str).str.upper().str.strip()
                pdd["order_rank"] = pdd["uid_norm"].map(PLATE_UID_RANK).fillna(10_000).astype(int)
                pdd = pdd.sort_values(["order_rank", "uid_norm", "size_range", "drawer"])
                st.markdown("##### Drawer view")
                st.dataframe(
                    pdd[[
                        "plate_uid", "proper_name", "screw_sizes", "size_range", "drawer",
                        "drawer_sizes", "drawer_count", "availability", "range_status",
                    ]].rename(columns={
                        "plate_uid": "UID",
                        "proper_name": "Proper Name",
                        "screw_sizes": "Screw Sizes",
                        "size_range": "Size Range",
                        "drawer": "Drawer",
                        "drawer_sizes": "Sizes In Drawer",
                        "drawer_count": "Count",
                        "availability": "Avail/Total",
                        "range_status": "Status",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )


# ── Powertools ────────────────────────────────────────────────────────────────
with inv_tabs[2]:
    pt_avail = pd.DataFrame(report["powertool_category_availability"])
    pt_del   = pd.DataFrame(report["powertool_delivered"])
    pt_uid   = pd.DataFrame(report.get("powertool_uid_availability", []))
    pt_u30   = pd.DataFrame(report.get("powertool_usage_30d", []))

    if pt_avail.empty:
        st.info("No powertool data.")
    else:
        if not pt_uid.empty:
            for col in ("powertool_uid", "category", "availability", "hospital", "surgery_date", "patient_doctor"):
                if col not in pt_uid.columns:
                    pt_uid[col] = ""
            pt_uid = pt_uid.copy()
            pt_uid["availability_norm"] = pt_uid["availability"].astype(str).str.upper().str.strip()
            pt_uid["uid_name"] = pt_uid["powertool_uid"].astype(str).str.strip()

            if search_query:
                pt_uid = pt_uid[
                    pt_uid["category"].str.contains(search_query, case=False, na=False)
                    | pt_uid["uid_name"].str.contains(search_query, case=False, na=False)
                    | pt_uid["hospital"].str.contains(search_query, case=False, na=False)
                    | pt_uid["surgery_date"].str.contains(search_query, case=False, na=False)
                ]

        if search_query:
            pt_avail = pt_avail[
                pt_avail["category"].str.contains(search_query, case=False, na=False)
            ]

        # Build lookup: uid → list of out details
        _pt_uid_out: dict[str, list[dict]] = {}
        if not pt_del.empty:
            for _, pr in pt_del.iterrows():
                _pt_uid_out.setdefault(str(pr.get("powertool_uid", "")).strip(), []).append({
                    "uid":     str(pr.get("powertool_uid", "") or "").strip(),
                    "hospital":str(pr.get("hospital", "") or "").strip(),
                    "surgery": str(pr.get("surgery_date", "") or "").strip(),
                    "patient": str(pr.get("patient_doctor", "") or "").strip(),
                })

        usage_by_uid: dict[str, int] = {}
        if not pt_u30.empty:
            for _, ur in pt_u30.iterrows():
                key = str(ur.get("powertool_uid", "")).strip()
                if not key:
                    continue
                usage_by_uid[key] = int(pd.to_numeric(ur.get("usage_30d", 0), errors="coerce") or 0)

    def _pt_row_html(row: pd.Series) -> str:
        availability = str(row.get("availability_norm", "")).upper().strip()
        if availability in {"AVAILABLE", "NA_HOLD"}:
            available, total = 1, 1
        else:
            available, total = 0, 1
        badge = avail_badge(available, total)
        uid = str(row.get("uid_name", "")).strip()
        state_note = {
            "AVAILABLE": "in office",
            "NA_HOLD": "on hold",
            "OUT": "out",
            "OUT_NA_HOLD": "out (on hold)",
        }.get(availability, availability.lower() or "unknown")
        use_30d = usage_by_uid.get(uid, 0)
        usage_note = f"<span style='color:#6b7280;font-size:12px;margin-left:8px'>30d use: {use_30d}</span>"

        out_lines = ""
        if availability.startswith("OUT"):
            for o in _pt_uid_out.get(uid, []):
                hosp = o["hospital"] or "—"
                surg = o["surgery"] or "—"
                patient_note = ""
                if o.get("patient"):
                    patient_note = (
                        "<span class='out-days' style='margin-left:8px'>"
                        f"{o['patient']}</span>"
                    )
                out_lines += (
                    f"<div class='out-line'>"
                    f"<span class='out-tag'>OUT</span> "
                    f"<span class='out-hosp'>{hosp}</span>"
                    f"<span class='out-sep'> · </span>"
                    f"<span class='out-surg'>surg {surg}</span>"
                    f"{patient_note}"
                    f"</div>"
                )
        else:
            out_lines = "<span style='color:#9ca3af;font-size:12px;font-style:italic'>all in office</span>"

        return (
            f"<div class='inv-row'>"
            f"<div style='display:flex;justify-content:space-between;align-items:flex-start'>"
            f"<div>"
            f"<span class='inv-name'>{row['category']}: {uid}</span>"
            f"<div class='inv-sub'> {uid} · {state_note}{usage_note}</div>"
            f"</div>"
            f"<span>{badge}</span>"
            f"</div>"
            f"<div style='margin-top:6px'>{out_lines}</div>"
            f"</div>"
        )

        if pt_uid.empty:
            st.info("No powertool unit data.")
        else:
            st.markdown(
                "".join(_pt_row_html(r) for _, r in pt_uid.sort_values(["category", "uid_name"]).iterrows()),
                unsafe_allow_html=True,
            )


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
