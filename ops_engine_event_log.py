"""
ops_engine_event_log.py
────────────────────────
Drop-in addition to ops_engine.py that reads the Event_Log sheet tab
and uses it to override inferred set locations.

USAGE — add to your ops_engine.py build_operations_report() call:

    from ops_engine_event_log import load_event_log, apply_event_log_overrides

    def build_operations_report(...):
        ...
        # After building set_office_status rows, apply event log overrides:
        event_log = load_event_log(cases_source=cases_source)
        set_office_status = apply_event_log_overrides(set_office_status, event_log)
        ...

The Event_Log sheet tab must be published as a second CSV from your Google Sheet:
  File → Share → Publish to web → Event_Log tab → CSV

Or pass the event_log_source URL directly to load_event_log().
"""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Optional
import pandas as pd


# ─── LOADER ──────────────────────────────────────────────────────────────────

def load_event_log(
    event_log_source: Optional[str] = None,
    cases_source: Optional[str] = None,
) -> pd.DataFrame:
    """
    Load the Event_Log sheet tab as a DataFrame.

    If event_log_source is None, tries to derive the URL from cases_source
    by replacing gid=0 with gid=<event_log_gid>.  You'll need to set the
    correct gid for your sheet.

    Returns an empty DataFrame if unavailable.
    """
    if not event_log_source:
        if cases_source and "gid=0" in cases_source:
            # Replace with Event_Log tab gid — update this number to match yours
            event_log_source = cases_source.replace("gid=0", "gid=EVENT_LOG_GID")
        else:
            return pd.DataFrame()

    try:
        df = pd.read_csv(event_log_source)
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        # Normalise column names to canonical set
        rename_map = {
            "event type": "event_type",
            "set uid": "set_uid",
            "from hospital": "from_hospital",
            "case ref": "case_ref",
        }
        df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
        for col in ("timestamp", "event_type", "set_uid", "hospital", "from_hospital", "notes", "actor"):
            if col not in df.columns:
                df[col] = ""
        df = df.fillna("")
        df["set_uid"] = df["set_uid"].str.strip().str.upper()
        df["hospital"] = df["hospital"].str.strip().str.upper()
        df["event_type"] = df["event_type"].str.strip().str.upper()
        return df
    except Exception:
        return pd.DataFrame()


# ─── OVERRIDE LOGIC ──────────────────────────────────────────────────────────

def build_latest_location_from_events(event_log: pd.DataFrame) -> dict[str, dict]:
    """
    For each set UID, find the most recent event and determine its current location.

    Returns dict: uid → {
        "location":    "HTI" | "OFFICE" | ...,
        "event_type":  "SEND" | "COLLECT" | "RETURN",
        "timestamp":   "18/04/2026 17:42",
        "from_event":  True,
    }
    """
    if event_log.empty:
        return {}

    # Sort by timestamp descending (string sort works for dd/MM/yyyy HH:mm if consistent)
    df = event_log.copy()
    df["_ts_sort"] = df["timestamp"].apply(_parse_timestamp_for_sort)
    df = df.sort_values("_ts_sort", ascending=False)

    latest: dict[str, dict] = {}
    for _, row in df.iterrows():
        uid = str(row.get("set_uid", "")).strip().upper()
        if not uid:
            continue
        if uid in latest:
            continue  # already have most recent for this UID

        evt  = str(row.get("event_type", "")).strip().upper()
        hosp = str(row.get("hospital", "")).strip().upper()

        if evt == "SEND":
            location = hosp  # out at hospital
        elif evt in {"COLLECT", "RETURN"}:
            location = "OFFICE"  # back in office
        else:
            continue  # AMEND, CANCEL etc don't affect location

        latest[uid] = {
            "location":   location,
            "event_type": evt,
            "timestamp":  str(row.get("timestamp", "")),
            "from_event": True,
        }

    return latest


def apply_event_log_overrides(
    set_office_status: list[dict],
    event_log: pd.DataFrame,
) -> list[dict]:
    """
    Given the existing set_office_status rows (built by ops_engine from case rows),
    override location_now with the most recent Event_Log entry for each set UID.

    Priority:
        Event_Log (most recent SEND/COLLECT/RETURN)
        > Case row inference
        > master_data home
    """
    if event_log.empty:
        return set_office_status

    overrides = build_latest_location_from_events(event_log)
    if not overrides:
        return set_office_status

    updated = []
    for row in set_office_status:
        uid = str(row.get("set_display", "") or row.get("uid", "")).strip().upper()
        if uid in overrides:
            override = overrides[uid]
            row = dict(row)
            row["location_now"]      = override["location"]
            row["event_log_override"] = True
            row["event_log_ts"]       = override["timestamp"]
            row["event_log_type"]     = override["event_type"]
            # If event log says OFFICE, clear the case assignment
            if override["location"] == "OFFICE":
                row["assignment_kind"]  = ""
                row["location_now"]     = "OFFICE"
        updated.append(row)

    return updated


def get_today_event_summary(event_log: pd.DataFrame, today: Optional[date] = None) -> dict:
    """
    Returns counts and details of today's events for the EOD summary.
    """
    if event_log.empty:
        return {"sent": [], "collected": [], "total": 0}

    if today is None:
        today = datetime.now().date()

    today_str = today.strftime("%d/%m/%Y")

    sent      = []
    collected = []

    for _, row in event_log.iterrows():
        ts = str(row.get("timestamp", "")).strip()
        if not ts.startswith(today_str):
            continue
        evt  = str(row.get("event_type", "")).strip().upper()
        uid  = str(row.get("set_uid", "")).strip()
        hosp = str(row.get("hospital", "")).strip()
        if evt == "SEND":
            sent.append({"uid": uid, "hospital": hosp, "timestamp": ts})
        elif evt in {"COLLECT", "RETURN"}:
            collected.append({"uid": uid, "from": str(row.get("from_hospital", "") or hosp), "timestamp": ts})

    return {
        "sent":      sent,
        "collected": collected,
        "total":     len(sent) + len(collected),
    }


# ─── RECONCILIATION REPORT ───────────────────────────────────────────────────

def build_eod_reconciliation(
    set_office_status: list[dict],
    event_log: pd.DataFrame,
    today: Optional[date] = None,
) -> dict:
    """
    Compare what the system thinks is in office vs. what event log says.
    Surfaces discrepancies for the physical EOD check.

    Returns:
    {
        "system_in_office":    [{"uid":..., "category":...}, ...],
        "event_log_overrides": [{"uid":..., "event_type":..., "location":...}, ...],
        "discrepancies":       [{"uid":..., "system_says":..., "event_says":...}, ...],
        "today_events":        {...},
    }
    """
    if today is None:
        today = datetime.now().date()

    overrides       = build_latest_location_from_events(event_log)
    today_summary   = get_today_event_summary(event_log, today)

    system_in_office: list[dict] = []
    discrepancies:    list[dict] = []

    for row in set_office_status:
        uid      = str(row.get("set_display", "") or row.get("uid", "")).strip().upper()
        loc_now  = str(row.get("location_now", "")).strip().upper()
        category = str(row.get("category", "")).strip()

        system_in_office_flag = (loc_now == "OFFICE")
        if system_in_office_flag:
            system_in_office.append({"uid": uid, "category": category})

        if uid in overrides:
            ev_loc = overrides[uid]["location"].upper()
            ev_ts  = overrides[uid]["timestamp"]
            ev_type= overrides[uid]["event_type"]
            if system_in_office_flag and ev_loc != "OFFICE":
                discrepancies.append({
                    "uid":         uid,
                    "category":    category,
                    "system_says": "OFFICE",
                    "event_says":  ev_loc,
                    "event_type":  ev_type,
                    "event_ts":    ev_ts,
                })
            elif not system_in_office_flag and ev_loc == "OFFICE":
                discrepancies.append({
                    "uid":         uid,
                    "category":    category,
                    "system_says": loc_now,
                    "event_says":  "OFFICE",
                    "event_type":  ev_type,
                    "event_ts":    ev_ts,
                })

    override_list = [
        {"uid": uid, **data}
        for uid, data in overrides.items()
    ]

    return {
        "system_in_office":    system_in_office,
        "event_log_overrides": override_list,
        "discrepancies":       discrepancies,
        "today_events":        today_summary,
    }


# ─── UTILS ───────────────────────────────────────────────────────────────────

def _parse_timestamp_for_sort(ts: str) -> str:
    """
    Convert dd/MM/yyyy HH:mm → yyyy-MM-dd HH:mm for string sort.
    Falls back to original string if format doesn't match.
    """
    ts = str(ts).strip()
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2})", ts)
    if m:
        d, mo, y, hh, mm = m.groups()
        return f"{y}-{mo}-{d} {hh}:{mm}"
    return ts