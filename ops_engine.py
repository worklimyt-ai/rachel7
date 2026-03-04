"""
ops_engine.py  –  Osteo Ops core logic
Timezone : Asia/Kuala_Lumpur
"""
from __future__ import annotations

import argparse
import csv
import importlib.util
import io
import json
import math
import re
import sys
import types
import urllib.error
import urllib.request
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

KL_TZ = ZoneInfo("Asia/Kuala_Lumpur")

OFFICE_COORDS = (3.1857858767423703, 101.63494178629432)
TBS_COORDS    = (3.0781643275980146, 101.71154614310385)

DEFAULT_CASES_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IY"
    "nirCfsnGHoMyn5xPoG5c/pub?gid=0&single=true&output=csv"
)
DEFAULT_ARCHIVE_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IY"
    "nirCfsnGHoMyn5xPoG5c/pub?gid=1320419668&single=true&output=csv"
)
DEFAULT_CASES_LOCAL   = Path("/Users/bellbell/Downloads/cases - cases.csv")
DEFAULT_ARCHIVE_LOCAL = Path("/Users/bellbell/Downloads/cases - archive.csv")

DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")

# Known CSV code aliases → master_data HOSPITALS keys
HOSPITAL_ALIASES: dict[str, str] = {
    "QE1":        "HQE1",
    "QE2":        "HQE2",
    "HRPZII":     "HRPZ",
    "UMSC":       "UKMSC",
    "KPMC":       "KPMCK",
    "KPJT":       "KPJTWK",
    "KPJJS":      "KPJJ",
    "HUKMHCTM":   "HUKM",
    "SMCI":       "SMSJ",
    "GLENEAGLES": "GHKL",
}

# ---------------------------------------------------------------------------
# Date / time helpers
# ---------------------------------------------------------------------------

def now_kl() -> datetime:
    return datetime.now(KL_TZ)

def now_kl_date() -> date:
    return now_kl().date()

def parse_date(value: Any) -> date | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None

def format_date(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else ""

# ---------------------------------------------------------------------------
# String / token normalisers
# ---------------------------------------------------------------------------

def normalize_code(value: Any) -> str:
    return str(value or "").strip().upper().replace(" ", "")

def normalize_set_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9.]", "", normalize_code(value))

def normalize_plate_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9.\-_*]", "", normalize_code(value))

def canonical_powertool_uid(value: Any) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", normalize_code(value))
    if not raw.startswith("P"):
        return raw
    m = re.match(r"^(P\d{4})(\d+)$", raw)
    if not m:
        return raw
    prefix, suffix = m.groups()
    return f"{prefix}{int(suffix)}"

def canonical_powertool_shorthand(value: Any) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", normalize_code(value))
    m = re.match(r"^(P\d{4})", raw)
    return m.group(1) if m else raw

def split_tokens(value: Any, pattern: str = r"[;]+") -> list[str]:
    text = str(value or "").strip()
    return [p.strip() for p in re.split(pattern, text) if p.strip()] if text else []

def split_plate_tokens(value: Any) -> list[str]:
    out: list[str] = []
    for part in split_tokens(value, pattern=r"[;]+"):
        for sub in split_tokens(part, pattern=r"[/]+"):
            out.append(sub)
    return out

def parse_locations(raw_location: Any) -> tuple[list[str], list[str]]:
    drawers, others = [], []
    for part in split_tokens(raw_location, pattern=r"[,]+"):
        token = normalize_code(part)
        (drawers if re.match(r"^D\d+$", token) else others).append(token) if token else None
    return drawers, others

def canonical_size_range(value: Any) -> str:
    token = normalize_code(value).replace("_", " ").replace("-", " ")
    if "EXTRA" in token and "LONG" in token:
        return "EXTRA LONG"
    if token == "LONG":
        return "LONG"
    if token in {"", "STD", "STANDARD"}:
        return "STANDARD"
    return token

def compact_set_id(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "?"
    return str(int(text)) if re.fullmatch(r"\d+", text) else text

def format_set_display(shorthand: Any, set_id: Any) -> str:
    short = str(shorthand or "").strip()
    if not short:
        return str(set_id or "").strip()
    return f"{short} ({compact_set_id(set_id)})"

def is_na_status(value: Any) -> bool:
    return "NA" in normalize_code(value)

def is_powertool_category(category: str) -> bool:
    return bool(re.match(r"^P\d", normalize_code(category)))

# ---------------------------------------------------------------------------
# Plate request parsing
# ---------------------------------------------------------------------------

def parse_plate_request(token: str) -> dict[str, Any] | None:
    raw = normalize_plate_code(token)
    if not raw:
        return None
    from_stock = "*" in raw
    cleaned = raw.replace("*", "")
    if cleaned.endswith("-EL"):
        base_uid      = cleaned[:-3]
        needed_ranges = ["STANDARD", "LONG", "EXTRA LONG"]
    elif cleaned.endswith("-L"):
        base_uid      = cleaned[:-2]
        needed_ranges = ["STANDARD", "LONG"]
    else:
        base_uid      = cleaned
        needed_ranges = ["STANDARD"]
    uid_norm = normalize_set_code(base_uid)
    if not uid_norm:
        return None
    return {
        "raw_token":    token,
        "base_uid":     uid_norm,
        "needed_ranges": needed_ranges,
        "from_stock":   from_stock,
    }

# ---------------------------------------------------------------------------
# Hospital resolution
# ---------------------------------------------------------------------------

def resolve_hospital_code(
    raw_code: Any, hospitals: dict[str, Any]
) -> tuple[str | None, str]:
    code = normalize_code(raw_code)
    if not code:
        return None, "empty"
    if code in hospitals:
        return code, "exact"
    if code in HOSPITAL_ALIASES:
        alias = HOSPITAL_ALIASES[code]
        if alias in hospitals:
            return alias, f"alias:{code}->{alias}"
    if f"H{code}" in hospitals:
        return f"H{code}", "prefixed_H"
    if code.startswith("H") and code[1:] in hospitals:
        return code[1:], "stripped_H"
    return None, "unresolved"

# ---------------------------------------------------------------------------
# Distance / drive estimates
# ---------------------------------------------------------------------------

def haversine_km(lat1: float, lng1: float, lat2: float, lng2: float) -> float:
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    a = (
        math.sin(math.radians(lat2 - lat1) / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(math.radians(lng2 - lng1) / 2) ** 2
    )
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def estimated_drive_km(straight_km: float) -> float:
    """Road-winding multiplier for Peninsular Malaysia."""
    return straight_km * 1.3

def estimated_drive_minutes(drive_km: float, avg_speed_kmh: float = 55.0) -> float:
    return (drive_km / avg_speed_kmh) * 60.0 if drive_km > 0 else 0.0

# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_master_data(module_path: str | Path) -> dict[str, Any]:
    module_file = Path(module_path)
    if not module_file.exists():
        raise FileNotFoundError(f"master_data not found: {module_file}")

    def _exec() -> Any:
        spec = importlib.util.spec_from_file_location("master_data_runtime", module_file)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Cannot load module: {module_file}")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)  # type: ignore[union-attr]
        return mod

    try:
        module = _exec()
    except ModuleNotFoundError as exc:
        if exc.name != "pandas":
            raise
        prev = sys.modules.get("pandas")
        sys.modules["pandas"] = types.SimpleNamespace(DataFrame=None, ExcelWriter=None)  # type: ignore[assignment]
        try:
            module = _exec()
        finally:
            if prev is None:
                del sys.modules["pandas"]
            else:
                sys.modules["pandas"] = prev

    sets      = getattr(module, "SETS", None)
    plates    = getattr(module, "PLATES", None)
    hospitals = getattr(module, "HOSPITALS", None)
    if not (isinstance(sets, list) and isinstance(plates, dict) and isinstance(hospitals, dict)):
        raise RuntimeError(
            "master_data.py must expose SETS (list), PLATES (dict), HOSPITALS (dict)"
        )
    return {"SETS": sets, "PLATES": plates, "HOSPITALS": hospitals}


def read_csv_records(source: str | Path) -> list[dict[str, str]]:
    src = str(source)
    if src.startswith("http://") or src.startswith("https://"):
        try:
            with urllib.request.urlopen(src, timeout=20) as resp:
                payload = resp.read().decode("utf-8-sig")
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Failed to fetch CSV: {src}") from exc
        return [dict(r) for r in csv.DictReader(io.StringIO(payload))]
    path = Path(src)
    if not path.exists():
        raise FileNotFoundError(f"CSV not found: {path}")
    with path.open("r", newline="", encoding="utf-8-sig") as fh:
        return [dict(r) for r in csv.DictReader(fh)]


def auto_source(explicit: str | None, local: Path, fallback_url: str) -> str:
    if explicit:
        return explicit
    return str(local) if local.exists() else fallback_url

# ---------------------------------------------------------------------------
# Index builders
# ---------------------------------------------------------------------------

def build_set_indexes(master_sets: list[dict[str, Any]]) -> dict[str, Any]:
    uid_map:       dict[str, list[dict[str, Any]]] = defaultdict(list)
    shorthand_map: dict[str, list[dict[str, Any]]] = defaultdict(list)
    category_map:  dict[str, list[dict[str, Any]]] = defaultdict(list)
    office_sets:   list[dict[str, Any]] = []
    normalized:    list[dict[str, Any]] = []

    for row in master_sets:
        item = dict(row)
        item["category"] = str(item.get("category", "")).strip()
        item["id"]       = str(item.get("id", "")).strip()
        item["home"]     = str(item.get("home", "")).strip().upper()
        item["uid"]      = str(item.get("uid", "")).strip()
        item["shorthand"]= str(item.get("shorthand", "")).strip()
        item["status"]   = str(item.get("status", "")).strip()
        item["_uid_norm"]       = normalize_set_code(item["uid"])
        item["_shorthand_norm"] = normalize_set_code(item["shorthand"])
        item["_category_norm"]  = normalize_set_code(item["category"])
        item["_set_key"]        = f"{item['_uid_norm']}|{item['category']}|{item['id']}"
        normalized.append(item)

        if item["_uid_norm"]:       uid_map[item["_uid_norm"]].append(item)
        if item["_shorthand_norm"]: shorthand_map[item["_shorthand_norm"]].append(item)
        if item["_category_norm"]:  category_map[item["_category_norm"]].append(item)
        if item["home"] == "OFFICE":
            # Only count as office-available if not NA status
            office_sets.append(item)

    return {
        "all_sets":     normalized,
        "uid_map":      uid_map,
        "shorthand_map":shorthand_map,
        "category_map": category_map,
        "office_sets":  office_sets,
    }


def build_plate_inventory(master_plates: dict[str, dict[str, Any]]) -> dict[str, Any]:
    size_buckets: dict[tuple[str, str], dict[str, Any]] = {}
    uid_ranges:   dict[str, set[str]] = defaultdict(set)
    uid_alias_candidates: dict[str, set[str]] = defaultdict(set)

    for sku_code, row in master_plates.items():
        if not isinstance(row, dict):
            continue
        uid = normalize_set_code(row.get("uid", ""))
        if not uid:
            continue
        size_range = canonical_size_range(row.get("size_range", "STANDARD"))
        key = (uid, size_range)
        drawers, others = parse_locations(row.get("location", ""))
        drawer_units = len(drawers) or (0 if others else 1)
        stock_units  = len(others)

        bucket = size_buckets.setdefault(key, {
            "plate_uid":    uid,
            "plate_name":   str(row.get("proper_name", "")).strip(),
            "set_category": str(row.get("set", "")).strip(),
            "size_range":   size_range,
            "screw_sizes_set": set(),
            "out_drawer_units": 0,
            "out_stock_units":  0,
            "drawer_locations": set(),
            "stock_locations":  set(),
            "out_details":  [],
        })
        screw_sizes = str(row.get("screw_sizes", "")).strip()
        if screw_sizes:
            bucket["screw_sizes_set"].add(screw_sizes)
        bucket["drawer_locations"].update(drawers)
        bucket["stock_locations"].update(others)
        if not drawers and not others:
            bucket["stock_locations"].add("UNSPECIFIED")

        uid_ranges[uid].add(size_range)
        uid_alias_candidates[uid].add(uid)
        stripped = re.sub(r"^[0-9.]+", "", uid)
        if stripped and stripped != uid:
            uid_alias_candidates[stripped].add(uid)

    uid_alias_map: dict[str, str] = {
        alias: next(iter(targets))
        for alias, targets in uid_alias_candidates.items()
        if len(targets) == 1
    }
    return {"size_buckets": size_buckets, "uid_ranges": uid_ranges, "uid_alias_map": uid_alias_map}

# ---------------------------------------------------------------------------
# Case summariser
# ---------------------------------------------------------------------------

def summarize_cases(
    cases_rows: list[dict[str, str]],
    set_indexes: dict[str, Any],
    today_kl: date,
) -> dict[str, Any]:
    uid_map      = set_indexes["uid_map"]
    shorthand_map= set_indexes["shorthand_map"]
    category_map = set_indexes["category_map"]

    parsed_cases:       list[dict[str, Any]] = []
    set_out_assignments:list[dict[str, Any]] = []
    unknown_set_tokens: Counter[str]         = Counter()

    for idx, row in enumerate(cases_rows, start=1):
        delivery_date  = parse_date(row.get("delivery_date", ""))
        surgery_date   = parse_date(row.get("surgery_date", ""))
        set_tokens     = split_tokens(row.get("sets", ""))
        plate_tokens   = split_plate_tokens(row.get("plates", ""))
        powertool_tokens = split_tokens(row.get("powertools") or row.get("powertool", ""))

        uid_hits:       list[dict[str, Any]] = []
        shorthand_hits: list[dict[str, Any]] = []
        set_display_tokens: list[str] = []

        for token in set_tokens:
            norm = normalize_set_code(token)
            if not norm:
                continue

            uid_matches, shorthand_matches = [], []
            if norm in uid_map:
                uid_matches = uid_map[norm]
                uid_hits.extend(uid_matches)
            elif norm in shorthand_map:
                shorthand_matches = shorthand_map[norm]
                shorthand_hits.extend(shorthand_matches)
            elif norm in category_map:
                shorthand_matches = category_map[norm]
                shorthand_hits.extend(shorthand_matches)
            elif len(norm) >= 4:
                fuzzy = [
                    item
                    for cat_key, items in category_map.items()
                    if cat_key.startswith(norm)
                    for item in items
                ]
                if fuzzy:
                    shorthand_matches = fuzzy
                    shorthand_hits.extend(fuzzy)
                else:
                    unknown_set_tokens[norm] += 1
                    set_display_tokens.append(token)
            else:
                unknown_set_tokens[norm] += 1
                set_display_tokens.append(token)

            if uid_matches:
                labels = sorted({
                    format_set_display(
                        item.get("shorthand") or item.get("category"),
                        item.get("id"),
                    )
                    for item in uid_matches
                })
                set_display_tokens.append("/".join(labels))
            elif shorthand_matches:
                labels = sorted({
                    f"{(item.get('shorthand') or item.get('category') or '').strip()} (?)"
                    for item in shorthand_matches
                })
                if labels:
                    set_display_tokens.append("/".join(l for l in labels if l.strip()))

        uid_hit_by_key     = {item["_set_key"]: item for item in uid_hits}
        shorthand_hit_keys = {
            item["_shorthand_norm"]
            for item in shorthand_hits
            if item.get("_shorthand_norm")
        }
        has_uid_set      = bool(uid_hit_by_key)
        has_shorthand_only = bool(shorthand_hit_keys) and not has_uid_set

        case_id       = f"C{idx:03d}"
        hospital_code = normalize_code(row.get("hospital", ""))

        case_record: dict[str, Any] = {
            "case_id":          case_id,
            "row_number":       idx,
            "prefix":           str(row.get("prefix", "")).strip(),
            "hospital":         hospital_code,
            "patient_doctor":   str(row.get("patient_doctor", "")).strip(),
            "delivery_date":    format_date(delivery_date),
            "surgery_date":     format_date(surgery_date),
            "sales_code":       str(row.get("sales_code", "")).strip(),
            "return_date":      str(row.get("return_date", "")).strip(),
            "status":           str(row.get("status", "")).strip(),
            "smart_status":     str(row.get("Smart Status", "")).strip(),
            "sets_raw":         str(row.get("sets", "")).strip(),
            "sets_display":     "; ".join(t for t in set_display_tokens if t)
                                or str(row.get("sets", "")).strip(),
            "plates_raw":       str(row.get("plates", "")).strip(),
            "powertools_raw":   str(row.get("powertools") or row.get("powertool", "")).strip(),
            "has_uid_set":      has_uid_set,
            "has_shorthand_only": has_shorthand_only,
            "set_uid_tokens":   sorted({item["_uid_norm"] for item in uid_hit_by_key.values()}),
            "set_shorthand_tokens": sorted(shorthand_hit_keys),
            # internal date objects (stripped before export)
            "delivery_date_obj": delivery_date,
            "surgery_date_obj":  surgery_date,
            "set_tokens":       set_tokens,
            "plate_tokens":     plate_tokens,
            "powertool_tokens": powertool_tokens,
        }
        parsed_cases.append(case_record)

        # Record which OFFICE sets are currently out for this case
        for item in uid_hit_by_key.values():
            if item.get("home") != "OFFICE":
                continue
            days_since = (today_kl - surgery_date).days if surgery_date else None
            set_out_assignments.append({
                "set_key":          item["_set_key"],
                "set_uid":          item["uid"],
                "category":         item["category"],
                "id":               item["id"],
                "set_display":      format_set_display(
                                        item.get("shorthand") or item.get("category"),
                                        item.get("id"),
                                    ),
                "home":             item["home"],
                "set_status":       item.get("status", ""),
                "location_now":     hospital_code or "UNKNOWN",
                "case_id":          case_id,
                "delivery_date":    format_date(delivery_date),
                "surgery_date":     format_date(surgery_date),
                "days_since_surgery": days_since,
                "case_status":      str(row.get("status", "")).strip(),
                "smart_status":     str(row.get("Smart Status", "")).strip(),
                "patient_doctor":   str(row.get("patient_doctor", "")).strip(),
            })

    return {
        "parsed_cases":       parsed_cases,
        "set_out_assignments":set_out_assignments,
        "unknown_set_tokens": unknown_set_tokens,
    }

# ---------------------------------------------------------------------------
# Output builders
# ---------------------------------------------------------------------------

def build_set_outputs(
    set_indexes: dict[str, Any], set_out_assignments: list[dict[str, Any]]
) -> dict[str, Any]:
    office_sets = set_indexes["office_sets"]

    # Exclude NA-status sets from totals
    active_office_sets = [s for s in office_sets if not is_na_status(s.get("status", ""))]

    total_by_category = Counter(item["category"] for item in active_office_sets)
    out_by_category   = Counter(
        item["category"]
        for item in set_out_assignments
        if not is_na_status(item.get("set_status", ""))
    )

    set_category_availability: list[dict[str, Any]] = []
    for category in sorted(total_by_category):
        total     = total_by_category[category]
        out       = out_by_category.get(category, 0)
        available = max(total - out, 0)
        set_category_availability.append({
            "category":    category,
            "total_office": total,
            "out_for_case": out,
            "available":    available,
            "availability": f"{available}/{total}",
        })

    # Per-set location status
    by_set_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in set_out_assignments:
        by_set_key[row["set_key"]].append(row)

    set_office_status: list[dict[str, Any]] = []
    for s in sorted(office_sets, key=lambda x: (x["category"], x["id"], x["uid"])):
        key         = s["_set_key"]
        assignments = by_set_key.get(key, [])
        base_row    = {
            "category":          s["category"],
            "id":                s["id"],
            "set_display":       format_set_display(
                                     s.get("shorthand") or s.get("category"), s.get("id")
                                 ),
            "home":              s["home"],
            "set_status":        s.get("status", ""),
        }
        if not assignments:
            set_office_status.append({
                **base_row,
                "location_now":      "OFFICE",
                "delivery_date":     "",
                "surgery_date":      "",
                "days_since_surgery":"",
                "case_status":       "",
                "case_id":           "",
                "patient_doctor":    "",
            })
        else:
            for row in assignments:
                set_office_status.append({
                    **base_row,
                    "location_now":      row["location_now"],
                    "delivery_date":     row["delivery_date"],
                    "surgery_date":      row["surgery_date"],
                    "days_since_surgery":row["days_since_surgery"],
                    "case_status":       row["case_status"],
                    "case_id":           row["case_id"],
                    "patient_doctor":    row["patient_doctor"],
                })

    return {
        "set_category_availability": set_category_availability,
        "set_office_status":         set_office_status,
    }


def build_plate_outputs(
    parsed_cases: list[dict[str, Any]],
    plate_inventory: dict[str, Any],
) -> dict[str, Any]:
    size_buckets  = plate_inventory["size_buckets"]
    uid_ranges    = plate_inventory["uid_ranges"]
    uid_alias_map = plate_inventory["uid_alias_map"]
    unknown_plate_tokens: Counter[str] = Counter()

    for case in parsed_cases:
        if not case["has_uid_set"]:
            continue
        for token in case["plate_tokens"]:
            parsed = parse_plate_request(token)
            if not parsed:
                continue
            uid = parsed["base_uid"]
            resolved = uid_alias_map.get(uid, uid)
            if resolved not in uid_ranges:
                unknown_plate_tokens[uid] += 1
                continue
            available_ranges = uid_ranges[resolved]
            for needed in parsed["needed_ranges"]:
                if needed not in available_ranges:
                    continue
                bucket = size_buckets[(resolved, needed)]
                if parsed["from_stock"] or not bucket["drawer_locations"]:
                    bucket["out_stock_units"] += 1
                else:
                    bucket["out_drawer_units"] += 1
                bucket["out_details"].append({
                    "case_id":        case["case_id"],
                    "hospital":       case["hospital"],
                    "delivery_date":  case["delivery_date"],
                    "surgery_date":   case["surgery_date"],
                    "raw_plate_token":parsed["raw_token"],
                    "from_stock":     parsed["from_stock"],
                })

    plate_size_range_availability: list[dict[str, Any]] = []
    plate_out_cases: list[dict[str, Any]] = []

    for (uid, size_range), bucket in sorted(size_buckets.items()):
        td = len(bucket["drawer_locations"])
        ts = len(bucket["stock_locations"])
        od = bucket["out_drawer_units"]
        os_ = bucket["out_stock_units"]
        total = td + ts
        out   = od + os_
        row = {
            "plate_uid":             uid,
            "proper_name":           bucket["plate_name"],
            "plate_name":            bucket["plate_name"],
            "screw_sizes":           ", ".join(sorted(bucket["screw_sizes_set"])),
            "set_category":          bucket["set_category"],
            "size_range":            size_range,
            "drawer_locations":      ", ".join(sorted(bucket["drawer_locations"])),
            "stock_locations":       ", ".join(sorted(bucket["stock_locations"])),
            "total_drawer_units":    td,
            "total_stock_units":     ts,
            "total_units":           total,
            "out_drawer_units":      od,
            "out_stock_units":       os_,
            "out_units":             out,
            "available_drawer_units":max(td - od, 0),
            "available_stock_units": max(ts - os_, 0),
            "available_units":       max(total - out, 0),
            "availability":          f"{max(total - out, 0)}/{total}",
        }
        plate_size_range_availability.append(row)
        for detail in bucket["out_details"]:
            plate_out_cases.append({
                "plate_uid":      uid,
                "size_range":     size_range,
                "hospital":       detail["hospital"],
                "case_id":        detail["case_id"],
                "delivery_date":  detail["delivery_date"],
                "surgery_date":   detail["surgery_date"],
                "from_stock":     detail["from_stock"],
                "raw_plate_token":detail["raw_plate_token"],
            })

    # UID-level summary
    uid_summary_map: dict[str, dict[str, Any]] = {}
    for row in plate_size_range_availability:
        uid = row["plate_uid"]
        s   = uid_summary_map.setdefault(uid, {
            "plate_uid":    uid,
            "proper_name":  row.get("proper_name", row["plate_name"]),
            "plate_name":   row["plate_name"],
            "screw_sizes_set": set(),
            "set_category": row["set_category"],
            "total_units":  0,
            "out_units":    0,
            "available_units": 0,
            "size_ranges":  [],
        })
        for ss in [p.strip() for p in str(row.get("screw_sizes", "")).split(",") if p.strip()]:
            s["screw_sizes_set"].add(ss)
        s["total_units"]    += row["total_units"]
        s["out_units"]      += row["out_units"]
        s["available_units"]+= row["available_units"]
        s["size_ranges"].append(f"{row['size_range']}({row['available_units']}/{row['total_units']})")

    plate_uid_summary = []
    for uid, s in sorted(uid_summary_map.items()):
        s["size_ranges"]  = ", ".join(s["size_ranges"])
        s["screw_sizes"] = ", ".join(sorted(s["screw_sizes_set"]))
        s.pop("screw_sizes_set", None)
        s["availability"] = f"{s['available_units']}/{s['total_units']}"
        plate_uid_summary.append(s)

    return {
        "plate_size_range_availability": plate_size_range_availability,
        "plate_uid_summary":             plate_uid_summary,
        "plate_out_cases":               sorted(
            plate_out_cases,
            key=lambda x: (x["plate_uid"], x["hospital"], x["case_id"]),
        ),
        "unknown_plate_tokens":          unknown_plate_tokens,
    }


def build_powertool_outputs(
    parsed_cases: list[dict[str, Any]],
    archive_rows: list[dict[str, str]],
    set_indexes: dict[str, Any],
    today_kl: date,
) -> dict[str, Any]:
    office_sets = set_indexes["office_sets"]
    office_powertools = [s for s in office_sets if is_powertool_category(s["category"])]

    power_uid_map:      dict[str, list[dict[str, Any]]] = defaultdict(list)
    power_shorthand_map:dict[str, list[dict[str, Any]]] = defaultdict(list)

    for item in office_powertools:
        item["_power_uid"]      = canonical_powertool_uid(item["uid"])
        sh_src                  = item.get("shorthand") or item.get("category")
        item["_power_shorthand"]= canonical_powertool_shorthand(sh_src)
        item["_power_key"]      = f"{item['_power_uid']}|{item['category']}|{item['id']}"
        item["_is_na_hold"]     = is_na_status(item.get("status", ""))
        if item["_power_uid"]:       power_uid_map[item["_power_uid"]].append(item)
        if item["_power_shorthand"]: power_shorthand_map[item["_power_shorthand"]].append(item)

    delivered_assignments: list[dict[str, Any]] = []
    unknown_powertool_tokens: Counter[str] = Counter()

    for case in parsed_cases:
        used_keys: set[str] = set()
        for token in case["powertool_tokens"]:
            uid_norm = canonical_powertool_uid(token)
            if uid_norm in power_uid_map:
                for item in power_uid_map[uid_norm]:
                    if item["_power_key"] in used_keys:
                        continue
                    used_keys.add(item["_power_key"])
                    delivered_assignments.append({
                        "powertool_uid":           item["uid"],
                        "powertool_uid_canonical": item["_power_uid"],
                        "category":                item["category"],
                        "id":                      item["id"],
                        "home":                    item["home"],
                        "status":                  item.get("status", ""),
                        "is_na_hold":              item["_is_na_hold"],
                        "case_id":                 case["case_id"],
                        "hospital":                case["hospital"],
                        "delivery_date":           case["delivery_date"],
                        "surgery_date":            case["surgery_date"],
                        "patient_doctor":          case["patient_doctor"],
                    })
                continue
            shorthand_norm = canonical_powertool_shorthand(token)
            if shorthand_norm in power_shorthand_map:
                continue   # pending delivery – not yet assigned
            unk = re.sub(r"[^A-Z0-9]", "", normalize_code(token))
            if unk:
                unknown_powertool_tokens[unk] += 1

    out_by_cat   = Counter(a["category"] for a in delivered_assignments if not a.get("is_na_hold"))
    total_by_cat = Counter(item["category"] for item in office_powertools)
    na_by_cat    = Counter(item["category"] for item in office_powertools if item.get("_is_na_hold"))

    powertool_category_availability: list[dict[str, Any]] = []
    for cat in sorted(total_by_cat):
        total    = total_by_cat[cat]
        na_hold  = na_by_cat.get(cat, 0)
        usable   = max(total - na_hold, 0)
        out      = out_by_cat.get(cat, 0)
        available= max(usable - out, 0)
        powertool_category_availability.append({
            "category":    cat,
            "total_office": total,
            "na_hold":     na_hold,
            "usable_total":usable,
            "out_for_case":out,
            "available":   available,
            "availability":f"{available}/{usable}",
        })

    assignment_by_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for a in delivered_assignments:
        key = f"{canonical_powertool_uid(a['powertool_uid'])}|{a['category']}|{a['id']}"
        assignment_by_key[key].append(a)

    powertool_uid_availability: list[dict[str, Any]] = []
    for item in sorted(office_powertools, key=lambda x: (x["category"], x["uid"], x["id"])):
        key   = item["_power_key"]
        rows  = assignment_by_key.get(key, [])
        base  = {
            "powertool_uid":item["uid"],
            "category":     item["category"],
            "id":           item["id"],
            "home":         item["home"],
            "status":       item.get("status", ""),
            "is_na_hold":   item.get("_is_na_hold", False),
        }
        if not rows:
            powertool_uid_availability.append({
                **base,
                "availability":"NA_HOLD" if item.get("_is_na_hold") else "AVAILABLE",
                "hospital":"", "case_id":"", "delivery_date":"", "surgery_date":"",
            })
        else:
            for a in rows:
                powertool_uid_availability.append({
                    **base,
                    "availability":"OUT_NA_HOLD" if item.get("_is_na_hold") else "OUT",
                    "hospital":     a["hospital"],
                    "case_id":      a["case_id"],
                    "delivery_date":a["delivery_date"],
                    "surgery_date": a["surgery_date"],
                })

    # 30-day usage from archive
    window_start = today_kl - timedelta(days=30)
    usage_counter: Counter[str] = Counter()
    for row in archive_rows:
        row_date = parse_date(row.get("surgery_date") or row.get("delivery_date") or "")
        if row_date is None or not (window_start <= row_date <= today_kl):
            continue
        for token in split_tokens(row.get("powertool") or row.get("powertools") or ""):
            uid_norm = canonical_powertool_uid(token)
            if uid_norm in power_uid_map:
                usage_counter[uid_norm] += 1

    powertool_usage_30d: list[dict[str, Any]] = []
    for item in sorted(office_powertools, key=lambda x: (x["category"], x["uid"], x["id"])):
        powertool_usage_30d.append({
            "powertool_uid":item["uid"],
            "category":     item["category"],
            "id":           item["id"],
            "is_na_hold":   item.get("_is_na_hold", False),
            "usage_30d":    usage_counter.get(item["_power_uid"], 0),
            "window_start": format_date(window_start),
            "window_end":   format_date(today_kl),
        })

    return {
        "powertool_category_availability": powertool_category_availability,
        "powertool_uid_availability":      powertool_uid_availability,
        "powertool_delivered":             delivered_assignments,
        "powertool_usage_30d":             powertool_usage_30d,
        "unknown_powertool_tokens":        unknown_powertool_tokens,
    }


def build_case_buckets(
    parsed_cases: list[dict[str, Any]], today_kl: date
) -> dict[str, list[dict[str, Any]]]:
    """
    Categorise each case into action buckets.
    A case can appear in multiple buckets (by design).
    """
    tomorrow = today_kl + timedelta(days=1)

    buckets: dict[str, list[dict[str, Any]]] = {
        "to_collect":          [],
        "to_deliver":          [],
        "to_check":            [],
        "to_top_up":           [],
        "to_follow_up":        [],
        "delivered_today":     [],
        "to_deliver_tomorrow": [],
    }

    for case in parsed_cases:
        delivery_date = case["delivery_date_obj"]
        surgery_date  = case["surgery_date_obj"]
        sales_code    = normalize_code(case["sales_code"])
        return_date   = normalize_code(case["return_date"])
        status        = normalize_code(case["status"])
        prefix        = normalize_code(case["prefix"])

        # Sales code recorded but equipment not yet returned
        if sales_code and not return_date:
            buckets["to_collect"].append(case)

        # Upcoming cases with only shorthand (set not yet assigned)
        if delivery_date and delivery_date >= today_kl and case["has_shorthand_only"]:
            buckets["to_deliver"].append(case)
            if delivery_date == tomorrow:
                buckets["to_deliver_tomorrow"].append(case)

        # ITO status → needs verification
        if "ITO" in status:
            buckets["to_check"].append(case)

        # Prefix P → plate top-up needed
        if prefix.startswith("P"):
            buckets["to_top_up"].append(case)

        # Surgery done but no sales code yet
        if surgery_date and surgery_date < today_kl and not sales_code:
            buckets["to_follow_up"].append(case)

        # Delivered today (UID confirmed)
        if delivery_date == today_kl and case["has_uid_set"]:
            buckets["delivered_today"].append(case)

    def _strip(case: dict[str, Any]) -> dict[str, Any]:
        return {
            "case_id":       case["case_id"],
            "prefix":        case["prefix"],
            "hospital":      case["hospital"],
            "patient_doctor":case["patient_doctor"],
            "delivery_date": case["delivery_date"],
            "surgery_date":  case["surgery_date"],
            "sales_code":    case["sales_code"],
            "return_date":   case["return_date"],
            "status":        case["status"],
            "smart_status":  case["smart_status"],
            "sets":          case["sets_display"],
            "sets_raw":      case["sets_raw"],
            "plates":        case["plates_raw"],
            "powertools":    case["powertools_raw"],
        }

    return {name: [_strip(c) for c in items] for name, items in buckets.items()}


def build_distance_rows(
    cases: list[dict[str, Any]],
    hospitals: dict[str, dict[str, Any]],
    route_tag: str,
) -> tuple[list[dict[str, Any]], Counter[str]]:
    unresolved: Counter[str] = Counter()
    rows: list[dict[str, Any]] = []

    for case in cases:
        raw_code = case["hospital"]
        resolved_code, resolved_by = resolve_hospital_code(raw_code, hospitals)

        if resolved_code is None:
            unresolved[raw_code] += 1
            rows.append({
                "route_group":          route_tag,
                "case_id":              case["case_id"],
                "hospital":             raw_code,
                "resolved_hospital":    "",
                "hospital_name":        "Unknown hospital code",
                "office_to_hospital_km":"",
                "office_est_drive_km":  "",
                "office_est_drive_min": "",
                "tbs_to_hospital_km":   "",
                "resolved_by":          resolved_by,
                "delivery_date":        case["delivery_date"],
                "surgery_date":         case["surgery_date"],
                "sets":                 case.get("sets", ""),
                "plates":               case.get("plates", ""),
                "sales_code":           case.get("sales_code", ""),
            })
            continue

        meta = hospitals[resolved_code]
        lat, lng = float(meta["lat"]), float(meta["lng"])
        straight  = haversine_km(OFFICE_COORDS[0], OFFICE_COORDS[1], lat, lng)
        tbs_dist  = haversine_km(TBS_COORDS[0], TBS_COORDS[1], lat, lng)
        drive_km  = estimated_drive_km(straight)
        drive_min = estimated_drive_minutes(drive_km)

        rows.append({
            "route_group":          route_tag,
            "case_id":              case["case_id"],
            "hospital":             raw_code,
            "resolved_hospital":    resolved_code,
            "hospital_name":        meta.get("name", ""),
            "office_to_hospital_km":round(straight, 2),
            "office_est_drive_km":  round(drive_km, 2),
            "office_est_drive_min": round(drive_min, 1),
            "tbs_to_hospital_km":   round(tbs_dist, 2),
            "resolved_by":          resolved_by,
            "delivery_date":        case["delivery_date"],
            "surgery_date":         case["surgery_date"],
            "sets":                 case.get("sets", ""),
            "plates":               case.get("plates", ""),
            "sales_code":           case.get("sales_code", ""),
        })

    rows.sort(key=lambda r: (
        parse_date(r.get("delivery_date", "")) or date.max,
        normalize_code(r.get("hospital", "")),
        r.get("case_id", ""),
    ))
    return rows, unresolved

# ---------------------------------------------------------------------------
# Main report builder
# ---------------------------------------------------------------------------

def build_operations_report(
    master_data_path: str | Path = "master_data.py",
    cases_source: str | None = None,
    archive_source: str | None = None,
    today_kl: date | None = None,
) -> dict[str, Any]:
    today = today_kl or now_kl_date()

    chosen_cases   = auto_source(cases_source,   DEFAULT_CASES_LOCAL,   DEFAULT_CASES_URL)
    chosen_archive = auto_source(archive_source, DEFAULT_ARCHIVE_LOCAL, DEFAULT_ARCHIVE_URL)

    master       = load_master_data(master_data_path)
    cases_rows   = read_csv_records(chosen_cases)
    archive_rows = read_csv_records(chosen_archive)

    set_indexes   = build_set_indexes(master["SETS"])
    case_summary  = summarize_cases(cases_rows, set_indexes, today)
    set_outputs   = build_set_outputs(set_indexes, case_summary["set_out_assignments"])

    plate_inventory = build_plate_inventory(master["PLATES"])
    plate_outputs   = build_plate_outputs(case_summary["parsed_cases"], plate_inventory)

    powertool_outputs = build_powertool_outputs(
        case_summary["parsed_cases"], archive_rows, set_indexes, today
    )

    case_buckets = build_case_buckets(case_summary["parsed_cases"], today)

    deliver_tomorrow_rows, unresolved_tomorrow = build_distance_rows(
        case_buckets["to_deliver_tomorrow"], master["HOSPITALS"], "to_deliver_tomorrow"
    )
    delivered_today_rows, unresolved_today = build_distance_rows(
        case_buckets["delivered_today"], master["HOSPITALS"], "delivered_today"
    )
    unknown_hospitals = unresolved_tomorrow + unresolved_today

    hospital_directory: list[dict[str, Any]] = [
        {
            "hosp_code":              code,
            "name":                   str(info.get("name", "")),
            "region":                 str(info.get("region", "")),
            "lat":                    float(info.get("lat", 0)),
            "lng":                    float(info.get("lng", 0)),
            "office_to_hospital_km":  round(
                haversine_km(OFFICE_COORDS[0], OFFICE_COORDS[1], float(info.get("lat",0)), float(info.get("lng",0))), 2
            ),
            "tbs_to_hospital_km":     round(
                haversine_km(TBS_COORDS[0], TBS_COORDS[1], float(info.get("lat",0)), float(info.get("lng",0))), 2
            ),
        }
        for code, info in sorted(master["HOSPITALS"].items())
    ]

    unknown = {
        "set_tokens": sorted(
            [{"token": t, "count": c} for t, c in case_summary["unknown_set_tokens"].items()],
            key=lambda x: (-x["count"], x["token"]),
        ),
        "plate_tokens": sorted(
            [{"token": t, "count": c} for t, c in plate_outputs["unknown_plate_tokens"].items()],
            key=lambda x: (-x["count"], x["token"]),
        ),
        "powertool_tokens": sorted(
            [{"token": t, "count": c} for t, c in powertool_outputs["unknown_powertool_tokens"].items()],
            key=lambda x: (-x["count"], x["token"]),
        ),
        "hospitals_for_routes": sorted(
            [{"token": t, "count": c} for t, c in unknown_hospitals.items()],
            key=lambda x: (-x["count"], x["token"]),
        ),
    }

    return {
        "meta": {
            "timezone":    "Asia/Kuala_Lumpur",
            "today_kl":    str(today),
            "tomorrow_kl": str(today + timedelta(days=1)),
            "office_coords": {"lat": OFFICE_COORDS[0], "lng": OFFICE_COORDS[1]},
            "tbs_coords":    {"lat": TBS_COORDS[0],    "lng": TBS_COORDS[1]},
            "sources": {
                "master_data": str(master_data_path),
                "cases":       chosen_cases,
                "archive":     chosen_archive,
            },
            "counts": {
                "cases_rows":      len(cases_rows),
                "archive_rows":    len(archive_rows),
                "master_sets":     len(master["SETS"]),
                "master_plates":   len(master["PLATES"]),
                "master_hospitals":len(master["HOSPITALS"]),
            },
        },
        "kpis": {
            "cases_total":        len(case_summary["parsed_cases"]),
            "sets_out":           len(case_summary["set_out_assignments"]),
            "powertools_out":     len(powertool_outputs["powertool_delivered"]),
            "to_collect":         len(case_buckets["to_collect"]),
            "to_deliver":         len(case_buckets["to_deliver"]),
            "to_check":           len(case_buckets["to_check"]),
            "to_top_up":          len(case_buckets["to_top_up"]),
            "to_follow_up":       len(case_buckets["to_follow_up"]),
            "delivered_today":    len(case_buckets["delivered_today"]),
            "to_deliver_tomorrow":len(case_buckets["to_deliver_tomorrow"]),
        },
        "set_category_availability":      set_outputs["set_category_availability"],
        "set_office_status":              set_outputs["set_office_status"],
        "plate_size_range_availability":  plate_outputs["plate_size_range_availability"],
        "plate_uid_summary":              plate_outputs["plate_uid_summary"],
        "plate_out_cases":                plate_outputs["plate_out_cases"],
        "powertool_category_availability":powertool_outputs["powertool_category_availability"],
        "powertool_uid_availability":     powertool_outputs["powertool_uid_availability"],
        "powertool_delivered":            powertool_outputs["powertool_delivered"],
        "powertool_usage_30d":            powertool_outputs["powertool_usage_30d"],
        "hospital_directory":             hospital_directory,
        "case_buckets":                   case_buckets,
        "distance_routes": {
            "to_deliver_tomorrow": deliver_tomorrow_rows,
            "delivered_today":     delivered_today_rows,
        },
        "unknown": unknown,
    }

# ---------------------------------------------------------------------------
# CSV / JSON file writers
# ---------------------------------------------------------------------------

def write_csv_table(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fieldnames: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row:
            if key not in seen:
                seen.add(key)
                fieldnames.append(key)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({
                k: json.dumps(v, ensure_ascii=True) if isinstance(v, (dict, list)) else v
                for k, v in row.items()
            })


def write_report_files(report: dict[str, Any], output_dir: str | Path) -> dict[str, str]:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    paths: dict[str, str] = {}

    json_path = out / "operations_report.json"
    json_path.write_text(json.dumps(report, indent=2, ensure_ascii=True), encoding="utf-8")
    paths["operations_report_json"] = str(json_path)

    tables: dict[str, list[dict[str, Any]]] = {
        "set_category_availability.csv":      report["set_category_availability"],
        "set_office_status.csv":              report["set_office_status"],
        "plate_size_range_availability.csv":  report["plate_size_range_availability"],
        "plate_uid_summary.csv":              report["plate_uid_summary"],
        "plate_out_cases.csv":                report["plate_out_cases"],
        "powertool_category_availability.csv":report["powertool_category_availability"],
        "powertool_uid_availability.csv":     report["powertool_uid_availability"],
        "powertool_delivered.csv":            report["powertool_delivered"],
        "powertool_usage_30d.csv":            report["powertool_usage_30d"],
        "hospital_directory.csv":             report["hospital_directory"],
        "distance_to_deliver_tomorrow.csv":   report["distance_routes"]["to_deliver_tomorrow"],
        "distance_delivered_today.csv":       report["distance_routes"]["delivered_today"],
        "unknown_set_tokens.csv":             report["unknown"]["set_tokens"],
        "unknown_plate_tokens.csv":           report["unknown"]["plate_tokens"],
        "unknown_powertool_tokens.csv":       report["unknown"]["powertool_tokens"],
        "unknown_route_hospitals.csv":        report["unknown"]["hospitals_for_routes"],
    }
    for bucket_name, rows in report["case_buckets"].items():
        tables[f"cases_{bucket_name}.csv"] = rows

    for filename, rows in tables.items():
        p = out / filename
        write_csv_table(p, rows)
        paths[filename] = str(p)

    return paths

# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def _cli() -> None:
    parser = argparse.ArgumentParser(
        description="Build operations report from master_data + cases/archive CSVs"
    )
    parser.add_argument("--master-data", default="master_data.py")
    parser.add_argument("--cases",   default=None, help="Cases CSV path or URL")
    parser.add_argument("--archive", default=None, help="Archive CSV path or URL")
    parser.add_argument("--out",     default="outputs", help="Output directory")
    args = parser.parse_args()

    report = build_operations_report(
        master_data_path=args.master_data,
        cases_source=args.cases,
        archive_source=args.archive,
    )
    paths = write_report_files(report, args.out)

    print("=== Osteo Ops Report ===")
    print(f"Date (KL)          : {report['meta']['today_kl']}")
    print(f"Cases total        : {report['kpis']['cases_total']}")
    print(f"Delivered today    : {report['kpis']['delivered_today']}")
    print(f"To deliver         : {report['kpis']['to_deliver']}")
    print(f"To deliver tomorrow: {report['kpis']['to_deliver_tomorrow']}")
    print(f"To collect         : {report['kpis']['to_collect']}")
    print(f"To follow up       : {report['kpis']['to_follow_up']}")
    print(f"Output JSON        : {paths['operations_report_json']}")


if __name__ == "__main__":
    _cli()
