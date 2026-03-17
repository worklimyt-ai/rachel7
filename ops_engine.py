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
MODULE_DIR = Path(__file__).resolve().parent
DEFAULT_MASTER_DATA_PATH = MODULE_DIR / "master_data.py"
DEFAULT_OUTPUT_DIR = MODULE_DIR / "outputs"

DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")
BOOKING_HOLD_DAYS = 3

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

def case_order_key(
    row_number: int,
    delivery_date: date | None,
    surgery_date: date | None,
) -> tuple[date, date, int]:
    return (
        delivery_date or date.min,
        surgery_date or date.min,
        row_number,
    )

# ---------------------------------------------------------------------------
# String / token normalisers
# ---------------------------------------------------------------------------

def normalize_code(value: Any) -> str:
    return str(value or "").strip().upper().replace(" ", "")

def normalize_set_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9.]", "", normalize_code(value))

def normalize_plate_code(value: Any) -> str:
    return re.sub(r"[^A-Z0-9.\-_*]", "", normalize_code(value))

def normalize_header_name(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())

def row_value(row: dict[str, Any], *names: str) -> str:
    for name in names:
        if name in row:
            return str(row.get(name, "") or "")
    normalized = {
        normalize_header_name(key): value
        for key, value in row.items()
    }
    for name in names:
        key = normalize_header_name(name)
        if key in normalized:
            return str(normalized[key] or "")
    return ""

def canonical_powertool_uid(value: Any) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", normalize_code(value))
    if not raw.startswith("P"):
        m = re.match(r"^(5503|5400|8400)(\d+)$", raw)
        if m:
            prefix, suffix = m.groups()
            return f"P{prefix}{suffix}"
        return raw
    m = re.match(r"^(P\d{4})(\d+)$", raw)
    if not m:
        return raw
    prefix, suffix = m.groups()
    return f"{prefix}{suffix}"

def canonical_powertool_shorthand(value: Any) -> str:
    raw = re.sub(r"[^A-Z0-9]", "", normalize_code(value))
    if not raw.startswith("P"):
        if re.match(r"^(5503|5400|8400)\d+$", raw):
            return f"P{raw[:4]}"
    m = re.match(r"^(P\d{4})", raw)
    return m.group(1) if m else raw

def split_tokens(value: Any, pattern: str = r"[;]+") -> list[str]:
    text = str(value or "").strip()
    return [p.strip() for p in re.split(pattern, text) if p.strip()] if text else []


def split_powertool_tokens(value: Any) -> list[str]:
    text = str(value or "").strip()
    return [
        p.strip() for p in re.split(r"[;,/\n\r\t ]+", text)
        if p.strip()
    ] if text else []

def split_extra_item_tokens(value: Any) -> list[str]:
    text = str(value or "").strip()
    return [
        p.strip() for p in re.split(r"[;,\n\r]+", text)
        if p.strip()
    ] if text else []

def split_plate_tokens(value: Any) -> list[str]:
    out: list[str] = []
    for part in split_tokens(value, pattern=r"[;]+"):
        for sub in split_tokens(part, pattern=r"[/]+"):
            out.append(sub)
    return out

def location_no_stock_state(value: Any) -> bool | None:
    token = normalize_code(value)
    if token in {"NOSTOCK", "NO_STOCK", "OUT", "OUTOFSTOCK", "OUT_OF_STOCK", "MISSING", "NS"}:
        return True
    if token in {"AVAILABLE", "IN", "INSTOCK", "IN_STOCK", "OK"}:
        return False
    return None

def parse_plate_locations(
    raw_location: Any,
    *,
    default_no_stock: bool = False,
) -> list[dict[str, Any]]:
    locations: list[dict[str, Any]] = []
    for part in split_tokens(raw_location, pattern=r"[,]+"):
        text = str(part or "").strip()
        if not text:
            continue
        match = re.match(r"^(.*?)\[(.*?)\]\s*$", text)
        token_text = match.group(1).strip() if match else text
        token = normalize_code(token_text)
        if not token:
            continue
        no_stock = default_no_stock
        if match:
            flag_text = match.group(2).strip()
            flags = [
                flag.strip() for flag in re.split(r"[|/;]+", flag_text)
                if flag.strip()
            ]
            explicit_state = None
            for flag in flags:
                state = location_no_stock_state(flag)
                if state is not None:
                    explicit_state = state
            if explicit_state is not None:
                no_stock = explicit_state
        locations.append({
            "name": token,
            "is_drawer": bool(re.match(r"^D\d+$", token)),
            "no_stock": no_stock,
        })
    return locations

def canonical_size_range(value: Any) -> str:
    token = normalize_code(value).replace("_", " ").replace("-", " ")
    if "EXTRA" in token and "LONG" in token:
        return "EXTRA LONG"
    if token == "LONG":
        return "LONG"
    if token in {"", "STD", "STANDARD"}:
        return "STANDARD"
    return token


def size_range_sort_key(value: Any) -> tuple[int, str]:
    token = canonical_size_range(value)
    order = {"SHORT": 0, "STANDARD": 1, "LONG": 2, "EXTRA LONG": 3}
    return (order.get(token, 99), token)


CLAVICLE_REVERSED_UIDS = {"DSC", "MSC", "DIA", "DPLH"}


def plate_uses_reversed_lr_sequence(uid: Any) -> bool:
    return normalize_set_code(uid) in CLAVICLE_REVERSED_UIDS


def plate_label_sort_key(value: Any, *, reverse_lr: bool = False) -> tuple[int, int, str]:
    token = normalize_code(value)
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
    return (side_rank, numeric_rank, token)


def plate_detail_sort_key(uid: Any, detail: dict[str, Any]) -> tuple[int, int, str]:
    return plate_label_sort_key(
        detail.get("label", ""),
        reverse_lr=plate_uses_reversed_lr_sequence(uid),
    )


def drawer_sort_key(value: Any) -> tuple[int, str]:
    token = normalize_code(value)
    if token.startswith("D") and token[1:].isdigit():
        return (int(token[1:]), token)
    return (10_000, token)

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


def format_case_item_label(category: Any, item_id: Any) -> str:
    category_text = str(category or "").strip()
    item_text = compact_set_id(item_id)
    if category_text and item_text:
        return f"{category_text} ({item_text})"
    return category_text or item_text


def normalize_bonegraft_token(value: Any) -> str:
    token = str(value or "").strip().upper()
    if not token:
        return ""
    token = token.replace("×", "X").replace("³", "^3")
    token = re.sub(r"\s+", "", token)
    token = token.replace("CC", "")
    token = token.replace("MM", "")
    token = token.replace(".", "")
    cube_match = re.fullmatch(r"(\d+)\^?3", token)
    if cube_match:
        return f"{cube_match.group(1)}CUBE"
    cube_dims = re.fullmatch(r"(\d+)X\1X\1", token)
    if cube_dims:
        return f"{cube_dims.group(1)}CUBE"
    return token


def format_bonegraft_label(name: Any, presentation: Any) -> str:
    name_text = str(name or "").strip()
    presentation_text = str(presentation or "").strip()
    if name_text and presentation_text:
        return f"{name_text} - {presentation_text}"
    return name_text or presentation_text


def is_booking_prefix(value: Any) -> bool:
    return normalize_code(value).startswith("BC")


def booking_hold_start(delivery_date: date | None) -> date | None:
    if delivery_date is None:
        return None
    return delivery_date - timedelta(days=BOOKING_HOLD_DAYS)


def is_booking_hold_active(delivery_date: date | None, today_kl: date) -> bool:
    hold_start = booking_hold_start(delivery_date)
    if hold_start is None or delivery_date is None:
        return False
    return hold_start <= today_kl <= delivery_date


def serialize_case_common(case: dict[str, Any]) -> dict[str, Any]:
    return {
        "case_id":              case["case_id"],
        "prefix":               case["prefix"],
        "hospital":             case["hospital"],
        "patient_doctor":       case["patient_doctor"],
        "delivery_date":        case["delivery_date"],
        "surgery_date":         case["surgery_date"],
        "sales_code":           case["sales_code"],
        "return_date":          case["return_date"],
        "status":               case["status"],
        "smart_status":         case["smart_status"],
        "sets_raw":             case["sets_raw"],
        "sets":                 case["sets_display"],
        "sets_returned_raw":    case.get("sets_returned_raw", ""),
        "sets_returned":        case.get("sets_returned_display", ""),
        "sets_outstanding_raw": case.get("sets_outstanding_raw", ""),
        "sets_outstanding":     case.get("sets_outstanding_display", ""),
        "set_categories":       case.get("set_categories", []),
        "plates_raw":           case.get("plates_raw", ""),
        "powertools_raw":       case.get("powertools_raw", ""),
        "bonegraft_raw":        case.get("bonegraft_raw", ""),
        "extra_items_raw":      case.get("extra_items_raw", ""),
        "suggested_sets":       case.get("suggested_sets", []),
        "suggested_sets_summary": case.get("suggested_sets_summary", ""),
    }

def is_na_status(value: Any) -> bool:
    return "NA" in normalize_code(value)


def is_standby_status(value: Any) -> bool:
    return "STANDBY" in normalize_code(value)


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
    suffix_rules = [
        ("-XLONLY", ["EXTRA LONG"]),
        ("XLONLY", ["EXTRA LONG"]),
        ("-ELONLY", ["EXTRA LONG"]),
        ("ELONLY", ["EXTRA LONG"]),
        ("-LONLY", ["LONG"]),
        ("LONLY", ["LONG"]),
        ("-SONLY", ["SHORT"]),
        ("SONLY", ["SHORT"]),
        ("-EL", ["STANDARD", "LONG", "EXTRA LONG"]),
        ("-L", ["STANDARD", "LONG"]),
        ("-S", ["STANDARD", "SHORT"]),
    ]
    base_uid = cleaned
    needed_ranges = ["STANDARD"]
    for suffix, ranges in suffix_rules:
        if cleaned.endswith(suffix):
            base_uid = cleaned[:-len(suffix)]
            needed_ranges = ranges
            break
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

    sets       = getattr(module, "SETS", None)
    plates     = getattr(module, "PLATES", None)
    hospitals  = getattr(module, "HOSPITALS", None)
    bonegraft  = getattr(module, "BONEGRAFT", [])
    if not (isinstance(sets, list) and isinstance(plates, dict) and isinstance(hospitals, dict)):
        raise RuntimeError(
            "master_data.py must expose SETS (list), PLATES (dict), HOSPITALS (dict)"
        )
    if not isinstance(bonegraft, list):
        bonegraft = []
    return {"SETS": sets, "PLATES": plates, "HOSPITALS": hospitals, "BONEGRAFT": bonegraft}


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
    standby_sets:  list[dict[str, Any]] = []
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
            office_sets.append(item)
        elif item["home"] == "STANDBY":
            standby_sets.append(item)

    return {
        "all_sets":     normalized,
        "uid_map":      uid_map,
        "shorthand_map":shorthand_map,
        "category_map": category_map,
        "office_sets":  office_sets,
        "standby_sets": standby_sets,
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

        bucket = size_buckets.setdefault(key, {
            "plate_uid":    uid,
            "plate_name":   str(row.get("proper_name", "")).strip(),
            "set_category": str(row.get("set", "")).strip(),
            "size_range":   size_range,
            "screw_sizes_set": set(),
            "total_drawer_units": 0,
            "total_stock_units": 0,
            "out_drawer_units": 0,
            "out_stock_units":  0,
            "drawer_locations": set(),
            "stock_locations":  set(),
            "drawer_size_map": defaultdict(list),
            "stock_size_map": defaultdict(list),
            # drawer_size_detail[drawer] = [{label, size_range, no_stock}]
            # Add "status": "no_stock" to a plate SKU in master_data.py to mark unavailable
            "drawer_size_detail": defaultdict(list),
            "stock_size_detail": defaultdict(list),
            "all_size_labels": [],
            "out_details":  [],
        })
        screw_sizes = str(row.get("screw_sizes", "")).strip()
        if screw_sizes:
            bucket["screw_sizes_set"].add(screw_sizes)
        size_label = str(row.get("size", "")).strip()
        lr_label = str(row.get("left_right", "")).strip()
        plate_label = " ".join(part for part in (size_label, lr_label) if part).strip() or str(sku_code).strip()
        # Detect "status": "no_stock" (or "out", "out_of_stock") on this SKU
        sku_status   = normalize_code(str(row.get("status", "")))  # uppercased, no spaces
        sku_no_stock = sku_status in {"NO_STOCK", "NOSTOCK", "OUT", "OUTOFSTOCK"}
        locations = parse_plate_locations(
            row.get("location", ""),
            default_no_stock=sku_no_stock,
        )
        drawers = [loc["name"] for loc in locations if loc["is_drawer"]]
        others = [loc["name"] for loc in locations if not loc["is_drawer"]]
        drawer_units = len(drawers)
        stock_units  = len(others) or (1 if not drawers and not others else 0)
        for location in locations:
            detail_row = {
                "label":      plate_label,
                "size_range": size_range,
                "no_stock":   bool(location["no_stock"]),
            }
            if location["is_drawer"]:
                drawer = location["name"]
                bucket["drawer_size_map"][drawer].append(plate_label)
                bucket["drawer_size_detail"][drawer].append(detail_row)
            else:
                stock_loc = location["name"]
                bucket["stock_size_map"][stock_loc].append(plate_label)
                bucket["stock_size_detail"][stock_loc].append(detail_row)
        bucket["total_drawer_units"] += drawer_units
        bucket["total_stock_units"] += stock_units
        bucket["drawer_locations"].update(drawers)
        bucket["stock_locations"].update(others)
        if not drawers and not others:
            bucket["stock_locations"].add("UNSPECIFIED")
        if plate_label not in bucket["all_size_labels"]:
            bucket["all_size_labels"].append(plate_label)

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


def finalize_set_assignments(
    assignments_by_key: dict[str, list[dict[str, Any]]],
    *,
    repeated_same_hospital_status: str,
) -> tuple[list[dict[str, Any]], dict[str, set[str]]]:
    finalized: list[dict[str, Any]] = []
    keys_by_case: dict[str, set[str]] = defaultdict(set)

    for set_key, rows in assignments_by_key.items():
        ordered_rows = sorted(rows, key=lambda row: row["_case_order_key"])
        winner = dict(ordered_rows[-1])
        related_case_ids = [row["case_id"] for row in ordered_rows]
        related_hospitals = sorted({
            row["location_now"]
            for row in ordered_rows
            if row.get("location_now")
        })
        winner["ownership_related_cases"] = ";".join(related_case_ids)
        winner["ownership_related_hospitals"] = ";".join(related_hospitals)
        if len(ordered_rows) > 1:
            winner["ownership_status"] = (
                "REVIEW_CONFLICT"
                if len(related_hospitals) > 1 else
                repeated_same_hospital_status
            )
        else:
            winner["ownership_status"] = ""
        winner.pop("_case_order_key", None)
        finalized.append(winner)
        keys_by_case[winner["case_id"]].add(set_key)

    finalized.sort(key=lambda row: (row["category"], row["id"], row["set_uid"]))
    return finalized, keys_by_case

# ---------------------------------------------------------------------------
# Case summariser
# ---------------------------------------------------------------------------

def summarize_cases(
    cases_rows: list[dict[str, str]],
    set_indexes: dict[str, Any],
    today_kl: date,
) -> dict[str, Any]:
    uid_map       = set_indexes["uid_map"]
    shorthand_map = set_indexes["shorthand_map"]
    category_map  = set_indexes["category_map"]

    parsed_cases: list[dict[str, Any]] = []
    unknown_set_tokens: Counter[str] = Counter()

    for idx, row in enumerate(cases_rows, start=1):
        delivery_date = parse_date(row_value(row, "delivery_date"))
        surgery_date = parse_date(row_value(row, "surgery_date"))
        prefix = row_value(row, "prefix").strip()
        is_booking_case = is_booking_prefix(prefix)
        booking_hold_from = booking_hold_start(delivery_date)
        booking_is_active = is_booking_case and is_booking_hold_active(delivery_date, today_kl)
        sets_raw = row_value(row, "sets").strip()
        sets_returned_raw = row_value(row, "sets_returned").strip()
        plate_tokens = split_plate_tokens(row_value(row, "plates"))
        powertools_raw = row_value(row, "powertools", "powertool").strip()
        powertool_tokens = split_powertool_tokens(powertools_raw)
        set_tokens = split_tokens(sets_raw)
        returned_set_tokens = split_tokens(sets_returned_raw)
        bonegraft_raw = row_value(row, "bonegraft").strip()
        extra_items_raw = row_value(row, "extra_items", "extra item", "extras").strip()
        bonegraft_tokens = split_tokens(bonegraft_raw)
        extra_item_tokens = split_extra_item_tokens(extra_items_raw)

        uid_hits: list[dict[str, Any]] = []
        shorthand_hits: list[dict[str, Any]] = []
        set_display_tokens: list[str] = []

        for token in set_tokens:
            norm = normalize_set_code(token)
            if not norm:
                continue

            uid_matches: list[dict[str, Any]] = []
            shorthand_matches: list[dict[str, Any]] = []
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
                    set_display_tokens.append("/".join(label for label in labels if label.strip()))

        uid_hit_by_key = {item["_set_key"]: item for item in uid_hits}
        shorthand_hit_keys = {
            item["_shorthand_norm"]
            for item in shorthand_hits
            if item.get("_shorthand_norm")
        }
        has_uid_set = bool(uid_hit_by_key)
        has_shorthand_only = bool(shorthand_hit_keys) and not has_uid_set
        set_categories = sorted({
            str(item.get("category", "")).strip()
            for item in list(uid_hit_by_key.values()) + shorthand_hits
            if str(item.get("category", "")).strip()
        })

        returned_hit_by_key: dict[str, dict[str, Any]] = {}
        returned_display_tokens: list[str] = []
        for token in returned_set_tokens:
            norm = normalize_set_code(token)
            if not norm:
                continue
            global_uid_matches = uid_map.get(norm, [])
            matched_items = [
                item for item in global_uid_matches
                if item["_set_key"] in uid_hit_by_key
            ]
            if matched_items:
                for item in matched_items:
                    returned_hit_by_key[item["_set_key"]] = item
                labels = sorted({
                    format_set_display(
                        item.get("shorthand") or item.get("category"),
                        item.get("id"),
                    )
                    for item in matched_items
                })
                if labels:
                    returned_display_tokens.append("/".join(labels))
            elif not global_uid_matches:
                unknown_set_tokens[norm] += 1
                returned_display_tokens.append(token)

        exact_set_rows = []
        for item in uid_hit_by_key.values():
            exact_set_rows.append({
                "set_key": item["_set_key"],
                "set_uid": item["uid"],
                "set_uid_norm": item["_uid_norm"],
                "category": item["category"],
                "id": item["id"],
                "home": item["home"],
                "set_status": item.get("status", ""),
                "set_display": format_set_display(
                    item.get("shorthand") or item.get("category"),
                    item.get("id"),
                ),
            })

        case_id = f"C{idx:03d}"
        hospital_code = normalize_code(row_value(row, "hospital"))

        case_record: dict[str, Any] = {
            "case_id":            case_id,
            "row_number":         idx,
            "prefix":             prefix,
            "hospital":           hospital_code,
            "patient_doctor":     row_value(row, "patient_doctor").strip(),
            "delivery_date":      format_date(delivery_date),
            "surgery_date":       format_date(surgery_date),
            "sales_code":         row_value(row, "sales_code").strip(),
            "return_date":        row_value(row, "return_date").strip(),
            "status":             row_value(row, "status").strip(),
            "smart_status":       row_value(row, "Smart Status").strip(),
            "sets_raw":           sets_raw,
            "sets_display":       "; ".join(token for token in set_display_tokens if token) or sets_raw,
            "sets_returned_raw":  sets_returned_raw,
            "sets_returned_display": (
                "; ".join(token for token in returned_display_tokens if token) or sets_returned_raw
            ),
            "sets_outstanding_raw": "",
            "sets_outstanding_display": "",
            "plates_raw":         row_value(row, "plates").strip(),
            "powertools_raw":     powertools_raw,
            "bonegraft_raw":      bonegraft_raw,
            "extra_items_raw":    extra_items_raw,
            "has_uid_set":        has_uid_set,
            "has_active_uid_set": False,
            "has_shorthand_only": has_shorthand_only,
            "is_booking_case":    is_booking_case,
            "booking_hold_active": booking_is_active,
            "booking_hold_start": format_date(booking_hold_from),
            "set_categories":     set_categories,
            "set_uid_tokens":     sorted({item["_uid_norm"] for item in uid_hit_by_key.values()}),
            "returned_set_uid_tokens": sorted({
                item["_uid_norm"] for item in returned_hit_by_key.values()
            }),
            "active_set_uid_tokens": [],
            "booked_set_uid_tokens": [],
            "set_shorthand_tokens": sorted(shorthand_hit_keys),
            "delivery_date_obj": delivery_date,
            "surgery_date_obj":  surgery_date,
            "set_tokens":        set_tokens,
            "plate_tokens":      plate_tokens,
            "powertool_tokens":  powertool_tokens,
            "bonegraft_tokens":  bonegraft_tokens,
            "extra_item_tokens": extra_item_tokens,
            "sent_set_items":    [],
            "_case_order_key":   case_order_key(idx, delivery_date, surgery_date),
            "_assigned_exact_sets": exact_set_rows,
            "_returned_set_keys":   set(returned_hit_by_key),
        }
        parsed_cases.append(case_record)

    out_candidates_by_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    booking_candidates_by_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for case in parsed_cases:
        days_since = (
            (today_kl - case["surgery_date_obj"]).days
            if case.get("surgery_date_obj") else None
        )
        for item in case["_assigned_exact_sets"]:
            if item["set_key"] in case["_returned_set_keys"]:
                continue
            if item.get("home") != "OFFICE":
                continue
            if case.get("is_booking_case") and not case.get("booking_hold_active"):
                continue
            target_map = (
                booking_candidates_by_key
                if case.get("booking_hold_active")
                else out_candidates_by_key
            )
            target_map[item["set_key"]].append({
                "set_key":            item["set_key"],
                "set_uid":            item["set_uid"],
                "category":           item["category"],
                "id":                 item["id"],
                "set_display":        item["set_display"],
                "home":               item["home"],
                "set_status":         item.get("set_status", ""),
                "assignment_kind":    "BOOKED" if case.get("booking_hold_active") else "OUT",
                "location_now":       case["hospital"] or "UNKNOWN",
                "case_id":            case["case_id"],
                "delivery_date":      case["delivery_date"],
                "surgery_date":       case["surgery_date"],
                "days_since_surgery": days_since,
                "case_status":        case["status"],
                "smart_status":       case["smart_status"],
                "patient_doctor":     case["patient_doctor"],
                "_case_order_key":    case["_case_order_key"],
            })

    set_out_assignments, active_office_keys_by_case = finalize_set_assignments(
        out_candidates_by_key,
        repeated_same_hospital_status="CARRY_OVER",
    )
    set_booking_assignments, active_booking_keys_by_case = finalize_set_assignments(
        booking_candidates_by_key,
        repeated_same_hospital_status="BOOKED_OVERLAP",
    )

    for case in parsed_cases:
        active_exact_sets = []
        booked_exact_sets = []
        for item in case["_assigned_exact_sets"]:
            if item["set_key"] in case["_returned_set_keys"]:
                continue
            if item.get("home") == "OFFICE":
                if case.get("booking_hold_active"):
                    if item["set_key"] in active_booking_keys_by_case.get(case["case_id"], set()):
                        booked_exact_sets.append(item)
                    continue
                if item["set_key"] not in active_office_keys_by_case.get(case["case_id"], set()):
                    continue
            active_exact_sets.append(item)
        active_exact_sets.sort(key=lambda item: (item["category"], item["id"], item["set_uid"]))
        booked_exact_sets.sort(key=lambda item: (item["category"], item["id"], item["set_uid"]))
        case["active_set_uid_tokens"] = sorted({
            item["set_uid_norm"] for item in active_exact_sets
        })
        case["booked_set_uid_tokens"] = sorted({
            item["set_uid_norm"] for item in booked_exact_sets
        })
        case["has_active_uid_set"] = bool(case["active_set_uid_tokens"])
        if active_exact_sets:
            case["sets_outstanding_raw"] = ";".join(item["set_uid"] for item in active_exact_sets)
            case["sets_outstanding_display"] = "; ".join(
                item["set_display"] for item in active_exact_sets
            )
        elif case["has_shorthand_only"]:
            case["sets_outstanding_raw"] = case["sets_raw"]
            case["sets_outstanding_display"] = case["sets_display"]
        active_keys = {item["set_key"] for item in active_exact_sets}
        booked_keys = {item["set_key"] for item in booked_exact_sets}
        sent_set_items = []
        for item in case.get("_assigned_exact_sets", []):
            assignment_kind = ""
            if item["set_key"] in booked_keys:
                assignment_kind = "BOOKED"
            elif item["set_key"] in active_keys:
                assignment_kind = "OUT"
            sent_set_items.append({
                "category": item["category"],
                "id": item["id"],
                "assignment_kind": assignment_kind,
                "label": format_case_item_label(item["category"], item["id"]),
            })
        case["sent_set_items"] = sorted(
            sent_set_items,
            key=lambda item: (
                item.get("assignment_kind") == "",
                item.get("assignment_kind") == "BOOKED",
                str(item.get("category", "")),
                compact_set_id(item.get("id", "")),
            ),
        )
        case.pop("_case_order_key", None)
        case.pop("_assigned_exact_sets", None)
        case.pop("_returned_set_keys", None)

    return {
        "parsed_cases":        parsed_cases,
        "set_out_assignments": set_out_assignments,
        "set_booking_assignments": set_booking_assignments,
        "unknown_set_tokens":  unknown_set_tokens,
    }

# ---------------------------------------------------------------------------
# Output builders
# ---------------------------------------------------------------------------

def build_set_outputs(
    set_indexes: dict[str, Any],
    set_out_assignments: list[dict[str, Any]],
    set_booking_assignments: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    office_sets = set_indexes["office_sets"]
    standby_sets = set_indexes.get("standby_sets", [])
    booking_assignments = set_booking_assignments or []

    # Exclude NA and STANDBY from availability totals.
    active_office_sets = [
        s for s in office_sets
        if not is_na_status(s.get("status", "")) and not is_standby_status(s.get("status", ""))
    ]

    total_by_category = Counter(item["category"] for item in active_office_sets)

    # Count physical sets once even if the same UID appears on multiple case rows.
    unique_out_by_key = {
        item["set_key"]: item
        for item in set_out_assignments
        if item.get("set_key")
        and not is_na_status(item.get("set_status", ""))
        and not is_standby_status(item.get("set_status", ""))
    }
    unique_booked_by_key = {
        item["set_key"]: item
        for item in booking_assignments
        if item.get("set_key")
        and not is_na_status(item.get("set_status", ""))
        and not is_standby_status(item.get("set_status", ""))
    }
    unique_reserved_by_key = dict(unique_booked_by_key)
    unique_reserved_by_key.update(unique_out_by_key)
    out_by_category = Counter(item["category"] for item in unique_out_by_key.values())
    booked_by_category = Counter(item["category"] for item in unique_booked_by_key.values())
    reserved_by_category = Counter(item["category"] for item in unique_reserved_by_key.values())
    booked_rows_by_category: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in unique_booked_by_key.values():
        booked_rows_by_category[item["category"]].append(item)

    set_category_availability: list[dict[str, Any]] = []
    for category in sorted(total_by_category):
        total     = total_by_category[category]
        out       = out_by_category.get(category, 0)
        booked    = booked_by_category.get(category, 0)
        reserved  = reserved_by_category.get(category, 0)
        available = max(total - reserved, 0)
        next_booking = None
        category_bookings = booked_rows_by_category.get(category, [])
        if category_bookings:
            next_booking = min(
                category_bookings,
                key=lambda item: (
                    parse_date(item.get("delivery_date")) or date.max,
                    parse_date(item.get("surgery_date")) or date.max,
                    str(item.get("case_id", "")),
                    str(item.get("id", "")),
                ),
            )
        set_category_availability.append({
            "category":    category,
            "total_office": total,
            "booked_for_case": booked,
            "out_for_case": out,
            "reserved_total": reserved,
            "available":    available,
            "availability": f"{available}/{total}",
            "next_booking_date": str(next_booking.get("delivery_date", "")) if next_booking else "",
            "next_booking_hospital": str(next_booking.get("location_now", "")) if next_booking else "",
            "next_booking_case_id": str(next_booking.get("case_id", "")) if next_booking else "",
            "next_booking_set": str(next_booking.get("set_display", "")) if next_booking else "",
        })

    # Per-set location status
    by_set_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in booking_assignments:
        by_set_key[row["set_key"]] = [row]
    for row in set_out_assignments:
        by_set_key[row["set_key"]] = [row]

    set_office_status: list[dict[str, Any]] = []
    visible_sets = office_sets + standby_sets
    for s in sorted(visible_sets, key=lambda x: (x["category"], x["id"], x["uid"])):
        key         = s["_set_key"]
        assignments = by_set_key.get(key, [])
        base_row    = {
            "set_key":           key,
            "category":          s["category"],
            "id":                s["id"],
            "set_display":       format_set_display(
                                     s.get("shorthand") or s.get("category"), s.get("id")
                                 ),
            "home":              s["home"],
            "set_status":        s.get("status", ""),
            "assignment_kind":   "",
        }
        if not assignments:
            in_standby = (
                str(s.get("home", "")).strip().upper() == "STANDBY"
                or is_standby_status(s.get("status", ""))
            )
            set_office_status.append({
                **base_row,
                "location_now":      "STANDBY" if in_standby else (
                    s["home"] if str(s.get("home", "")).strip().upper() in {"OFFICE", "STANDBY"} else "OFFICE"
                ),
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
                    "assignment_kind":   row.get("assignment_kind", ""),
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
                    "case_status":    case.get("status", ""),
                    "from_stock":     parsed["from_stock"],
                })

    plate_size_range_availability: list[dict[str, Any]] = []
    plate_out_cases: list[dict[str, Any]] = []
    plate_drawer_detail: list[dict[str, Any]] = []

    for (uid, size_range), bucket in sorted(
        size_buckets.items(),
        key=lambda item: (item[0][0], size_range_sort_key(item[0][1])),
    ):
        drawer_keys = sorted(bucket["drawer_size_map"].keys(), key=drawer_sort_key)
        stock_keys = sorted(bucket["stock_size_map"].keys(), key=drawer_sort_key)
        td = len(bucket["drawer_locations"])
        ts = len(bucket["stock_locations"])
        od        = bucket["out_drawer_units"]
        out_stock = bucket["out_stock_units"]
        total     = td + ts
        out       = od + out_stock
        available_units = max(total - out, 0)
        if available_units == total:
            range_status = "READY"
        elif available_units == 0:
            range_status = "OUT OF STOCK"
        else:
            range_status = "PARTIAL"

        out_case_details: list[dict[str, Any]] = []
        drawer_case_map: dict[str, list[dict[str, Any]]] = {drawer: [] for drawer in drawer_keys}
        stock_case_map: dict[str, list[dict[str, Any]]] = {stock: [] for stock in stock_keys}
        next_drawer_idx = 0
        next_stock_idx = 0
        for d in bucket["out_details"]:
            case_detail = {
                "case_id":      d["case_id"],
                "hospital":     d["hospital"],
                "surgery_date": d["surgery_date"],
                "case_status":  d.get("case_status", ""),
                "from_stock":   d["from_stock"],
            }
            out_case_details.append(case_detail)
            if d["from_stock"] or not drawer_keys:
                if stock_keys:
                    target_stock = stock_keys[min(next_stock_idx, len(stock_keys) - 1)]
                    stock_case_map[target_stock].append(case_detail)
                    if next_stock_idx < len(stock_keys) - 1:
                        next_stock_idx += 1
                continue
            target_drawer = drawer_keys[min(next_drawer_idx, len(drawer_keys) - 1)]
            drawer_case_map[target_drawer].append(case_detail)
            if next_drawer_idx < len(drawer_keys) - 1:
                next_drawer_idx += 1

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
            "out_stock_units":       out_stock,
            "out_units":             out,
            "available_drawer_units":max(td - od, 0),
            "available_stock_units": max(ts - out_stock, 0),
            "available_units":       available_units,
            "availability":          f"{available_units}/{total}",
            "range_status":          range_status,
            "out_case_details":      out_case_details,
        }
        plate_size_range_availability.append(row)

        for drawer in drawer_keys:
            # drawer_size_detail carries {label, size_range, no_stock} per chip
            raw_detail = sorted(
                bucket["drawer_size_detail"].get(drawer, []),
                key=lambda x: plate_detail_sort_key(uid, x),
            )
            detail = [
                {
                    "label": d["label"],
                    "size_range": d["size_range"],
                    "no_stock": d["no_stock"],
                }
                for d in raw_detail
            ]
            labels = [d["label"] for d in detail]
            plate_drawer_detail.append({
                "plate_uid":         uid,
                "proper_name":       bucket["plate_name"],
                "screw_sizes":       ", ".join(sorted(bucket["screw_sizes_set"])),
                "size_range":        size_range,
                "drawer":            drawer,
                "drawer_sizes":      ", ".join(labels),
                "drawer_size_detail": detail,         # [{label, size_range, no_stock}]
                "drawer_count":      len(labels),
                "available_units":   available_units,
                "total_units":       total,
                "availability":      f"{available_units}/{total}",
                "range_status":      range_status,
                "out_case_details":  out_case_details,
                "drawer_out_case_details": drawer_case_map.get(drawer, []),
            })
        for stock_loc in stock_keys:
            raw_detail = sorted(
                bucket["stock_size_detail"].get(stock_loc, []),
                key=lambda x: plate_detail_sort_key(uid, x),
            )
            detail = [
                {
                    "label": d["label"],
                    "size_range": d["size_range"],
                    "no_stock": d["no_stock"],
                }
                for d in raw_detail
            ]
            labels = [d["label"] for d in detail]
            plate_drawer_detail.append({
                "plate_uid":         uid,
                "proper_name":       bucket["plate_name"],
                "screw_sizes":       ", ".join(sorted(bucket["screw_sizes_set"])),
                "size_range":        size_range,
                "drawer":            stock_loc,
                "drawer_sizes":      ", ".join(labels),
                "drawer_size_detail": detail,
                "drawer_count":      len(labels),
                "available_units":   available_units,
                "total_units":       total,
                "availability":      f"{available_units}/{total}",
                "range_status":      range_status,
                "out_case_details":  out_case_details,
                "drawer_out_case_details": stock_case_map.get(stock_loc, []),
            })

        for detail in bucket["out_details"]:
            plate_out_cases.append({
                "plate_uid":       uid,
                "size_range":      size_range,
                "hospital":        detail["hospital"],
                "case_id":         detail["case_id"],
                "delivery_date":   detail["delivery_date"],
                "surgery_date":    detail["surgery_date"],
                "from_stock":      detail["from_stock"],
                "raw_plate_token": detail["raw_plate_token"],
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
            "aggregate_total_units":  0,
            "aggregate_out_units":    0,
            "aggregate_available_units": 0,
            "size_ranges":  [],
            "size_ranges_detail": [],
            "missing_ranges": [],
            "partial_ranges": [],
            "summary_basis": "ALL_RANGES",
            "_standard_row": None,
        })
        for ss in [p.strip() for p in str(row.get("screw_sizes", "")).split(",") if p.strip()]:
            s["screw_sizes_set"].add(ss)
        s["aggregate_total_units"]    += row["total_units"]
        s["aggregate_out_units"]      += row["out_units"]
        s["aggregate_available_units"]+= row["available_units"]
        s["size_ranges_detail"].append({
            "size_range": row["size_range"],
            "text": f"{row['size_range']}({row['available_units']}/{row['total_units']})",
        })
        if row["available_units"] == 0:
            s["missing_ranges"].append(row["size_range"])
        elif row["available_units"] < row["total_units"]:
            s["partial_ranges"].append(row["size_range"])
        if row["size_range"] == "STANDARD":
            s["_standard_row"] = row
            s["summary_basis"] = "STANDARD"

    plate_uid_summary = []
    for uid_key, s in sorted(uid_summary_map.items()):
        s["size_ranges"] = ", ".join(
            item["text"]
            for item in sorted(s["size_ranges_detail"], key=lambda item: size_range_sort_key(item["size_range"]))
        )
        s.pop("size_ranges_detail", None)
        s["screw_sizes"] = ", ".join(sorted(s["screw_sizes_set"]))
        s.pop("screw_sizes_set", None)
        standard_row = s.pop("_standard_row", None)
        if standard_row:
            s["total_units"] = standard_row["total_units"]
            s["out_units"] = standard_row["out_units"]
            s["available_units"] = standard_row["available_units"]
            s["availability"] = standard_row["availability"]
            extra_missing = [
                size_range for size_range in s["missing_ranges"]
                if size_range != "STANDARD"
            ]
            extra_partial = [
                size_range for size_range in s["partial_ranges"]
                if size_range != "STANDARD"
            ]
            if standard_row["available_units"] == 0:
                s["status_note"] = "OUT OF STOCK: STANDARD"
            elif standard_row["available_units"] < standard_row["total_units"]:
                s["status_note"] = "PARTIAL: STANDARD"
            elif extra_missing:
                s["status_note"] = "READY (extras missing: " + ", ".join(
                    sorted(extra_missing, key=size_range_sort_key)
                ) + ")"
            elif extra_partial:
                s["status_note"] = "READY (extras partial: " + ", ".join(
                    sorted(extra_partial, key=size_range_sort_key)
                ) + ")"
            else:
                s["status_note"] = "READY"
        else:
            s["total_units"] = s["aggregate_total_units"]
            s["out_units"] = s["aggregate_out_units"]
            s["available_units"] = s["aggregate_available_units"]
            s["availability"] = f"{s['available_units']}/{s['total_units']}"
            if s["missing_ranges"]:
                s["status_note"] = "OUT OF STOCK: " + ", ".join(
                    sorted(s["missing_ranges"], key=size_range_sort_key)
                )
            elif s["partial_ranges"]:
                s["status_note"] = "PARTIAL: " + ", ".join(
                    sorted(s["partial_ranges"], key=size_range_sort_key)
                )
            else:
                s["status_note"] = "READY"
        s.pop("aggregate_total_units", None)
        s.pop("aggregate_out_units", None)
        s.pop("aggregate_available_units", None)
        plate_uid_summary.append(s)

    return {
        "plate_size_range_availability": plate_size_range_availability,
        "plate_drawer_detail":           plate_drawer_detail,
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
        item["_is_standby_hold"] = is_standby_status(item.get("status", ""))
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
                        "_power_key":              item["_power_key"],
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

    unique_out_by_key = {
        a["_power_key"]: a
        for a in delivered_assignments
        if a.get("_power_key")
        and not a.get("is_na_hold")
        and not is_standby_status(a.get("status", ""))
    }
    out_by_cat = Counter(a["category"] for a in unique_out_by_key.values())
    total_by_cat = Counter(item["category"] for item in office_powertools)
    na_by_cat    = Counter(item["category"] for item in office_powertools if item.get("_is_na_hold"))
    standby_by_cat = Counter(item["category"] for item in office_powertools if item.get("_is_standby_hold"))

    powertool_category_availability: list[dict[str, Any]] = []
    for cat in sorted(total_by_cat):
        total    = total_by_cat[cat]
        na_hold  = na_by_cat.get(cat, 0)
        standby_hold = standby_by_cat.get(cat, 0)
        usable   = max(total - na_hold - standby_hold, 0)
        out      = out_by_cat.get(cat, 0)
        available= max(usable - out, 0)
        powertool_category_availability.append({
            "category":    cat,
            "total_office": total,
            "na_hold":     na_hold,
            "standby_hold":standby_hold,
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
            "is_standby_hold": item.get("_is_standby_hold", False),
        }
        if not rows:
            powertool_uid_availability.append({
                **base,
                "availability":(
                    "NA_HOLD" if item.get("_is_na_hold")
                    else "STANDBY_HOLD" if item.get("_is_standby_hold")
                    else "AVAILABLE"
                ),
                "hospital":"", "case_id":"", "delivery_date":"", "surgery_date":"",
            })
        else:
            for a in rows:
                powertool_uid_availability.append({
                    **base,
                    "availability":(
                        "OUT_NA_HOLD" if item.get("_is_na_hold")
                        else "OUT_STANDBY_HOLD" if item.get("_is_standby_hold")
                        else "OUT"
                    ),
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
        for token in split_powertool_tokens(row.get("powertool") or row.get("powertools") or ""):
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


def build_bonegraft_index(master_bonegraft: list[dict[str, Any]]) -> dict[str, dict[str, dict[str, str]]]:
    by_ref: dict[str, dict[str, str]] = {}
    by_alias: dict[str, dict[str, str]] = {}
    for item in master_bonegraft:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        presentation = str(item.get("presentation", "")).strip()
        ref = str(item.get("ref", "")).strip()
        label = format_bonegraft_label(name, presentation)
        payload = {
            "name": name,
            "presentation": presentation,
            "label": label,
            "ref": ref,
        }
        ref_norm = normalize_bonegraft_token(ref)
        if ref_norm:
            by_ref[ref_norm] = payload
        aliases: set[str] = set()
        for raw in (item.get("shorthand", ""), presentation):
            norm = normalize_bonegraft_token(raw)
            if norm:
                aliases.add(norm)
        shorthand_norm = normalize_bonegraft_token(item.get("shorthand", ""))
        if shorthand_norm.endswith("5") and len(shorthand_norm) > 2:
            aliases.add(shorthand_norm[:-1])
        presentation_token = normalize_bonegraft_token(presentation)
        if presentation.lower().endswith("cc"):
            bare_match = re.match(r"^([0-9.]+)\s*cc$", presentation.strip(), flags=re.IGNORECASE)
            if bare_match:
                bare_value = normalize_bonegraft_token(bare_match.group(1))
                if bare_value:
                    aliases.add(bare_value)
        if presentation_token:
            aliases.add(presentation_token)
        for alias in aliases:
            by_alias.setdefault(alias, payload)
    return {"by_ref": by_ref, "by_alias": by_alias}


def build_case_sent_item_details(
    parsed_cases: list[dict[str, Any]],
    plate_inventory: dict[str, Any],
    powertool_outputs: dict[str, Any],
    master_bonegraft: list[dict[str, Any]] | None = None,
) -> dict[str, dict[str, Any]]:
    size_buckets = plate_inventory["size_buckets"]
    uid_alias_map = plate_inventory["uid_alias_map"]
    plate_meta_by_uid: dict[str, dict[str, Any]] = {}
    for (uid, size_range), bucket in size_buckets.items():
        meta = plate_meta_by_uid.setdefault(uid, {
            "proper_name": str(bucket.get("plate_name", "")).strip(),
            "ranges": set(),
        })
        meta["ranges"].add(size_range)
        if not meta["proper_name"]:
            meta["proper_name"] = str(bucket.get("proper_name", "")).strip()

    powertool_items_by_case: dict[str, dict[str, dict[str, Any]]] = defaultdict(dict)
    resolved_powertool_tokens_by_case: dict[str, set[str]] = defaultdict(set)
    for item in powertool_outputs.get("powertool_delivered", []):
        case_id = str(item.get("case_id", "")).strip()
        if not case_id:
            continue
        key = str(item.get("_power_key", "")).strip() or f"{item.get('category', '')}|{item.get('id', '')}"
        label = format_case_item_label(item.get("category", ""), item.get("id", ""))
        powertool_items_by_case[case_id][key] = {
            "category": str(item.get("category", "")).strip(),
            "id": str(item.get("id", "")).strip(),
            "label": label,
            "is_resolved": True,
        }
        resolved_powertool_tokens_by_case[case_id].add(
            canonical_powertool_uid(item.get("powertool_uid", ""))
        )
        resolved_powertool_tokens_by_case[case_id].add(
            canonical_powertool_shorthand(item.get("category", ""))
        )

    bonegraft_index = build_bonegraft_index(master_bonegraft or [])
    case_items: dict[str, dict[str, Any]] = {}
    for case in parsed_cases:
        case_id = str(case.get("case_id", "")).strip()
        if not case_id:
            continue

        sent_sets = [
            dict(item)
            for item in case.get("sent_set_items", [])
            if isinstance(item, dict)
        ]

        plate_items_by_key: dict[str, dict[str, Any]] = {}
        for token in case.get("plate_tokens", []):
            parsed = parse_plate_request(str(token))
            if not parsed:
                continue
            raw_token = str(parsed.get("raw_token", "")).strip()
            resolved_uid = uid_alias_map.get(parsed["base_uid"], parsed["base_uid"])
            meta = plate_meta_by_uid.get(resolved_uid)
            if meta:
                item = plate_items_by_key.setdefault(resolved_uid, {
                    "plate_uid": resolved_uid,
                    "proper_name": meta["proper_name"] or raw_token,
                    "size_ranges": set(),
                    "is_resolved": True,
                })
                resolved_ranges = [
                    size_range
                    for size_range in parsed["needed_ranges"]
                    if (resolved_uid, size_range) in size_buckets
                ]
                item["size_ranges"].update(resolved_ranges or parsed["needed_ranges"])
                continue
            fallback_key = normalize_plate_code(raw_token) or raw_token
            plate_items_by_key.setdefault(fallback_key, {
                "plate_uid": parsed["base_uid"],
                "proper_name": raw_token,
                "size_ranges": set(),
                "is_resolved": False,
            })

        sent_plates: list[dict[str, Any]] = []
        for item in plate_items_by_key.values():
            size_ranges = sorted(item["size_ranges"], key=size_range_sort_key)
            size_ranges_label = ", ".join(size_ranges)
            label = item["plate_uid"] if item.get("is_resolved") else item["proper_name"]
            if size_ranges_label:
                label = f"{label} ({size_ranges_label})"
            sent_plates.append({
                "plate_uid": item["plate_uid"],
                "proper_name": item["proper_name"],
                "size_ranges": size_ranges,
                "size_ranges_label": size_ranges_label,
                "label": label,
                "is_resolved": item["is_resolved"],
            })
        sent_plates.sort(key=lambda item: (str(item.get("plate_uid", "")), str(item.get("size_ranges_label", ""))))

        powertool_items = dict(powertool_items_by_case.get(case_id, {}))
        resolved_powertool_tokens = resolved_powertool_tokens_by_case.get(case_id, set())
        for token in case.get("powertool_tokens", []):
            raw_token = str(token).strip()
            if not raw_token:
                continue
            canonical_uid = canonical_powertool_uid(raw_token)
            canonical_short = canonical_powertool_shorthand(raw_token)
            if canonical_uid in resolved_powertool_tokens or canonical_short in resolved_powertool_tokens:
                continue
            fallback_key = normalize_code(raw_token)
            if not fallback_key:
                continue
            powertool_items.setdefault(fallback_key, {
                "category": raw_token,
                "id": "",
                "label": raw_token,
                "is_resolved": False,
            })
        sent_powertools = sorted(
            powertool_items.values(),
            key=lambda item: (not item.get("is_resolved", False), str(item.get("category", "")), str(item.get("id", ""))),
        )

        bonegraft_items_by_key: dict[str, dict[str, Any]] = {}
        for token in case.get("bonegraft_tokens", []):
            raw_token = str(token).strip()
            if not raw_token:
                continue
            norm = normalize_bonegraft_token(raw_token)
            resolved = bonegraft_index["by_ref"].get(norm) or bonegraft_index["by_alias"].get(norm)
            if resolved:
                bonegraft_items_by_key.setdefault(resolved["label"], {
                    **resolved,
                    "raw_token": raw_token,
                    "is_resolved": True,
                })
                continue
            bonegraft_items_by_key.setdefault(raw_token, {
                "name": raw_token,
                "presentation": "",
                "label": raw_token,
                "raw_token": raw_token,
                "is_resolved": False,
            })
        sent_bonegraft = sorted(
            bonegraft_items_by_key.values(),
            key=lambda item: (not item.get("is_resolved", False), str(item.get("name", "")), str(item.get("presentation", "")), str(item.get("label", ""))),
        )

        sent_extra_items = [
            {"label": str(token).strip()}
            for token in case.get("extra_item_tokens", [])
            if str(token).strip()
        ]

        case_items[case_id] = {
            "sent_sets": sent_sets,
            "sent_plates": sent_plates,
            "sent_powertools": sent_powertools,
            "sent_bonegraft": sent_bonegraft,
            "sent_extra_items": sent_extra_items,
        }

    return case_items


def is_cancelled_case(case: dict[str, Any]) -> bool:
    status_token = normalize_code(case.get("status", ""))
    smart_status_token = normalize_code(case.get("smart_status", ""))
    prefix_token = normalize_code(case.get("prefix", ""))
    cancel_tokens = [status_token, smart_status_token, prefix_token]
    cancel_markers = ("CANCEL", "CANCELLED", "CANCELED", "CXL", "CNCL", "CNX")
    if any(any(marker in token for marker in cancel_markers) for token in cancel_tokens if token):
        return True

    has_sales = bool(normalize_code(case.get("sales_code", "")))
    if has_sales:
        return False

    return False


def infer_set_reusable_date(case: dict[str, Any]) -> date | None:
    return (
        parse_date(case.get("return_date"))
        or parse_date(case.get("surgery_date"))
        or parse_date(case.get("delivery_date"))
    )


def build_set_suggestion_summary(suggestion: dict[str, Any]) -> str:
    parts = [f"{suggestion['category']}: {suggestion['set_display']}"]
    reason = str(suggestion.get("suggestion_reason", "")).strip()
    if reason:
        parts.append(reason)
    if not suggestion.get("confirmed", False):
        parts.append("not confirmed")
    return " | ".join(parts)


def attach_upcoming_set_suggestions(
    parsed_cases: list[dict[str, Any]],
    set_indexes: dict[str, Any],
    set_out_assignments: list[dict[str, Any]],
    set_booking_assignments: list[dict[str, Any]] | None,
    today_kl: date,
) -> None:
    case_by_id = {
        str(case.get("case_id", "")).strip(): case
        for case in parsed_cases
        if str(case.get("case_id", "")).strip()
    }
    assignments_by_key: dict[str, dict[str, Any]] = {}
    for row in set_booking_assignments or []:
        key = str(row.get("set_key", "")).strip()
        if key:
            assignments_by_key[key] = dict(row, assignment_kind=row.get("assignment_kind") or "BOOKED")
    for row in set_out_assignments:
        key = str(row.get("set_key", "")).strip()
        if key:
            assignments_by_key[key] = dict(row, assignment_kind=row.get("assignment_kind") or "OUT")

    candidate_sets_by_category: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in set_indexes["office_sets"] + set_indexes.get("standby_sets", []):
        if is_na_status(item.get("status", "")):
            continue
        category = str(item.get("category", "")).strip()
        if not category:
            continue
        candidate_sets_by_category[category].append(item)

    suggested_counts: Counter[str] = Counter()
    ordered_cases = sorted(
        (
            case for case in parsed_cases
            if case.get("has_shorthand_only")
            and not is_cancelled_case(case)
            and case.get("delivery_date_obj")
            and case["delivery_date_obj"] >= today_kl
        ),
        key=lambda case: (
            case.get("delivery_date_obj") or date.max,
            case.get("surgery_date_obj") or date.max,
            str(case.get("case_id", "")),
        ),
    )

    for case in parsed_cases:
        case["suggested_sets"] = []
        case["suggested_sets_summary"] = ""

    for case in ordered_cases:
        target_delivery = case.get("delivery_date_obj") or date.max
        suggestions: list[dict[str, Any]] = []

        for category in case.get("set_categories", []):
            candidates = candidate_sets_by_category.get(str(category).strip(), [])
            if not candidates:
                continue

            ranked_candidates: list[tuple[tuple[Any, ...], dict[str, Any]]] = []
            for set_row in candidates:
                set_key = str(set_row.get("_set_key", "")).strip()
                assignment = assignments_by_key.get(set_key)
                source_case = case_by_id.get(str(assignment.get("case_id", "")).strip(), {}) if assignment else {}
                current_state = "OFFICE"
                current_location = "OFFICE"
                if assignment:
                    current_state = str(assignment.get("assignment_kind", "")).strip().upper() or "OUT"
                    current_location = str(assignment.get("location_now", "")).strip() or current_state
                elif (
                    str(set_row.get("home", "")).strip().upper() == "STANDBY"
                    or is_standby_status(set_row.get("status", ""))
                ):
                    current_state = "STANDBY"
                    current_location = "STANDBY"

                source_delivery = parse_date(assignment.get("delivery_date")) if assignment else None
                source_surgery = parse_date(assignment.get("surgery_date")) if assignment else None
                source_return = parse_date(source_case.get("return_date", "")) if source_case else None
                source_sales_code = str(source_case.get("sales_code", "")).strip()
                source_cancelled = bool(source_case) and is_cancelled_case(source_case)
                reusable_date = infer_set_reusable_date(source_case) if source_case else (
                    source_surgery or source_delivery
                )
                is_reused_candidate = suggested_counts.get(set_key, 0) > 0

                if current_state == "OFFICE":
                    suggestion_kind = "IN_OFFICE"
                    suggestion_rank = 0
                    suggestion_reason = (
                        "only set in office"
                        if len(candidates) == 1 else
                        "next available in office"
                    )
                    confirmed = not is_reused_candidate
                elif current_state == "STANDBY":
                    suggestion_kind = "IN_STANDBY"
                    suggestion_rank = 1
                    suggestion_reason = "available in standby"
                    confirmed = False
                elif current_state == "BOOKED":
                    ready_by_target = reusable_date is not None and reusable_date <= target_delivery
                    suggestion_kind = "BOOKED_READY" if ready_by_target else "BOOKED_LATER"
                    suggestion_rank = 3 if ready_by_target else 5
                    if source_delivery and source_delivery <= target_delivery:
                        suggestion_reason = (
                            f"booked for {source_case.get('case_id', assignment.get('case_id', ''))}"
                            f" on {format_date(source_delivery)}"
                        )
                    else:
                        suggestion_reason = (
                            f"booked for {source_case.get('case_id', assignment.get('case_id', ''))}"
                            if source_case or assignment else
                            "booked for another case"
                        )
                    confirmed = False
                else:
                    ready_by_target = reusable_date is not None and reusable_date <= target_delivery
                    if source_cancelled:
                        suggestion_kind = "OUT_CANCELLED"
                        suggestion_rank = 2
                        suggestion_reason = (
                            f"currently at {current_location}; case cancelled"
                        )
                    elif source_return:
                        suggestion_kind = "OUT_RETURNED"
                        suggestion_rank = 2
                        suggestion_reason = (
                            f"currently at {current_location}; return dated {format_date(source_return)}"
                        )
                    elif source_sales_code:
                        suggestion_kind = "OUT_SALES_POSTED"
                        suggestion_rank = 2
                        suggestion_reason = (
                            f"currently at {current_location}; sales posted on {source_case.get('case_id', assignment.get('case_id', ''))}"
                        )
                    elif ready_by_target and source_surgery:
                        suggestion_kind = "OUT_SURGERY_DONE"
                        suggestion_rank = 2
                        suggestion_reason = (
                            f"currently at {current_location}; surgery done {format_date(source_surgery)}"
                        )
                    elif ready_by_target and source_delivery:
                        suggestion_kind = "OUT_DELIVERED"
                        suggestion_rank = 3
                        suggestion_reason = (
                            f"currently at {current_location}; delivered {format_date(source_delivery)}"
                        )
                    else:
                        suggestion_kind = "OUT_WAITING"
                        suggestion_rank = 4
                        if source_surgery:
                            suggestion_reason = (
                                f"currently at {current_location}; surgery {format_date(source_surgery)}"
                            )
                        elif source_delivery:
                            suggestion_reason = (
                                f"currently at {current_location}; delivered {format_date(source_delivery)}"
                            )
                        else:
                            suggestion_reason = f"currently at {current_location}"
                    confirmed = False

                if is_reused_candidate:
                    suggestion_reason += "; already suggested earlier"

                rank_key = (
                    suggestion_rank,
                    1 if is_reused_candidate else 0,
                    reusable_date or date.max,
                    compact_set_id(set_row.get("id", "")),
                    str(set_row.get("uid", "")),
                )
                ranked_candidates.append((rank_key, {
                    "category":             str(category).strip(),
                    "set_key":              set_key,
                    "set_uid":              str(set_row.get("uid", "")).strip(),
                    "set_id":               str(set_row.get("id", "")).strip(),
                    "set_display":          format_set_display(
                        set_row.get("shorthand") or set_row.get("category"),
                        set_row.get("id"),
                    ),
                    "current_state":        current_state,
                    "current_location":     current_location,
                    "suggestion_kind":      suggestion_kind,
                    "suggestion_reason":    suggestion_reason,
                    "confirmed":            confirmed,
                    "source_case_id":       str(source_case.get("case_id", assignment.get("case_id", "")) if source_case or assignment else "").strip(),
                    "source_hospital":      str(source_case.get("hospital", assignment.get("location_now", "")) if source_case or assignment else "").strip(),
                    "source_delivery_date": format_date(source_delivery),
                    "source_surgery_date":  format_date(source_surgery),
                    "source_sales_code":    source_sales_code,
                    "source_return_date":   format_date(source_return),
                    "source_case_status":   str(source_case.get("status", assignment.get("case_status", "")) if source_case or assignment else "").strip(),
                    "source_case_cancelled": source_cancelled,
                }))

            if not ranked_candidates:
                continue

            _, suggestion = min(ranked_candidates, key=lambda item: item[0])
            suggestion["summary"] = build_set_suggestion_summary(suggestion)
            suggestions.append(suggestion)
            suggested_counts[suggestion["set_key"]] += 1

        case["suggested_sets"] = suggestions
        case["suggested_sets_summary"] = "; ".join(
            item["summary"] for item in suggestions if item.get("summary")
        )


def build_case_region_summary(
    parsed_cases: list[dict[str, Any]],
    hospitals: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    summary_by_region: dict[str, dict[str, Any]] = {}
    case_region_meta: dict[str, dict[str, Any]] = {}

    for case in parsed_cases:
        raw_hospital = case.get("hospital", "")
        resolved_code, resolved_by = resolve_hospital_code(raw_hospital, hospitals)
        hospital_meta = hospitals.get(resolved_code, {}) if resolved_code else {}
        region = str(hospital_meta.get("region", "")).strip() or "Unknown"
        hospital_name = str(hospital_meta.get("name", "")).strip()
        sales_recorded = bool(normalize_code(case.get("sales_code", "")))
        cancelled = is_cancelled_case(case)

        bucket = summary_by_region.setdefault(region, {
            "region": region,
            "total_cases": 0,
            "cancelled_cases": 0,
            "sales_cases": 0,
        })
        bucket["total_cases"] += 1
        if cancelled:
            bucket["cancelled_cases"] += 1
        if sales_recorded:
            bucket["sales_cases"] += 1

        case_id = str(case.get("case_id", "")).strip()
        if case_id:
            case_region_meta[case_id] = {
                "region": region,
                "resolved_hospital": resolved_code or "",
                "resolved_by": resolved_by,
                "hospital_name": hospital_name,
                "is_cancelled_case": cancelled,
                "has_sales": sales_recorded,
            }

    summary_rows: list[dict[str, Any]] = []
    for row in sorted(summary_by_region.values(), key=lambda item: (-item["total_cases"], item["region"])):
        total = int(row["total_cases"])
        sales = int(row["sales_cases"])
        summary_rows.append({
            **row,
            "sales_total_cases": f"{sales}/{total}" if total else "0/0",
        })

    return summary_rows, case_region_meta


def build_archive_30d_summary(
    archive_rows: list[dict[str, str]],
    hospitals: dict[str, dict[str, Any]],
    set_indexes: dict[str, Any],
    today_kl: date,
) -> dict[str, Any]:
    uid_map = set_indexes["uid_map"]
    window_start = today_kl - timedelta(days=30)
    cases_by_region: Counter[str] = Counter()
    cancelled_by_region: Counter[str] = Counter()
    total_cases = 0
    total_cancelled = 0
    sets_delivered = 0

    for row in archive_rows:
        row_date = parse_date(row_value(row, "surgery_date")) or parse_date(row_value(row, "delivery_date"))
        if row_date is None or not (window_start <= row_date <= today_kl):
            continue
        raw_hospital = row_value(row, "hospital")
        resolved_code, _ = resolve_hospital_code(raw_hospital, hospitals)
        hospital_meta = hospitals.get(resolved_code, {}) if resolved_code else {}
        region = str(hospital_meta.get("region", "")).strip() or "Unknown"
        cancelled = is_cancelled_case({
            "status": row_value(row, "status"),
            "smart_status": row_value(row, "Smart Status"),
            "prefix": row_value(row, "prefix"),
            "sales_code": row_value(row, "sales_code"),
        })
        total_cases += 1
        cases_by_region[region] += 1
        if cancelled:
            total_cancelled += 1
            cancelled_by_region[region] += 1
        matched_set_count = 0
        for token in split_tokens(row_value(row, "sets")):
            uid_norm = normalize_set_code(token)
            if uid_norm not in uid_map:
                continue
            matched_set_count += 1
        sets_delivered += matched_set_count

    return {
        "window_start": format_date(window_start),
        "window_end": format_date(today_kl),
        "total_cases_30d": total_cases,
        "total_cancelled_cases_30d": total_cancelled,
        "sets_delivered_30d": sets_delivered,
        "cases_by_region_30d": [
            {"region": region, "cases": count}
            for region, count in sorted(cases_by_region.items(), key=lambda item: (-item[1], item[0]))
        ],
        "cancelled_by_region_30d": [
            {"region": region, "cancelled_cases": count}
            for region, count in sorted(cancelled_by_region.items(), key=lambda item: (-item[1], item[0]))
        ],
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
        cancelled     = is_cancelled_case(case)

        # Sales code recorded but equipment not yet returned
        if sales_code and not return_date:
            buckets["to_collect"].append(case)

        if cancelled:
            continue

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
        if delivery_date == today_kl and case["has_uid_set"] and not case.get("is_booking_case"):
            buckets["delivered_today"].append(case)

    def _strip(case: dict[str, Any]) -> dict[str, Any]:
        return {
            **serialize_case_common(case),
            "plates":     case["plates_raw"],
            "powertools": case["powertools_raw"],
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
        base_row = {
            "route_group":           route_tag,
            "case_id":               case["case_id"],
            "hospital":              raw_code,
            "resolved_by":           resolved_by,
            "delivery_date":         case["delivery_date"],
            "surgery_date":          case["surgery_date"],
            "sets":                  case.get("sets", ""),
            "plates":                case.get("plates", ""),
            "sales_code":            case.get("sales_code", ""),
        }

        if resolved_code is None:
            unresolved[raw_code] += 1
            rows.append({
                **base_row,
                "resolved_hospital":    "",
                "hospital_name":        "Unknown hospital code",
                "office_to_hospital_km":"",
                "office_est_drive_km":  "",
                "office_est_drive_min": "",
                "tbs_to_hospital_km":   "",
            })
            continue

        meta = hospitals[resolved_code]
        lat, lng = float(meta["lat"]), float(meta["lng"])
        straight  = haversine_km(OFFICE_COORDS[0], OFFICE_COORDS[1], lat, lng)
        tbs_dist  = haversine_km(TBS_COORDS[0], TBS_COORDS[1], lat, lng)
        drive_km  = estimated_drive_km(straight)
        drive_min = estimated_drive_minutes(drive_km)

        rows.append({
            **base_row,
            "resolved_hospital":    resolved_code,
            "hospital_name":        meta.get("name", ""),
            "office_to_hospital_km":round(straight, 2),
            "office_est_drive_km":  round(drive_km, 2),
            "office_est_drive_min": round(drive_min, 1),
            "tbs_to_hospital_km":   round(tbs_dist, 2),
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
    master_data_path: str | Path = DEFAULT_MASTER_DATA_PATH,
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
    set_outputs   = build_set_outputs(
        set_indexes,
        case_summary["set_out_assignments"],
        case_summary.get("set_booking_assignments", []),
    )
    attach_upcoming_set_suggestions(
        case_summary["parsed_cases"],
        set_indexes,
        case_summary["set_out_assignments"],
        case_summary.get("set_booking_assignments", []),
        today,
    )

    plate_inventory = build_plate_inventory(master["PLATES"])
    plate_outputs   = build_plate_outputs(case_summary["parsed_cases"], plate_inventory)

    powertool_outputs = build_powertool_outputs(
        case_summary["parsed_cases"], archive_rows, set_indexes, today
    )
    case_sent_items = build_case_sent_item_details(
        case_summary["parsed_cases"],
        plate_inventory,
        powertool_outputs,
        master.get("BONEGRAFT", []),
    )
    case_region_summary, case_region_meta = build_case_region_summary(
        case_summary["parsed_cases"],
        master["HOSPITALS"],
    )
    archive_30d_summary = build_archive_30d_summary(
        archive_rows,
        master["HOSPITALS"],
        set_indexes,
        today,
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

    cases_all = [
        {
            **serialize_case_common(c),
            "has_uid_set":             c["has_uid_set"],
            "has_active_uid_set":      c.get("has_active_uid_set", False),
            "has_shorthand_only":      c["has_shorthand_only"],
            "is_booking_case":         c.get("is_booking_case", False),
            "booking_hold_active":     c.get("booking_hold_active", False),
            "booking_hold_start":      c.get("booking_hold_start", ""),
            "returned_set_uid_tokens": c.get("returned_set_uid_tokens", []),
            "active_set_uid_tokens":   c.get("active_set_uid_tokens", []),
            "booked_set_uid_tokens":   c.get("booked_set_uid_tokens", []),
            **case_region_meta.get(c["case_id"], {
                "region": "Unknown",
                "resolved_hospital": "",
                "resolved_by": "unresolved",
                "hospital_name": "",
                "is_cancelled_case": False,
                "has_sales": False,
            }),
            **case_sent_items.get(c["case_id"], {
                "sent_sets": [],
                "sent_plates": [],
                "sent_powertools": [],
                "sent_bonegraft": [],
                "sent_extra_items": [],
            }),
        }
        for c in case_summary["parsed_cases"]
    ]

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
        "plate_drawer_detail":            plate_outputs["plate_drawer_detail"],
        "plate_uid_summary":              plate_outputs["plate_uid_summary"],
        "plate_out_cases":                plate_outputs["plate_out_cases"],
        "powertool_category_availability":powertool_outputs["powertool_category_availability"],
        "powertool_uid_availability":     powertool_outputs["powertool_uid_availability"],
        "powertool_delivered":            powertool_outputs["powertool_delivered"],
        "powertool_usage_30d":            powertool_outputs["powertool_usage_30d"],
        "hospital_directory":             hospital_directory,
        "case_region_summary":            case_region_summary,
        "archive_30d_summary":           archive_30d_summary,
        "cases_all":                      cases_all,
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
        "plate_drawer_detail.csv":            report["plate_drawer_detail"],
        "plate_uid_summary.csv":              report["plate_uid_summary"],
        "plate_out_cases.csv":                report["plate_out_cases"],
        "powertool_category_availability.csv":report["powertool_category_availability"],
        "powertool_uid_availability.csv":     report["powertool_uid_availability"],
        "powertool_delivered.csv":            report["powertool_delivered"],
        "powertool_usage_30d.csv":            report["powertool_usage_30d"],
        "hospital_directory.csv":             report["hospital_directory"],
        "cases_all.csv":                      report.get("cases_all", []),
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
    parser.add_argument("--master-data", default=str(DEFAULT_MASTER_DATA_PATH))
    parser.add_argument("--cases",   default=None, help="Cases CSV path or URL")
    parser.add_argument("--archive", default=None, help="Archive CSV path or URL")
    parser.add_argument("--out",     default=str(DEFAULT_OUTPUT_DIR), help="Output directory")
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
