import argparse
import importlib.util
import random
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd


def load_master_sets(master_path: str):
    spec = importlib.util.spec_from_file_location("master_data", master_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    sets = getattr(module, "SETS", None)
    if sets is None:
        raise ValueError("master_data.py must define SETS = [...]")
    return sets


def bucket_from_set(row: dict) -> str:
    """Map a master_data set row into a rough simulation bucket.
    This is intentionally coarse. It is a starting model, not a full set-family map.
    """
    cat = str(row.get("category", "")).upper()
    short = str(row.get("shorthand", "")).upper()

    if "PFN II 170-240" in cat:
        return "PFN II"
    if "PFN II 340-420 SYSTEM" in cat or "PFN II 340-420 IMPLANT" in cat:
        return "PFN II"
    if cat == "PFN" or short == "PFN":
        return "PFN"
    if "LONG PFN" in cat or short == "LPFN":
        return "PFN"
    if "ILN FEMUR" in cat or short == "ILNF":
        return "FEMORAL NAIL"
    if "ILN TIBIA" in cat or short in {"ILNT", "SUPSUB"}:
        return "TIBIAL NAIL"
    if "ILN HUMERUS" in cat or short == "ILNH":
        return "HUMERAL NAIL"
    if "ILN RADIUS ULNA" in cat or short == "ILNRU":
        return "ULNA/RADIUS NAIL"
    if "FIBULAR NAIL" in cat or short == "FIBN":
        return "FIBULA NAIL"
    if "FNS" in cat or short == "FNS":
        return "FNS"
    if "RFN" in cat or short == "RFN":
        return "DISTAL FEMORAL NAIL"
    if short == "TENS" or "TENS" in cat:
        return "ELASTIC NAIL"
    if short.startswith("C") or "CANNA" in cat:
        return "CANNULATED"
    if short == "REAM" or "REAMER" in cat:
        return "REAMER"
    if short in {"1.5", "2.0", "2.4", "2.7", "3.5", "P5503", "P5400", "P8400", "FOOT", "ROI", "COATL"}:
        return "PLATE"
    return short or cat or "OTHER"


FAMILY_RULES = [
    ("Proximal Femoral Nail II", "PFN II"),
    ("Distal Femoral Nail II", "DISTAL FEMORAL NAIL"),
    ("Proximal Femoral Nail", "PFN"),
    ("Femoral Nail", "FEMORAL NAIL"),
    ("Tibial Nail", "TIBIAL NAIL"),
    ("Humeral Nail", "HUMERAL NAIL"),
    ("Fibula Nail", "FIBULA NAIL"),
    ("Ulna Nail", "ULNA/RADIUS NAIL"),
    ("Radius Nail", "ULNA/RADIUS NAIL"),
    ("Elastic Nail", "ELASTIC NAIL"),
    ("Calcaneal", "PLATE"),
    ("Clavicle", "PLATE"),
    ("Femur Plate", "PLATE"),
    ("Femoral Plate", "PLATE"),
    ("Fibular Plate", "PLATE"),
    ("Humeral Plate", "PLATE"),
    ("Tibia Plate", "PLATE"),
    ("Tibia Straight Plate", "PLATE"),
    ("Ulna Radius Straight Plate", "PLATE"),
    ("Reconstruction Plate", "PLATE"),
    ("Tubular Plate", "PLATE"),
    ("Olecranon", "PLATE"),
]


def bucket_from_family_text(text: str) -> list[str]:
    """Split one case's product_families into one or more rough demand buckets."""
    if pd.isna(text) or not str(text).strip():
        return []
    parts = [p.strip() for p in str(text).split("|") if p.strip()]
    out = []
    for p in parts:
        bucket = None
        for needle, label in FAMILY_RULES:
            if needle.lower() in p.lower():
                bucket = label
                break
        if bucket is None:
            bucket = "OTHER"
        out.append(bucket)
    return out


def office_inventory_from_master(sets: list[dict], include_non_office: bool = False) -> Counter:
    inv = Counter()
    for row in sets:
        home = str(row.get("home", "")).upper()
        status = str(row.get("status", "")).upper().strip()
        if "NA" in status:
            continue
        if include_non_office or home == "OFFICE":
            inv[bucket_from_set(row)] += 1
    return inv


def demand_probabilities(cases: pd.DataFrame):
    cases = cases.copy()
    cases["case_date"] = pd.to_datetime(cases["case_date"], errors="coerce")
    cases = cases.dropna(subset=["case_date", "hospital"])
    days = max(1, cases["case_date"].nunique())
    per_hospital_bucket = Counter()
    for _, row in cases.iterrows():
        hospital = str(row["hospital"]).strip().upper()
        for bucket in bucket_from_family_text(row.get("product_families", "")):
            per_hospital_bucket[(hospital, bucket)] += 1
    probs = {}
    for key, count in per_hospital_bucket.items():
        probs[key] = count / days
    return probs, days


def run_simulation(probs, office_inventory: Counter, days: int = 90, turnaround_days: int = 2, seed: int = 42):
    random.seed(seed)
    hospitals = sorted({h for h, _ in probs})
    buckets = sorted({b for _, b in probs})
    available = Counter(office_inventory)
    in_use = []  # (return_day, bucket)
    stats = Counter()
    shortages_by_bucket = Counter()
    demand_by_bucket = Counter()

    for day in range(1, days + 1):
        new_in_use = []
        for return_day, bucket in in_use:
            if return_day <= day:
                available[bucket] += 1
                stats["returned_sets"] += 1
            else:
                new_in_use.append((return_day, bucket))
        in_use = new_in_use

        for hospital in hospitals:
            for bucket in buckets:
                lam = probs.get((hospital, bucket), 0)
                whole = int(lam)
                frac = lam - whole
                events = whole + (1 if random.random() < frac else 0)
                for _ in range(events):
                    stats["demand_events"] += 1
                    demand_by_bucket[bucket] += 1
                    if available[bucket] > 0:
                        available[bucket] -= 1
                        in_use.append((day + turnaround_days, bucket))
                        stats["served"] += 1
                    else:
                        stats["shortages"] += 1
                        shortages_by_bucket[bucket] += 1

    rows = []
    for bucket in sorted(demand_by_bucket):
        demand = demand_by_bucket[bucket]
        short = shortages_by_bucket[bucket]
        rows.append(
            {
                "bucket": bucket,
                "initial_office_sets": office_inventory.get(bucket, 0),
                "demand_events": demand,
                "shortages": short,
                "shortage_rate": round(short / demand, 4) if demand else 0,
            }
        )
    return pd.DataFrame(rows).sort_values(["shortages", "demand_events"], ascending=[False, False]), stats


def main():
    ap = argparse.ArgumentParser(description="Starter simulation for OFFICE implant circulation.")
    ap.add_argument("--cases", default="whatsapp_implant_cases.csv")
    ap.add_argument("--master", default="master_data.py")
    ap.add_argument("--days", type=int, default=120)
    ap.add_argument("--turnaround", type=int, default=2)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--include-non-office", action="store_true", help="Treat all sets as part of the circulation pool.")
    args = ap.parse_args()

    cases = pd.read_csv(args.cases)
    sets = load_master_sets(args.master)
    office_inventory = office_inventory_from_master(sets, include_non_office=args.include_non_office)
    probs, historical_days = demand_probabilities(cases)
    summary, stats = run_simulation(
        probs,
        office_inventory,
        days=args.days,
        turnaround_days=args.turnaround,
        seed=args.seed,
    )

    out_csv = Path("simulation_summary.csv")
    summary.to_csv(out_csv, index=False)

    print("=== SIMULATION COMPLETE ===")
    print(f"Historical days learned from: {historical_days}")
    print(f"Simulated days: {args.days}")
    print(f"Turnaround days: {args.turnaround}")
    print(f"Seed: {args.seed}")
    print()
    print("Initial OFFICE pool by bucket:")
    for bucket, count in sorted(office_inventory.items()):
        print(f"  {bucket}: {count}")
    print()
    print("Stats:")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    print()
    print("Top shortage buckets:")
    if len(summary):
        print(summary.head(10).to_string(index=False))
    else:
        print("  No demand buckets were recognized.")
    print()
    print(f"Saved summary to: {out_csv.resolve()}")
    print("\nImportant: this is a starting model. It mainly tests the OFFICE circulation pool.")
    print("It does not yet model exact surgeon-owned/local parked sets serving their own hospitals.")


if __name__ == "__main__":
    main()
