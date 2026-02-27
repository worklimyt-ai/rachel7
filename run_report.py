from __future__ import annotations

import argparse

from ops_engine import build_operations_report, write_report_files


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate operations JSON/CSV outputs")
    parser.add_argument("--master-data", default="master_data.py", help="Path to master_data.py")
    parser.add_argument("--cases", default=None, help="Cases CSV path or URL")
    parser.add_argument("--archive", default=None, help="Archive CSV path or URL")
    parser.add_argument("--out", default="outputs", help="Output folder")
    args = parser.parse_args()

    report = build_operations_report(
        master_data_path=args.master_data,
        cases_source=args.cases,
        archive_source=args.archive,
    )
    paths = write_report_files(report, args.out)

    print("Generated operations report")
    print(f"today_kl: {report['meta']['today_kl']}")
    print(f"cases_total: {report['kpis']['cases_total']}")
    print(f"to_deliver: {report['kpis']['to_deliver']}")
    print(f"delivered_today: {report['kpis']['delivered_today']}")
    print(f"to_deliver_tomorrow: {report['kpis']['to_deliver_tomorrow']}")
    print(f"json: {paths['operations_report_json']}")


if __name__ == "__main__":
    main()
