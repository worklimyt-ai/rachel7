# Osteo Ops: Script + Dashboard

This workspace now includes:
- `ops_engine.py`: core logic for sets, plates, powertools, case buckets, and distance estimates.
- `run_report.py`: CLI to generate JSON/CSV outputs.
- `dashboard.py`: Streamlit dashboard for commander view.

## Data sources
The engine supports either local CSV files or Google Sheet CSV URLs.

Default behavior:
- `run_report.py` uses local files first:
  - `/Users/bellbell/Downloads/cases - cases.csv`
  - `/Users/bellbell/Downloads/cases - archive.csv`
  - then falls back to Google CSV URLs in `ops_engine.py`.
- `dashboard.py` starts with Google CSV URLs in the sidebar and can be switched to local paths.

Master data is loaded from `master_data.py` and uses:
- `SETS`
- `PLATES`
- `HOSPITALS`

## Setup (uv on macOS)
```bash
cd /Users/bellbell/osteo_track/rachel
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
```

## Run script report
```bash
uv run python run_report.py --out outputs
```

Optional explicit sources:
```bash
uv run python run_report.py \
  --master-data /Users/bellbell/osteo_track/rachel/master_data.py \
  --cases "https://docs.google.com/spreadsheets/d/e/2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IYnirCfsnGHoMyn5xPoG5c/pub?gid=0&single=true&output=csv" \
  --archive "https://docs.google.com/spreadsheets/d/e/2PACX-1vQrbm_5s59966ZVWFmrqkg1vQ21YR1YEd1h_J0M7Fc6FjO0ai3l-aWns0IYnirCfsnGHoMyn5xPoG5c/pub?gid=1320419668&single=true&output=csv" \
  --out outputs
```

Main output:
- `outputs/operations_report.json`
- plus CSVs for each section (sets, plates, powertools, buckets, routes, unknowns)

Google Sheets entry guide:
- `GOOGLE_SHEETS_INPUT_GUIDE.md`

## Run dashboard
```bash
uv run streamlit run dashboard.py
```

In the sidebar, you can change source paths/URLs and refresh.

## Notes
- Timezone is fixed to `Asia/Kuala_Lumpur`.
- Route distances are estimated from hospital coordinates in `master_data.py`.
- Hospital code aliases (e.g. `QE1 -> HQE1`) are handled in `ops_engine.py`.
