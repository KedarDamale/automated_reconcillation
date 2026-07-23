# GST Reconciliation Demo

Small Flask interface around the reconciliation rules in `automated_reconciliation.ipynb`.

## Run locally

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
flask --app app run --debug
```

Open <http://127.0.0.1:5000>. Gunicorn is intended for Linux/macOS deployment:

```bash
gunicorn --workers 2 --bind 0.0.0.0:8000 app:app
```

Uploaded inputs and generated workbooks are stored under `instance/jobs/` for this demo.

Each completed reconciliation provides one `gst_reconciliation_report.xlsx` download with
nine sheets: the two original registers, a combined missing/invalid GSTIN review sheet,
matched pairs, unmatched rows in both directions with probable matches, date-only
differences, and the two normalized reconciled result tables shown in the application.

## Performance architecture

The data pipeline uses Polars/fastexcel for CSV and XLSX ingestion, NumPy and
RapidFuzz for deterministic one-to-one matching, and XlsxWriter in
constant-memory mode for report generation. Pandas and openpyxl are not runtime
dependencies.

Run a repeatable generated workload benchmark with:

```powershell
python benchmarks\benchmark_reconciliation.py --rows 10000 --workload exact --export
python benchmarks\benchmark_reconciliation.py --rows 100000 --workload exact
```

`--workload` also accepts `fuzzy` and `unmatched`. Workbook export is intentionally
opt-in because the nine-sheet XLSX is usually the slowest stage.
