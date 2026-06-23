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
