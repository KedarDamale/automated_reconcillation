from __future__ import annotations

import json
import math
import os
import uuid
from pathlib import Path

import pandas as pd
from flask import Flask, abort, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from reconciliation import ReconciliationInputError, run_reconciliation


ALLOWED_EXTENSIONS = {".csv", ".xlsx"}
PAGE_SIZE = 10


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.update(
        MAX_CONTENT_LENGTH=20 * 1024 * 1024,
        JOBS_DIR=Path(app.instance_path) / "jobs",
    )
    app.config["JOBS_DIR"].mkdir(parents=True, exist_ok=True)

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.post("/reconcile")
    def reconcile_uploads():
        pr_file = request.files.get("pr_file")
        gstr2b_file = request.files.get("gstr2b_file")

        try:
            _validate_upload(pr_file, "Purchase Register")
            _validate_upload(gstr2b_file, "GSTR-2B")

            job_id = uuid.uuid4().hex
            job_dir = app.config["JOBS_DIR"] / job_id
            job_dir.mkdir()

            pr_path = job_dir / f"pr_input{Path(secure_filename(pr_file.filename)).suffix.lower()}"
            gstr2b_path = job_dir / f"gstr2b_input{Path(secure_filename(gstr2b_file.filename)).suffix.lower()}"
            pr_file.save(pr_path)
            gstr2b_file.save(gstr2b_path)

            pr_result, gstr2b_result = run_reconciliation(pr_path, gstr2b_path)
            _save_result(pr_result, job_dir / "pr_result.json")
            _save_result(gstr2b_result, job_dir / "gstr2b_result.json")
        except ReconciliationInputError as exc:
            if "job_dir" in locals():
                _remove_job(job_dir)
            return render_template("index.html", error=str(exc)), 400
        except Exception:
            if "job_dir" in locals():
                _remove_job(job_dir)
            app.logger.exception("Reconciliation failed")
            return render_template(
                "index.html", error="Reconciliation failed. Check the uploaded data and try again."
            ), 400

        return redirect(url_for("results", job_id=job_id))

    @app.get("/results/<job_id>")
    def results(job_id: str):
        job_dir = _job_dir(app, job_id)
        pr_rows = _load_rows(job_dir / "pr_result.json")
        gstr2b_rows = _load_rows(job_dir / "gstr2b_result.json")

        pr_page = _positive_int(request.args.get("pr_page"), 1)
        gstr2b_page = _positive_int(request.args.get("gstr2b_page"), 1)
        pr_view = _paginate(pr_rows, pr_page)
        gstr2b_view = _paginate(gstr2b_rows, gstr2b_page)

        return render_template(
            "index.html",
            job_id=job_id,
            pr=pr_view,
            gstr2b=gstr2b_view,
        )

    @app.get("/export/<job_id>/<kind>")
    def export(job_id: str, kind: str):
        names = {
            "pr": ("pr_reconciled.xlsx", "pr_result.json"),
            "gstr2b": ("gstr2b_reconciled.xlsx", "gstr2b_result.json"),
        }
        if kind not in names:
            abort(404)
        excel_name, json_name = names[kind]
        job_dir = _job_dir(app, job_id)
        excel_path = job_dir / excel_name
        if not excel_path.is_file():
            rows = _load_rows(job_dir / json_name)
            pd.DataFrame(rows).to_excel(excel_path)
        return send_file(excel_path, as_attachment=True, download_name=excel_name)

    @app.errorhandler(413)
    def file_too_large(_error):
        return render_template("index.html", error="Uploads must be 20 MB or smaller."), 413

    return app


def _validate_upload(upload, label: str) -> None:
    if upload is None or not upload.filename:
        raise ReconciliationInputError(f"Select a {label} file.")
    extension = Path(secure_filename(upload.filename)).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        raise ReconciliationInputError(f"{label} must be a .csv or .xlsx file.")


def _save_result(dataframe: pd.DataFrame, json_path: Path) -> None:
    json_path.write_text(
        dataframe.to_json(orient="records", date_format="iso"),
        encoding="utf-8",
    )


def _load_rows(path: Path) -> list[dict]:
    if not path.is_file():
        abort(404)
    return json.loads(path.read_text(encoding="utf-8"))


def _job_dir(app: Flask, job_id: str) -> Path:
    if len(job_id) != 32 or any(character not in "0123456789abcdef" for character in job_id):
        abort(404)
    job_dir = app.config["JOBS_DIR"] / job_id
    if not job_dir.is_dir():
        abort(404)
    return job_dir


def _positive_int(raw_value: str | None, default: int) -> int:
    try:
        return max(1, int(raw_value or default))
    except ValueError:
        return default


def _paginate(rows: list[dict], page: int) -> dict:
    pages = max(1, math.ceil(len(rows) / PAGE_SIZE))
    page = min(page, pages)
    start = (page - 1) * PAGE_SIZE
    return {
        "rows": rows[start : start + PAGE_SIZE],
        "columns": list(rows[0].keys()) if rows else [],
        "page": page,
        "pages": pages,
        "total": len(rows),
    }


def _remove_job(job_dir: Path) -> None:
    for path in job_dir.iterdir():
        path.unlink(missing_ok=True)
    job_dir.rmdir()


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
