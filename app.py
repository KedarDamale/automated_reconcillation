from __future__ import annotations

import json
import logging
import math
import uuid
from datetime import datetime
from pathlib import Path

import polars as pl
from flask import Flask, abort, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from reconciliation import (
    DEFAULT_DATE_TOLERANCE_DAYS,
    MAX_DATE_TOLERANCE_DAYS,
    MIN_DATE_TOLERANCE_DAYS,
    PREFERRED_COLUMNS,
    REQUIRED_COLUMNS,
    ReconciliationInputError,
    read_uploaded_columns,
    run_reconciliation_report,
    suggest_column_mapping,
)
from workbook_export import write_reconciliation_workbook


if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")

ALLOWED_EXTENSIONS = {".csv", ".xlsx"}
PAGE_SIZE = 10
DISPLAY_AMOUNT_COLUMNS = {
    "Taxable Value",
    "IGST",
    "CGST",
    "SGST",
    "Total GST",
}
DISPLAY_SCORE_COLUMNS = {"Best score", "Best probable score"}


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

    @app.post("/columns")
    def preview_columns():
        upload = request.files.get("file")
        try:
            _validate_upload(upload, "File")
            columns = read_uploaded_columns(upload, upload.filename)
        except ReconciliationInputError as exc:
            return {"error": str(exc)}, 400

        return {
            "columns": columns,
            "preferred_columns": PREFERRED_COLUMNS,
            "required_columns": REQUIRED_COLUMNS,
            "suggestions": suggest_column_mapping(columns, PREFERRED_COLUMNS),
        }

    @app.post("/reconcile")
    def reconcile_uploads():
        pr_file = request.files.get("pr_file")
        gstr2b_file = request.files.get("gstr2b_file")

        try:
            _validate_upload(pr_file, "Purchase Register")
            _validate_upload(gstr2b_file, "GSTR-2B")
            pr_mapping = _parse_mapping(request.form.get("pr_mapping"), "Purchase Register")
            gstr2b_mapping = _parse_mapping(request.form.get("gstr2b_mapping"), "GSTR-2B")
            date_tolerance_days = _parse_date_tolerance(request.form.get("date_tolerance_days"))

            job_id = uuid.uuid4().hex
            job_dir = app.config["JOBS_DIR"] / job_id
            job_dir.mkdir()

            pr_path = job_dir / f"pr_input{Path(secure_filename(pr_file.filename)).suffix.lower()}"
            gstr2b_path = job_dir / f"gstr2b_input{Path(secure_filename(gstr2b_file.filename)).suffix.lower()}"
            pr_file.save(pr_path)
            gstr2b_file.save(gstr2b_path)

            report = run_reconciliation_report(
                pr_path,
                gstr2b_path,
                pr_mapping=pr_mapping,
                gstr2b_mapping=gstr2b_mapping,
                date_tolerance_days=date_tolerance_days,
            )
            _save_result(report.pr_result, job_dir / "pr_result.json")
            _save_result(report.gstr2b_result, job_dir / "gstr2b_result.json")
            write_reconciliation_workbook(
                report, job_dir / "gst_reconciliation_report.xlsx"
            )
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

    @app.get("/export/<job_id>")
    def export(job_id: str):
        job_dir = _job_dir(app, job_id)
        excel_path = job_dir / "gst_reconciliation_report.xlsx"
        if not excel_path.is_file():
            abort(404)
        return send_file(
            excel_path,
            as_attachment=True,
            download_name="gst_reconciliation_report.xlsx",
        )

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


def _parse_mapping(raw: str | None, label: str) -> dict | None:
    """Parse a user-confirmed column mapping submitted alongside a file.

    Absent entirely, this signals "no mapping was confirmed" (e.g. JS
    disabled, or the column-preview call failed) and reconciliation falls
    back to automatic fuzzy matching. Present but malformed is treated as a
    tampered/broken request and rejected immediately, same as any other
    invalid input.
    """
    if not raw:
        return None
    try:
        mapping = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ReconciliationInputError(f"{label} column mapping was malformed.") from exc
    if not isinstance(mapping, dict):
        raise ReconciliationInputError(f"{label} column mapping was malformed.")
    return mapping


def _parse_date_tolerance(raw: str | None) -> int:
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return DEFAULT_DATE_TOLERANCE_DAYS
    return min(max(value, MIN_DATE_TOLERANCE_DAYS), MAX_DATE_TOLERANCE_DAYS)


def _save_result(dataframe: pl.DataFrame, json_path: Path) -> None:
    dataframe.write_json(json_path)


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
        "rows": [_format_result_row(row) for row in rows[start : start + PAGE_SIZE]],
        "columns": list(rows[0].keys()) if rows else [],
        "page": page,
        "pages": pages,
        "total": len(rows),
    }


def _format_result_row(row: dict) -> dict:
    return {
        column: _format_result_value(column, value)
        for column, value in row.items()
    }


def _format_result_value(column: str, value):
    if value is None:
        return value
    if column == "Invoice Date":
        try:
            compact_date = str(int(float(value)))
            return datetime.strptime(compact_date, "%Y%m%d").strftime("%d-%b-%Y")
        except (TypeError, ValueError, OverflowError):
            return value
    if column in DISPLAY_AMOUNT_COLUMNS | DISPLAY_SCORE_COLUMNS:
        try:
            return f"{float(value):,.2f}"
        except (TypeError, ValueError):
            return value
    return value


def _remove_job(job_dir: Path) -> None:
    for path in job_dir.iterdir():
        path.unlink(missing_ok=True)
    job_dir.rmdir()


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
