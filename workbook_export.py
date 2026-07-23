"""Streaming nine-sheet GST reconciliation workbook writer."""

from __future__ import annotations

import logging
import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable

import numpy as np
import polars as pl
import xlsxwriter

from reconciliation import MatchPair, ReconciliationReport

logger = logging.getLogger(__name__)

SHEET_NAMES = [
    "1 Purchase Register", "2 GSTR-2B", "3 Missing & Invalid GSTIN",
    "4 Matched Entries", "5 PR Not in 2B", "6 2B Not in PR", "7 Date Only",
    "8 PR Reconciled", "9 GSTR-2B Reconciled",
]
NAVY = "#17324D"
BLUE = "#DCEAF7"
PALE_BLUE = "#EEF5FB"
GOLD = "#F8D66D"
PALE_GOLD = "#FFF5CC"
GREEN = "#DDEEDB"
WHITE = "#FFFFFF"
GRID = "#B8C6D1"
DATE_NUMBER_FORMAT = "dd-mmm-yyyy"
AMOUNT_NUMBER_FORMAT = "#,##0.00;[Red]-#,##0.00"
EXCEL_MAX_ROWS = 1_048_576
WIDTH_SAMPLE_ROWS = 250


def _formats(workbook: xlsxwriter.Workbook) -> dict[str, xlsxwriter.format.Format]:
    return {
        "header": workbook.add_format({
            "bold": True, "font_color": WHITE, "bg_color": NAVY,
            "align": "center", "valign": "vcenter", "text_wrap": True,
            "bottom": 1, "bottom_color": GRID,
        }),
        "date": workbook.add_format({"num_format": DATE_NUMBER_FORMAT}),
        "amount": workbook.add_format({"num_format": AMOUNT_NUMBER_FORMAT}),
        "score": workbook.add_format({
            "bold": True, "font_color": NAVY, "bg_color": PALE_GOLD,
            "align": "center", "valign": "vcenter",
            "left": 5, "right": 5, "left_color": NAVY, "right_color": NAVY,
            "num_format": "0.00",
        }),
        "score_header": workbook.add_format({
            "bold": True, "font_color": NAVY, "bg_color": GOLD,
            "align": "center", "valign": "vcenter",
            "left": 5, "right": 5, "left_color": NAVY, "right_color": NAVY,
        }),
        "date_difference": workbook.add_format({
            "bold": True, "font_color": NAVY, "bg_color": PALE_BLUE,
            "align": "center", "valign": "vcenter", "num_format": "0",
        }),
        "group_blue": workbook.add_format({
            "bold": True, "font_color": NAVY, "bg_color": BLUE,
            "font_size": 12, "align": "center", "valign": "vcenter",
        }),
        "group_gold": workbook.add_format({
            "bold": True, "font_color": NAVY, "bg_color": GOLD,
            "font_size": 12, "align": "center", "valign": "vcenter",
        }),
        "group_green": workbook.add_format({
            "bold": True, "font_color": NAVY, "bg_color": GREEN,
            "font_size": 12, "align": "center", "valign": "vcenter",
        }),
    }


def write_reconciliation_workbook(
    report: ReconciliationReport, output_path: str | Path
) -> Path:
    started = time.perf_counter()
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(
        str(output_path),
        {"constant_memory": True, "strings_to_numbers": False, "strings_to_formulas": False},
    )
    workbook.set_properties({
        "title": "GST Reconciliation Report",
        "subject": "Purchase Register and GSTR-2B reconciliation",
    })
    formats = _formats(workbook)
    try:
        _write_simple_sheet(workbook, SHEET_NAMES[0], report.pr_raw, formats)
        _write_simple_sheet(workbook, SHEET_NAMES[1], report.gstr2b_raw, formats)
        _write_simple_sheet(workbook, SHEET_NAMES[2], _gstin_review_dataframe(report), formats)
        _write_grouped_sheet(
            workbook, SHEET_NAMES[3], "Purchase Register", "PR Index", report.pr_raw,
            "Match Score", "GSTR-2B", "2B Index", report.gstr2b_raw,
            ((p.pr_index, p.score, p.gstr2b_index) for p in report.matched_pairs),
            formats, True,
        )
        _write_grouped_sheet(
            workbook, SHEET_NAMES[4], "Purchase Register - Not in 2B", "PR Index",
            report.pr_raw, "Probable Score", "Probable GSTR-2B Match", "2B Index",
            report.gstr2b_raw,
            (_probable_row(i, report.pr_probable_matches.get(i), False)
             for i in report.pr_unmatched_indexes),
            formats, True,
        )
        _write_grouped_sheet(
            workbook, SHEET_NAMES[5], "GSTR-2B - Not in Purchase Register", "2B Index",
            report.gstr2b_raw, "Probable Score", "Probable Purchase Register Match",
            "PR Index", report.pr_raw,
            (_probable_row(i, report.gstr2b_probable_matches.get(i), True)
             for i in report.gstr2b_unmatched_indexes),
            formats, True,
        )
        _write_grouped_sheet(
            workbook, SHEET_NAMES[6], "Purchase Register", "PR Index", report.pr_raw,
            "Date Difference (Days)", "GSTR-2B", "2B Index", report.gstr2b_raw,
            ((p.pr_index, _date_difference_days(report, p), p.gstr2b_index)
             for p in report.date_only_pairs),
            formats, False,
        )
        _write_simple_sheet(workbook, SHEET_NAMES[7], report.pr_result, formats)
        _write_simple_sheet(workbook, SHEET_NAMES[8], report.gstr2b_result, formats)
    finally:
        workbook.close()
    logger.info(
        "xlsx export: path=%s size=%d elapsed=%.3fs",
        output_path, output_path.stat().st_size, time.perf_counter() - started,
    )
    return output_path


def _guard_rows(sheet_name: str, data_rows: int, header_rows: int) -> None:
    if data_rows + header_rows > EXCEL_MAX_ROWS:
        raise ValueError(
            f'Worksheet "{sheet_name}" needs {data_rows + header_rows:,} rows; '
            f"Excel supports at most {EXCEL_MAX_ROWS:,}."
        )


def _write_simple_sheet(
    workbook: xlsxwriter.Workbook, name: str, frame: pl.DataFrame,
    formats: dict[str, xlsxwriter.format.Format],
) -> None:
    _guard_rows(name, frame.height, 1)
    worksheet = workbook.add_worksheet(name)
    headers = frame.columns
    widths = [max(12, min(40, len(str(header)) + 2)) for header in headers]
    worksheet.set_row(0, 24)
    worksheet.write_row(0, 0, headers, formats["header"])
    date_columns = {i for i, header in enumerate(headers) if _is_date_header(header)}
    amount_columns = {i for i, header in enumerate(headers) if _is_amount_header(header)}
    for column in date_columns:
        worksheet.set_column(column, column, None, formats["date"])
    for column in amount_columns:
        worksheet.set_column(column, column, None, formats["amount"])
    for row_number, values in enumerate(frame.iter_rows(), start=1):
        clean = [_display_value(value, headers[i]) for i, value in enumerate(values)]
        worksheet.write_row(row_number, 0, clean)
        if row_number <= WIDTH_SAMPLE_ROWS:
            for column, value in enumerate(clean):
                if value is not None:
                    widths[column] = min(40, max(widths[column], len(str(value)) + 2))
    for column, width in enumerate(widths):
        fmt = formats["date"] if column in date_columns else (
            formats["amount"] if column in amount_columns else None
        )
        worksheet.set_column(column, column, width, fmt)
    if headers:
        worksheet.autofilter(0, 0, max(0, frame.height), len(headers) - 1)
    worksheet.freeze_panes(1, 0)


def _write_grouped_sheet(
    workbook: xlsxwriter.Workbook, name: str,
    left_label: str, left_index_label: str, left_frame: pl.DataFrame,
    center_label: str, right_label: str, right_index_label: str,
    right_frame: pl.DataFrame, rows: Iterable[tuple[int, float | int | None, int | None]],
    formats: dict[str, xlsxwriter.format.Format], emphasize_center: bool,
) -> None:
    materialized_rows = list(rows)
    _guard_rows(name, len(materialized_rows), 2)
    worksheet = workbook.add_worksheet(name)
    left_headers = [left_index_label, *left_frame.columns]
    right_headers = [right_index_label, *right_frame.columns]
    left_end = len(left_headers) - 1
    center = len(left_headers)
    right_start = center + 1
    right_end = right_start + len(right_headers) - 1
    worksheet.merge_range(0, 0, 0, left_end, left_label, formats["group_blue"])
    worksheet.write(0, center, "Match", formats["group_gold"])
    worksheet.merge_range(0, right_start, 0, right_end, right_label, formats["group_green"])
    headers = [*left_headers, center_label, *right_headers]
    worksheet.write_row(1, 0, headers, formats["header"])
    worksheet.write(1, center, center_label, formats["score_header"])
    for column, header in enumerate(headers):
        if _is_date_header(header):
            worksheet.set_column(column, column, None, formats["date"])
        elif _is_amount_header(header):
            worksheet.set_column(column, column, None, formats["amount"])
    widths = [max(12, min(40, len(str(header)) + 2)) for header in headers]
    for output_row, (left_index, center_value, right_index) in enumerate(
        materialized_rows, start=2
    ):
        left = [left_index, *left_frame.row(left_index)]
        right = (
            [right_index, *right_frame.row(right_index)]
            if right_index is not None else [None] * len(right_headers)
        )
        values = [*left, center_value, *right]
        clean = [_display_value(value, headers[i]) for i, value in enumerate(values)]
        worksheet.write_row(output_row, 0, clean)
        if center_value is not None:
            worksheet.write(
                output_row, center, center_value,
                formats["score"] if emphasize_center else formats["date_difference"],
            )
        if output_row < WIDTH_SAMPLE_ROWS + 2:
            for column, value in enumerate(clean):
                if value is not None:
                    widths[column] = min(40, max(widths[column], len(str(value)) + 2))
    for column, width in enumerate(widths):
        header = headers[column]
        fmt = formats["date"] if _is_date_header(header) else (
            formats["amount"] if _is_amount_header(header) else None
        )
        worksheet.set_column(column, column, width, fmt)
    worksheet.autofilter(1, 0, max(1, len(materialized_rows) + 1), right_end)
    worksheet.freeze_panes(2, 0)
    worksheet.set_row(0, 25)
    worksheet.set_row(1, 24)


def _gstin_review_dataframe(report: ReconciliationReport) -> pl.DataFrame:
    metadata = ["Source", "Original Index", "GSTIN Issue", "Original GSTIN"]
    source_columns = [_review_column_name(column, metadata) for column in report.pr_raw.columns]
    for column in report.gstr2b_raw.columns:
        reviewed = _review_column_name(column, metadata)
        if reviewed not in source_columns:
            source_columns.append(reviewed)
    records: list[dict] = []
    for source, frame, issues, originals in (
        ("Purchase Register", report.pr_raw, report.pr_gstin_issues, report.pr_original_gstin_values),
        ("GSTR-2B", report.gstr2b_raw, report.gstr2b_gstin_issues,
         report.gstr2b_original_gstin_values),
    ):
        for index, issue in sorted(issues.items()):
            record = {
                "Source": source, "Original Index": index, "GSTIN Issue": issue,
                "Original GSTIN": originals.get(index),
            }
            record.update({
                _review_column_name(column, metadata): value
                for column, value in frame.row(index, named=True).items()
            })
            records.append(record)
    schema = [*metadata, *source_columns]
    if not records:
        return pl.DataFrame({column: [] for column in schema})
    return pl.DataFrame(records, infer_schema_length=None).select(
        [pl.col(column) if column in records[0] or any(column in r for r in records)
         else pl.lit(None).alias(column) for column in schema]
    )


def _review_column_name(column: object, metadata_columns: list[str]) -> str:
    name = str(column)
    return f"Raw {name}" if name in metadata_columns else name


def _probable_row(
    unmatched_index: int, probable: MatchPair | None, reverse: bool,
) -> tuple[int, float | None, int | None]:
    if probable is None:
        return unmatched_index, None, None
    return (
        unmatched_index, probable.score,
        probable.pr_index if reverse else probable.gstr2b_index,
    )


def _date_difference_days(report: ReconciliationReport, pair: MatchPair) -> int | None:
    pr = _workbook_date_value(report.pr_result["Invoice Date"][pair.pr_index])
    twob = _workbook_date_value(report.gstr2b_result["Invoice Date"][pair.gstr2b_index])
    return None if pr is None or twob is None else abs((twob - pr).days)


def _is_date_header(header: str | None) -> bool:
    normalized = " ".join((header or "").lower().split())
    return "date" in normalized and "difference" not in normalized


def _is_amount_header(header: str | None) -> bool:
    normalized = " ".join((header or "").lower().split())
    return (
        normalized in {"taxable value", "igst", "cgst", "sgst", "total gst"}
        or "taxable" in normalized
    )


def _display_value(value: object, header: str | None) -> object:
    value = _excel_value(value)
    if _is_date_header(header):
        parsed = _workbook_date_value(value)
        return parsed if parsed is not None else value
    if _is_amount_header(header):
        parsed = _workbook_amount_value(value)
        return parsed if parsed is not None else value
    return value


def _workbook_date_value(value: object) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float, np.number)) and not isinstance(value, bool):
        number = float(value)
        whole = int(number)
        if number.is_integer() and 19000101 <= whole <= 29991231:
            try:
                return datetime.strptime(str(whole), "%Y%m%d")
            except ValueError:
                return None
        if 1 <= number <= 100000:
            return datetime(1899, 12, 30) + timedelta(days=number)
        return None
    text = str(value).strip()
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        return _workbook_date_value(float(text))
    for fmt in (
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y",
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _workbook_amount_value(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float, np.number)) and not isinstance(value, bool):
        number = float(value)
        return None if np.isnan(number) else number
    text = str(value).strip().replace("\u00a0", "")
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    text = re.sub(r"[,₹$\s]", "", text)
    if text.endswith("-"):
        text = f"-{text[:-1]}"
    try:
        amount = float(text)
    except ValueError:
        return None
    return -amount if negative else amount


def _excel_value(value: object) -> object:
    if value is None:
        return None
    if isinstance(value, np.generic):
        value = value.item()
    if isinstance(value, float) and np.isnan(value):
        return None
    if isinstance(value, (list, tuple, set)):
        return ", ".join(str(item) for item in value)
    if isinstance(value, (datetime, date, str, int, float, bool)):
        return value
    return str(value)
