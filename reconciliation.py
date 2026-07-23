"""High-throughput GST reconciliation using Polars, NumPy, and RapidFuzz."""

from __future__ import annotations

import logging
import os
import re
import time
import unicodedata
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import numpy as np
import polars as pl
from rapidfuzz import fuzz, process

logger = logging.getLogger(__name__)

PREFERRED_COLUMNS = [
    "Supplier Name", "GSTIN", "Invoice No", "Invoice Date",
    "HSN Code (optional)", "Taxable Value", "IGST", "CGST", "SGST", "Total GST",
]
REQUIRED_COLUMNS = ["Supplier Name", "GSTIN", "Invoice No", "Invoice Date", "Taxable Value"]
NUMERIC_COLUMNS = {"Taxable Value", "IGST", "CGST", "SGST", "Total GST"}
DEFAULT_MATCH_THRESHOLD = 95
MIN_DATE_TOLERANCE_DAYS = 0
MAX_DATE_TOLERANCE_DAYS = 60
DEFAULT_DATE_TOLERANCE_DAYS = 10
MISSING_GSTIN = "Missing GSTIN"
INVALID_GSTIN = "Invalid GSTIN"
GSTIN_TOKEN_PATTERN = re.compile(r"\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[A-Z0-9]")


@dataclass(frozen=True)
class MatchPair:
    pr_index: int
    gstr2b_index: int
    score: float


@dataclass
class ReconciliationReport:
    pr_raw: pl.DataFrame
    gstr2b_raw: pl.DataFrame
    pr_result: pl.DataFrame
    gstr2b_result: pl.DataFrame
    matched_pairs: list[MatchPair]
    date_only_pairs: list[MatchPair]
    pr_gstin_issues: dict[int, str]
    gstr2b_gstin_issues: dict[int, str]
    pr_original_gstin_values: dict[int, object]
    gstr2b_original_gstin_values: dict[int, object]
    pr_unmatched_indexes: list[int]
    gstr2b_unmatched_indexes: list[int]
    pr_probable_matches: dict[int, MatchPair]
    gstr2b_probable_matches: dict[int, MatchPair]

    @property
    def pr_missing_gstin_indexes(self) -> list[int]:
        return [i for i, issue in self.pr_gstin_issues.items() if issue == MISSING_GSTIN]


class ReconciliationInputError(ValueError):
    pass


def handle_multitype_import(
    file_path: str | Path, *, preserve_raw_values: bool = False
) -> pl.DataFrame:
    if file_path is None:
        raise ReconciliationInputError("No file was provided.")
    extension = os.path.splitext(str(file_path))[1].lower()
    try:
        if extension == ".csv":
            options: dict[str, Any] = {
                "try_parse_dates": False,
                "truncate_ragged_lines": False,
            }
            if preserve_raw_values:
                options["infer_schema"] = False
                options["null_values"] = []
            return pl.read_csv(file_path, **options)
        if extension == ".xlsx":
            return pl.read_excel(file_path, engine="calamine", infer_schema_length=None)
    except Exception as exc:
        raise ReconciliationInputError(f"Could not read {Path(file_path).name}: {exc}") from exc
    raise ReconciliationInputError("Only .csv and .xlsx files are supported.")


def read_uploaded_columns(file, filename: str) -> list[str]:
    extension = Path(filename).suffix.lower()
    try:
        if extension == ".csv":
            return pl.read_csv(file, n_rows=0).columns
        if extension == ".xlsx":
            frame = pl.read_excel(file, engine="calamine", read_options={"n_rows": 0})
            return frame.columns
    except Exception as exc:
        raise ReconciliationInputError(f"Could not read {filename}: {exc}") from exc
    raise ReconciliationInputError("Only .csv and .xlsx files are supported.")


def normalize_col_names(column_name: object) -> str:
    return (
        str(column_name).lower().strip().replace("_", " ").replace("-", " ").replace(".", "")
    )


def suggest_column_mapping(
    source_columns: list[str],
    preferred_columns: list[str] = PREFERRED_COLUMNS,
    threshold: int = DEFAULT_MATCH_THRESHOLD,
) -> dict[str, dict]:
    normalized = {column: normalize_col_names(column) for column in source_columns}
    used: set[str] = set()
    result: dict[str, dict] = {}
    for preferred in preferred_columns:
        candidates = {key: value for key, value in normalized.items() if key not in used}
        if not candidates:
            result[preferred] = {"source": None, "score": 0}
            continue
        best, score, _ = process.extractOne(
            normalize_col_names(preferred), candidates.values(), scorer=fuzz.token_set_ratio
        )
        source = next((key for key, value in candidates.items() if value == best), None)
        if score >= threshold and source is not None:
            used.add(source)
            result[preferred] = {"source": source, "score": float(score)}
        else:
            result[preferred] = {"source": None, "score": float(score)}
    return result


def _taxable_value_sources(
    source_value: str | list[str] | None, label: str
) -> list[str]:
    if source_value is None or source_value == "":
        return []
    if isinstance(source_value, str):
        return [source_value]
    if not isinstance(source_value, (list, tuple)) or not source_value:
        raise ReconciliationInputError(f"{label} column mapping for Taxable Value was malformed.")
    if any(not isinstance(item, str) or not item for item in source_value):
        raise ReconciliationInputError(f"{label} column mapping for Taxable Value was malformed.")
    if len(set(source_value)) != len(source_value):
        raise ReconciliationInputError(
            f"{label} column mapping selects the same Taxable Value column more than once."
        )
    return list(source_value)


def _numeric_expr(column: str) -> pl.Expr:
    text = (
        pl.col(column).cast(pl.String, strict=False).str.strip_chars()
        .str.replace_all("\u00a0", "")
        .str.replace_all(r"[,₹$\s]", "")
        .str.replace(r"^\((.*)\)$", r"-${1}")
        .str.replace(r"^(.*)-$", r"-${1}")
    )
    return text.cast(pl.Float64, strict=False)


def apply_column_mapping(
    df: pl.DataFrame,
    mapping: dict[str, str | list[str] | None],
    preferred_columns: list[str] = PREFERRED_COLUMNS,
    label: str = "File",
) -> pl.DataFrame:
    unknown = sorted(set(mapping) - set(preferred_columns))
    if unknown:
        raise ReconciliationInputError(
            f"{label} column mapping references unknown field(s): {', '.join(unknown)}."
        )
    used: dict[str, str] = {}
    renames: dict[str, str] = {}
    taxable_sources: list[str] = []
    for preferred in preferred_columns:
        value = mapping.get(preferred)
        if preferred == "Taxable Value":
            sources = _taxable_value_sources(value, label)
        else:
            if isinstance(value, (list, tuple)):
                raise ReconciliationInputError(
                    f"{label} can only map multiple source columns to Taxable Value."
                )
            sources = [value] if value else []
        for source in sources:
            if source not in df.columns:
                raise ReconciliationInputError(
                    f'{label} column mapping points "{preferred}" at a column '
                    f'("{source}") that isn\'t in the uploaded file.'
                )
            if source in used:
                raise ReconciliationInputError(
                    f'{label} column mapping uses "{source}" for both "{used[source]}" '
                    f'and "{preferred}". Each column can only be mapped once.'
                )
            used[source] = preferred
        if preferred == "Taxable Value":
            taxable_sources = sources
        elif sources:
            renames[sources[0]] = preferred
    missing = [column for column in REQUIRED_COLUMNS if column not in used.values()]
    if missing:
        raise ReconciliationInputError(
            f"{label} is missing a mapping for required column(s): {', '.join(missing)}."
        )

    result = df.rename(renames)
    if taxable_sources:
        invalid = pl.lit(False)
        values: list[pl.Expr] = []
        for source in taxable_sources:
            present = pl.col(source).cast(pl.String, strict=False).str.strip_chars().ne("")
            parsed = _numeric_expr(source)
            invalid = invalid | (present & parsed.is_null())
            values.append(parsed)
        invalid_count = result.select(invalid.sum()).item()
        if invalid_count:
            raise ReconciliationInputError(
                f"{label} selected Taxable Value columns contain {invalid_count} non-numeric row(s)."
            )
        result = result.with_columns(
            pl.sum_horizontal(values, ignore_nulls=False).alias("Taxable Value")
        )
    result = _ensure_columns(result, preferred_columns)
    return result.select(preferred_columns)


def extract_preferred_columns_from_df(
    df: pl.DataFrame,
    preferred_columns: list[str],
    threshold: int = DEFAULT_MATCH_THRESHOLD,
) -> tuple[pl.DataFrame, dict]:
    suggestions = suggest_column_mapping(df.columns, preferred_columns, threshold)
    renames = {
        info["source"]: preferred
        for preferred, info in suggestions.items()
        if info["source"] is not None
    }
    result = _ensure_columns(df.rename(renames), preferred_columns)
    return result.select(preferred_columns), {
        column: info["source"] for column, info in suggestions.items()
    }


def _ensure_columns(df: pl.DataFrame, columns: list[str]) -> pl.DataFrame:
    missing = [pl.lit(None).alias(column) for column in columns if column not in df.columns]
    return df.with_columns(missing) if missing else df


def _is_missing_value(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return isinstance(value, float) and np.isnan(value)


def normalize_supplier_name(value: object) -> str | None:
    return None if _is_missing_value(value) else " ".join(str(value).lower().split())


def _canonical_gstin(value: object) -> tuple[str | None, str | None]:
    if _is_missing_value(value):
        return None, MISSING_GSTIN
    text = unicodedata.normalize("NFKC", str(value)).upper()
    compact = re.sub(r"[^A-Z0-9]", "", re.sub(r"[\u200B-\u200D\uFEFF]", "", text))
    tokens = GSTIN_TOKEN_PATTERN.findall(compact)
    return (tokens[0], None) if len(tokens) == 1 else (None, INVALID_GSTIN)


def normalize_gstin(value: object) -> str | None:
    return _canonical_gstin(value)[0]


def _normalize_identifier(value: object) -> str:
    text = str(value).strip()
    match = re.fullmatch(r"(\d+)\.0+", text)
    return match.group(1) if match else text


def normalize_invoice_no(value: object) -> str | None:
    return None if _is_missing_value(value) else _normalize_identifier(value).lower()


def normalize_hsn_code(value: object) -> str | None:
    return None if _is_missing_value(value) else _normalize_identifier(value)


def _normalize_numeric_value(value: object) -> float | None:
    if _is_missing_value(value):
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
        number = float(text)
    except (TypeError, ValueError):
        return None
    return -number if negative else number


def normalize_numeric_col(value: object) -> float | None:
    return _normalize_numeric_value(value)


def _parse_date_value(value: object) -> date | None:
    if _is_missing_value(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float, np.number)) and not isinstance(value, bool):
        return _parse_numeric_date(float(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        return _parse_numeric_date(float(text))
    for fmt in (
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y",
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _parse_numeric_date(value: float) -> date | None:
    if np.isnan(value):
        return None
    whole = int(value)
    if value.is_integer() and 19000101 <= whole <= 29991231:
        try:
            return datetime.strptime(str(whole), "%Y%m%d").date()
        except ValueError:
            return None
    if 1 <= value <= 100000:
        try:
            return date(1899, 12, 30) + timedelta(days=value)
        except (OverflowError, ValueError):
            return None
    return None


def _compact_date(value: object) -> int | None:
    parsed = _parse_date_value(value)
    return None if parsed is None else parsed.year * 10000 + parsed.month * 100 + parsed.day


def normalize_invoice_date(series: pl.Series) -> pl.Series:
    return pl.Series(series.name, [_compact_date(value) for value in series], dtype=pl.Int32)


def normalize_values(df: pl.DataFrame) -> pl.DataFrame:
    supplier = (
        pl.col("Supplier Name").cast(pl.String, strict=False).str.strip_chars()
        .str.to_lowercase().str.replace_all(r"\s+", " ")
    )
    invoice = (
        pl.col("Invoice No").cast(pl.String, strict=False).str.strip_chars()
        .str.replace(r"\.0+$", "").str.to_lowercase()
    )
    hsn = (
        pl.col("HSN Code (optional)").cast(pl.String, strict=False)
        .str.strip_chars().str.replace(r"\.0+$", "")
    )
    gstin_pairs = [_canonical_gstin(value) for value in df["GSTIN"].to_list()]
    result = df.with_columns(
        supplier.replace("", None).alias("Supplier Name"),
        pl.Series("GSTIN", [pair[0] for pair in gstin_pairs], dtype=pl.String),
        pl.Series("_GSTIN Issue", [pair[1] for pair in gstin_pairs], dtype=pl.String),
        invoice.replace("", None).alias("Invoice No"),
        normalize_invoice_date(df["Invoice Date"]),
        hsn.replace("", None).alias("HSN Code (optional)"),
        *[_numeric_expr(column).alias(column) for column in NUMERIC_COLUMNS],
    )
    return result


def _validate_reconciliation_fields(df: pl.DataFrame, label: str) -> None:
    empty = [
        column for column in REQUIRED_COLUMNS
        if column != "GSTIN" and df[column].null_count() == df.height
    ]
    if empty:
        raise ReconciliationInputError(
            f"{label} is missing usable columns: {', '.join(empty)}. Check the column names and data."
        )


def prepare_dataframe(
    file_path: str | Path,
    label: str,
    column_mapping: dict[str, str | list[str] | None] | None = None,
) -> pl.DataFrame:
    return prepare_dataframe_from_raw(
        handle_multitype_import(file_path, preserve_raw_values=True), label, column_mapping
    )


def prepare_dataframe_from_raw(
    dataframe: pl.DataFrame,
    label: str,
    column_mapping: dict[str, str | list[str] | None] | None = None,
) -> pl.DataFrame:
    mapped = (
        apply_column_mapping(dataframe, column_mapping, PREFERRED_COLUMNS, label)
        if column_mapping is not None
        else extract_preferred_columns_from_df(dataframe, PREFERRED_COLUMNS)[0]
    )
    original_gstin = dict(enumerate(mapped["GSTIN"].to_list()))
    normalized = normalize_values(mapped).with_row_index("_row_index")
    issues = {
        int(index): issue
        for index, issue in normalized.select("_row_index", "_GSTIN Issue").iter_rows()
        if issue is not None
    }
    prepared = normalized.select("_row_index", *PREFERRED_COLUMNS, "_GSTIN Issue").with_columns(
        pl.Series("_Original GSTIN", list(original_gstin.values()))
    )
    _validate_reconciliation_fields(prepared, label)
    return prepared


def _metadata(df: pl.DataFrame, key: str, default: Any) -> Any:
    if key == "gstin_issues" and "_GSTIN Issue" in df.columns:
        return {
            int(index): issue
            for index, issue in df.select("_row_index", "_GSTIN Issue").iter_rows()
            if issue is not None
        }
    if key == "original_gstin_values" and "_Original GSTIN" in df.columns:
        return dict(enumerate(df["_Original GSTIN"].to_list()))
    return default


def _gstin_issues_from_dataframe(df: pl.DataFrame) -> dict[int, str]:
    stored = _metadata(df, "gstin_issues", None)
    if stored is not None:
        return stored
    return {
        i: MISSING_GSTIN for i, value in enumerate(df["GSTIN"].to_list()) if _is_missing_value(value)
    }


def _date_ordinal(value: object) -> int:
    parsed = _parse_date_value(value)
    return np.iinfo(np.int32).min if parsed is None else parsed.toordinal()


def _frame_arrays(df: pl.DataFrame) -> dict[str, np.ndarray]:
    return {
        "supplier": np.asarray(df["Supplier Name"].fill_null("").to_list(), dtype=object),
        "gstin": np.asarray(df["GSTIN"].fill_null("").to_list(), dtype=object),
        "invoice": np.asarray(df["Invoice No"].fill_null("").to_list(), dtype=object),
        "date": np.asarray([_date_ordinal(v) for v in df["Invoice Date"]], dtype=np.int32),
        "taxable": np.asarray(
            [np.nan if v is None else float(v) for v in df["Taxable Value"]], dtype=np.float64
        ),
    }


def _comparison_key(columns: list[str], row: dict[str, object]) -> tuple:
    return tuple(row.get(column) for column in columns)


def _eligible_row(row: dict[str, object]) -> bool:
    return all(row.get(column) is not None for column in ("GSTIN", "Invoice Date", "Taxable Value"))


def _reserve_exact_pairs(
    pur: pl.DataFrame, twob: pl.DataFrame, eligible_pr: list[int], available: np.ndarray
) -> list[MatchPair]:
    groups: dict[tuple, deque[int]] = defaultdict(deque)
    for row in twob.iter_rows(named=True):
        index = int(row["_row_index"])
        if available[index] and _eligible_row(row):
            groups[_comparison_key(PREFERRED_COLUMNS, row)].append(index)
    pairs: list[MatchPair] = []
    eligible = set(eligible_pr)
    for row in pur.iter_rows(named=True):
        index = int(row["_row_index"])
        if index not in eligible:
            continue
        if _eligible_row(row):
            matches = groups.get(_comparison_key(PREFERRED_COLUMNS, row))
            if matches:
                pairs.append(MatchPair(index, matches.popleft(), 100.0))
    return pairs


def _reserve_date_only_pairs(
    pur: pl.DataFrame, twob: pl.DataFrame, pr_indexes: list[int],
    available: np.ndarray, date_tolerance_days: int,
) -> list[MatchPair]:
    if date_tolerance_days <= 0:
        return []
    columns = [column for column in PREFERRED_COLUMNS if column != "Invoice Date"]
    groups: dict[tuple, list[int]] = defaultdict(list)
    for row in twob.iter_rows(named=True):
        index = int(row["_row_index"])
        if available[index] and _eligible_row(row):
            groups[_comparison_key(columns, row)].append(index)
    used = np.zeros(twob.height, dtype=bool)
    pairs: list[MatchPair] = []
    pr_set = set(pr_indexes)
    for row in pur.iter_rows(named=True):
        index = int(row["_row_index"])
        if index not in pr_set:
            continue
        if not _eligible_row(row):
            continue
        pr_date = _date_ordinal(row["Invoice Date"])
        possible: list[tuple[int, int]] = []
        for candidate in groups.get(_comparison_key(columns, row), []):
            if used[candidate]:
                continue
            difference = abs(_date_ordinal(twob["Invoice Date"][candidate]) - pr_date)
            if 0 < difference <= date_tolerance_days:
                possible.append((difference, candidate))
        if possible:
            _, candidate = min(possible)
            used[candidate] = True
            pairs.append(MatchPair(index, candidate, 100.0))
    return pairs


def _candidate_scores(
    pr_index: int, pr_arrays: dict[str, np.ndarray], twob_arrays: dict[str, np.ndarray],
    group_indexes: dict[object, np.ndarray], available: np.ndarray,
    numeric_tolerance: float, date_tolerance_days: int,
) -> list[tuple[int, float]]:
    taxable = pr_arrays["taxable"][pr_index]
    pr_date = pr_arrays["date"][pr_index]
    if np.isnan(taxable) or pr_date == np.iinfo(np.int32).min:
        return []
    candidates = group_indexes.get(pr_arrays["gstin"][pr_index])
    if candidates is None or not candidates.size:
        return []
    candidates = candidates[available[candidates]]
    if not candidates.size:
        return []
    mask = (
        (np.abs(twob_arrays["date"][candidates] - pr_date) <= date_tolerance_days)
        & (np.abs(twob_arrays["taxable"][candidates] - taxable)
           <= numeric_tolerance * max(abs(taxable), 1.0))
    )
    candidates = candidates[mask]
    if not candidates.size:
        return []
    supplier_scores = process.cdist(
        [str(pr_arrays["supplier"][pr_index])],
        [str(twob_arrays["supplier"][i]) for i in candidates],
        scorer=fuzz.WRatio, dtype=np.float64,
    )[0]
    invoice_scores = process.cdist(
        [str(pr_arrays["invoice"][pr_index])],
        [str(twob_arrays["invoice"][i]) for i in candidates],
        scorer=fuzz.WRatio, dtype=np.float64,
    )[0]
    return [
        (int(index), float(score))
        for index, score in zip(candidates, 0.5 * supplier_scores + 0.5 * invoice_scores, strict=True)
    ]


def _classify_reconciliation(
    pur: pl.DataFrame, twob: pl.DataFrame, *, pr_raw: pl.DataFrame, gstr2b_raw: pl.DataFrame,
    pr_gstin_issues: dict[int, str] | None = None,
    gstr2b_gstin_issues: dict[int, str] | None = None,
    pr_original_gstin_values: dict[int, object] | None = None,
    gstr2b_original_gstin_values: dict[int, object] | None = None,
    similarity_threshold: int = 80, numeric_tolerance: float = 0.01,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> ReconciliationReport:
    stage = time.perf_counter()
    pr_issues = pr_gstin_issues or _gstin_issues_from_dataframe(pur)
    twob_issues = gstr2b_gstin_issues or _gstin_issues_from_dataframe(twob)
    pr_original = pr_original_gstin_values or dict(enumerate(pur["GSTIN"].to_list()))
    twob_original = gstr2b_original_gstin_values or dict(enumerate(twob["GSTIN"].to_list()))
    pr_available = np.ones(pur.height, dtype=bool)
    twob_available = np.ones(twob.height, dtype=bool)
    for index in pr_issues:
        pr_available[index] = False
    for index in twob_issues:
        twob_available[index] = False
    eligible_pr = np.flatnonzero(pr_available).tolist()

    exact = _reserve_exact_pairs(pur, twob, eligible_pr, twob_available)
    for pair in exact:
        pr_available[pair.pr_index] = False
        twob_available[pair.gstr2b_index] = False
    logger.info("exact matching: pairs=%d elapsed=%.3fs", len(exact), time.perf_counter() - stage)

    stage = time.perf_counter()
    date_only = _reserve_date_only_pairs(
        pur, twob, np.flatnonzero(pr_available).tolist(), twob_available, date_tolerance_days
    )
    for pair in date_only:
        pr_available[pair.pr_index] = False
        twob_available[pair.gstr2b_index] = False
    logger.info("date-only matching: pairs=%d elapsed=%.3fs", len(date_only), time.perf_counter() - stage)

    stage = time.perf_counter()
    pr_arrays, twob_arrays = _frame_arrays(pur), _frame_arrays(twob)
    grouped: dict[object, np.ndarray] = {}
    for gstin in set(twob_arrays["gstin"]):
        grouped[gstin] = np.flatnonzero(twob_arrays["gstin"] == gstin).astype(np.int32)
    cached: dict[int, list[tuple[int, float]]] = {}
    fuzzy: list[MatchPair] = []
    for pr_index in np.flatnonzero(pr_available):
        scores = _candidate_scores(
            int(pr_index), pr_arrays, twob_arrays, grouped, twob_available,
            numeric_tolerance, date_tolerance_days,
        )
        cached[int(pr_index)] = scores
        if scores:
            candidate, score = max(scores, key=lambda item: (item[1], -item[0]))
            if score >= similarity_threshold:
                pair = MatchPair(int(pr_index), candidate, score)
                fuzzy.append(pair)
                pr_available[pr_index] = False
                twob_available[candidate] = False
    logger.info("fuzzy matching: pairs=%d elapsed=%.3fs", len(fuzzy), time.perf_counter() - stage)

    unmatched_pr = np.flatnonzero(pr_available).tolist()
    unmatched_twob = np.flatnonzero(twob_available).tolist()
    unmatched_twob_set = set(unmatched_twob)
    probable_pr: dict[int, MatchPair] = {}
    probable_twob_candidates: dict[int, list[MatchPair]] = defaultdict(list)
    probable_lists_pr: dict[int, list[int]] = {}
    for pr_index in unmatched_pr:
        scores = cached.get(pr_index)
        if scores is None:
            scores = _candidate_scores(
                pr_index, pr_arrays, twob_arrays, grouped, twob_available,
                numeric_tolerance, date_tolerance_days,
            )
        below = [
            MatchPair(pr_index, candidate, score) for candidate, score in scores
            if candidate in unmatched_twob_set and score < similarity_threshold
        ]
        below.sort(key=lambda pair: (-pair.score, pair.gstr2b_index))
        if below:
            probable_pr[pr_index] = below[0]
            probable_lists_pr[pr_index] = [pair.gstr2b_index for pair in below]
            for pair in below:
                probable_twob_candidates[pair.gstr2b_index].append(pair)
    probable_twob: dict[int, MatchPair] = {}
    probable_lists_twob: dict[int, list[int]] = {}
    for index, pairs in probable_twob_candidates.items():
        pairs.sort(key=lambda pair: (-pair.score, pair.pr_index))
        probable_twob[index] = pairs[0]
        probable_lists_twob[index] = [pair.pr_index for pair in pairs]

    matched = [*exact, *fuzzy]
    pr_match = {pair.pr_index: pair for pair in matched}
    twob_match = {pair.gstr2b_index: pair for pair in matched}
    pr_date = {pair.pr_index: pair for pair in date_only}
    twob_date = {pair.gstr2b_index: pair for pair in date_only}

    def result_frame(frame: pl.DataFrame, side: str) -> pl.DataFrame:
        height = frame.height
        score: list[float | None] = [None] * height
        best: list[int | None] = [None] * height
        probable_indexes: list[list[int] | None] = [None] * height
        probable_score: list[float | None] = [None] * height
        probable_best: list[int | None] = [None] * height
        category = ["Unmatched"] * height
        issues = pr_issues if side == "pr" else twob_issues
        matches = pr_match if side == "pr" else twob_match
        dates = pr_date if side == "pr" else twob_date
        probable = probable_pr if side == "pr" else probable_twob
        lists = probable_lists_pr if side == "pr" else probable_lists_twob
        for index, issue in issues.items():
            category[index] = issue
        for index, pair in matches.items():
            score[index] = pair.score
            best[index] = pair.gstr2b_index if side == "pr" else pair.pr_index
            category[index] = "Matched"
        for index, pair in dates.items():
            score[index] = pair.score
            best[index] = pair.gstr2b_index if side == "pr" else pair.pr_index
            category[index] = "Date only"
        for index, pair in probable.items():
            probable_indexes[index] = lists[index]
            probable_score[index] = pair.score
            probable_best[index] = pair.gstr2b_index if side == "pr" else pair.pr_index
        opposite = "2B" if side == "pr" else "PR"
        hidden = [
            column for column in ("_row_index", "_GSTIN Issue", "_Original GSTIN")
            if column in frame.columns
        ]
        return frame.drop(hidden).with_columns(
            pl.Series("Best score", score, dtype=pl.Float64),
            pl.Series(f"Best match {opposite} index", best, dtype=pl.Int32),
            pl.Series(f"Probable {opposite} indexes", probable_indexes, dtype=pl.List(pl.Int32)),
            pl.Series("Best probable score", probable_score, dtype=pl.Float64),
            pl.Series(f"Best probable {opposite} index", probable_best, dtype=pl.Int32),
            pl.Series("Match category", category, dtype=pl.String),
        )

    return ReconciliationReport(
        pr_raw, gstr2b_raw, result_frame(pur, "pr"), result_frame(twob, "twob"),
        matched, date_only, pr_issues, twob_issues, pr_original, twob_original,
        unmatched_pr, unmatched_twob, probable_pr, probable_twob,
    )


def reconcile(
    pur: pl.DataFrame, twob: pl.DataFrame, similarity_threshold: int = 80,
    numeric_tolerance: float = 0.01,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> tuple[pl.DataFrame, pl.DataFrame]:
    if "_row_index" not in pur.columns:
        pur = pur.with_row_index("_row_index")
    if "_row_index" not in twob.columns:
        twob = twob.with_row_index("_row_index")
    report = _classify_reconciliation(
        pur, twob,
        pr_raw=pur.drop([c for c in ("_row_index", "_GSTIN Issue", "_Original GSTIN") if c in pur.columns]),
        gstr2b_raw=twob.drop([c for c in ("_row_index", "_GSTIN Issue", "_Original GSTIN") if c in twob.columns]),
        similarity_threshold=similarity_threshold, numeric_tolerance=numeric_tolerance,
        date_tolerance_days=date_tolerance_days,
    )
    return report.pr_result, report.gstr2b_result


def run_reconciliation_report(
    pr_path: str | Path, gstr2b_path: str | Path,
    pr_mapping: dict[str, str | list[str] | None] | None = None,
    gstr2b_mapping: dict[str, str | list[str] | None] | None = None,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> ReconciliationReport:
    total = time.perf_counter()
    stage = total
    pr_raw = handle_multitype_import(pr_path, preserve_raw_values=True)
    gstr2b_raw = handle_multitype_import(gstr2b_path, preserve_raw_values=True)
    logger.info(
        "read inputs: pr_rows=%d gstr2b_rows=%d elapsed=%.3fs",
        pr_raw.height, gstr2b_raw.height, time.perf_counter() - stage,
    )
    stage = time.perf_counter()
    pur = prepare_dataframe_from_raw(pr_raw, "Purchase Register", pr_mapping)
    twob = prepare_dataframe_from_raw(gstr2b_raw, "GSTR-2B", gstr2b_mapping)
    logger.info("normalization: elapsed=%.3fs", time.perf_counter() - stage)
    report = _classify_reconciliation(
        pur, twob, pr_raw=pr_raw, gstr2b_raw=gstr2b_raw,
        pr_gstin_issues=_metadata(pur, "gstin_issues", {}),
        gstr2b_gstin_issues=_metadata(twob, "gstin_issues", {}),
        pr_original_gstin_values=_metadata(pur, "original_gstin_values", {}),
        gstr2b_original_gstin_values=_metadata(twob, "original_gstin_values", {}),
        date_tolerance_days=date_tolerance_days,
    )
    logger.info("reconciliation complete: elapsed=%.3fs", time.perf_counter() - total)
    return report


def run_reconciliation(
    pr_path: str | Path, gstr2b_path: str | Path,
    pr_mapping: dict[str, str | list[str] | None] | None = None,
    gstr2b_mapping: dict[str, str | list[str] | None] | None = None,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> tuple[pl.DataFrame, pl.DataFrame]:
    report = run_reconciliation_report(
        pr_path, gstr2b_path, pr_mapping, gstr2b_mapping, date_tolerance_days
    )
    return report.pr_result, report.gstr2b_result
