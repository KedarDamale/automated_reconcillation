"""Reconciliation logic lifted from automated_reconciliation.ipynb."""

from __future__ import annotations

import logging
import os
import time
from pathlib import Path

import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process

logger = logging.getLogger(__name__)

PREFERRED_COLUMNS = [
    "Supplier Name",
    "GSTIN",
    "Invoice No",
    "Invoice Date",
    "HSN Code (optional)",
    "Taxable Value",
    "IGST",
    "CGST",
    "SGST",
    "Total GST",
]

REQUIRED_COLUMNS = ["Supplier Name", "GSTIN", "Invoice No", "Invoice Date", "Taxable Value"]

DEFAULT_MATCH_THRESHOLD = 95

MIN_DATE_TOLERANCE_DAYS = 0
MAX_DATE_TOLERANCE_DAYS = 60
DEFAULT_DATE_TOLERANCE_DAYS = 10


class ReconciliationInputError(ValueError):
    """Raised when an uploaded file cannot safely enter the notebook pipeline."""


def handle_multitype_import(file_path: str | Path) -> pd.DataFrame:
    if file_path is None:
        raise ReconciliationInputError("No file was provided.")

    extension = os.path.splitext(str(file_path))[1].lower()
    try:
        if extension == ".csv":
            return pd.read_csv(file_path)
        if extension == ".xlsx":
            return pd.read_excel(file_path)
    except Exception as exc:
        raise ReconciliationInputError(
            f"Could not read {Path(file_path).name}: {exc}"
        ) from exc

    raise ReconciliationInputError("Only .csv and .xlsx files are supported.")


def read_uploaded_columns(file, filename: str) -> list[str]:
    """Read just the header row of an upload, for the column-mapping preview."""
    extension = os.path.splitext(filename)[1].lower()
    try:
        if extension == ".csv":
            return list(pd.read_csv(file, nrows=0).columns)
        if extension == ".xlsx":
            return list(
                pd.read_excel(file, nrows=0, engine_kwargs={"read_only": True}).columns
            )
    except Exception as exc:
        raise ReconciliationInputError(f"Could not read {filename}: {exc}") from exc

    raise ReconciliationInputError("Only .csv and .xlsx files are supported.")


def normalize_col_names(column_name: object) -> str:
    return (
        str(column_name)
        .lower()
        .strip()
        .replace("_", " ")
        .replace("-", " ")
        .replace(".", "")
    )


def suggest_column_mapping(
    source_columns: list[str],
    preferred_columns: list[str] = PREFERRED_COLUMNS,
    threshold: int = DEFAULT_MATCH_THRESHOLD,
) -> dict[str, dict]:
    """Fuzzy-match uploaded column headers against the preferred schema.

    Returns ``{preferred_column: {"source": matched_source_or_None, "score": float}}``
    so a caller (e.g. the column-mapping UI) can show the system's best guess
    and let a person confirm or override it, rather than silently trusting a
    fuzzy match on financial data.
    """
    source_cols_norm = {column: normalize_col_names(column) for column in source_columns}
    preferred_cols_norm = {
        preferred: normalize_col_names(preferred) for preferred in preferred_columns
    }

    used_source_columns: set[str] = set()
    suggestions: dict[str, dict] = {}

    for preferred_column, preferred_norm in preferred_cols_norm.items():
        candidates = {
            source_column: normalized_column
            for source_column, normalized_column in source_cols_norm.items()
            if source_column not in used_source_columns
        }

        if not candidates:
            suggestions[preferred_column] = {"source": None, "score": 0}
            continue

        best_norm, score, _ = process.extractOne(
            preferred_norm,
            candidates.values(),
            scorer=fuzz.token_set_ratio,
        )

        if score >= threshold:
            matched_source_column = next(
                source
                for source, normalized in candidates.items()
                if normalized == best_norm
            )
            suggestions[preferred_column] = {"source": matched_source_column, "score": float(score)}
            used_source_columns.add(matched_source_column)
        else:
            suggestions[preferred_column] = {"source": None, "score": float(score)}

    return suggestions


def apply_column_mapping(
    df: pd.DataFrame,
    mapping: dict[str, str | None],
    preferred_columns: list[str] = PREFERRED_COLUMNS,
    label: str = "File",
) -> pd.DataFrame:
    """Rename ``df``'s columns using a person-confirmed preferred -> source mapping.

    This is deliberately strict: a mapping a user assembled by hand (however
    much they rearranged it first) is trusted verbatim, but if it is
    inconsistent -- it references a column that doesn't exist, reuses one
    source column for two preferred fields, or leaves a required field
    unmapped -- reconciliation must not silently guess. It fails immediately
    with a message identifying the problem.
    """
    unknown_preferred = sorted(set(mapping) - set(preferred_columns))
    if unknown_preferred:
        raise ReconciliationInputError(
            f"{label} column mapping references unknown field(s): {', '.join(unknown_preferred)}."
        )

    used_sources: dict[str, str] = {}
    for preferred_column in preferred_columns:
        source_column = mapping.get(preferred_column) or None
        if source_column is None:
            continue
        if source_column not in df.columns:
            raise ReconciliationInputError(
                f'{label} column mapping points "{preferred_column}" at a column '
                f'("{source_column}") that isn\'t in the uploaded file.'
            )
        if source_column in used_sources:
            raise ReconciliationInputError(
                f'{label} column mapping uses "{source_column}" for both '
                f'"{used_sources[source_column]}" and "{preferred_column}". '
                "Each column can only be mapped once."
            )
        used_sources[source_column] = preferred_column

    missing_required = [
        column
        for column in REQUIRED_COLUMNS
        if column in preferred_columns and column not in used_sources.values()
    ]
    if missing_required:
        raise ReconciliationInputError(
            f"{label} is missing a mapping for required column(s): {', '.join(missing_required)}."
        )

    rename_map = {source: preferred for source, preferred in used_sources.items()}
    renamed_df = df.rename(columns=rename_map)

    for column in preferred_columns:
        if column not in renamed_df.columns:
            renamed_df[column] = None

    return renamed_df[preferred_columns]


def extract_preferred_columns_from_df(
    df: pd.DataFrame,
    preferred_columns: list[str],
    threshold: int = DEFAULT_MATCH_THRESHOLD,
) -> tuple[pd.DataFrame, dict]:
    """Automatically map and rename columns (no person in the loop).

    Used when a caller hasn't supplied a confirmed mapping, e.g. existing
    scripts/tests. Missing a column or two is tolerated here -- callers that
    want a person to catch and fix that should collect a mapping via
    :func:`suggest_column_mapping` and apply it with :func:`apply_column_mapping`.
    """
    suggestions = suggest_column_mapping(list(df.columns), preferred_columns, threshold)
    mapping = {column: info["source"] for column, info in suggestions.items()}

    rename_map = {source: preferred for preferred, source in mapping.items() if source}
    renamed_df = df.rename(columns=rename_map)

    for column in preferred_columns:
        if column not in renamed_df.columns:
            renamed_df[column] = None

    return renamed_df[preferred_columns], mapping


def normalize_supplier_name(value):
    if pd.isna(value):
        return None
    return str(value).lower().strip()


def normalize_gstin(value):
    if pd.isna(value):
        return None
    return str(value).strip().upper()


def normalize_invoice_no(value):
    if pd.isna(value):
        return None
    return str(value).strip().lower()


def normalize_invoice_date(series: pd.Series) -> pd.Series:
    date = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return date.dt.year * 10000 + date.dt.month * 100 + date.dt.day


def normalize_hsn_code(value):
    if pd.isna(value):
        return None
    return str(value).strip()


def normalize_numeric_col(value):
    if pd.isna(value):
        return None
    return value


def normalize_values(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Supplier Name"] = df["Supplier Name"].apply(normalize_supplier_name)
    df["GSTIN"] = df["GSTIN"].apply(normalize_gstin)
    df["Invoice No"] = df["Invoice No"].apply(normalize_invoice_no)
    df["Invoice Date"] = normalize_invoice_date(df["Invoice Date"])
    df["HSN Code (optional)"] = df["HSN Code (optional)"].apply(normalize_hsn_code)

    for column in ["Taxable Value", "IGST", "CGST", "SGST", "Total GST"]:
        df[column] = df[column].apply(normalize_numeric_col)
    return df


def _validate_reconciliation_fields(df: pd.DataFrame, label: str) -> None:
    empty = [column for column in REQUIRED_COLUMNS if df[column].isna().all()]
    if empty:
        raise ReconciliationInputError(
            f"{label} is missing usable columns: {', '.join(empty)}. "
            "Check the column names and data."
        )

    numeric = pd.to_numeric(df["Taxable Value"], errors="coerce")
    invalid_count = int((df["Taxable Value"].notna() & numeric.isna()).sum())
    if invalid_count:
        raise ReconciliationInputError(
            f"{label} contains {invalid_count} non-numeric Taxable Value row(s)."
        )
    df["Taxable Value"] = numeric


def prepare_dataframe(
    file_path: str | Path,
    label: str,
    column_mapping: dict[str, str | None] | None = None,
) -> pd.DataFrame:
    dataframe = handle_multitype_import(file_path)
    if column_mapping is not None:
        dataframe = apply_column_mapping(dataframe, column_mapping, PREFERRED_COLUMNS, label)
    else:
        dataframe, _ = extract_preferred_columns_from_df(dataframe, PREFERRED_COLUMNS)
    dataframe = normalize_values(dataframe)
    dataframe = dataframe[PREFERRED_COLUMNS]
    _validate_reconciliation_fields(dataframe, label)
    return dataframe


def reconcile(
    pur: pd.DataFrame,
    twob: pd.DataFrame,
    similarity_threshold: int = 80,
    numeric_tolerance: float = 0.01,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    pur = pur.copy()
    twob = twob.copy()

    pur["Best score"] = pd.NA
    pur["Best match 2B index"] = pd.NA
    pur["Probable 2B indexes"] = pd.NA

    twob["Best score"] = pd.NA
    twob["Best match PR index"] = pd.NA
    twob["Probable PR indexes"] = pd.NA

    matched_2b_indexes = set()

    twob_invoice_dates = pd.to_datetime(
        twob["Invoice Date"], format="%Y%m%d", errors="coerce"
    )

    # Pre-split GSTR-2B rows by GSTIN once so each Purchase Register row only
    # scans the rows that could plausibly match it, instead of re-masking the
    # entire GSTR-2B table on every iteration (that O(n*m) scan was slow
    # enough on real-sized uploads to trip the web server's request timeout).
    twob_by_gstin = {gstin: group for gstin, group in twob.groupby("GSTIN", sort=False)}
    twob_missing_gstin = twob[twob["GSTIN"].isna()]

    for pur_index, row in pur.iterrows():
        if pd.isna(row["GSTIN"]):
            candidate_pool = twob_missing_gstin
        else:
            candidate_pool = twob_by_gstin.get(row["GSTIN"])
            if candidate_pool is None:
                continue

        if candidate_pool.empty:
            continue

        row_invoice_date = pd.to_datetime(
            row["Invoice Date"], format="%Y%m%d", errors="coerce"
        )
        pool_dates = twob_invoice_dates.loc[candidate_pool.index]
        candidates = candidate_pool[
            ((pool_dates - row_invoice_date).abs() <= pd.Timedelta(days=date_tolerance_days))
            & (
                abs(candidate_pool["Taxable Value"] - row["Taxable Value"])
                <= numeric_tolerance * max(row["Taxable Value"], 1)
            )
        ]

        if candidates.empty:
            continue

        candidate_indexes = candidates.index.tolist()

        supplier_scores = process.cdist(
            [row["Supplier Name"]],
            candidates["Supplier Name"].tolist(),
            scorer=fuzz.WRatio,
            dtype=np.float64,
        )[0]
        invoice_scores = process.cdist(
            [str(row["Invoice No"])],
            [str(value) for value in candidates["Invoice No"].tolist()],
            scorer=fuzz.WRatio,
            dtype=np.float64,
        )[0]
        scores = 0.5 * supplier_scores + 0.5 * invoice_scores

        eligible = np.array(
            [candidate_index not in matched_2b_indexes for candidate_index in candidate_indexes]
        )
        eligible_scores = np.where(eligible, scores, -1.0)

        best_score = 0
        best_candidate_index = None
        if eligible_scores.size and eligible_scores.max() > 0:
            best_position = int(np.argmax(eligible_scores))
            best_score = float(eligible_scores[best_position])
            best_candidate_index = candidate_indexes[best_position]

        if best_candidate_index is not None and best_score >= similarity_threshold:
            pur.at[pur_index, "Best score"] = best_score
            pur.at[pur_index, "Best match 2B index"] = best_candidate_index

            twob.at[best_candidate_index, "Best score"] = best_score
            twob.at[best_candidate_index, "Best match PR index"] = pur_index
            matched_2b_indexes.add(best_candidate_index)
        else:
            pur.at[pur_index, "Probable 2B indexes"] = candidate_indexes
            for candidate_index in candidate_indexes:
                existing = twob.at[candidate_index, "Probable PR indexes"]
                if isinstance(existing, list):
                    existing.append(pur_index)
                elif pd.isna(existing):
                    twob.at[candidate_index, "Probable PR indexes"] = [pur_index]

    return pur, twob


def run_reconciliation(
    pr_path: str | Path,
    gstr2b_path: str | Path,
    pr_mapping: dict[str, str | None] | None = None,
    gstr2b_mapping: dict[str, str | None] | None = None,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    purchase_register = prepare_dataframe(pr_path, "Purchase Register", pr_mapping)
    gstr2b = prepare_dataframe(gstr2b_path, "GSTR-2B", gstr2b_mapping)

    logger.info(
        "reconciliation starting: pr_rows=%d gstr2b_rows=%d date_tolerance_days=%d",
        len(purchase_register), len(gstr2b), date_tolerance_days,
    )
    start = time.perf_counter()
    pur_result, gstr2b_result = reconcile(
        purchase_register, gstr2b, date_tolerance_days=date_tolerance_days
    )
    elapsed = time.perf_counter() - start

    matched = int(pur_result["Best match 2B index"].notna().sum())
    probable_only = int(pur_result["Probable 2B indexes"].notna().sum())
    logger.info(
        "reconciliation finished: pr_rows=%d gstr2b_rows=%d matched=%d probable_only=%d elapsed=%.2fs",
        len(purchase_register), len(gstr2b), matched, probable_only, elapsed,
    )
    return pur_result, gstr2b_result
