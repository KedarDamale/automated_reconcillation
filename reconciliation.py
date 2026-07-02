"""Reconciliation logic lifted from automated_reconciliation.ipynb."""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz, process

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


def normalize_col_names(column_name: object) -> str:
    return (
        str(column_name)
        .lower()
        .strip()
        .replace("_", " ")
        .replace("-", " ")
        .replace(".", "")
    )


def extract_preferred_columns_from_df(
    df: pd.DataFrame,
    preferred_columns: list[str],
    threshold: int = 95,
) -> tuple[pd.DataFrame, dict]:
    df_cols = list(df.columns)
    df_cols_norm = {column: normalize_col_names(column) for column in df_cols}
    preferred_cols_norm = {
        preferred: normalize_col_names(preferred) for preferred in preferred_columns
    }

    used_source_columns = set()
    rename_map: dict = {}

    for preferred_column, preferred_norm in preferred_cols_norm.items():
        candidates = {
            source_column: normalized_column
            for source_column, normalized_column in df_cols_norm.items()
            if source_column not in used_source_columns
        }

        if not candidates:
            rename_map[preferred_column] = None
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
            rename_map[matched_source_column] = preferred_column
            used_source_columns.add(matched_source_column)
        else:
            rename_map[preferred_column] = None

    final_rename_map = {
        source: target
        for source, target in rename_map.items()
        if source in df.columns and target is not None
    }
    renamed_df = df.rename(columns=final_rename_map)

    for column in preferred_columns:
        if column not in renamed_df.columns:
            renamed_df[column] = None

    return renamed_df[preferred_columns], rename_map


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
    required = ["Supplier Name", "GSTIN", "Invoice No", "Invoice Date", "Taxable Value"]
    empty = [column for column in required if df[column].isna().all()]
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


def prepare_dataframe(file_path: str | Path, label: str) -> pd.DataFrame:
    dataframe = handle_multitype_import(file_path)
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
            ((pool_dates - row_invoice_date).abs() <= pd.Timedelta(days=10))
            & (
                abs(candidate_pool["Taxable Value"] - row["Taxable Value"])
                <= numeric_tolerance * max(row["Taxable Value"], 1)
            )
        ]

        if candidates.empty:
            continue

        best_score = 0
        best_candidate_index = None
        candidate_indexes = []

        for candidate_index, candidate in candidates.iterrows():
            candidate_indexes.append(candidate_index)
            if candidate_index in matched_2b_indexes:
                continue

            score = 0.5 * fuzz.WRatio(
                row["Supplier Name"], candidate["Supplier Name"]
            ) + 0.5 * fuzz.WRatio(str(row["Invoice No"]), str(candidate["Invoice No"]))

            if score > best_score:
                best_score = score
                best_candidate_index = candidate_index

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
    pr_path: str | Path, gstr2b_path: str | Path
) -> tuple[pd.DataFrame, pd.DataFrame]:
    purchase_register = prepare_dataframe(pr_path, "Purchase Register")
    gstr2b = prepare_dataframe(gstr2b_path, "GSTR-2B")
    return reconcile(purchase_register, gstr2b)
