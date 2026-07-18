"""Reconciliation logic lifted from automated_reconciliation.ipynb."""

from __future__ import annotations

import logging
import os
import time
from collections import defaultdict, deque
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
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


@dataclass(frozen=True)
class MatchPair:
    """A one-to-one link between a Purchase Register and GSTR-2B row."""

    pr_index: int
    gstr2b_index: int
    score: float


@dataclass
class ReconciliationReport:
    """All raw and classified data needed by the web UI and workbook export."""

    pr_raw: pd.DataFrame
    gstr2b_raw: pd.DataFrame
    pr_result: pd.DataFrame
    gstr2b_result: pd.DataFrame
    matched_pairs: list[MatchPair]
    date_only_pairs: list[MatchPair]
    pr_missing_gstin_indexes: list[int]
    pr_unmatched_indexes: list[int]
    gstr2b_unmatched_indexes: list[int]
    pr_probable_matches: dict[int, MatchPair]
    gstr2b_probable_matches: dict[int, MatchPair]


class ReconciliationInputError(ValueError):
    """Raised when an uploaded file cannot safely enter the notebook pipeline."""


def handle_multitype_import(
    file_path: str | Path,
    *,
    preserve_raw_values: bool = False,
) -> pd.DataFrame:
    if file_path is None:
        raise ReconciliationInputError("No file was provided.")

    extension = os.path.splitext(str(file_path))[1].lower()
    raw_options = {"dtype": object, "keep_default_na": False} if preserve_raw_values else {}
    try:
        if extension == ".csv":
            return pd.read_csv(file_path, **raw_options)
        if extension == ".xlsx":
            return pd.read_excel(file_path, **raw_options)
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
    mapping: dict[str, str | list[str] | None],
    preferred_columns: list[str] = PREFERRED_COLUMNS,
    label: str = "File",
) -> pd.DataFrame:
    """Rename ``df``'s columns using a person-confirmed preferred -> source mapping.

    This is deliberately strict: a mapping a user assembled by hand (however
    much they rearranged it first) is trusted verbatim, but if it is
    inconsistent -- it references a column that doesn't exist, reuses one
    source column for two preferred fields, or leaves a required field
    unmapped -- reconciliation must not silently guess. ``Taxable Value`` is
    the single exception to one-source-per-field: it may contain a list of
    source columns, which are summed row by row.
    """
    unknown_preferred = sorted(set(mapping) - set(preferred_columns))
    if unknown_preferred:
        raise ReconciliationInputError(
            f"{label} column mapping references unknown field(s): {', '.join(unknown_preferred)}."
        )

    used_sources: dict[str, str] = {}
    mapped_sources: dict[str, str] = {}
    taxable_value_sources: list[str] = []
    for preferred_column in preferred_columns:
        source_value = mapping.get(preferred_column)
        if preferred_column == "Taxable Value":
            source_columns = _taxable_value_sources(source_value, label)
        else:
            if isinstance(source_value, (list, tuple)):
                raise ReconciliationInputError(
                    f"{label} can only map multiple source columns to Taxable Value."
                )
            source_columns = [source_value] if source_value else []

        if not source_columns:
            continue

        for source_column in source_columns:
            if not isinstance(source_column, str) or not source_column:
                raise ReconciliationInputError(
                    f"{label} column mapping for {preferred_column} was malformed."
                )
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

        if preferred_column == "Taxable Value":
            taxable_value_sources = source_columns
        else:
            mapped_sources[source_columns[0]] = preferred_column

    missing_required = [
        column
        for column in REQUIRED_COLUMNS
        if column in preferred_columns and column not in used_sources.values()
    ]
    if missing_required:
        raise ReconciliationInputError(
            f"{label} is missing a mapping for required column(s): {', '.join(missing_required)}."
        )

    renamed_df = df.rename(columns=mapped_sources)
    if taxable_value_sources:
        renamed_df["Taxable Value"] = _sum_taxable_value_columns(
            df, taxable_value_sources, label
        )

    for column in preferred_columns:
        if column not in renamed_df.columns:
            renamed_df[column] = None

    return renamed_df[preferred_columns]


def _taxable_value_sources(
    source_value: str | list[str] | None,
    label: str,
) -> list[str]:
    if source_value is None or source_value == "":
        return []
    if isinstance(source_value, str):
        return [source_value]
    if not isinstance(source_value, (list, tuple)):
        raise ReconciliationInputError(
            f"{label} column mapping for Taxable Value was malformed."
        )
    if not source_value:
        return []
    if any(not isinstance(source_column, str) or not source_column for source_column in source_value):
        raise ReconciliationInputError(
            f"{label} column mapping for Taxable Value was malformed."
        )
    if len(set(source_value)) != len(source_value):
        raise ReconciliationInputError(
            f"{label} column mapping selects the same Taxable Value column more than once."
        )
    return list(source_value)


def _sum_taxable_value_columns(
    df: pd.DataFrame,
    source_columns: list[str],
    label: str,
) -> pd.Series:
    numeric_columns: list[pd.Series] = []
    for source_column in source_columns:
        source = df[source_column]
        numeric = pd.to_numeric(source, errors="coerce")
        present = source.notna() & source.astype(str).str.strip().ne("")
        invalid_count = int((present & numeric.isna()).sum())
        if invalid_count:
            raise ReconciliationInputError(
                f'{label} column "{source_column}" contains {invalid_count} non-numeric '
                "row(s) selected for Taxable Value."
            )
        numeric_columns.append(numeric)

    numeric_frame = pd.concat(numeric_columns, axis=1)
    total = numeric_frame.sum(axis=1, min_count=1)
    return total


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
    if _is_missing_value(value):
        return None
    return str(value).lower().strip()


def normalize_gstin(value):
    if _is_missing_value(value):
        return None
    return str(value).strip().upper()


def normalize_invoice_no(value):
    if _is_missing_value(value):
        return None
    return str(value).strip().lower()


def normalize_invoice_date(series: pd.Series) -> pd.Series:
    date = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return date.dt.year * 10000 + date.dt.month * 100 + date.dt.day


def normalize_hsn_code(value):
    if _is_missing_value(value):
        return None
    return str(value).strip()


def normalize_numeric_col(value):
    if _is_missing_value(value):
        return None
    return value


def _is_missing_value(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


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
    column_mapping: dict[str, str | list[str] | None] | None = None,
) -> pd.DataFrame:
    dataframe = handle_multitype_import(file_path, preserve_raw_values=True)
    return prepare_dataframe_from_raw(dataframe, label, column_mapping)


def prepare_dataframe_from_raw(
    dataframe: pd.DataFrame,
    label: str,
    column_mapping: dict[str, str | list[str] | None] | None = None,
) -> pd.DataFrame:
    """Create a normalized reconciliation copy without changing raw upload data."""
    dataframe = dataframe.copy()
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
    report = _classify_reconciliation(
        pur,
        twob,
        pr_raw=pur,
        gstr2b_raw=twob,
        similarity_threshold=similarity_threshold,
        numeric_tolerance=numeric_tolerance,
        date_tolerance_days=date_tolerance_days,
    )
    return report.pr_result, report.gstr2b_result


def _classify_reconciliation(
    pur: pd.DataFrame,
    twob: pd.DataFrame,
    *,
    pr_raw: pd.DataFrame,
    gstr2b_raw: pd.DataFrame,
    similarity_threshold: int = 80,
    numeric_tolerance: float = 0.01,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> ReconciliationReport:
    pur = pur.copy()
    twob = twob.copy()

    pur["Best score"] = pd.NA
    pur["Best match 2B index"] = pd.NA
    pur["Probable 2B indexes"] = pd.NA
    pur["Best probable score"] = pd.NA
    pur["Best probable 2B index"] = pd.NA
    pur["Match category"] = "Unmatched"

    twob["Best score"] = pd.NA
    twob["Best match PR index"] = pd.NA
    twob["Probable PR indexes"] = pd.NA
    twob["Best probable score"] = pd.NA
    twob["Best probable PR index"] = pd.NA
    twob["Match category"] = "Unmatched"

    pr_missing_gstin_indexes = [
        int(index) for index in pur.index if _is_missing_value(pur.at[index, "GSTIN"])
    ]
    for index in pr_missing_gstin_indexes:
        pur.at[index, "Match category"] = "Missing GSTIN"

    missing_gstin_set = set(pr_missing_gstin_indexes)
    eligible_pr_indexes = [
        int(index) for index in pur.index if index not in missing_gstin_set
    ]
    available_pr_indexes = set(eligible_pr_indexes)
    available_2b_indexes = {int(index) for index in twob.index}
    twob_indexes_by_gstin = {
        gstin: [int(index) for index in group.index]
        for gstin, group in twob.groupby("GSTIN", sort=False)
    }

    matched_pairs: list[MatchPair] = []
    date_only_pairs: list[MatchPair] = []

    exact_pairs = _reserve_full_exact_pairs(
        pur, twob, eligible_pr_indexes, available_2b_indexes
    )
    for pair in exact_pairs:
        _record_pair(pur, twob, pair, "Matched")
        matched_pairs.append(pair)
        available_pr_indexes.discard(pair.pr_index)
        available_2b_indexes.discard(pair.gstr2b_index)

    reserved_date_pairs = _reserve_date_only_pairs(
        pur,
        twob,
        sorted(available_pr_indexes),
        available_2b_indexes,
        date_tolerance_days,
    )
    for pair in reserved_date_pairs:
        _record_pair(pur, twob, pair, "Date only")
        date_only_pairs.append(pair)
        available_pr_indexes.discard(pair.pr_index)
        available_2b_indexes.discard(pair.gstr2b_index)

    # Preserve the existing deterministic, Purchase-Register-order greedy
    # matching behavior for all rows not handled by the exact-match passes.
    for pur_index in sorted(available_pr_indexes):
        candidates = _candidate_rows(
            pur.loc[pur_index],
            twob,
            available_2b_indexes,
            numeric_tolerance,
            date_tolerance_days,
            twob_indexes_by_gstin,
        )
        scored_candidates = _score_candidates(pur.loc[pur_index], candidates)
        if not scored_candidates:
            continue

        best_candidate_index, best_score = max(
            scored_candidates, key=lambda candidate: (candidate[1], -candidate[0])
        )
        if best_score >= similarity_threshold:
            pair = MatchPair(pur_index, best_candidate_index, float(best_score))
            _record_pair(pur, twob, pair, "Matched")
            matched_pairs.append(pair)
            available_pr_indexes.discard(pur_index)
            available_2b_indexes.discard(best_candidate_index)

    pr_unmatched_indexes = sorted(available_pr_indexes)
    gstr2b_unmatched_indexes = sorted(available_2b_indexes)
    pr_probable_matches, gstr2b_probable_matches = _record_probable_matches(
        pur,
        twob,
        pr_unmatched_indexes,
        gstr2b_unmatched_indexes,
        similarity_threshold,
        numeric_tolerance,
        date_tolerance_days,
        twob_indexes_by_gstin,
    )

    return ReconciliationReport(
        pr_raw=pr_raw.copy(),
        gstr2b_raw=gstr2b_raw.copy(),
        pr_result=pur,
        gstr2b_result=twob,
        matched_pairs=matched_pairs,
        date_only_pairs=date_only_pairs,
        pr_missing_gstin_indexes=pr_missing_gstin_indexes,
        pr_unmatched_indexes=pr_unmatched_indexes,
        gstr2b_unmatched_indexes=gstr2b_unmatched_indexes,
        pr_probable_matches=pr_probable_matches,
        gstr2b_probable_matches=gstr2b_probable_matches,
    )


def _reserve_full_exact_pairs(
    pur: pd.DataFrame,
    twob: pd.DataFrame,
    pr_indexes: list[int],
    available_2b_indexes: set[int],
) -> list[MatchPair]:
    twob_by_key: dict[tuple, deque[int]] = defaultdict(deque)
    for index in sorted(available_2b_indexes):
        row = twob.loc[index]
        if _has_required_candidate_values(row):
            twob_by_key[_row_comparison_key(row, PREFERRED_COLUMNS)].append(index)

    pairs: list[MatchPair] = []
    for pr_index in pr_indexes:
        row = pur.loc[pr_index]
        if not _has_required_candidate_values(row):
            continue
        matches = twob_by_key.get(_row_comparison_key(row, PREFERRED_COLUMNS))
        if matches:
            pairs.append(MatchPair(pr_index, matches.popleft(), 100.0))
    return pairs


def _reserve_date_only_pairs(
    pur: pd.DataFrame,
    twob: pd.DataFrame,
    pr_indexes: list[int],
    available_2b_indexes: set[int],
    date_tolerance_days: int,
) -> list[MatchPair]:
    if date_tolerance_days <= 0:
        return []

    comparison_columns = [column for column in PREFERRED_COLUMNS if column != "Invoice Date"]
    twob_by_key: dict[tuple, list[int]] = defaultdict(list)
    for index in sorted(available_2b_indexes):
        row = twob.loc[index]
        if _has_required_candidate_values(row):
            twob_by_key[_row_comparison_key(row, comparison_columns)].append(index)

    pairs: list[MatchPair] = []
    used_2b_indexes: set[int] = set()
    for pr_index in pr_indexes:
        pr_row = pur.loc[pr_index]
        if not _has_required_candidate_values(pr_row):
            continue
        pr_date = _invoice_timestamp(pr_row["Invoice Date"])
        possible: list[tuple[int, int]] = []
        for gstr2b_index in twob_by_key.get(
            _row_comparison_key(pr_row, comparison_columns), []
        ):
            if gstr2b_index in used_2b_indexes:
                continue
            twob_date = _invoice_timestamp(twob.at[gstr2b_index, "Invoice Date"])
            if pr_date is None or twob_date is None:
                continue
            difference = abs((twob_date - pr_date).days)
            if 0 < difference <= date_tolerance_days:
                possible.append((difference, gstr2b_index))

        if possible:
            _, gstr2b_index = min(possible, key=lambda candidate: (candidate[0], candidate[1]))
            used_2b_indexes.add(gstr2b_index)
            pairs.append(MatchPair(pr_index, gstr2b_index, 100.0))
    return pairs


def _candidate_rows(
    row: pd.Series,
    twob: pd.DataFrame,
    available_2b_indexes: set[int] | list[int],
    numeric_tolerance: float,
    date_tolerance_days: int,
    twob_indexes_by_gstin: dict[object, list[int]],
) -> pd.DataFrame:
    if not available_2b_indexes or not _has_required_candidate_values(row):
        return twob.iloc[0:0]

    available = (
        available_2b_indexes
        if isinstance(available_2b_indexes, set)
        else set(available_2b_indexes)
    )
    candidate_indexes = [
        index
        for index in twob_indexes_by_gstin.get(row["GSTIN"], [])
        if index in available
    ]
    if not candidate_indexes:
        return twob.iloc[0:0]
    candidate_pool = twob.loc[candidate_indexes]

    row_date = _invoice_timestamp(row["Invoice Date"])
    pool_dates = pd.to_datetime(
        candidate_pool["Invoice Date"], format="%Y%m%d", errors="coerce"
    )
    taxable_values = pd.to_numeric(candidate_pool["Taxable Value"], errors="coerce")
    taxable_value = float(row["Taxable Value"])
    return candidate_pool[
        ((pool_dates - row_date).abs() <= pd.Timedelta(days=date_tolerance_days))
        & (
            abs(taxable_values - taxable_value)
            <= numeric_tolerance * max(taxable_value, 1)
        )
    ]


def _score_candidates(row: pd.Series, candidates: pd.DataFrame) -> list[tuple[int, float]]:
    if candidates.empty:
        return []

    supplier_scores = process.cdist(
        [_score_text(row["Supplier Name"])],
        [_score_text(value) for value in candidates["Supplier Name"].tolist()],
        scorer=fuzz.WRatio,
        dtype=np.float64,
    )[0]
    invoice_scores = process.cdist(
        [_score_text(row["Invoice No"])],
        [_score_text(value) for value in candidates["Invoice No"].tolist()],
        scorer=fuzz.WRatio,
        dtype=np.float64,
    )[0]
    scores = 0.5 * supplier_scores + 0.5 * invoice_scores
    return [
        (int(index), float(score))
        for index, score in zip(candidates.index.tolist(), scores, strict=True)
    ]


def _record_probable_matches(
    pur: pd.DataFrame,
    twob: pd.DataFrame,
    pr_unmatched_indexes: list[int],
    gstr2b_unmatched_indexes: list[int],
    similarity_threshold: int,
    numeric_tolerance: float,
    date_tolerance_days: int,
    twob_indexes_by_gstin: dict[object, list[int]],
) -> tuple[dict[int, MatchPair], dict[int, MatchPair]]:
    pr_probable_matches: dict[int, MatchPair] = {}
    twob_candidates: dict[int, list[MatchPair]] = defaultdict(list)

    for pr_index in pr_unmatched_indexes:
        candidates = _candidate_rows(
            pur.loc[pr_index],
            twob,
            gstr2b_unmatched_indexes,
            numeric_tolerance,
            date_tolerance_days,
            twob_indexes_by_gstin,
        )
        below_threshold = [
            MatchPair(pr_index, candidate_index, score)
            for candidate_index, score in _score_candidates(pur.loc[pr_index], candidates)
            if score < similarity_threshold
        ]
        below_threshold.sort(key=lambda pair: (-pair.score, pair.gstr2b_index))
        if not below_threshold:
            continue

        best = below_threshold[0]
        pr_probable_matches[pr_index] = best
        pur.at[pr_index, "Probable 2B indexes"] = [
            pair.gstr2b_index for pair in below_threshold
        ]
        pur.at[pr_index, "Best probable score"] = best.score
        pur.at[pr_index, "Best probable 2B index"] = best.gstr2b_index
        for pair in below_threshold:
            twob_candidates[pair.gstr2b_index].append(pair)

    gstr2b_probable_matches: dict[int, MatchPair] = {}
    for gstr2b_index, candidates in twob_candidates.items():
        candidates.sort(key=lambda pair: (-pair.score, pair.pr_index))
        best = candidates[0]
        gstr2b_probable_matches[gstr2b_index] = best
        twob.at[gstr2b_index, "Probable PR indexes"] = [
            pair.pr_index for pair in candidates
        ]
        twob.at[gstr2b_index, "Best probable score"] = best.score
        twob.at[gstr2b_index, "Best probable PR index"] = best.pr_index

    return pr_probable_matches, gstr2b_probable_matches


def _record_pair(
    pur: pd.DataFrame,
    twob: pd.DataFrame,
    pair: MatchPair,
    category: str,
) -> None:
    pur.at[pair.pr_index, "Best score"] = pair.score
    pur.at[pair.pr_index, "Best match 2B index"] = pair.gstr2b_index
    pur.at[pair.pr_index, "Match category"] = category
    twob.at[pair.gstr2b_index, "Best score"] = pair.score
    twob.at[pair.gstr2b_index, "Best match PR index"] = pair.pr_index
    twob.at[pair.gstr2b_index, "Match category"] = category


def _row_comparison_key(row: pd.Series, columns: list[str]) -> tuple:
    return tuple(_comparison_value(column, row[column]) for column in columns)


def _comparison_value(column: str, value: object):
    if _is_missing_value(value):
        return None
    if column in {"Invoice Date", "Taxable Value", "IGST", "CGST", "SGST", "Total GST"}:
        try:
            return Decimal(str(value)).normalize()
        except (InvalidOperation, ValueError):
            return str(value).strip()
    return str(value).strip()


def _has_required_candidate_values(row: pd.Series) -> bool:
    return all(
        not _is_missing_value(row[column])
        for column in ("GSTIN", "Invoice Date", "Taxable Value")
    )


def _invoice_timestamp(value: object) -> pd.Timestamp | None:
    if _is_missing_value(value):
        return None
    try:
        compact_date = str(int(float(value)))
    except (TypeError, ValueError, OverflowError):
        return None
    timestamp = pd.to_datetime(compact_date, format="%Y%m%d", errors="coerce")
    return None if pd.isna(timestamp) else timestamp


def _score_text(value: object) -> str:
    return "" if _is_missing_value(value) else str(value)


def run_reconciliation_report(
    pr_path: str | Path,
    gstr2b_path: str | Path,
    pr_mapping: dict[str, str | list[str] | None] | None = None,
    gstr2b_mapping: dict[str, str | list[str] | None] | None = None,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> ReconciliationReport:
    pr_raw = handle_multitype_import(pr_path, preserve_raw_values=True).reset_index(drop=True)
    gstr2b_raw = handle_multitype_import(
        gstr2b_path, preserve_raw_values=True
    ).reset_index(drop=True)
    purchase_register = prepare_dataframe_from_raw(
        pr_raw, "Purchase Register", pr_mapping
    )
    gstr2b = prepare_dataframe_from_raw(gstr2b_raw, "GSTR-2B", gstr2b_mapping)

    logger.info(
        "reconciliation starting: pr_rows=%d gstr2b_rows=%d date_tolerance_days=%d",
        len(purchase_register), len(gstr2b), date_tolerance_days,
    )
    start = time.perf_counter()
    report = _classify_reconciliation(
        purchase_register,
        gstr2b,
        pr_raw=pr_raw,
        gstr2b_raw=gstr2b_raw,
        date_tolerance_days=date_tolerance_days,
    )
    elapsed = time.perf_counter() - start

    logger.info(
        "reconciliation finished: pr_rows=%d gstr2b_rows=%d matched=%d date_only=%d "
        "pr_unmatched=%d gstr2b_unmatched=%d elapsed=%.2fs",
        len(purchase_register),
        len(gstr2b),
        len(report.matched_pairs),
        len(report.date_only_pairs),
        len(report.pr_unmatched_indexes),
        len(report.gstr2b_unmatched_indexes),
        elapsed,
    )
    return report


def run_reconciliation(
    pr_path: str | Path,
    gstr2b_path: str | Path,
    pr_mapping: dict[str, str | list[str] | None] | None = None,
    gstr2b_mapping: dict[str, str | list[str] | None] | None = None,
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    report = run_reconciliation_report(
        pr_path,
        gstr2b_path,
        pr_mapping=pr_mapping,
        gstr2b_mapping=gstr2b_mapping,
        date_tolerance_days=date_tolerance_days,
    )
    return report.pr_result, report.gstr2b_result
