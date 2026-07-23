"""Repeatable reconciliation benchmarks for 10k/100k-row workloads."""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
import time
from pathlib import Path

import polars as pl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from reconciliation import PREFERRED_COLUMNS, run_reconciliation_report
from workbook_export import write_reconciliation_workbook


def build_frame(rows: int, workload: str) -> pl.DataFrame:
    frame = pl.DataFrame({"index": range(rows)})
    supplier = (
        pl.lit("Acme Supplier")
        if workload != "fuzzy"
        else pl.when(pl.col("index") % 2 == 0)
        .then(pl.lit("Acme Supplier")).otherwise(pl.lit("Acme Suppliers"))
    )
    return frame.with_columns(
        supplier.alias("Supplier Name"),
        pl.lit("22AAAAA0000A1Z5").alias("GSTIN"),
        pl.col("index").cast(pl.String).str.pad_start(8, "0").alias("Invoice No"),
        pl.lit("01/06/2026").alias("Invoice Date"),
        pl.lit("1001").alias("HSN Code (optional)"),
        (pl.col("index").cast(pl.Float64) + 1000).alias("Taxable Value"),
        pl.lit(0).alias("IGST"),
        pl.lit(90).alias("CGST"),
        pl.lit(90).alias("SGST"),
        pl.lit(180).alias("Total GST"),
    ).select(PREFERRED_COLUMNS)


def peak_rss_mb() -> float | None:
    if os.name != "nt":
        import resource
        return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024
    try:
        import ctypes
        from ctypes import wintypes

        class Counters(ctypes.Structure):
            _fields_ = [
                ("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t), ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]
        counters = Counters()
        counters.cb = ctypes.sizeof(counters)
        get_process = ctypes.windll.kernel32.GetCurrentProcess
        get_process.restype = wintypes.HANDLE
        get_memory = ctypes.windll.psapi.GetProcessMemoryInfo
        get_memory.argtypes = [wintypes.HANDLE, ctypes.POINTER(Counters), wintypes.DWORD]
        get_memory.restype = wintypes.BOOL
        if not get_memory(get_process(), ctypes.byref(counters), counters.cb):
            return None
        return counters.PeakWorkingSetSize / (1024 * 1024)
    except Exception:
        return None


def run(rows: int, workload: str, export: bool) -> None:
    with tempfile.TemporaryDirectory() as temporary:
        directory = Path(temporary)
        pr = build_frame(rows, workload)
        twob = build_frame(rows, workload)
        if workload == "unmatched":
            twob = twob.with_columns(pl.lit("33BBBBB1111B2Z6").alias("GSTIN"))
        elif workload == "fuzzy":
            twob = twob.with_columns(
                (pl.col("Supplier Name") + pl.lit(" Pvt Ltd")).alias("Supplier Name"),
                (pl.col("Invoice No") + pl.lit("-A")).alias("Invoice No"),
            )
        pr_path, twob_path = directory / "pr.csv", directory / "2b.csv"
        pr.write_csv(pr_path)
        twob.write_csv(twob_path)
        started = time.perf_counter()
        report = run_reconciliation_report(pr_path, twob_path)
        reconciliation_seconds = time.perf_counter() - started
        workbook_seconds = None
        workbook_size = None
        if export:
            started = time.perf_counter()
            output = write_reconciliation_workbook(report, directory / "report.xlsx")
            workbook_seconds = time.perf_counter() - started
            workbook_size = output.stat().st_size
        peak_mb = peak_rss_mb()
        print({
            "rows_per_register": rows,
            "workload": workload,
            "reconciliation_seconds": round(reconciliation_seconds, 3),
            "workbook_seconds": None if workbook_seconds is None else round(workbook_seconds, 3),
            "workbook_bytes": workbook_size,
            "peak_rss_mb": None if peak_mb is None else round(peak_mb, 1),
        })


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, choices=(10_000, 100_000), default=10_000)
    parser.add_argument("--workload", choices=("exact", "fuzzy", "unmatched"), default="exact")
    parser.add_argument("--export", action="store_true")
    args = parser.parse_args()
    run(args.rows, args.workload, args.export)
