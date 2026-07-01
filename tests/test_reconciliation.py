import tempfile
import unittest
from pathlib import Path

import pandas as pd

from reconciliation import ReconciliationInputError, run_reconciliation

HEADER = (
    "Supplier Name,GSTIN,Invoice No,Invoice Date,HSN Code (optional),"
    "Taxable Value,IGST,CGST,SGST,Total GST\n"
)


def _row(supplier, gstin, invoice_no, date, taxable, igst=0, cgst=90, sgst=90, total_gst=180):
    return f"{supplier},{gstin},{invoice_no},{date},1001,{taxable},{igst},{cgst},{sgst},{total_gst}"


def _write_csv(directory: Path, filename: str, header: str, rows: list[str]) -> Path:
    path = directory / filename
    path.write_text(header + "\n".join(rows), encoding="utf-8")
    return path


class RunReconciliationEndToEndTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_exact_match_scores_100_and_links_both_sides(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertEqual(pur.loc[0, "Best score"], 100.0)
        self.assertEqual(pur.loc[0, "Best match 2B index"], 0)
        self.assertTrue(pd.isna(pur.loc[0, "Probable 2B indexes"]))
        self.assertEqual(twob.loc[0, "Best match PR index"], 0)

    def test_fuzzy_supplier_name_still_matches_above_threshold(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Suppliers Pvt Ltd", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertGreaterEqual(pur.loc[0, "Best score"], 80)
        self.assertEqual(pur.loc[0, "Best match 2B index"], 0)

    def test_invoice_date_within_10_days_matches_beyond_does_not(self):
        # Two unrelated GSTINs keep the pairs isolated so neither can
        # cross-match the other's 2B row.
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
            _row("Acme Supplier", "33BBBBB1111B2Z6", "INV-002", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "11/06/2026", 1000),  # exactly 10 days later
            _row("Acme Supplier", "33BBBBB1111B2Z6", "INV-002", "12/06/2026", 1000),  # 11 days later
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertEqual(pur.loc[0, "Best match 2B index"], 0)
        self.assertTrue(pd.isna(pur.loc[1, "Best match 2B index"]))
        self.assertTrue(pd.isna(pur.loc[1, "Probable 2B indexes"]))

    def test_taxable_value_tolerance_boundary(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-002", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1010),  # within 1% tolerance
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-002", "01/06/2026", 1020),  # beyond 1% tolerance
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertEqual(pur.loc[0, "Best match 2B index"], 0)
        self.assertTrue(pd.isna(pur.loc[1, "Best match 2B index"]))

    def test_weak_match_is_recorded_as_probable_on_both_sides(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Globex Corp", "22AAAAA0000A1Z5", "INV-999", "01/06/2026", 1000),
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertTrue(pd.isna(pur.loc[0, "Best match 2B index"]))
        self.assertEqual(pur.loc[0, "Probable 2B indexes"], [0])
        self.assertEqual(twob.loc[0, "Probable PR indexes"], [0])

    def test_different_gstin_produces_no_candidates(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "33BBBBB1111B2Z6", "INV-001", "01/06/2026", 1000),
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertTrue(pd.isna(pur.loc[0, "Best match 2B index"]))
        self.assertTrue(pd.isna(pur.loc[0, "Probable 2B indexes"]))

    def test_second_best_candidate_is_left_unmatched(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),  # exact match
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "02/06/2026", 1000),  # also a candidate
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertEqual(pur.loc[0, "Best match 2B index"], 0)
        self.assertTrue(pd.isna(twob.loc[1, "Best match PR index"]))

    def test_header_name_variants_are_reconciled_end_to_end(self):
        alt_header = (
            "SUPPLIER_NAME,Gstin,Invoice_No,Invoice_Date,HSN_Code_optional,"
            "TAXABLE_VALUE,igst,Cgst,sgst,TOTAL_GST\n"
        )
        pr = _write_csv(self.dir, "pr.csv", alt_header, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])

        pur, twob = run_reconciliation(pr, gstr2b)

        self.assertEqual(pur.loc[0, "Best score"], 100.0)
        self.assertEqual(pur.loc[0, "Best match 2B index"], 0)

    def test_unsupported_extension_raises(self):
        pr = self.dir / "purchase.txt"
        pr.write_text("not a real file", encoding="utf-8")
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])

        with self.assertRaises(ReconciliationInputError):
            run_reconciliation(pr, gstr2b)

    def test_non_numeric_taxable_value_raises(self):
        pr = _write_csv(self.dir, "pr.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", "not-a-number"),
        ])
        gstr2b = _write_csv(self.dir, "2b.csv", HEADER, [
            _row("Acme Supplier", "22AAAAA0000A1Z5", "INV-001", "01/06/2026", 1000),
        ])

        with self.assertRaises(ReconciliationInputError):
            run_reconciliation(pr, gstr2b)


if __name__ == "__main__":
    unittest.main()
