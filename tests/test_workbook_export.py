import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

import fastexcel
import polars as pl

from reconciliation import INVALID_GSTIN, MISSING_GSTIN, run_reconciliation_report
from workbook_export import SHEET_NAMES, write_reconciliation_workbook


HEADER = (
    "Supplier Name,GSTIN,Invoice No,Invoice Date,HSN Code (optional),"
    "Taxable Value,IGST,CGST,SGST,Total GST\n"
)


def sheet(path: Path, name: str, *, has_header: bool = True) -> pl.DataFrame:
    return pl.read_excel(
        path, sheet_name=name, has_header=has_header,
        drop_empty_rows=False, drop_empty_cols=False, infer_schema_length=None,
    )


def row(supplier, gstin, invoice, invoice_date, hsn="0012", taxable=1000,
        igst=0, cgst=90, sgst=90, total=180):
    return (
        f"{supplier},{gstin},{invoice},{invoice_date},{hsn},"
        f"{taxable},{igst},{cgst},{sgst},{total}"
    )


class NineSheetWorkbookTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def _write_input(self, name, rows):
        path = self.directory / name
        path.write_text(HEADER + "\n".join(rows), encoding="utf-8")
        return path

    def _build_report(self, *, date_tolerance_days=10):
        purchase_register = self._write_input(
            "purchase.csv",
            [
                row("ACME Exact", "22AAAAA0000A1Z5", "000123", "01/06/2026"),
                row("Date Only", "33BBBBB1111B2Z6", "DATE-1", "01/06/2026"),
                row("Purchase Weak", "44CCCCC2222C3Z7", "PR-001", "03/06/2026"),
                row("No GSTIN", " ", "NULL-1", "04/06/2026"),
                row("PR Only", "55DDDDD3333D4Z8", "ONLY-PR", "05/06/2026"),
            ],
        )
        gstr2b = self._write_input(
            "gstr2b.csv",
            [
                row("ACME Exact", "22AAAAA0000A1Z5", "000123", "01/06/2026"),
                row("Date Only", "33BBBBB1111B2Z6", "DATE-1", "06/06/2026"),
                row("Unrelated Vendor", "44CCCCC2222C3Z7", "OTHER-999", "03/06/2026"),
                row("2B Only", "66EEEEE4444E5Z9", "ONLY-2B", "07/06/2026"),
            ],
        )
        return run_reconciliation_report(
            purchase_register,
            gstr2b,
            date_tolerance_days=date_tolerance_days,
        )

    def test_report_categories_are_mutually_exclusive(self):
        report = self._build_report()

        self.assertEqual([(pair.pr_index, pair.gstr2b_index) for pair in report.matched_pairs], [(0, 0)])
        self.assertEqual([(pair.pr_index, pair.gstr2b_index) for pair in report.date_only_pairs], [(1, 1)])
        self.assertEqual(report.pr_missing_gstin_indexes, [3])
        self.assertEqual(report.pr_unmatched_indexes, [2, 4])
        self.assertEqual(report.gstr2b_unmatched_indexes, [2, 3])
        self.assertEqual(report.pr_probable_matches[2].gstr2b_index, 2)
        self.assertLess(report.pr_probable_matches[2].score, 80)
        self.assertNotIn(4, report.pr_probable_matches)
        self.assertEqual(report.gstr2b_probable_matches[2].pr_index, 2)
        self.assertNotIn(3, report.gstr2b_probable_matches)

    def test_raw_values_and_workbook_layout_are_preserved(self):
        report = self._build_report()
        self.assertEqual(report.pr_raw["Supplier Name"][0], "ACME Exact")
        self.assertEqual(report.pr_raw["Invoice No"][0], "000123")
        self.assertEqual(report.pr_raw["HSN Code (optional)"][0], "0012")

        output = write_reconciliation_workbook(report, self.directory / "report.xlsx")
        self.assertEqual(fastexcel.read_excel(output).sheet_names, SHEET_NAMES)

        original = sheet(output, SHEET_NAMES[0])
        self.assertEqual(original["Supplier Name"][0], "ACME Exact")
        self.assertEqual(original["Invoice No"][0], "000123")
        self.assertEqual(original["HSN Code (optional)"][0], "0012")
        self.assertEqual(original["Invoice Date"][0].isoformat(), "2026-06-01")
        self.assertEqual(original["Taxable Value"][0], 1000)

        missing = sheet(output, SHEET_NAMES[2])
        self.assertEqual(missing.row(0)[:5], ("Purchase Register", 3, MISSING_GSTIN, " ", "No GSTIN"))

        matched = sheet(output, SHEET_NAMES[3], has_header=False)
        self.assertEqual(matched.row(0)[0], "Purchase Register")
        self.assertEqual(matched.row(1)[11], "Match Score")
        self.assertEqual(int(matched.row(2)[0]), 0)
        self.assertEqual(float(matched.row(2)[11]), 100)
        self.assertEqual(int(matched.row(2)[12]), 0)

        pr_unmatched = sheet(output, SHEET_NAMES[4], has_header=False)
        self.assertEqual(int(pr_unmatched.row(2)[0]), 2)
        self.assertLess(float(pr_unmatched.row(2)[11]), 80)
        self.assertEqual(int(pr_unmatched.row(2)[12]), 2)
        self.assertEqual(int(pr_unmatched.row(3)[0]), 4)
        self.assertIsNone(pr_unmatched.row(3)[11])

        date_only = sheet(output, SHEET_NAMES[6], has_header=False)
        self.assertEqual(date_only.row(1)[11], "Date Difference (Days)")
        self.assertEqual(int(date_only.row(2)[0]), 1)
        self.assertEqual(int(date_only.row(2)[11]), 5)
        self.assertEqual(int(date_only.row(2)[12]), 1)

        pr_reconciled = sheet(output, SHEET_NAMES[7])
        self.assertEqual(pr_reconciled["Supplier Name"][0], "acme exact")
        self.assertIn("Best score", pr_reconciled.columns)
        self.assertEqual(pr_reconciled["Match category"][0], "Matched")
        gstr2b_reconciled = sheet(output, SHEET_NAMES[8])
        self.assertEqual(gstr2b_reconciled["Supplier Name"][0], "acme exact")
        self.assertIn("Best match PR index", gstr2b_reconciled.columns)
        self.assertIn("Probable PR indexes", gstr2b_reconciled.columns)

        with ZipFile(output) as archive:
            styles = archive.read("xl/styles.xml").decode()
        self.assertIn('formatCode="dd-mmm-yyyy"', styles)
        self.assertIn("thick", styles)

    def test_noisy_and_invalid_gstins_are_classified_once_for_review(self):
        purchase_register = self._write_input(
            "gstin-pr.csv",
            [
                row("Noisy GSTIN", "27ADZFS0848J1Z8_X000D_", "NOISY-1", "01/06/2026"),
                row("Blank GSTIN", " ", "BLANK-1", "02/06/2026"),
                row("Malformed GSTIN", "GST: NOT-VALID", "BAD-1", "03/06/2026"),
            ],
        )
        gstr2b = self._write_input(
            "gstin-2b.csv",
            [
                row("Noisy GSTIN", "27ADZFS0848J1Z8", "NOISY-1", "01/06/2026"),
                row("Invalid 2B", "27ADZFS0848J1Z8 / 22AAAAA0000A1Z5", "BAD-2", "04/06/2026"),
            ],
        )

        report = run_reconciliation_report(purchase_register, gstr2b)

        self.assertEqual(report.pr_result["GSTIN"][0], "27ADZFS0848J1Z8")
        self.assertEqual([(pair.pr_index, pair.gstr2b_index) for pair in report.matched_pairs], [(0, 0)])
        self.assertEqual(report.pr_gstin_issues, {1: MISSING_GSTIN, 2: INVALID_GSTIN})
        self.assertEqual(report.gstr2b_gstin_issues, {1: INVALID_GSTIN})
        self.assertEqual(report.pr_unmatched_indexes, [])
        self.assertEqual(report.gstr2b_unmatched_indexes, [])

        output = write_reconciliation_workbook(report, self.directory / "gstin-review.xlsx")
        review = sheet(output, SHEET_NAMES[2])
        rows = review.rows()
        self.assertEqual(
            [(row_values[0], row_values[1], row_values[2]) for row_values in rows],
            [
                ("Purchase Register", 1, MISSING_GSTIN),
                ("Purchase Register", 2, INVALID_GSTIN),
                ("GSTR-2B", 1, INVALID_GSTIN),
            ],
        )
        self.assertEqual(rows[0][3], " ")
        self.assertEqual(rows[1][3], "GST: NOT-VALID")
        self.assertEqual(rows[2][3], "27ADZFS0848J1Z8 / 22AAAAA0000A1Z5")
        self.assertEqual(sheet(output, SHEET_NAMES[4], has_header=False).height, 2)
        self.assertEqual(sheet(output, SHEET_NAMES[5], has_header=False).height, 2)

    def test_date_only_requires_all_other_fields_and_obeys_tolerance(self):
        report = self._build_report(date_tolerance_days=4)
        self.assertEqual(report.date_only_pairs, [])
        self.assertIn(1, report.pr_unmatched_indexes)
        self.assertIn(1, report.gstr2b_unmatched_indexes)

        pr = self._write_input(
            "tax-mismatch-pr.csv",
            [row("Tax Difference", "77FFFFF5555F6Z1", "TAX-1", "01/06/2026", cgst=90)],
        )
        twob = self._write_input(
            "tax-mismatch-2b.csv",
            [row("Tax Difference", "77FFFFF5555F6Z1", "TAX-1", "03/06/2026", cgst=91)],
        )
        tax_mismatch = run_reconciliation_report(pr, twob, date_tolerance_days=10)
        self.assertEqual(tax_mismatch.date_only_pairs, [])
        self.assertEqual(len(tax_mismatch.matched_pairs), 1)


if __name__ == "__main__":
    unittest.main()
