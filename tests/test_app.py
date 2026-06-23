import io
import json
import tempfile
import unittest
from pathlib import Path

from app import create_app


COLUMNS = (
    "Supplier Name,GSTIN,Invoice No,Invoice Date,HSN Code (optional),"
    "Taxable Value,IGST,CGST,SGST,Total GST\n"
)


def csv_upload(prefix: str, count: int = 12) -> io.BytesIO:
    rows = [
        f"Acme Supplier,22AAAAA0000A1Z5,{prefix}-{number},01/06/2026,1001,"
        f"{1000 + number},0,90,90,180"
        for number in range(count)
    ]
    return io.BytesIO((COLUMNS + "\n".join(rows)).encode("utf-8"))


class ReconciliationAppTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.app = create_app()
        self.app.config.update(TESTING=True, JOBS_DIR=Path(self.temp_dir.name))
        self.client = self.app.test_client()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_upload_results_pagination_and_exports(self):
        response = self.client.post(
            "/reconcile",
            data={
                "pr_file": (csv_upload("INV"), "purchase.csv"),
                "gstr2b_file": (csv_upload("INV"), "gstr2b.csv"),
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 302)
        result_url = response.headers["Location"]
        page = self.client.get(result_url)
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"12 rows", page.data)
        self.assertIn(b"Page 1 of 2", page.data)

        job_id = result_url.rstrip("/").split("/")[-1]
        for kind in ("pr", "gstr2b"):
            export = self.client.get(f"/export/{job_id}/{kind}")
            self.assertEqual(export.status_code, 200)
            self.assertGreater(len(export.data), 1000)
            export.close()

        pr_rows = json.loads(
            (Path(self.temp_dir.name) / job_id / "pr_result.json").read_text("utf-8")
        )
        self.assertEqual(len(pr_rows), 12)
        self.assertEqual(pr_rows[0]["Best score"], 100.0)
        self.assertEqual(pr_rows[0]["Best match 2B index"], 0)

    def test_rejects_wrong_extension(self):
        response = self.client.post(
            "/reconcile",
            data={
                "pr_file": (io.BytesIO(b"bad"), "purchase.txt"),
                "gstr2b_file": (csv_upload("INV", 1), "gstr2b.csv"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn(b"must be a .csv or .xlsx file", response.data)


if __name__ == "__main__":
    unittest.main()
