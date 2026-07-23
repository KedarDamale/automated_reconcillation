import json
import tempfile
import unittest
from pathlib import Path

import fastexcel

from app import create_app
from reconciliation import PREFERRED_COLUMNS
from workbook_export import SHEET_NAMES


SAMPLE_DIRECTORY = Path(__file__).resolve().parents[1] / "sample_test_files"
PURCHASE_REGISTER = SAMPLE_DIRECTORY / "input1.xlsx"
GSTR2B = SAMPLE_DIRECTORY / "input2.xlsx"


class UploadedWorkbookEndToEndTest(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.app = create_app()
        self.app.config.update(
            TESTING=True,
            JOBS_DIR=Path(self.temporary.name),
        )
        self.client = self.app.test_client()

    def tearDown(self):
        self.temporary.cleanup()

    def _preview(self, path: Path) -> dict:
        with path.open("rb") as upload:
            response = self.client.post(
                "/columns",
                data={"file": (upload, path.name)},
                content_type="multipart/form-data",
            )
        self.assertEqual(response.status_code, 200, response.get_data(as_text=True))
        return response.get_json()

    def test_uploaded_workbooks_complete_full_web_flow(self):
        self.assertTrue(PURCHASE_REGISTER.is_file())
        self.assertTrue(GSTR2B.is_file())

        pr_preview = self._preview(PURCHASE_REGISTER)
        gstr2b_preview = self._preview(GSTR2B)

        pr_mapping = {
            preferred: pr_preview["suggestions"][preferred]["source"]
            for preferred in PREFERRED_COLUMNS
        }
        gstr2b_mapping = {
            preferred: gstr2b_preview["suggestions"][preferred]["source"]
            for preferred in PREFERRED_COLUMNS
        }
        self.assertTrue(all(pr_mapping.values()))
        self.assertTrue(all(gstr2b_mapping.values()))
        self.assertEqual(gstr2b_mapping["HSN Code (optional)"], "HSN Code")

        with PURCHASE_REGISTER.open("rb") as pr_upload, GSTR2B.open("rb") as gstr2b_upload:
            response = self.client.post(
                "/reconcile",
                data={
                    "pr_file": (pr_upload, PURCHASE_REGISTER.name),
                    "gstr2b_file": (gstr2b_upload, GSTR2B.name),
                    "pr_mapping": json.dumps(pr_mapping),
                    "gstr2b_mapping": json.dumps(gstr2b_mapping),
                    "date_tolerance_days": "10",
                },
                content_type="multipart/form-data",
            )

        self.assertEqual(response.status_code, 302, response.get_data(as_text=True))
        result_url = response.headers["Location"]
        result_page = self.client.get(result_url)
        self.assertEqual(result_page.status_code, 200)
        self.assertIn(b"558 rows", result_page.data)
        self.assertIn(b"579 rows", result_page.data)

        job_id = result_url.rstrip("/").split("/")[-1]
        job_directory = Path(self.temporary.name) / job_id
        pr_result = json.loads((job_directory / "pr_result.json").read_text("utf-8"))
        gstr2b_result = json.loads((job_directory / "gstr2b_result.json").read_text("utf-8"))
        self.assertEqual(len(pr_result), 558)
        self.assertEqual(len(gstr2b_result), 579)
        self.assertIn("Match category", pr_result[0])
        self.assertIn("Best match 2B index", pr_result[0])
        self.assertIn("Best match PR index", gstr2b_result[0])

        export = self.client.get(f"/export/{job_id}")
        self.assertEqual(export.status_code, 200)
        self.assertGreater(len(export.data), 10_000)
        export_bytes = bytes(export.data)
        export.close()
        workbook = fastexcel.read_excel(export_bytes)
        self.assertEqual(workbook.sheet_names, SHEET_NAMES)
        self.assertEqual(
            workbook.load_sheet(SHEET_NAMES[0]).total_height,
            558,
        )
        self.assertEqual(
            workbook.load_sheet(SHEET_NAMES[1]).total_height,
            579,
        )
