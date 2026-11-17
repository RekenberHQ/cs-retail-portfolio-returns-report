from datetime import datetime
from zoneinfo import ZoneInfo
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch

from cs_returns_report.cli import publish_saved_report
from cs_returns_report.publishing import CONTENT_TYPE, publish_report


class PublishingTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "report.xlsx"
        self.path.write_bytes(b"test workbook")
        self.env = patch.dict(os.environ, {
            "REKENBER_APP_DB_URL": "postgresql://test",
            "REKENBER_REPORT_ORGANIZATION_ID": "2",
            "REKENBER_REPORT_S3_BUCKET": "rekenber-generated-reports",
            "REKENBER_REPORT_S3_ROOT": "production",
            "REKENBER_REPORT_S3_ACCESS_KEY_ID": "test",
            "REKENBER_REPORT_S3_SECRET_ACCESS_KEY": "test",
        })
        self.env.start()
        self.addCleanup(self.env.stop)
        self.db = patch("psycopg.connect").start()
        self.s3 = patch("boto3.client").start()
        self.addCleanup(patch.stopall)
        self.cursor = self.db.return_value.__enter__.return_value.cursor.return_value.__enter__.return_value
        self.cursor.fetchone.side_effect = [(2,), (9,)]

    def test_plain_upload_and_upsert(self):
        result = publish_report(self.path, "2026-11", 12)
        upload = self.s3.return_value.put_object.call_args.kwargs
        self.assertEqual(upload["Bucket"], "rekenber-generated-reports")
        self.assertEqual(upload["Body"], self.path.read_bytes())
        self.assertEqual(upload["ContentType"], CONTENT_TYPE)
        upload_date = datetime.now(ZoneInfo("Europe/Prague")).date().isoformat()
        self.assertIn(f"/reports/organizations/2/cs-retail-portfolio-returns/{upload_date}/", upload["Key"])
        sql, params = self.cursor.execute.call_args.args
        self.assertIn("ON CONFLICT (organization_id, report_type, report_name) DO UPDATE", sql)
        self.assertEqual(params[-1], 12)
        self.assertEqual(params[2], upload_date)
        self.assertTrue(params[3].endswith(upload_date))
        self.assertEqual(params[4], f"cs-retail-portfolio-returns-{upload_date}.xlsx")
        self.assertEqual(result["id"], 9)

    def test_upload_failure_does_not_register_report(self):
        self.s3.return_value.put_object.side_effect = RuntimeError("storage unavailable")
        with self.assertRaisesRegex(RuntimeError, "storage unavailable"):
            publish_report(self.path, "2026-11", 12)
        self.assertFalse(any("INSERT" in call.args[0] for call in self.cursor.execute.call_args_list))

    def test_disabled_organization_does_not_upload(self):
        self.cursor.fetchone.side_effect = [None]
        with self.assertRaisesRegex(ValueError, "must be enabled"):
            publish_report(self.path, "2026-11", 12)
        self.s3.return_value.put_object.assert_not_called()

    def test_republish_local_report_records_failure_and_success(self):
        previous = Path.cwd()
        os.chdir(self.directory.name)
        self.addCleanup(os.chdir, previous)
        workspace = Path("workspace/2026-11")
        workspace.mkdir(parents=True)
        summary = workspace / "summary.json"
        summary.write_text(json.dumps({"status": "complete", "rows": 12}))
        with patch("cs_returns_report.cli.publish_report", side_effect=RuntimeError("fail")):
            with self.assertRaises(RuntimeError):
                publish_saved_report("2026-11")
        self.assertEqual(json.loads(summary.read_text())["publication_status"], "failed")
        with patch("cs_returns_report.cli.publish_report", return_value={"id": 9}):
            publish_saved_report("2026-11")
        self.assertEqual(json.loads(summary.read_text())["publication_status"], "published")
        summary.write_text(json.dumps({"status": "needs_annotations", "rows": 12}))
        with self.assertRaisesRegex(ValueError, "successfully completed"):
            publish_saved_report("2026-11")


if __name__ == "__main__":
    unittest.main()
