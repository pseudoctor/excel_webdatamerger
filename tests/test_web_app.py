import io
import json
import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

from web_app.app import WebConfig, create_app


class WebAppTestCase(unittest.TestCase):
    def setUp(self):
        self.tmpdir = Path(tempfile.mkdtemp(prefix="excelmerger-tests-"))
        self.addCleanup(shutil.rmtree, self.tmpdir, True)
        self.original_upload_root = WebConfig.UPLOAD_ROOT
        self.original_username = WebConfig.USERNAME
        self.original_password = WebConfig.PASSWORD
        self.original_secret = WebConfig.SECRET_KEY
        self.original_merge_async = getattr(WebConfig, "MERGE_ASYNC", True)

        WebConfig.UPLOAD_ROOT = self.tmpdir
        WebConfig.USERNAME = "tester"
        WebConfig.PASSWORD = "secret"
        WebConfig.SECRET_KEY = "test-secret-key"
        WebConfig.MERGE_ASYNC = False

    def tearDown(self):
        WebConfig.UPLOAD_ROOT = self.original_upload_root
        WebConfig.USERNAME = self.original_username
        WebConfig.PASSWORD = self.original_password
        WebConfig.SECRET_KEY = self.original_secret
        WebConfig.MERGE_ASYNC = self.original_merge_async

    def make_client(self):
        app = create_app()
        app.config.update(TESTING=True)
        client = app.test_client()
        with client.session_transaction() as session:
            session["user"] = WebConfig.USERNAME
        return app, client

    def test_login_page_preserves_next_parameter(self):
        app = create_app()
        app.config.update(TESTING=True)
        client = app.test_client()

        response = client.get("/login?next=/mapping")

        self.assertEqual(response.status_code, 200)
        self.assertIn('name="next" value="/mapping"', response.get_data(as_text=True))

    def test_merge_result_download_survives_new_app_instance(self):
        _, client = self.make_client()

        response = client.post(
            "/merge",
            data={
                "files": (io.BytesIO(b"col1,col2\n1,2\n"), "sample.csv"),
                "output_format": "csv",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 202)
        payload = response.get_json()
        self.assertTrue(payload["ok"])
        task_id = payload["task_id"]
        self.assertEqual(payload["status"], "completed")

        _, new_client = self.make_client()
        download = new_client.get(f"/download/{task_id}")

        self.assertEqual(download.status_code, 200)
        self.assertEqual(download.mimetype, "text/csv")
        self.assertIn("col1", download.get_data(as_text=True))
        download.close()

    def test_merge_returns_suggested_filename(self):
        _, client = self.make_client()

        response = client.post(
            "/merge",
            data={
                "files": (io.BytesIO(b"col1,col2\n1,2\n"), "sample.csv"),
                "output_format": "csv",
            },
            content_type="multipart/form-data",
        )

        self.assertEqual(response.status_code, 202)
        payload = response.get_json()
        self.assertTrue(payload["ok"])
        self.assertRegex(payload["suggested_filename"], r"^merged_\d{8}_\d{6}$")
        self.assertIn("/task/", payload["status_url"])

    def test_task_status_reports_completed_download(self):
        _, client = self.make_client()

        response = client.post(
            "/merge",
            data={
                "files": (io.BytesIO(b"col1,col2\n1,2\n"), "sample.csv"),
                "output_format": "csv",
            },
            content_type="multipart/form-data",
        )

        payload = response.get_json()
        task_id = payload["task_id"]
        status = client.get(f"/task/{task_id}")

        self.assertEqual(status.status_code, 200)
        status_payload = status.get_json()
        self.assertTrue(status_payload["ok"])
        self.assertEqual(status_payload["status"], "completed")
        self.assertEqual(status_payload["format"], "csv")
        self.assertIn(f"/download/{task_id}", status_payload["download_url"])

    def test_download_uses_custom_filename_with_actual_extension(self):
        _, client = self.make_client()

        merge_response = client.post(
            "/merge",
            data={
                "files": (io.BytesIO(b"col1,col2\n1,2\n"), "sample.csv"),
                "output_format": "csv",
            },
            content_type="multipart/form-data",
        )
        task_id = merge_response.get_json()["task_id"]

        download = client.get(f"/download/{task_id}?filename=final_report.xlsx")

        self.assertEqual(download.status_code, 200)
        self.assertEqual(
            download.headers["Content-Disposition"],
            'attachment; filename=final_report.csv',
        )
        download.close()

    def test_download_sanitizes_invalid_custom_filename(self):
        _, client = self.make_client()

        merge_response = client.post(
            "/merge",
            data={
                "files": (io.BytesIO(b"col1,col2\n1,2\n"), "sample.csv"),
                "output_format": "csv",
            },
            content_type="multipart/form-data",
        )
        task_id = merge_response.get_json()["task_id"]

        download = client.get(f"/download/{task_id}?filename=../bad:name")

        self.assertEqual(download.status_code, 200)
        disposition = download.headers["Content-Disposition"]
        self.assertIn("filename=bad_name.csv", disposition)
        self.assertNotIn("..", disposition)
        download.close()

    def test_download_preserves_multi_dot_filename(self):
        _, client = self.make_client()

        merge_response = client.post(
            "/merge",
            data={
                "files": (io.BytesIO(b"col1,col2\n1,2\n"), "sample.csv"),
                "output_format": "csv",
            },
            content_type="multipart/form-data",
        )
        task_id = merge_response.get_json()["task_id"]

        download = client.get(f"/download/{task_id}?filename=sales.v2.final")

        self.assertEqual(download.status_code, 200)
        self.assertIn(
            "filename=sales.v2.final.csv",
            download.headers["Content-Disposition"],
        )
        download.close()

    def test_cleanup_temp_only_removes_expired_jobs(self):
        _, client = self.make_client()
        active_dir = self.tmpdir / "active-task"
        expired_dir = self.tmpdir / "expired-task"
        orphan_dir = self.tmpdir / "orphan-task"
        active_dir.mkdir()
        expired_dir.mkdir()
        orphan_dir.mkdir()

        (active_dir / "metadata.json").write_text(
            json.dumps(
                {
                    "path": "merged.csv",
                    "format": "csv",
                    "created_at": datetime.now(timezone.utc).isoformat(),
                }
            ),
            encoding="utf-8",
        )
        (expired_dir / "metadata.json").write_text(
            json.dumps(
                {
                    "path": "merged.csv",
                    "format": "csv",
                    "created_at": (
                        datetime.now(timezone.utc) - timedelta(minutes=240)
                    ).isoformat(),
                }
            ),
            encoding="utf-8",
        )
        old_timestamp = (
            datetime.now(timezone.utc) - timedelta(minutes=240)
        ).timestamp()
        os.utime(orphan_dir, (old_timestamp, old_timestamp))

        response = client.post("/cleanup", data={"target": "temp"})

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload["ok"])
        self.assertIn(payload["removed"], {0, 2})
        self.assertEqual(payload["skipped"], 1)
        self.assertTrue(active_dir.exists())
        self.assertFalse(expired_dir.exists())
        self.assertFalse(orphan_dir.exists())


if __name__ == "__main__":
    unittest.main()
