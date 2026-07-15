import os
from pathlib import Path
import tempfile
import unittest
from urllib.parse import parse_qs, urlsplit

from roof_intelligence_jobs import RoofIntelligenceJobStore, SUPPORTED_ROOF_TYPES


class RoofIntelligenceJobStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "roof-jobs.sqlite3"
        self.store = RoofIntelligenceJobStore(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_individual_job_requires_full_address_with_zip(self):
        with self.assertRaisesRegex(ValueError, "five-digit ZIP"):
            self.store.create_individual_job("65 N Yuma St, Denver, CO")

        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")

        self.assertEqual(job["status"], "queued")
        self.assertEqual(job["job_type"], "individual_address")
        self.assertEqual(job["input"]["property_address"], "65 N Yuma St, Denver, CO 80223")

    def test_zip_job_applies_defaults_and_all_roof_types(self):
        job = self.store.create_zip_job("80223", "10", "10000", "", ["All"])

        self.assertEqual(job["report_limit"], 10)
        self.assertEqual(job["minimum_roof_size"], 10000)
        self.assertIsNone(job["minimum_age"])
        self.assertEqual(job["roof_types"], list(SUPPORTED_ROOF_TYPES))

    def test_duplicate_property_keeps_one_canonical_record_and_two_reports(self):
        first = self.store.create_individual_job("65 N Yuma Street, Denver, CO 80223")
        second = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        result = {
            "address": "65 N Yuma St",
            "city": "Denver",
            "state": "CO",
            "zip": "80223",
            "county": "Denver",
            "parcel": "0508500065000",
            "roof_squares": 125,
            "report_path": "",
        }

        self.store.start_job(first["id"])
        self.store.complete_individual_job(first["id"], result)
        self.store.start_job(second["id"])
        self.store.complete_individual_job(second["id"], {**result, "roof_squares": 126})

        with self.store.connect() as connection:
            property_count = connection.execute("SELECT count(*) FROM properties").fetchone()[0]
            report_count = connection.execute("SELECT count(*) FROM roof_intelligence_reports").fetchone()[0]
            current_squares = connection.execute("SELECT roof_squares FROM properties").fetchone()[0]

        self.assertEqual(property_count, 1)
        self.assertEqual(report_count, 2)
        self.assertEqual(current_squares, 126)
        self.assertEqual(len(self.store.list_notifications()), 2)

    def test_failure_stores_concise_error_and_notification(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")

        failed = self.store.fail_job(
            job["id"],
            "imagery_unavailable",
            "County imagery service did not return a usable image.",
            retryable=True,
        )

        self.assertEqual(failed["status"], "failed")
        self.assertTrue(failed["retryable"])
        self.assertEqual(failed["error_code"], "imagery_unavailable")
        self.assertEqual(self.store.list_notifications()[0]["kind"], "job_failed")


class RoofIntelligenceRouteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import pcs_proposal_web

        cls.web = pcs_proposal_web

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.previous_db = os.environ.get("ROOF_INTELLIGENCE_DB_PATH")
        self.previous_worker = os.environ.get("ROOF_INTELLIGENCE_LOCAL_WORKER")
        os.environ["ROOF_INTELLIGENCE_DB_PATH"] = str(Path(self.temp_dir.name) / "route-jobs.sqlite3")
        os.environ["ROOF_INTELLIGENCE_LOCAL_WORKER"] = "0"
        self.client = self.web.app.test_client()

    def tearDown(self):
        if self.previous_db is None:
            os.environ.pop("ROOF_INTELLIGENCE_DB_PATH", None)
        else:
            os.environ["ROOF_INTELLIGENCE_DB_PATH"] = self.previous_db
        if self.previous_worker is None:
            os.environ.pop("ROOF_INTELLIGENCE_LOCAL_WORKER", None)
        else:
            os.environ["ROOF_INTELLIGENCE_LOCAL_WORKER"] = self.previous_worker
        self.temp_dir.cleanup()

    def test_individual_submission_redirects_to_persistent_job_and_status_api(self):
        response = self.client.post(
            "/roof-intelligence/jobs/individual",
            data={"property_address": "65 N Yuma St, Denver, CO 80223"},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        job_id = query["job_id"][0]

        status = self.client.get(f"/api/roof-intelligence/jobs/{job_id}")
        self.assertEqual(status.status_code, 200)
        payload = status.get_json()
        self.assertEqual(payload["status"], "queued")
        self.assertEqual(payload["input"]["property_address"], "65 N Yuma St, Denver, CO 80223")

        page = self.client.get(response.headers["Location"])
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Individual Address", page.data)
        self.assertIn(b"ZIP Code Batch", page.data)
        self.assertIn(b"job-clock", page.data)
        self.assertIn(b"waiting for an available PilotPoint IQ worker", page.data)

    def test_zip_submission_persists_filters(self):
        response = self.client.post(
            "/roof-intelligence/jobs/zip",
            data={
                "zip_code": "80223",
                "report_limit": "25",
                "minimum_roof_size": "10000",
                "minimum_age": "20",
                "roof_types": ["TPO", "Metal"],
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        payload = self.client.get(
            f"/api/roof-intelligence/jobs/{query['job_id'][0]}"
        ).get_json()
        self.assertEqual(payload["minimum_age"], 20)
        self.assertEqual(payload["roof_types"], ["TPO", "Metal"])
        self.assertEqual(payload["report_limit"], 25)


if __name__ == "__main__":
    unittest.main()
