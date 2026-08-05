import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from contact_store import ContactStore, ContactStoreError
from roof_intelligence_jobs import RoofIntelligenceJobStore
from tenant_settings_store import TenantSettingsStore


class _Response:
    def __init__(self, value):
        self.value = value

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self):
        return json.dumps(self.value).encode("utf-8")


class TenantPostgrestTests(unittest.TestCase):
    def setUp(self):
        self.store = ContactStore(
            "https://example.supabase.co",
            "sb_publishable_test",
            "user-access-token",
            "tenant-1",
        )

    def test_reads_are_explicitly_filtered_to_active_tenant(self):
        with patch("contact_store.urlopen", return_value=_Response([])) as request:
            self.store._request("organization", params={"select": "id"})
        self.assertIn("tenant_id=eq.tenant-1", request.call_args.args[0].full_url)

    def test_writes_receive_server_selected_tenant(self):
        with patch("contact_store.urlopen", return_value=_Response([])) as request:
            self.store._request("organization", method="POST", payload={"name": "Example"})
        payload = json.loads(request.call_args.args[0].data.decode("utf-8"))
        self.assertEqual(payload["tenant_id"], "tenant-1")
        self.assertEqual(request.call_args.args[0].headers["Authorization"], "Bearer user-access-token")

    def test_cross_tenant_payload_is_rejected_before_network_request(self):
        with self.assertRaisesRegex(ContactStoreError, "Cross-company"):
            self.store._request(
                "organization",
                method="POST",
                payload={"tenant_id": "tenant-2", "name": "Blocked"},
            )

    def test_storage_prefix_contains_tenant_folder_report_and_revision(self):
        settings = TenantSettingsStore(
            "https://example.supabase.co", "key", "token", "tenant-1"
        )
        with patch.object(settings, "get_settings", return_value={"default_report_folder_id": "folder-1"}):
            self.assertEqual(
                settings.storage_prefix("report-1", 3),
                "tenant-1/folders/folder-1/reports/report-1/revisions/3",
            )


class LocalRoofTenantCompatibilityTests(unittest.TestCase):
    def test_report_reads_are_scoped_by_job_user_key(self):
        with tempfile.TemporaryDirectory() as directory:
            store = RoofIntelligenceJobStore(Path(directory) / "jobs.sqlite3")
            job = store.create_individual_job(
                "101 Test Ave, Denver, CO 80202", user_key="tenant-a:user-a"
            )
            now = "2026-08-05T12:00:00+00:00"
            with store.connect() as connection:
                connection.execute(
                    "INSERT INTO properties (id,canonical_key,normalized_address,data_json,created_at,updated_at) "
                    "VALUES ('property-1','property-1','101 TEST AVE','{}',?,?)",
                    (now, now),
                )
                connection.execute(
                    "INSERT INTO roof_intelligence_reports "
                    "(id,property_id,job_id,result_json,created_at) VALUES "
                    "('report-1','property-1',?,'{}',?)",
                    (job["id"], now),
                )
            self.assertIsNotNone(store.get_report("report-1", "tenant-a:user-a"))
            self.assertIsNone(store.get_report("report-1", "tenant-b:user-b"))


if __name__ == "__main__":
    unittest.main()

