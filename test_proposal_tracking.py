import datetime
import os
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

import pcs_proposal_web as web
from proposal_tracking_cutover_flags import (
    MASTER_FLAG,
    READ_FLAG,
    SHADOW_WRITE_FLAG,
    WRITE_FLAG,
    load_proposal_tracking_cutover_flags,
)
from proposal_tracking_store import ProposalTrackingStore


PROPOSAL_ID = "11111111-1111-4111-8111-111111111111"
RELATIONSHIP_ID = "22222222-2222-4222-8222-222222222222"


CURRENT_TRACKER_HEADERS = [
    "Customer",
    "Contact",
    "Email Address",
    "Lead Generated",
    "Submitted By",
    "Estimate Dt",
    "Proposal Dt",
    "Follow-Up",
    "Estimated By",
    "Response",
]


class ProposalTrackingCutoverFlagTests(unittest.TestCase):
    def test_flags_are_disabled_by_default(self):
        flags = load_proposal_tracking_cutover_flags({})
        self.assertFalse(flags.master_enabled)
        self.assertFalse(flags.reads_enabled)
        self.assertFalse(flags.writes_enabled)
        self.assertTrue(flags.spreadsheet_reads_active)
        self.assertTrue(flags.spreadsheet_writes_active)

    def test_capability_flags_require_master_flag(self):
        flags = load_proposal_tracking_cutover_flags({
            READ_FLAG: "1",
            WRITE_FLAG: "1",
            SHADOW_WRITE_FLAG: "1",
        })
        self.assertFalse(flags.reads_enabled)
        self.assertFalse(flags.writes_enabled)
        self.assertFalse(flags.shadow_writes_enabled)

    def test_shadow_mode_keeps_spreadsheet_writes_active(self):
        flags = load_proposal_tracking_cutover_flags({
            MASTER_FLAG: "true",
            WRITE_FLAG: "true",
            SHADOW_WRITE_FLAG: "true",
        })
        self.assertTrue(flags.writes_enabled)
        self.assertTrue(flags.spreadsheet_writes_active)
        self.assertFalse(flags.fully_cut_over)

    def test_persistent_shadow_mode_is_used_when_environment_is_absent(self):
        persisted = {
            MASTER_FLAG: "true",
            READ_FLAG: "false",
            WRITE_FLAG: "true",
            SHADOW_WRITE_FLAG: "true",
        }
        with patch(
            "pcs_local_settings.proposal_tracking_cutover_environment",
            return_value=persisted,
        ), patch.dict(os.environ, {}, clear=True):
            flags = load_proposal_tracking_cutover_flags()
        self.assertTrue(flags.master_enabled)
        self.assertFalse(flags.reads_enabled)
        self.assertTrue(flags.writes_enabled)
        self.assertTrue(flags.shadow_writes_enabled)
        self.assertTrue(flags.spreadsheet_reads_active)
        self.assertTrue(flags.spreadsheet_writes_active)

    def test_process_environment_overrides_persistent_cutover_setting(self):
        persisted = {
            MASTER_FLAG: "true",
            READ_FLAG: "false",
            WRITE_FLAG: "true",
            SHADOW_WRITE_FLAG: "true",
        }
        with patch(
            "pcs_local_settings.proposal_tracking_cutover_environment",
            return_value=persisted,
        ), patch.dict(os.environ, {MASTER_FLAG: "false"}, clear=True):
            flags = load_proposal_tracking_cutover_flags()
        self.assertFalse(flags.master_enabled)
        self.assertFalse(flags.writes_enabled)
        self.assertFalse(flags.shadow_writes_enabled)


class ProposalTrackingStoreTests(unittest.TestCase):
    def setUp(self):
        self.store = ProposalTrackingStore("https://example.supabase.co", "test-key")

    @staticmethod
    def proposal_row(**updates):
        row = {
            "id": PROPOSAL_ID,
            "customer_name": "Example Roofing",
            "project_street_address": "123 Main St",
            "display_name": "Example Roofing - 123 Main St",
            "lead_source": "Referral",
            "submitted_by": "David",
            "estimated_by": "Vern",
            "estimate_completed_date": None,
            "proposal_sent_date": "2026-07-15",
            "follow_up_date": None,
            "response_notes": None,
            "proposal_contact": [{
                "is_primary": True,
                "organization_contact": {
                    "id": RELATIONSHIP_ID,
                    "business_email": "casey@example.com",
                    "is_current": True,
                    "contact": {"id": "contact-id", "full_name": "Casey Smith"},
                },
            }],
        }
        row.update(updates)
        return row

    def test_screen_entry_uses_generated_display_name_and_contact_tables(self):
        entry = self.store._screen_entry(self.proposal_row())
        self.assertEqual(entry["customer"], "Example Roofing - 123 Main St")
        self.assertEqual(entry["contact"], "Casey Smith")
        self.assertEqual(entry["email_address"], "casey@example.com")
        self.assertEqual(entry["lead_source"], "Referral")

    def test_missing_entries_include_missing_estimate_date(self):
        with patch.object(self.store, "list_proposals", return_value=[self.proposal_row()]):
            entries = self.store.list_missing_entries()
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["estimate_date_input"], "")

    def test_completed_entry_is_not_returned_as_missing(self):
        row = self.proposal_row(estimate_completed_date="2026-07-14")
        with patch.object(self.store, "list_proposals", return_value=[row]):
            self.assertEqual(self.store.list_missing_entries(), [])

    def test_numeric_spreadsheet_row_resolves_to_migrated_proposal(self):
        with patch.object(
            self.store,
            "_request",
            return_value=[{"id": PROPOSAL_ID}],
        ) as request:
            resolved = self.store._resolve_proposal_id("42")
        self.assertEqual(resolved, PROPOSAL_ID)
        self.assertEqual(request.call_args.kwargs["params"]["source_row_number"], "eq.42")

    def test_mark_followups_updates_resolved_ids(self):
        responses = [[{"id": PROPOSAL_ID}], [{"id": PROPOSAL_ID}]]
        with patch.object(self.store, "_request", side_effect=responses) as request:
            count = self.store.mark_follow_ups(["42"], datetime.date(2026, 8, 3))
        self.assertEqual(count, 1)
        self.assertEqual(request.call_args.kwargs["payload"], {
            "follow_up_date": "2026-08-03",
            "status": "sent",
        })

    def test_editing_dates_does_not_reopen_a_closed_proposal(self):
        payload = self.store._editable_payload({
            "proposal_date": "8/1/2026",
            "follow_up_date": "8/3/2026",
        })
        self.assertNotIn("status", payload)

    def test_legacy_statuses_map_to_new_lifecycle(self):
        self.assertEqual(
            self.store._editable_payload({"status": "won"})["status"],
            "under_contract",
        )
        self.assertEqual(
            self.store._editable_payload({"status": "withdrawn"})["status"],
            "dead",
        )

    def test_new_proposal_save_sets_estimate_date_without_setting_sent_date(self):
        responses = [[], [], [{"id": PROPOSAL_ID}]]
        with patch.object(self.store, "_request", side_effect=responses) as request:
            proposal_id = self.store.upsert_from_proposal_save(
                created_date="08/03/2026",
                customer_name="Example Roofing",
                street_address="123 Main St",
                city="Denver",
                state="co",
                zip_code="80202",
                submitted_by="David",
                folder_name="Example Roofing - 123 Main St",
                lead_value="Referral",
            )
        self.assertEqual(proposal_id, PROPOSAL_ID)
        payload = request.call_args.kwargs["payload"]
        self.assertEqual(payload["estimate_completed_date"], "2026-08-03")
        self.assertNotIn("proposal_sent_date", payload)
        self.assertEqual(payload["project_state"], "CO")


class ProposalTrackingSpreadsheetColumnTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.tracker_path = os.path.join(
            self.temporary_directory.name, "Proposal Tracking.xlsx"
        )

    def tearDown(self):
        self.temporary_directory.cleanup()

    def write_tracker(self, rows):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Tracking"
        worksheet.append(CURRENT_TRACKER_HEADERS)
        for row in rows:
            worksheet.append(row)
        workbook.save(self.tracker_path)
        workbook.close()

    def read_row(self, row_number=2):
        workbook = load_workbook(self.tracker_path, data_only=True)
        try:
            worksheet = workbook.active
            return [
                worksheet.cell(row=row_number, column=column).value
                for column in range(1, 11)
            ]
        finally:
            workbook.close()

    def test_tracker_screen_reads_reordered_columns_by_header(self):
        self.write_tracker([[
            "Anchor Roofing - 10 Oak Ave",
            "",
            "joel.anchorroofing@gmail.com",
            "Referral",
            "Mark",
            datetime.date(2026, 7, 30),
            "-",
            "-",
            "Mark",
            "Waiting",
        ]])

        entries, error = web._load_proposal_tracker_missing_entries_spreadsheet(
            self.tracker_path
        )

        self.assertIsNone(error)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["estimated_by"], "Mark")
        self.assertEqual(entries[0]["estimate_date_input"], "7/30/2026")
        self.assertEqual(entries[0]["proposal_date_input"], "-")
        self.assertEqual(entries[0]["follow_up_date_input"], "-")

    def test_tracker_screen_save_writes_reordered_columns_by_header(self):
        self.write_tracker([[
            "Example Roofing - 123 Main St",
            "Casey",
            "casey@example.com",
            "Referral",
            "David",
            "",
            "",
            "",
            "Vern",
            "Keep this response",
        ]])

        count = web._update_proposal_tracker_missing_entries_spreadsheet([{
            "row_number": "2",
            "contact": "Casey Smith",
            "email_address": "casey@example.com",
            "lead_source": "Website",
            "submitted_by": "Mark",
            "estimate_date": "7/30/2026",
            "proposal_date": "8/1/2026",
            "follow_up_date": "8/15/2026",
            "estimated_by": "Mark",
        }], self.tracker_path)

        self.assertEqual(count, 1)
        row = self.read_row()
        self.assertEqual(row[5], "7/30/2026")
        self.assertEqual(row[6], "8/1/2026")
        self.assertEqual(row[7], "8/15/2026")
        self.assertEqual(row[8], "Mark")
        self.assertEqual(row[9], "Keep this response")

    def test_weekly_follow_up_uses_proposal_and_follow_up_headers(self):
        self.write_tracker([[
            "Example Roofing - 123 Main St",
            "Casey",
            "casey@example.com",
            "Referral",
            "David",
            datetime.date(2026, 7, 30),
            datetime.date(2026, 8, 1),
            "",
            "Vern",
            "Keep this response",
        ]])

        entries, _, error = web._load_weekly_follow_up_entries_spreadsheet(
            self.tracker_path, datetime.date(2026, 8, 2)
        )
        self.assertIsNone(error)
        self.assertEqual([entry["row_number"] for entry in entries], [2])

        count = web._update_weekly_follow_up_dates_spreadsheet(
            [2], datetime.date(2026, 8, 3), self.tracker_path
        )
        self.assertEqual(count, 1)
        row = self.read_row()
        self.assertEqual(row[6], datetime.datetime(2026, 8, 1, 0, 0))
        self.assertEqual(row[7], datetime.datetime(2026, 8, 3, 0, 0))
        self.assertEqual(row[8], "Vern")
        self.assertEqual(row[9], "Keep this response")

    def test_proposal_save_refresh_preserves_dates_and_response_columns(self):
        self.write_tracker([[
            "Example Roofing - 123 Main St",
            "Casey",
            "casey@example.com",
            "Referral",
            "David",
            "",
            datetime.date(2026, 8, 1),
            "",
            "Mark",
            "Keep this response",
        ]])

        updated = web.update_existing_tracker_row(
            "Example Roofing - 123 Main St",
            "Website",
            "David",
            datetime.date(2026, 7, 30),
            self.tracker_path,
        )

        self.assertTrue(updated)
        row = self.read_row()
        self.assertEqual(row[5], datetime.datetime(2026, 7, 30, 0, 0))
        self.assertEqual(row[6], datetime.datetime(2026, 8, 1, 0, 0))
        self.assertEqual(row[8], "Vern")
        self.assertEqual(row[9], "Keep this response")


if __name__ == "__main__":
    unittest.main()
