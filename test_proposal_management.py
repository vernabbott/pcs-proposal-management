import datetime
import unittest
from unittest.mock import Mock, patch

import pcs_proposal_web as web
from proposal_tracking_store import ProposalContactOrganizationRequired


PROPOSAL_ID = "11111111-1111-4111-8111-111111111111"
CONTACT_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"


class ProposalManagementPageTests(unittest.TestCase):
    def setUp(self):
        web.app.config["TESTING"] = True
        self.client = web.app.test_client()
        now = datetime.datetime.now()
        self.rows = [
            {
                "id": PROPOSAL_ID,
                "name": "Open Customer - 123 Main St",
                "folder_name": "Open Folder",
                "status": "sent",
                "submitted_by": "Mark",
                "estimated_by": "Vern",
                "estimate_completed_date_display": "7/31/2026",
                "proposal_sent_date": "2026-08-01",
                "proposal_sent_date_display": "8/1/2026",
                "follow_up_date_display": "8/15/2026",
                "last_modified": now,
                "last_modified_display": now.strftime("%m/%d/%Y"),
                "organization_contact_id": "relationship-id",
                "contact_id": CONTACT_ID,
                "contact_name": "Casey Smith",
                "contact_email": "casey@example.com",
                "organization_name": "Example Roofing",
            },
            {
                "id": "22222222-2222-4222-8222-222222222222",
                "name": "Contract Customer - 456 Oak Ave",
                "folder_name": "Contract Folder",
                "status": "under_contract",
                "proposal_sent_date": "2026-07-15",
                "last_modified": now,
                "last_modified_display": now.strftime("%m/%d/%Y"),
                "organization_contact_id": "",
                "contact_name": "",
                "contact_email": "",
                "organization_name": "",
            },
            {
                "id": "33333333-3333-4333-8333-333333333333",
                "name": "Draft Customer - 789 Pine Rd",
                "folder_name": "Draft Folder",
                "status": "draft",
                "proposal_sent_date": None,
                "last_modified": now,
                "last_modified_display": now.strftime("%m/%d/%Y"),
                "organization_contact_id": "draft-relationship-id",
                "contact_name": "Taylor Jones",
                "contact_email": "taylor@example.com",
                "organization_name": "Draft Roofing",
            },
            {
                "id": "44444444-4444-4444-8444-444444444444",
                "name": "Unlinked Draft - 321 Cedar Ln",
                "folder_name": "Unlinked Draft Folder",
                "status": "draft",
                "proposal_sent_date": None,
                "last_modified": now,
                "last_modified_display": now.strftime("%m/%d/%Y"),
                "organization_contact_id": "",
                "contact_id": "",
                "contact_name": "",
                "contact_email": "",
                "organization_name": "",
            },
        ]

    def test_all_filter_uses_proposal_rows_and_database_folder_name(self):
        store = Mock()
        store.list_management_proposals.return_value = self.rows
        store.list_management_contact_options.return_value = []
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.get("/proposals?filter=all")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Open Customer - 123 Main St", response.data)
        self.assertIn(b"Draft Customer - 789 Pine Rd", response.data)
        self.assertIn(b"Contract Customer - 456 Oak Ave", response.data)
        self.assertIn(b"folder_name=Open+Folder", response.data)
        self.assertIn(f"proposal_id={PROPOSAL_ID}".encode(), response.data)
        self.assertIn(b"read_only=No", response.data)
        self.assertIn(b"casey@example.com", response.data)
        self.assertIn(b"Attach Contact", response.data)
        self.assertIn(b"Sales Person", response.data)
        self.assertIn(b"Estimated By", response.data)
        self.assertIn(b"Estimated Dt", response.data)
        self.assertIn(b"Sent Dt", response.data)
        self.assertIn(b"Follow-Up Dt", response.data)
        self.assertIn(b'data-label="Sales Person">Mark</div>', response.data)
        self.assertIn(b'data-label="Estimated By">Vern</div>', response.data)
        self.assertIn(b'data-label="Estimated Dt">7/31/2026</div>', response.data)
        self.assertIn(b'data-label="Sent Dt">8/1/2026</div>', response.data)
        self.assertIn(b'data-label="Follow-Up Dt">8/15/2026</div>', response.data)
        self.assertNotIn(b"proposal-contact-name", response.data)
        self.assertNotIn(b"proposal-contact-email", response.data)
        self.assertNotIn(b"save-contact-btn", response.data)
        self.assertNotIn(b"bi bi-floppy", response.data)
        self.assertIn(b'class="proposal-name"', response.data)
        self.assertIn(
            b'aria-label="Open proposal details for Open Customer - 123 Main St"',
            response.data,
        )
        self.assertNotIn(b'title="Edit proposal"', response.data)
        self.assertNotIn(b"bi bi-pencil", response.data)
        self.assertIn(
            b"action-btn icon-only under-contract-btn", response.data
        )
        self.assertIn(b"bi bi-check-lg", response.data)
        self.assertIn(
            b'aria-label="Move Open Customer - 123 Main St to under contract"',
            response.data,
        )
        self.assertNotIn(b"bi bi-check2-circle", response.data)
        self.assertIn(b"action-btn icon-only dead-btn", response.data)
        store.list_management_proposals.assert_called_once_with(
            {"draft", "sent", "under_contract", "finished", "dead"}
        )

    def test_database_only_proposal_opens_populated_detail_page(self):
        store = Mock()
        store.get_management_proposal.return_value = {
            "id": PROPOSAL_ID,
            "customer_name": "Open Customer",
            "project_street_address": "123 Main St",
            "project_city": "Denver",
            "project_state": "CO",
            "project_zip_code": "80202",
            "submitted_by": "Mark",
            "lead_source": "Referral",
            "response_notes": "Call next week",
            "draft_detail": {
                "flat_roof_squares": "125",
                "current_roof": "TPO/EPDM",
                "product": "Gaco",
                "proposal_language": "Preserve this proposal detail",
                "office_fee_pct": "5.0%",
            },
        }
        with patch.object(web, "PROPOSAL_DATABASE_SOURCE_ENABLED", True), patch.object(
            web, "_resolve_existing_proposal_folder", return_value=None
        ), patch.object(web, "current_tenant_context"), patch.object(
            web, "get_proposal_tracking_store", return_value=store
        ):
            response = self.client.get(
                "/proposal_details",
                query_string={
                    "proposal_id": PROPOSAL_ID,
                    "folder_name": "Open Folder",
                    "read_only": "No",
                },
            )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'value="Open Customer"', response.data)
        self.assertIn(b'value="123 Main St"', response.data)
        self.assertIn(b'value="Denver"', response.data)
        self.assertIn(b'Create Proposal Files', response.data)
        self.assertIn(b'id="contactButton"', response.data)
        self.assertIn(
            b'id="contactButton" name="action" value="contact" formnovalidate',
            response.data,
        )
        self.assertIn(b'name="flat_roof_squares"', response.data)
        self.assertIn(b'value="125"', response.data)
        self.assertIn(b'<option value="TPO/EPDM" selected>', response.data)
        self.assertIn(b'name="product" value="Gaco" checked', response.data)
        self.assertIn(b'value="Preserve this proposal detail"', response.data)
        self.assertIn(b'name="office_fee_pct"', response.data)
        self.assertIn(b'value="5.0%"', response.data)

    def test_new_proposal_contact_button_creates_database_draft(self):
        store = Mock()
        store.upsert_from_proposal_save.return_value = PROPOSAL_ID
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.post(
                "/update-proposal/NEW",
                data={
                    "action": "contact",
                    "customer_name": "New Customer",
                    "street_address": "500 New St",
                    "city": "Denver",
                    "state": "CO",
                    "zip_code": "80202",
                    "submitted_by": "Mark",
                    "lead": "Referral",
                },
            )
        self.assertEqual(response.status_code, 302)
        self.assertIn("/contacts?", response.location)
        self.assertIn(f"attach_to_proposal={PROPOSAL_ID}", response.location)
        self.assertIn("q=New+Customer", response.location)
        store.upsert_from_proposal_save.assert_called_once_with(
            created_date=None,
            customer_name="New Customer",
            street_address="500 New St",
            city="Denver",
            state="CO",
            zip_code="80202",
            submitted_by="Mark",
            folder_name="New Customer - 500 New St",
            lead_value="Referral",
            estimated_by="",
        )
        store.save_proposal_draft_detail.assert_called_once()
        self.assertEqual(
            store.save_proposal_draft_detail.call_args.args[1]["customer_name"],
            "New Customer",
        )

    def test_new_proposal_contact_button_creates_named_draft_when_form_is_blank(self):
        store = Mock()
        store.upsert_from_proposal_save.return_value = PROPOSAL_ID
        with patch.object(web, "get_proposal_tracking_store", return_value=store), patch.object(
            web.uuid, "uuid4"
        ) as uuid4:
            uuid4.return_value.hex = "abc12345deadbeef"
            response = self.client.post(
                "/update-proposal/NEW",
                data={"action": "contact"},
            )

        self.assertEqual(response.status_code, 302)
        self.assertIn("/contacts?", response.location)
        self.assertIn(f"attach_to_proposal={PROPOSAL_ID}", response.location)
        self.assertIn("proposal_name=New+Proposal+ABC12345", response.location)
        self.assertNotIn("&q=", response.location)
        store.upsert_from_proposal_save.assert_called_once_with(
            created_date=None,
            customer_name="New Proposal ABC12345",
            street_address="",
            city="",
            state="",
            zip_code="",
            submitted_by="",
            folder_name="New Proposal ABC12345",
            lead_value="",
            estimated_by="",
        )
        store.save_proposal_draft_detail.assert_called_once_with(PROPOSAL_ID, {})

    def test_new_proposal_contact_button_is_enabled(self):
        response = self.client.get("/proposal_details/new")
        self.assertEqual(response.status_code, 200)
        self.assertIn(
            b'id="contactButton" name="action" value="contact" formnovalidate',
            response.data,
        )
        self.assertNotIn(b'id="contactButton" disabled', response.data)

    def test_customer_autocomplete_uses_active_tenant_organizations(self):
        store = Mock()
        store.list_organizations.return_value = [
            {"id": "org-2", "name": "  Zenith   Roofing  "},
            {"id": "org-1", "name": "Acme Management"},
            {"id": "org-3", "name": "acme management"},
            {"id": "org-4", "name": ""},
        ]
        with patch.object(web, "get_contact_store", return_value=store):
            response = self.client.get("/proposal_details/new")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b'<option value="Acme Management"></option>', response.data)
        self.assertIn(b'<option value="Zenith Roofing"></option>', response.data)
        self.assertEqual(response.data.count(b'value="Acme Management"'), 1)
        self.assertNotIn(b'<option value="Advanced Roofing"></option>', response.data)
        store.list_organizations.assert_called_once_with()

    def test_create_finalizes_contact_linked_draft_instead_of_inserting_duplicate(self):
        store = Mock()
        payload = {
            "action": "create",
            "database_proposal_id": PROPOSAL_ID,
            "customer_name": "Boulder County",
            "street_address": "132 Main St",
            "city": "Denver",
            "state": "CO",
            "zip_code": "88888",
            "flat_roof_squares": "100",
            "current_roof": "TPO/EPDM",
            "product": "Gaco",
            "submitted_by": "Vern",
        }
        with patch.object(
            web,
            "create_proposal_from_fields",
            return_value="Boulder County - 132 Main St",
        ) as create, patch.object(
            web, "get_proposal_tracking_store", return_value=store
        ), patch.object(
            web, "copy_proposal_to_submitter_destination"
        ), patch.object(
            web, "move_selected_proposal_files_to_folder"
        ):
            response = self.client.post(
                "/update-proposal/__blank__",
                data=payload,
            )

        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith("/proposals"))
        self.assertFalse(create.call_args.kwargs["update_tracking"])
        store.upsert_from_proposal_save.assert_called_once_with(
            proposal_id=PROPOSAL_ID,
            created_date=datetime.date.today(),
            customer_name="Boulder County",
            street_address="132 Main St",
            city="Denver",
            state="CO",
            zip_code="88888",
            submitted_by="Vern",
            folder_name="Boulder County - 132 Main St",
            lead_value=None,
        )
        store.clear_proposal_draft_detail.assert_called_once_with(PROPOSAL_ID)

    def test_all_filter_uses_contact_cards_for_sent_and_draft_proposals(self):
        store = Mock()
        store.list_management_proposals.return_value = self.rows
        store.list_management_contact_options.return_value = []
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.get("/proposals?filter=all")
        page = response.data
        self.assertNotIn(b"Not Sent", page)
        self.assertNotIn(b"Sent Proposals", page)
        self.assertEqual(page.count(b'class="proposal-table"'), 1)
        self.assertIn(b"All Proposals", page)
        self.assertNotIn(b'class="proposal-contact-name"', page)
        self.assertNotIn(b'class="proposal-contact-email"', page)
        self.assertIn(b'<div class="contact-card-name">Taylor Jones</div>', page)
        self.assertIn(b'<div class="contact-card-email">taylor@example.com</div>', page)
        self.assertIn(b'<div class="contact-card-name">Casey Smith</div>', page)
        self.assertIn(
            b'<div class="contact-card-email">casey@example.com</div>', page
        )
        self.assertIn(f'/contacts?edit={CONTACT_ID}'.encode(), page)
        self.assertIn(b'aria-label="Edit contact Casey Smith"', page)
        self.assertIn(
            b'aria-label="Attach contact for Unlinked Draft - 321 Cedar Ln"',
            page,
        )
        self.assertIn(
            b'class="contact-card-empty contact-card-attach-link"',
            page,
        )
        self.assertIn(
            b'attach_to_proposal=44444444-4444-4444-8444-444444444444',
            page,
        )
        self.assertNotIn(b"save-contact-btn", page)

    def test_funnel_filter_menu_replaces_status_buttons_and_includes_finished(self):
        store = Mock()
        store.list_management_proposals.return_value = self.rows
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.get("/proposals")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b'id="proposalFilterButton"', response.data)
        self.assertIn(b'id="proposalFilterMenu"', response.data)
        self.assertIn(b'aria-haspopup="menu"', response.data)
        self.assertIn(b'aria-expanded="false"', response.data)
        self.assertIn(b'Current filter: All', response.data)
        self.assertIn(b'bi bi-funnel"', response.data)
        self.assertNotIn(b'bi bi-funnel-fill', response.data)
        self.assertIn(b'href="/proposals?filter=all"', response.data)
        self.assertIn(b'href="/proposals?filter=draft"', response.data)
        self.assertIn(b'href="/proposals?filter=sent"', response.data)
        self.assertIn(b'href="/proposals?filter=under_contract"', response.data)
        self.assertIn(b'href="/proposals?filter=finished"', response.data)
        self.assertIn(b'href="/proposals?filter=dead"', response.data)
        self.assertNotIn(b'href="/proposals?filter=open"', response.data)
        self.assertNotIn(b'href="/proposals?filter=unsent"', response.data)
        filter_links = response.data.decode().split('id="proposalFilterMenu"', 1)[1]
        self.assertLess(filter_links.index("filter=all"), filter_links.index("filter=draft"))
        self.assertLess(filter_links.index("filter=draft"), filter_links.index("filter=sent"))
        self.assertLess(filter_links.index("filter=sent"), filter_links.index("filter=under_contract"))
        self.assertLess(filter_links.index("filter=under_contract"), filter_links.index("filter=finished"))
        self.assertLess(filter_links.index("filter=finished"), filter_links.index("filter=dead"))
        self.assertIn(b'aria-current="true"', response.data)
        self.assertNotIn(b'<select', response.data)
        self.assertNotIn(b'id="openProposals"', response.data)
        self.assertNotIn(b'id="underContract"', response.data)
        store.list_management_proposals.assert_called_once_with(
            {"draft", "sent", "under_contract", "finished", "dead"}
        )

    def test_draft_unsent_filter_shows_only_draft_proposals(self):
        store = Mock()
        sent_without_date = dict(
            self.rows[0],
            id="55555555-5555-4555-8555-555555555555",
            name="Sent Without Date - 987 Spruce St",
            proposal_sent_date=None,
        )
        store.list_management_proposals.return_value = self.rows + [sent_without_date]
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.get("/proposals?filter=draft")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Draft Customer - 789 Pine Rd", response.data)
        self.assertIn(b"Unlinked Draft - 321 Cedar Ln", response.data)
        self.assertNotIn(b"Open Customer - 123 Main St", response.data)
        self.assertNotIn(b"Sent Without Date - 987 Spruce St", response.data)
        self.assertNotIn(b"Contract Customer - 456 Oak Ave", response.data)
        self.assertIn(b"Draft Unsent Proposals", response.data)
        store.list_management_proposals.assert_called_once_with({"draft"})

    def test_all_filter_uses_one_list_for_every_lifecycle_status(self):
        store = Mock()
        finished = dict(self.rows[0], id="finished-id", name="Finished Proposal")
        finished["status"] = "finished"
        dead = dict(self.rows[0], id="dead-id", name="Dead Proposal")
        dead["status"] = "dead"
        store.list_management_proposals.return_value = self.rows + [finished, dead]
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.get("/proposals?filter=all")

        self.assertEqual(response.status_code, 200)
        for name in (
            b"Open Customer - 123 Main St",
            b"Contract Customer - 456 Oak Ave",
            b"Draft Customer - 789 Pine Rd",
            b"Finished Proposal",
            b"Dead Proposal",
        ):
            self.assertIn(name, response.data)
        self.assertEqual(response.data.count(b'class="proposal-table"'), 1)
        store.list_management_proposals.assert_called_once_with(
            {"draft", "sent", "under_contract", "finished", "dead"}
        )

    def test_under_contract_tab_uses_joined_tracking_status(self):
        store = Mock()
        store.list_management_proposals.return_value = self.rows
        store.list_management_contact_options.return_value = []
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.get("/proposals?status=under")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Contract Customer - 456 Oak Ave", response.data)
        self.assertNotIn(b"Open Customer - 123 Main St", response.data)
        self.assertIn(b"folder_name=Contract+Folder", response.data)
        self.assertIn(b"read_only=Yes", response.data)
        self.assertIn(
            b'aria-label="Open proposal details for Contract Customer - 456 Oak Ave"',
            response.data,
        )
        self.assertIn(b"No contact assigned", response.data)
        self.assertNotIn(b"Attach Contact", response.data)

    def test_contact_api_assigns_or_creates_contact(self):
        store = Mock()
        store.assign_or_create_primary_contact.return_value = {
            "id": "relationship-id",
            "name": "Casey Smith",
            "email": "casey@example.com",
            "organization": "Example Roofing",
            "created": False,
        }
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.post(
                f"/api/proposals/{PROPOSAL_ID}/primary-contact",
                json={"organization_contact_id": "relationship-id"},
            )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["contact"]["name"], "Casey Smith")

    def test_contact_api_requests_organization_name_for_unknown_domain(self):
        store = Mock()
        store.assign_or_create_primary_contact.side_effect = (
            ProposalContactOrganizationRequired("new-roofer.com")
        )
        with patch.object(web, "get_proposal_tracking_store", return_value=store):
            response = self.client.post(
                f"/api/proposals/{PROPOSAL_ID}/primary-contact",
                json={
                    "contact_name": "Casey Smith",
                    "email": "casey@new-roofer.com",
                },
            )
        self.assertEqual(response.status_code, 409)
        self.assertTrue(response.get_json()["organization_required"])
        self.assertEqual(response.get_json()["domain"], "new-roofer.com")


if __name__ == "__main__":
    unittest.main()
