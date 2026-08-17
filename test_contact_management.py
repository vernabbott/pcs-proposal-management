import unittest
import uuid
from unittest.mock import Mock, patch

import pcs_proposal_web as web
from contact_store import ContactConfigurationError, ContactStore, ContactStoreError


CONTACT_ID = str(uuid.uuid4())
ORGANIZATION_ID = str(uuid.uuid4())


def sample_contact(current=True):
    return {
        "id": str(uuid.uuid4()),
        "title": "Regional Manager",
        "business_email": "casey@example.com",
        "business_phone": "303-555-0100",
        "mobile_phone": "",
        "branch_address_line_1": "250 Branch Way",
        "branch_address_line_2": "Suite 200",
        "branch_city": "Aurora",
        "branch_state": "CO",
        "branch_zip_code": "80012",
        "is_current": current,
        "do_not_contact": False,
        "contact": {
            "id": CONTACT_ID,
            "full_name": "Casey Morgan",
            "first_name": "Casey",
            "last_name": "Morgan",
            "linkedin_url": None,
            "notes": None,
        },
        "organization": {
            "id": ORGANIZATION_ID,
            "name": "Example Management",
            "organization_type": "Property Management",
            "main_office_address_line_1": "100 Main Street",
            "main_office_address_line_2": "Floor 4",
            "main_office_city": "Denver",
            "main_office_state": "CO",
            "main_office_zip_code": "80202",
        },
    }


class FakeContactStore:
    def __init__(self):
        self.created = []
        self.updated = []
        self.archived = []
        self.organizations_created = []
        self.organizations_resolved_from_email = []
        self.organizations_updated = []
        self.duplicate_matches = []
        self.contact_searches = []
        self.row = sample_contact()

    def list_contacts(self, *, search="", status="active"):
        self.contact_searches.append((search, status))
        return [self.row]

    def list_organizations(self):
        return [self.row["organization"]]

    def get_contact(self, contact_id):
        return self.row if contact_id == CONTACT_ID else None

    def find_duplicate_contacts(self, values):
        return self.duplicate_matches

    def create_contact(self, values):
        self.created.append(values)
        return CONTACT_ID

    def update_contact(self, contact_id, values):
        self.updated.append((contact_id, values))

    def archive_contact(self, contact_id):
        self.archived.append(contact_id)

    def create_organization(self, name, organization_type, values=None):
        self.organizations_created.append((name, organization_type, values))
        return ORGANIZATION_ID

    def find_organization_by_name(self, name):
        organization = self.row["organization"]
        return organization if organization["name"].casefold() == name.casefold() else None

    def resolve_organization_from_email(self, email):
        self.organizations_resolved_from_email.append(email)
        return ORGANIZATION_ID

    def update_organization(self, organization_id, name, organization_type, values):
        self.organizations_updated.append((organization_id, name, organization_type, values))


class ContactManagementRouteTests(unittest.TestCase):
    def setUp(self):
        web.app.config.update(TESTING=True, SECRET_KEY="test")
        self.client = web.app.test_client()
        self.store = FakeContactStore()
        self.store_patch = patch.object(web, "get_contact_store", return_value=self.store)
        self.store_patch.start()

    def tearDown(self):
        self.store_patch.stop()

    def test_landing_page_has_contact_management_option(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Contact Management", response.data)
        self.assertIn(b'href="/contacts"', response.data)

    def test_contact_page_lists_contact_and_edit_form(self):
        response = self.client.get(f"/contacts?edit={CONTACT_ID}")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Casey Morgan", response.data)
        self.assertIn(b"Example Management", response.data)
        self.assertIn(b"casey@example.com", response.data)
        self.assertIn(b"Business Phone", response.data)
        self.assertIn(b"303-555-0100", response.data)
        self.assertIn(b"Main Office Address", response.data)
        self.assertIn(b"100 Main Street", response.data)
        self.assertIn(b"Branch Address", response.data)
        self.assertIn(b"250 Branch Way", response.data)
        self.assertIn(b"Save Changes", response.data)
        self.assertIn(b'list="organization-options"', response.data)
        self.assertIn(b"Choose an organization or enter a new one", response.data)

    def test_contact_page_has_history_back_button_with_home_fallback(self):
        response = self.client.get("/contacts")
        self.assertEqual(response.status_code, 200)
        self.assertIn(
            b'id="contact-back-button" href="/"',
            response.data,
        )
        self.assertIn(b"window.history.back()", response.data)
        self.assertIn(b"previousPage.origin === window.location.origin", response.data)
        self.assertIn(b"bi bi-house-door", response.data)

    def test_contact_name_links_to_edit_form_and_chain_button_replaces_edit_button(self):
        response = self.client.get("/contacts")
        self.assertEqual(response.status_code, 200)
        self.assertIn(
            f'href="/contacts?edit={CONTACT_ID}&amp;status=active&amp;q=#contact-form" aria-label="Edit Casey Morgan"'.encode(),
            response.data,
        )
        self.assertIn(b'class="name contact-name-link"', response.data)
        self.assertIn(b'class="icon-button contact-link-button"', response.data)
        self.assertIn(f'data-contact-id="{CONTACT_ID}"'.encode(), response.data)
        self.assertIn(b'class="bi bi-link-45deg"', response.data)
        self.assertNotIn(b"bi-pencil", response.data)

    def test_assignment_mode_can_select_an_existing_contact(self):
        proposal_id = "11111111-1111-4111-8111-111111111111"
        relationship_id = self.store.row["id"]
        response = self.client.get(
            "/contacts",
            query_string={
                "attach_to_proposal": proposal_id,
                "proposal_name": "Example Roofing - 100 Main St",
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Attach a contact to Example Roofing - 100 Main St", response.data)
        self.assertIn(
            f'action="/proposals/{proposal_id}/contacts/{relationship_id}/attach"'.encode(),
            response.data,
        )
        self.assertIn(b"Select an active contact below", response.data)
        self.assertIn(b'name="attach_to_proposal"', response.data)

    def test_assignment_mode_uses_customer_name_as_organization_search(self):
        proposal_id = "11111111-1111-4111-8111-111111111111"
        response = self.client.get(
            "/contacts",
            query_string={
                "attach_to_proposal": proposal_id,
                "proposal_name": "Example Roofing - 100 Main St",
                "q": "Example Roofing",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn(b'value="Example Roofing"', response.data)
        self.assertIn(("Example Roofing", "active"), self.store.contact_searches)

    def test_existing_contact_selection_attaches_and_returns_to_proposals(self):
        proposal_id = "11111111-1111-4111-8111-111111111111"
        relationship_id = self.store.row["id"]
        tracking_store = Mock()
        tracking_store.assign_or_create_primary_contact.return_value = {
            "name": "Casey Morgan"
        }
        with patch.object(web, "get_proposal_tracking_store", return_value=tracking_store):
            response = self.client.post(
                f"/proposals/{proposal_id}/contacts/{relationship_id}/attach",
                data={"proposal_name": "Example Roofing - 100 Main St"},
            )
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith("/proposals"))
        tracking_store.assign_or_create_primary_contact.assert_called_once_with(
            proposal_id,
            organization_contact_id=relationship_id,
        )

    def test_detail_contact_selection_returns_to_detail_and_prefills_blank_customer(self):
        proposal_id = "11111111-1111-4111-8111-111111111111"
        relationship_id = self.store.row["id"]
        tracking_store = Mock()
        tracking_store.assign_or_create_primary_contact.return_value = {
            "name": "Casey Morgan",
            "organization": "Example Management",
        }
        with patch.object(web, "get_proposal_tracking_store", return_value=tracking_store):
            start = self.client.get(
                "/contacts",
                query_string={
                    "attach_to_proposal": proposal_id,
                    "proposal_name": "New Proposal ABC12345",
                    "return_to_detail": "1",
                    "proposal_folder_name": "New Proposal ABC12345",
                    "customer_was_blank": "1",
                },
            )
            self.assertEqual(start.status_code, 200)
            response = self.client.post(
                f"/proposals/{proposal_id}/contacts/{relationship_id}/attach",
                data={"proposal_name": "New Proposal ABC12345"},
            )

        self.assertEqual(response.status_code, 302)
        self.assertIn("/proposal_details?", response.location)
        self.assertIn("folder_name=New+Proposal+ABC12345", response.location)
        self.assertIn(f"proposal_id={proposal_id}", response.location)
        tracking_store.update_proposal_customer_name.assert_called_once_with(
            proposal_id,
            "Example Management",
        )

    def test_detail_contact_without_organization_returns_with_customer_blank(self):
        proposal_id = "11111111-1111-4111-8111-111111111111"
        relationship_id = self.store.row["id"]
        tracking_store = Mock()
        tracking_store.assign_or_create_primary_contact.return_value = {
            "name": "Casey Morgan",
            "organization": "",
        }
        with patch.object(web, "get_proposal_tracking_store", return_value=tracking_store):
            self.client.get(
                "/contacts",
                query_string={
                    "attach_to_proposal": proposal_id,
                    "return_to_detail": "1",
                    "proposal_folder_name": "New Proposal ABC12345",
                    "customer_was_blank": "1",
                },
            )
            response = self.client.post(
                f"/proposals/{proposal_id}/contacts/{relationship_id}/attach"
            )

        self.assertEqual(response.status_code, 302)
        self.assertIn("customer_was_blank=1", response.location)
        tracking_store.update_proposal_customer_name.assert_not_called()

    def test_new_contact_is_created_attached_and_returns_to_proposals(self):
        proposal_id = "11111111-1111-4111-8111-111111111111"
        tracking_store = Mock()
        tracking_store.assign_or_create_primary_contact.return_value = {
            "name": "Avery Lee"
        }
        with patch.object(web, "get_proposal_tracking_store", return_value=tracking_store):
            response = self.client.post(
                "/contacts",
                data={
                    "first_name": "Avery",
                    "last_name": "Lee",
                    "organization_id": ORGANIZATION_ID,
                    "organization_name": "Example Management",
                    "organization_type": "Property Management",
                    "business_email": "avery@example.com",
                    "attach_to_proposal": proposal_id,
                    "proposal_name": "Example Roofing - 100 Main St",
                },
            )
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith("/proposals"))
        tracking_store.assign_or_create_primary_contact.assert_called_once_with(
            proposal_id,
            organization_contact_id=self.store.row["id"],
        )

    def test_create_contact_uses_existing_organization(self):
        response = self.client.post(
            "/contacts",
            data={
                "first_name": "Avery",
                "last_name": "Lee",
                "organization_id": ORGANIZATION_ID,
                "organization_name": "Example Management",
                "organization_type": "Property Management",
                "business_email": "avery@example.com",
                "business_phone": "303-555-0110",
                "main_office_address_line_1": "100 Main Street",
                "main_office_city": "Denver",
                "main_office_state": "CO",
                "main_office_zip_code": "80202",
                "branch_address_line_1": "250 Branch Way",
                "branch_city": "Aurora",
                "branch_state": "CO",
                "branch_zip_code": "80012",
                "title": "Manager",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.store.created[0]["organization_id"], ORGANIZATION_ID)
        self.assertEqual(self.store.created[0]["first_name"], "Avery")
        self.assertEqual(self.store.created[0]["business_phone"], "303-555-0110")
        self.assertEqual(self.store.created[0]["branch_address_line_1"], "250 Branch Way")
        self.assertEqual(self.store.organizations_updated[0][0], ORGANIZATION_ID)
        self.assertEqual(self.store.organizations_updated[0][1:3], ("Example Management", "Property Management"))
        self.assertEqual(
            self.store.organizations_updated[0][3]["main_office_address_line_1"],
            "100 Main Street",
        )

    def test_create_contact_can_create_organization_inline(self):
        self.client.post(
            "/contacts",
            data={
                "first_name": "Avery",
                "organization_name": "New Roofers",
                "organization_type": "Roofing Contractor",
                "main_office_address_line_1": "500 Roofers Road",
                "branch_address_line_1": "600 Branch Road",
            },
        )
        self.assertEqual(self.store.organizations_created[0][0:2], ("New Roofers", "Roofing Contractor"))
        self.assertEqual(
            self.store.organizations_created[0][2]["main_office_address_line_1"],
            "500 Roofers Road",
        )
        self.assertEqual(self.store.created[0]["branch_address_line_1"], "600 Branch Road")
        self.assertEqual(self.store.created[0]["organization_id"], ORGANIZATION_ID)

    def test_create_contact_uses_email_domain_when_organization_is_not_selected(self):
        response = self.client.post(
            "/contacts",
            data={
                "first_name": "Avery",
                "business_email": "avery@Example-Roofing.com",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(
            self.store.organizations_resolved_from_email,
            ["avery@Example-Roofing.com"],
        )
        self.assertEqual(self.store.created[0]["organization_id"], ORGANIZATION_ID)

    def test_possible_duplicate_shows_three_choices_without_writing(self):
        self.store.duplicate_matches = [self.store.row]
        response = self.client.post(
            "/contacts",
            data={
                "first_name": "Casey",
                "last_name": "Morgan",
                "organization_id": ORGANIZATION_ID,
                "organization_name": "Example Management",
                "organization_type": "Property Management",
                "business_email": "casey@example.com",
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Possible duplicate contact", response.data)
        self.assertIn(b"Replace Selected", response.data)
        self.assertIn(b"Keep Both", response.data)
        self.assertIn(b">Stop<", response.data)
        self.assertEqual(self.store.created, [])
        self.assertEqual(self.store.updated, [])
        self.assertEqual(self.store.organizations_updated, [])

    def test_keep_both_creates_a_separate_contact(self):
        self.store.duplicate_matches = [self.store.row]
        response = self.client.post(
            "/contacts",
            data={
                "first_name": "Casey",
                "last_name": "Morgan",
                "organization_id": ORGANIZATION_ID,
                "organization_name": "Example Management",
                "organization_type": "Property Management",
                "business_email": "casey@example.com",
                "duplicate_action": "keep",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(len(self.store.created), 1)
        self.assertEqual(self.store.updated, [])

    def test_replace_updates_selected_existing_contact(self):
        self.store.duplicate_matches = [self.store.row]
        response = self.client.post(
            "/contacts",
            data={
                "first_name": "Casey",
                "last_name": "Morgan",
                "organization_id": ORGANIZATION_ID,
                "organization_name": "Example Management",
                "organization_type": "Property Management",
                "business_email": "casey@example.com",
                "duplicate_action": "replace",
                "duplicate_contact_id": CONTACT_ID,
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.store.created, [])
        self.assertEqual(self.store.updated[0][0], CONTACT_ID)

    def test_edit_and_delete_routes_call_store(self):
        edit_response = self.client.post(
            f"/contacts/{CONTACT_ID}/edit",
            data={
                "first_name": "Casey",
                "last_name": "Morgan",
                "organization_id": ORGANIZATION_ID,
                "organization_name": "Example Management",
                "organization_type": "Property Management",
            },
        )
        delete_response = self.client.post(f"/contacts/{CONTACT_ID}/delete")
        self.assertEqual(edit_response.status_code, 302)
        self.assertEqual(delete_response.status_code, 302)
        self.assertEqual(self.store.updated[0][0], CONTACT_ID)
        self.assertEqual(self.store.archived, [CONTACT_ID])

    def test_unconfigured_page_shows_settings_link_without_failing(self):
        with patch.object(web, "get_contact_store", side_effect=ContactConfigurationError("Not configured")):
            response = self.client.get("/contacts")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Contact data is unavailable", response.data)
        self.assertIn(b"Open Local Settings", response.data)


class ContactStoreDuplicateTests(unittest.TestCase):
    def setUp(self):
        self.store = ContactStore("https://example.supabase.co", "test-key")
        self.store.list_contacts = lambda **kwargs: [sample_contact()]

    def test_exact_email_match_is_case_insensitive(self):
        matches = self.store.find_duplicate_contacts({
            "first_name": "Different",
            "last_name": "Person",
            "business_email": "CASEY@EXAMPLE.COM",
        })
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["duplicate_reasons"], ["Same business email"])

    def test_exact_first_and_last_name_is_a_possible_duplicate(self):
        matches = self.store.find_duplicate_contacts({
            "first_name": " casey ",
            "last_name": "MORGAN",
            "business_email": "different@example.com",
        })
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["duplicate_reasons"], ["Same first and last name"])


class ContactStoreRelationshipStatusTests(unittest.TestCase):
    def setUp(self):
        self.store = ContactStore("https://example.supabase.co", "test-key")

    def test_archive_marks_current_relationship_inactive(self):
        with patch.object(self.store, "_request") as request:
            self.store.archive_contact(CONTACT_ID)

        request.assert_called_once_with(
            "organization_contact",
            method="PATCH",
            params={"contact_id": f"eq.{CONTACT_ID}", "is_current": "eq.true"},
            payload={"is_current": False},
        )


class ContactStoreEmailOrganizationTests(unittest.TestCase):
    def setUp(self):
        self.store = ContactStore("https://example.supabase.co", "test-key")

    def test_resolve_organization_reuses_existing_domain_record(self):
        with patch.object(
            self.store,
            "_request",
            return_value=[{
                "id": ORGANIZATION_ID,
                "normalized_name": "example.com",
                "email_domain": "example.com",
                "source_name": "Contact email domain",
            }],
        ) as request:
            organization_id = self.store.resolve_organization_from_email(
                "Casey@EXAMPLE.COM"
            )

        self.assertEqual(organization_id, ORGANIZATION_ID)
        request.assert_called_once_with(
            "organization",
            params={
                "select": "id,normalized_name,email_domain,source_name",
                "or": "(normalized_name.eq.example.com,email_domain.ilike.example.com)",
                "limit": "20",
            },
        )

    def test_resolve_organization_requires_an_email_domain(self):
        with self.assertRaisesRegex(
            ContactStoreError,
            "Enter a business email address or select an organization",
        ):
            self.store.resolve_organization_from_email("")

    def test_find_organization_for_email_returns_unique_domain_match(self):
        organization = {
            "id": ORGANIZATION_ID,
            "name": "Example Roofing",
            "normalized_name": "example roofing",
            "email_domain": "example.com",
            "organization_type": "Roofing Company",
        }
        with patch.object(self.store, "_request", return_value=[organization]) as request:
            result = self.store.find_organization_for_email("casey@EXAMPLE.com")
        self.assertEqual(result, organization)
        self.assertEqual(request.call_args.args[0], "organization")
        self.assertEqual(
            request.call_args.kwargs["params"]["or"],
            "(normalized_name.eq.example.com,email_domain.ilike.example.com)",
        )

    def test_shared_email_domain_requires_user_named_organization(self):
        with patch.object(self.store, "_request") as request:
            self.assertIsNone(
                self.store.find_organization_for_email("casey@gmail.com")
            )
        request.assert_not_called()

    def test_named_organization_creation_records_nonshared_email_domain(self):
        with patch.object(self.store, "find_organization_by_name", return_value=None), patch.object(
            self.store, "_request", return_value=[{"id": ORGANIZATION_ID}]
        ) as request:
            organization_id = self.store.resolve_named_organization_for_email(
                "Example Roofing", "casey@example.com"
            )
        self.assertEqual(organization_id, ORGANIZATION_ID)
        self.assertEqual(request.call_args.kwargs["payload"]["name"], "Example Roofing")
        self.assertEqual(request.call_args.kwargs["payload"]["email_domain"], "example.com")

    def test_named_organization_does_not_claim_shared_email_domain(self):
        with patch.object(self.store, "find_organization_by_name", return_value=None), patch.object(
            self.store, "_request", return_value=[{"id": ORGANIZATION_ID}]
        ) as request:
            self.store.resolve_named_organization_for_email(
                "Example Roofing", "casey@gmail.com"
            )
        self.assertNotIn("email_domain", request.call_args.kwargs["payload"])

    def test_selected_shared_domains_use_unknown_organization(self):
        for domain in (
            "aol.com",
            "comcast.net",
            "fastmail.com",
            "gmail.com",
            "hotmail.com",
            "icloud.com",
            "mail.com",
            "me.com",
            "m.knck.io",
            "msn.com",
            "yahoo.com",
        ):
            with self.subTest(domain=domain), patch.object(
                self.store,
                "_request",
                return_value=[{"id": ORGANIZATION_ID}],
            ) as request:
                organization_id = self.store.resolve_organization_from_email(
                    f"casey@{domain}"
                )

            self.assertEqual(organization_id, ORGANIZATION_ID)
            request.assert_called_once_with(
                "organization",
                params={
                    "select": "id",
                    "normalized_name": "eq.unknown",
                    "limit": "1",
                },
            )

    def test_resolve_organization_reuses_domain_record_after_it_is_renamed(self):
        renamed_organization = {
            "id": ORGANIZATION_ID,
            "normalized_name": "example roofing",
            "email_domain": "example.com",
            "source_name": "Contact email domain",
        }
        with patch.object(
            self.store,
            "_request",
            return_value=[renamed_organization],
        ):
            organization_id = self.store.resolve_organization_from_email(
                "casey@example.com"
            )

        self.assertEqual(organization_id, ORGANIZATION_ID)

    def test_resolve_organization_creates_normalized_domain_record(self):
        responses = [
            [],
            [{"id": ORGANIZATION_ID}],
        ]
        with patch.object(self.store, "_request", side_effect=responses) as request:
            organization_id = self.store.resolve_organization_from_email(
                "Casey@Example.COM."
            )

        self.assertEqual(organization_id, ORGANIZATION_ID)
        self.assertEqual(
            request.call_args_list[1].kwargs["payload"],
            {
                "name": "example.com",
                "organization_type": "Unknown",
                "email_domain": "example.com",
            },
        )


if __name__ == "__main__":
    unittest.main()
