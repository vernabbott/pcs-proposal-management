import unittest
from unittest.mock import patch

import pcs_proposal_web


class TenantInvitationRouteTests(unittest.TestCase):
    def setUp(self):
        pcs_proposal_web.app.config.update(TESTING=True)
        self.client = pcs_proposal_web.app.test_client()

    def test_invitation_page_is_available_without_an_authenticated_session(self):
        with patch(
            "pcs_proposal_web.supabase_configuration",
            return_value=("https://example.supabase.co", "sb_publishable_example"),
        ):
            response = self.client.get(
                "/auth/accept-invite#access_token=browser-only-token&type=invite"
            )

        self.assertEqual(response.status_code, 200)
        body = response.get_data(as_text=True)
        self.assertIn("Set your PCS password", body)
        self.assertIn("window.history.replaceState", body)
        self.assertIn("/auth/v1/user", body)
        self.assertIn("sb_publishable_example", body)
        self.assertNotIn("browser-only-token", body)

    def test_invitation_page_fails_closed_without_supabase_configuration(self):
        with patch(
            "pcs_proposal_web.supabase_configuration", return_value=("", "")
        ):
            response = self.client.get("/auth/accept-invite")

        self.assertEqual(response.status_code, 503)
        self.assertIn(
            "PCS is not connected to Supabase", response.get_data(as_text=True)
        )

    def test_sign_in_uses_runtime_application_name(self):
        response = self.client.get("/sign-in")

        self.assertEqual(response.status_code, 200)
        body = response.get_data(as_text=True)
        self.assertIn("Sign in · PCS Proposal", body)
        self.assertIn("Sign in to PCS Proposal", body)
        self.assertNotIn("PCS Beta", body)
        self.assertNotIn('<span class="beta">BETA</span>', body)


if __name__ == "__main__":
    unittest.main()
