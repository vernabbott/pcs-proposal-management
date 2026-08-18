import unittest
from email.message import EmailMessage
from unittest.mock import patch

import pcs_proposal_web
import run_app


class DesktopLifecycleDecisionTests(unittest.TestCase):
    def test_server_stays_alive_while_browser_session_is_active(self):
        self.assertFalse(
            pcs_proposal_web._desktop_session_should_stop(
                100,
                server_started_at=0,
                last_heartbeat=99,
                close_requested_at=None,
            )
        )

    def test_server_stops_after_browser_close_grace_period(self):
        self.assertFalse(
            pcs_proposal_web._desktop_session_should_stop(
                104,
                server_started_at=0,
                last_heartbeat=100,
                close_requested_at=100,
            )
        )
        self.assertTrue(
            pcs_proposal_web._desktop_session_should_stop(
                105,
                server_started_at=0,
                last_heartbeat=100,
                close_requested_at=100,
            )
        )

    def test_server_stops_when_browser_never_connects(self):
        self.assertTrue(
            pcs_proposal_web._desktop_session_should_stop(
                60,
                server_started_at=0,
                last_heartbeat=None,
                close_requested_at=None,
            )
        )

    def test_browser_lifecycle_handles_internal_navigation_and_safari_cache(self):
        script = pcs_proposal_web._DESKTOP_LIFECYCLE_SCRIPT
        self.assertIn("destination.origin === window.location.origin", script)
        self.assertIn("document.addEventListener('submit', markInternalNavigation", script)
        self.assertIn("window.addEventListener('pageshow'", script)
        self.assertIn("startHeartbeat();", script)
        self.assertIn("if (!internalNavigationPending)", script)


class LauncherEnvironmentTests(unittest.TestCase):
    @patch("run_app.subprocess.Popen")
    @patch("run_app.open", create=True)
    @patch("run_app.os.makedirs")
    def test_server_process_enables_desktop_lifecycle(self, _makedirs, open_mock, popen_mock):
        open_mock.return_value.close.return_value = None
        popen_mock.return_value.pid = 123

        run_app._start_server_process("127.0.0.1", 5050)

        environment = popen_mock.call_args.kwargs["env"]
        self.assertEqual(environment[run_app.SERVER_ONLY_ENV], "1")
        self.assertEqual(environment[run_app.DESKTOP_LIFECYCLE_ENV], "1")
        self.assertEqual(environment[run_app.LOCAL_ROOF_WORKER_ENV], "1")


class RequestLoggingSecurityTests(unittest.TestCase):
    def test_request_form_redacts_all_supported_credentials(self):
        logged = pcs_proposal_web._redacted_request_form(
            {
                "action": "save_supabase_configuration",
                "supabase_publishable_key": "sb_publishable_example",
                "supabase_service_role_key": "sb_secret_example",
                "google_maps_api_key": "AIza-example",
                "password": "example-password",
                "customer_name": "Visible Customer",
            }
        )

        self.assertEqual(logged["supabase_publishable_key"], "[REDACTED]")
        self.assertEqual(logged["supabase_service_role_key"], "[REDACTED]")
        self.assertEqual(logged["google_maps_api_key"], "[REDACTED]")
        self.assertEqual(logged["password"], "[REDACTED]")
        self.assertEqual(logged["customer_name"], "Visible Customer")

    def test_request_form_redacts_secret_value_under_unexpected_name(self):
        logged = pcs_proposal_web._redacted_request_form(
            {"configuration_value": "sb_secret_example"}
        )

        self.assertEqual(logged["configuration_value"], "[REDACTED]")


class WeeklyFollowUpEmailTests(unittest.TestCase):
    def test_follow_up_sender_is_always_vern(self):
        self.assertEqual(
            pcs_proposal_web.get_weekly_follow_up_sender_email(),
            "vern@procoatingsystems.com",
        )

    def test_new_outlook_sender_stamp_removes_stale_identity_metadata(self):
        message = EmailMessage()
        message["From"] = "richard@procoatingsystems.com"
        message["X-MS-Exchange-MessageSentRepresentingType"] = "1"
        message["X-MS-TNEF-Correlator"] = "stale"

        pcs_proposal_web._stamp_and_verify_new_outlook_sender(
            message,
            "vern@procoatingsystems.com",
        )

        self.assertEqual(message["From"], "vern@procoatingsystems.com")
        self.assertEqual(message["Sender"], "vern@procoatingsystems.com")
        self.assertEqual(message["Reply-To"], "vern@procoatingsystems.com")
        self.assertEqual(message["X-Unsent"], "1")
        self.assertNotIn("X-MS-Exchange-MessageSentRepresentingType", message)
        self.assertNotIn("X-MS-TNEF-Correlator", message)

    def test_follow_up_bcc_recipients_include_mark(self):
        self.assertEqual(
            pcs_proposal_web.get_weekly_follow_up_bcc_recipients(),
            ["mark@procoatingsystems.com"],
        )

    @patch("pcs_proposal_web._open_new_outlook_template_draft")
    @patch("pcs_proposal_web._is_running_new_outlook", return_value=True)
    @patch("pcs_proposal_web.sys.platform", "darwin")
    def test_new_outlook_follow_up_draft_passes_mark_as_bcc(
        self,
        _new_outlook_mock,
        open_draft_mock,
    ):
        open_draft_mock.return_value = "fallback:new-outlook-template"

        pcs_proposal_web._open_outlook_html_draft_for_submitter(
            "Follow-Up List",
            "Plain body",
            "<p>HTML body</p>",
            "David",
            ["david@procoatingsystems.com"],
        )

        open_draft_mock.assert_called_once_with(
            "Follow-Up List",
            "Plain body",
            "<p>HTML body</p>",
            ["david@procoatingsystems.com"],
            ["mark@procoatingsystems.com"],
            "vern@procoatingsystems.com",
        )

    @patch("pcs_proposal_web.subprocess.run")
    @patch("pcs_proposal_web._is_running_new_outlook", return_value=False)
    @patch("pcs_proposal_web.sys.platform", "darwin")
    def test_classic_outlook_follow_up_draft_adds_mark_as_bcc(
        self,
        _new_outlook_mock,
        run_mock,
    ):
        run_mock.return_value.stdout = "matched"

        pcs_proposal_web._open_outlook_html_draft_for_submitter(
            "Follow-Up List",
            "Plain body",
            "<p>HTML body</p>",
            "Richard",
            ["richard@procoatingsystems.com"],
        )

        command = run_mock.call_args.args[0]
        self.assertTrue(
            any(
                line.startswith("make new bcc recipient at end of bcc recipients of newMessage")
                for line in command
            )
        )
        self.assertEqual(
            command[-5:],
            [
                "vern@procoatingsystems.com",
                "richard@procoatingsystems.com",
                "Vern",
                "1",
                "mark@procoatingsystems.com",
            ],
        )
        self.assertEqual(command[-1], "mark@procoatingsystems.com")

    def test_all_proposal_summary_senders_are_vern(self):
        for submitter in ("David", "Lydia", "Mark", "Richard", "Randy", "Vern", ""):
            with self.subTest(submitter=submitter):
                self.assertEqual(
                    pcs_proposal_web.get_sender_email_for_submitter(submitter),
                    "vern@procoatingsystems.com",
                )

    @patch("pcs_proposal_web.build_proposal_summary_email_text", return_value="Plain body")
    @patch("pcs_proposal_web.build_proposal_summary_email_html", return_value="<p>HTML body</p>")
    @patch("pcs_proposal_web.build_proposal_folder_link", return_value="https://example.test/folder")
    @patch("pcs_proposal_web._open_new_outlook_template_draft")
    @patch("pcs_proposal_web._is_running_new_outlook", return_value=True)
    @patch("pcs_proposal_web.sys.platform", "darwin")
    def test_new_outlook_proposal_summary_explicitly_uses_vern(
        self,
        _new_outlook_mock,
        open_draft_mock,
        _folder_link_mock,
        _html_mock,
        _text_mock,
    ):
        open_draft_mock.return_value = "fallback:new-outlook-template"

        pcs_proposal_web.create_outlook_proposal_summary_draft(
            "Test Customer",
            "123 Test St",
            "David",
            10,
            1,
            0,
            "Metal",
            1000,
            "",
            "",
            "Test Folder",
        )

        self.assertEqual(
            open_draft_mock.call_args.kwargs["sender_email"],
            "vern@procoatingsystems.com",
        )


if __name__ == "__main__":
    unittest.main()
