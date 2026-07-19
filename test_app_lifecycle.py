import unittest
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


if __name__ == "__main__":
    unittest.main()
