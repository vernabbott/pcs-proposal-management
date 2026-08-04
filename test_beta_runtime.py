import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from beta_runtime import apply_beta_environment, beta_environment
from pcs_local_settings import supabase_configuration


class BetaRuntimeIsolationTests(unittest.TestCase):
    def test_beta_paths_are_separate_from_production(self):
        values = beta_environment(
            home=Path("/Users/tester"),
            pilotpoint_project_dir=Path("/projects/PilotPoint Beta"),
        )
        self.assertEqual(values["PCS_DEFAULT_PORT"], "5051")
        self.assertEqual(values["PCS_SERVER_STARTUP_TIMEOUT"], "90")
        self.assertEqual(values["PCS_APP_ENV"], "beta")
        self.assertIn("PCS Proposal Management Beta", values["PCS_DATA_DIR"])
        self.assertIn("PilotPoint Beta", values["ROOF_INTELLIGENCE_PROJECT_DIR"])
        self.assertNotIn("OneDrive-ProfessionalCoatingSystems", values["PCS_PROPOSAL_TRACKER"])
        self.assertEqual(values["PROPOSAL_TRACKING_SUPABASE_ENABLED"], "0")

    def test_beta_does_not_inherit_production_supabase_credentials(self):
        environment = {
            "PCS_SUPABASE_URL": "https://production.supabase.co",
            "PCS_SUPABASE_SERVICE_ROLE_KEY": "production-secret",
        }
        apply_beta_environment(environment)
        self.assertNotIn("PCS_SUPABASE_URL", environment)
        self.assertNotIn("PCS_SUPABASE_SERVICE_ROLE_KEY", environment)

    def test_explicit_beta_supabase_credentials_are_mapped(self):
        environment = {
            "PCS_BETA_SUPABASE_URL": "https://beta.supabase.co",
            "PCS_BETA_SUPABASE_SERVICE_ROLE_KEY": "beta-secret",
        }
        apply_beta_environment(environment)
        self.assertEqual(environment["PCS_SUPABASE_URL"], "https://beta.supabase.co")
        self.assertEqual(environment["PCS_SUPABASE_SERVICE_ROLE_KEY"], "beta-secret")

    def test_beta_settings_have_no_default_supabase_project(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            settings = str(Path(temporary_directory) / "settings.json")
            with patch.dict(
                os.environ,
                {"PCS_SETTINGS_PATH": settings},
                clear=True,
            ):
                self.assertEqual(supabase_configuration(), ("", ""))


if __name__ == "__main__":
    unittest.main()
