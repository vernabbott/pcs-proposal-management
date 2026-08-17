from pathlib import Path
import unittest

from integration_runtime import apply_integration_environment, integration_environment


class IntegrationRuntimeTests(unittest.TestCase):
    def test_integration_app_is_isolated_and_production_capable(self):
        values = integration_environment(
            home=Path("/Users/tester"),
            pilotpoint_project_dir=Path("/projects/PilotPoint Beta"),
        )
        self.assertEqual(values["PCS_APP_ENV"], "integration")
        self.assertEqual(values["PCS_DEFAULT_PORT"], "5052")
        self.assertEqual(values["PCS_MULTI_TENANT_ENABLED"], "1")
        self.assertEqual(values["PCS_PROPOSAL_STORAGE_MODE"], "supabase")
        self.assertEqual(values["PCS_SUPABASE_ONLY"], "1")
        self.assertIn("PCS Proposal Integration", values["PCS_DATA_DIR"])
        self.assertIn("PilotPoint Beta", values["ROOF_INTELLIGENCE_PROJECT_DIR"])
        self.assertNotIn(
            "OneDrive-ProfessionalCoatingSystems", values["PCS_PROPOSAL_TRACKER"]
        )

    def test_integration_runtime_uses_its_own_persistent_supabase_settings(self):
        environment = {
            "PCS_SUPABASE_URL": "https://production.example.supabase.co",
            "PCS_SUPABASE_PUBLISHABLE_KEY": "production-key",
        }

        apply_integration_environment(environment)

        self.assertNotIn("PCS_SUPABASE_URL", environment)
        self.assertNotIn("PCS_SUPABASE_PUBLISHABLE_KEY", environment)
        self.assertEqual(environment["PCS_ALLOW_LOCAL_SUPABASE"], "1")

    def test_explicit_integration_project_overrides_persistent_settings(self):
        environment = {
            "PCS_INTEGRATION_SUPABASE_URL": "https://integration.example.supabase.co",
            "PCS_INTEGRATION_SUPABASE_PUBLISHABLE_KEY": "integration-key",
        }

        apply_integration_environment(environment)

        self.assertEqual(
            environment["PCS_SUPABASE_URL"],
            "https://integration.example.supabase.co",
        )
        self.assertEqual(
            environment["PCS_SUPABASE_PUBLISHABLE_KEY"], "integration-key"
        )


if __name__ == "__main__":
    unittest.main()
