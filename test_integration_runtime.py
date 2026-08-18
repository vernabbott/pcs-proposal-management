from pathlib import Path
import tempfile
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
        self.assertEqual(values["PCS_PROPOSAL_STORAGE_MODE"], "supabase_shadow")
        self.assertEqual(values["PCS_SUPABASE_ONLY"], "0")
        self.assertEqual(
            values["PROPOSAL_TRACKING_SUPABASE_READS_ENABLED"], "1"
        )
        self.assertEqual(
            values["PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED"], "1"
        )
        self.assertIn("PCS Proposal Integration", values["PCS_DATA_DIR"])
        self.assertIn("PilotPoint Beta", values["ROOF_INTELLIGENCE_PROJECT_DIR"])
        self.assertEqual(
            values["PCS_EMAIL_TEMPLATE_DIR"],
            "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
            "PCS/Marketing/Email Templates",
        )
        self.assertEqual(
            values["PCS_EMAIL_LIST_DIR"],
            "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
            "PCS/Marketing/Email Lists",
        )
        self.assertEqual(
            values["PCS_PROPOSAL_TRACKER"],
            "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
            "PCS/1 - Open Proposals/Proposal Tracking.xlsx",
        )

    def test_integration_runtime_uses_its_own_persistent_supabase_settings(self):
        environment = {
            "PCS_SUPABASE_URL": "https://production.example.supabase.co",
            "PCS_SUPABASE_PUBLISHABLE_KEY": "production-key",
            "PCS_INTEGRATION_WORKER_ENV_FILE": "/missing/worker.env",
        }

        apply_integration_environment(environment)

        self.assertNotIn("PCS_SUPABASE_URL", environment)
        self.assertNotIn("PCS_SUPABASE_PUBLISHABLE_KEY", environment)
        self.assertEqual(environment["PCS_ALLOW_LOCAL_SUPABASE"], "1")

    def test_explicit_integration_project_overrides_persistent_settings(self):
        environment = {
            "PCS_INTEGRATION_SUPABASE_URL": "https://integration.example.supabase.co",
            "PCS_INTEGRATION_SUPABASE_PUBLISHABLE_KEY": "integration-key",
            "PCS_INTEGRATION_WORKER_ENV_FILE": "/missing/worker.env",
        }

        apply_integration_environment(environment)

        self.assertEqual(
            environment["PCS_SUPABASE_URL"],
            "https://integration.example.supabase.co",
        )
        self.assertEqual(
            environment["PCS_SUPABASE_PUBLISHABLE_KEY"], "integration-key"
        )

    def test_private_worker_environment_loads_only_approved_credentials(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            worker_environment = Path(temporary_dir) / "worker.env"
            worker_environment.write_text(
                "DATABASE_URL=postgresql://worker.example/test\n"
                "OPENAI_API_KEY='private-openai-key'\n"
                "PCS_SUPABASE_SERVICE_ROLE_KEY=must-not-load\n",
                encoding="utf-8",
            )
            environment = {
                "PCS_INTEGRATION_WORKER_ENV_FILE": str(worker_environment),
            }

            apply_integration_environment(environment)

            self.assertEqual(
                environment["DATABASE_URL"], "postgresql://worker.example/test"
            )
            self.assertEqual(environment["OPENAI_API_KEY"], "private-openai-key")
            self.assertNotIn("PCS_SUPABASE_SERVICE_ROLE_KEY", environment)

    def test_private_worker_environment_does_not_override_process_values(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            worker_environment = Path(temporary_dir) / "worker.env"
            worker_environment.write_text(
                "DATABASE_URL=postgresql://file.example/test\n",
                encoding="utf-8",
            )
            environment = {
                "PCS_INTEGRATION_WORKER_ENV_FILE": str(worker_environment),
                "DATABASE_URL": "postgresql://process.example/test",
            }

            apply_integration_environment(environment)

            self.assertEqual(
                environment["DATABASE_URL"], "postgresql://process.example/test"
            )


if __name__ == "__main__":
    unittest.main()
