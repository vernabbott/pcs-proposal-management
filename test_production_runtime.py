import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from production_runtime import apply_production_environment, production_environment


class ProductionRuntimeTests(unittest.TestCase):
    def test_production_environment_uses_cutover_configuration(self):
        values = production_environment(home=Path("/tmp/test-home"))

        self.assertEqual(values["PCS_APP_ENV"], "production")
        self.assertEqual(values["PCS_MULTI_TENANT_ENABLED"], "1")
        self.assertEqual(values["PCS_PROPOSAL_STORAGE_MODE"], "supabase_shadow")
        self.assertEqual(values["PCS_DEFAULT_PORT"], "5050")
        self.assertIn("PCS Proposal Management", values["PCS_DATA_DIR"])
        self.assertTrue(values["PCS_PROPOSAL_TRACKER"].endswith("Proposal Tracking.xlsx"))

    def test_apply_production_environment_removes_privileged_client_key(self):
        environment = {
            "PCS_SUPABASE_SERVICE_ROLE_KEY": "must-not-survive",
            "PCS_SUPABASE_URL": "https://wrong.example",
            "PCS_SUPABASE_PUBLISHABLE_KEY": "wrong-key",
            "PCS_PRODUCTION_WORKER_ENV_FILE": "/missing/worker.env",
        }

        apply_production_environment(environment)

        self.assertNotIn("PCS_SUPABASE_SERVICE_ROLE_KEY", environment)
        self.assertNotIn("PCS_SUPABASE_URL", environment)
        self.assertNotIn("PCS_SUPABASE_PUBLISHABLE_KEY", environment)

    def test_explicit_production_client_configuration_is_supported(self):
        environment = {
            "PCS_PRODUCTION_SUPABASE_URL": "https://production.example",
            "PCS_PRODUCTION_SUPABASE_PUBLISHABLE_KEY": "publishable-key",
            "PCS_PRODUCTION_WORKER_ENV_FILE": "/missing/worker.env",
        }

        apply_production_environment(environment)

        self.assertEqual(
            environment["PCS_SUPABASE_URL"], "https://production.example"
        )
        self.assertEqual(
            environment["PCS_SUPABASE_PUBLISHABLE_KEY"], "publishable-key"
        )

    def test_worker_environment_is_loaded_from_private_file(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            worker_environment = Path(temporary_dir) / "worker.env"
            worker_environment.write_text(
                "DATABASE_URL=postgresql://worker\nOPENAI_API_KEY=test-key\n",
                encoding="utf-8",
            )
            environment = {
                "PCS_PRODUCTION_WORKER_ENV_FILE": str(worker_environment),
            }

            with patch.dict(os.environ, {}, clear=True):
                apply_production_environment(environment)

            self.assertEqual(environment["DATABASE_URL"], "postgresql://worker")
            self.assertEqual(environment["OPENAI_API_KEY"], "test-key")


if __name__ == "__main__":
    unittest.main()
