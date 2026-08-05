import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pcs_local_settings import save_supabase_configuration


class SupabaseConfigurationValidationTests(unittest.TestCase):
    def test_beta_accepts_loopback_local_supabase(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            settings_path = Path(temporary_directory) / "settings.json"
            environment = {
                "PCS_APP_ENV": "beta",
                "PCS_SETTINGS_PATH": str(settings_path),
            }
            with patch.dict(os.environ, environment, clear=False):
                save_supabase_configuration(
                    "http://127.0.0.1:54321",
                    "eyJ" + "x" * 40,
                )

            saved = json.loads(settings_path.read_text(encoding="utf-8"))
            self.assertEqual(saved["supabase_url"], "http://127.0.0.1:54321")

    def test_production_rejects_local_supabase(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            environment = {
                "PCS_APP_ENV": "production",
                "PCS_SETTINGS_PATH": str(Path(temporary_directory) / "settings.json"),
            }
            with patch.dict(os.environ, environment, clear=False):
                with self.assertRaises(ValueError):
                    save_supabase_configuration(
                        "http://127.0.0.1:54321",
                        "eyJ" + "x" * 40,
                    )


if __name__ == "__main__":
    unittest.main()
