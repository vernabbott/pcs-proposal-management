import unittest

from pcs_runtime_config import (
    load_runtime_configuration,
    proposal_storage_environment,
    proposal_storage_mode,
)


class RuntimeConfigurationTests(unittest.TestCase):
    def test_production_defaults_are_safe_and_backward_compatible(self):
        configuration = load_runtime_configuration({})
        self.assertEqual(configuration.app_variant, "production")
        self.assertFalse(configuration.multi_tenant_enabled)
        self.assertEqual(configuration.proposal_storage_mode, "spreadsheet")
        self.assertFalse(configuration.proposal_database_source_enabled)

    def test_tenant_security_is_independent_of_beta_name(self):
        configuration = load_runtime_configuration({
            "PCS_APP_ENV": "integration",
            "PCS_MULTI_TENANT_ENABLED": "1",
            "PCS_PROPOSAL_STORAGE_MODE": "supabase",
        })
        self.assertEqual(configuration.app_variant, "integration")
        self.assertTrue(configuration.multi_tenant_enabled)
        self.assertTrue(configuration.proposal_database_source_enabled)

    def test_storage_modes_map_to_legacy_cutover_flags(self):
        self.assertEqual(
            proposal_storage_environment("spreadsheet")[
                "PROPOSAL_TRACKING_SUPABASE_ENABLED"
            ],
            "0",
        )
        shadow = proposal_storage_environment("shadow")
        self.assertEqual(shadow["PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED"], "1")
        self.assertEqual(
            shadow["PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED"], "1"
        )
        supabase = proposal_storage_environment("supabase")
        self.assertEqual(supabase["PROPOSAL_TRACKING_SUPABASE_READS_ENABLED"], "1")
        self.assertEqual(supabase["PCS_SUPABASE_ONLY"], "1")

    def test_supabase_only_alias_selects_supabase_mode(self):
        self.assertEqual(proposal_storage_mode({"PCS_SUPABASE_ONLY": "true"}), "supabase")

    def test_invalid_storage_mode_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "PCS_PROPOSAL_STORAGE_MODE"):
            proposal_storage_mode({"PCS_PROPOSAL_STORAGE_MODE": "both"})


if __name__ == "__main__":
    unittest.main()
