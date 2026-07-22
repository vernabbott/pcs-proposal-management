import unittest
from pathlib import Path

from roof_intelligence_cutover_flags import (
    MASTER_FLAG,
    READ_FLAG,
    SHADOW_WRITE_FLAG,
    WORKER_FLAG,
    WRITE_FLAG,
    load_cutover_flags,
)


class RoofIntelligenceCutoverFlagTests(unittest.TestCase):
    def test_current_local_workflow_is_the_default(self):
        flags = load_cutover_flags({})
        self.assertTrue(flags.local_workflow_active)
        self.assertFalse(flags.master_enabled)
        self.assertFalse(flags.reads_enabled)
        self.assertFalse(flags.writes_enabled)
        self.assertFalse(flags.worker_enabled)
        self.assertFalse(flags.shadow_writes_enabled)
        self.assertTrue(flags.local_reads_active)
        self.assertTrue(flags.local_writes_active)
        self.assertTrue(flags.local_worker_active)
        self.assertFalse(flags.fully_cut_over)

    def test_subordinate_flags_are_inert_without_master(self):
        flags = load_cutover_flags(
            {
                READ_FLAG: "1",
                WRITE_FLAG: "1",
                WORKER_FLAG: "1",
                SHADOW_WRITE_FLAG: "1",
            }
        )
        self.assertTrue(flags.local_workflow_active)
        self.assertFalse(flags.reads_enabled)
        self.assertFalse(flags.writes_enabled)
        self.assertFalse(flags.worker_enabled)
        self.assertFalse(flags.shadow_writes_enabled)

    def test_master_supports_staged_activation(self):
        flags = load_cutover_flags(
            {MASTER_FLAG: "true", READ_FLAG: "true", WRITE_FLAG: "false"}
        )
        self.assertTrue(flags.local_workflow_active)
        self.assertFalse(flags.local_reads_active)
        self.assertTrue(flags.local_writes_active)
        self.assertTrue(flags.local_worker_active)
        self.assertTrue(flags.reads_enabled)
        self.assertFalse(flags.writes_enabled)

    def test_shadow_writes_preserve_local_authority(self):
        flags = load_cutover_flags(
            {
                MASTER_FLAG: "1",
                WRITE_FLAG: "1",
                SHADOW_WRITE_FLAG: "1",
            }
        )
        self.assertTrue(flags.writes_enabled)
        self.assertTrue(flags.shadow_writes_enabled)
        self.assertTrue(flags.local_writes_active)
        self.assertFalse(flags.fully_cut_over)

    def test_full_cutover_disables_all_local_paths(self):
        flags = load_cutover_flags(
            {
                MASTER_FLAG: "1",
                READ_FLAG: "1",
                WRITE_FLAG: "1",
                WORKER_FLAG: "1",
            }
        )
        self.assertTrue(flags.fully_cut_over)
        self.assertFalse(flags.local_workflow_active)

    def test_current_pcs_entry_points_do_not_import_cutover_flags(self):
        project_dir = Path(__file__).resolve().parent
        for file_name in ("pcs_proposal_web.py", "run_app.py", "roof_intelligence_jobs.py"):
            source = (project_dir / file_name).read_text(encoding="utf-8")
            self.assertNotIn("roof_intelligence_cutover_flags", source)


if __name__ == "__main__":
    unittest.main()
