"""Disabled-by-default flags for the staged reporting cutover.

PCS may inspect these flags, but all new behavior remains inactive unless the
master flag and the corresponding capability flag are both explicitly enabled.
"""

from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Mapping


MASTER_FLAG = "ROOF_INTELLIGENCE_SUPABASE_ENABLED"
READ_FLAG = "ROOF_INTELLIGENCE_SUPABASE_READS_ENABLED"
WRITE_FLAG = "ROOF_INTELLIGENCE_SUPABASE_WRITES_ENABLED"
WORKER_FLAG = "ROOF_INTELLIGENCE_SUPABASE_WORKER_ENABLED"
SHADOW_WRITE_FLAG = "ROOF_INTELLIGENCE_SUPABASE_SHADOW_WRITES_ENABLED"
EDITING_FLAG = "ROOF_INTELLIGENCE_REPORT_EDITING_ENABLED"

_TRUE_VALUES = {"1", "true", "yes", "on"}


def _enabled(environment: Mapping[str, str], name: str) -> bool:
    return str(environment.get(name, "0")).strip().lower() in _TRUE_VALUES


@dataclass(frozen=True)
class RoofIntelligenceCutoverFlags:
    master_enabled: bool
    reads_enabled: bool
    writes_enabled: bool
    worker_enabled: bool
    shadow_writes_enabled: bool
    editing_enabled: bool

    @property
    def local_reads_active(self) -> bool:
        return not self.reads_enabled

    @property
    def local_writes_active(self) -> bool:
        return self.shadow_writes_enabled or not self.writes_enabled

    @property
    def local_worker_active(self) -> bool:
        return not self.worker_enabled

    @property
    def local_workflow_active(self) -> bool:
        return (
            self.local_reads_active
            or self.local_writes_active
            or self.local_worker_active
        )

    @property
    def fully_cut_over(self) -> bool:
        return self.reads_enabled and self.writes_enabled and self.worker_enabled


def load_cutover_flags(
    environment: Mapping[str, str] | None = None,
) -> RoofIntelligenceCutoverFlags:
    values = environment if environment is not None else os.environ
    master = _enabled(values, MASTER_FLAG)
    return RoofIntelligenceCutoverFlags(
        master_enabled=master,
        reads_enabled=master and _enabled(values, READ_FLAG),
        writes_enabled=master and _enabled(values, WRITE_FLAG),
        worker_enabled=master and _enabled(values, WORKER_FLAG),
        shadow_writes_enabled=master and _enabled(values, SHADOW_WRITE_FLAG),
        editing_enabled=master and _enabled(values, EDITING_FLAG),
    )


__all__ = [
    "EDITING_FLAG",
    "MASTER_FLAG",
    "READ_FLAG",
    "SHADOW_WRITE_FLAG",
    "WRITE_FLAG",
    "WORKER_FLAG",
    "RoofIntelligenceCutoverFlags",
    "load_cutover_flags",
]
