"""Disabled-by-default flags for the proposal-tracking Supabase cutover."""

from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Mapping


MASTER_FLAG = "PROPOSAL_TRACKING_SUPABASE_ENABLED"
READ_FLAG = "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED"
WRITE_FLAG = "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED"
SHADOW_WRITE_FLAG = "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED"

_TRUE_VALUES = {"1", "true", "yes", "on"}


def _enabled(environment: Mapping[str, str], name: str) -> bool:
    return str(environment.get(name, "0")).strip().lower() in _TRUE_VALUES


@dataclass(frozen=True)
class ProposalTrackingCutoverFlags:
    master_enabled: bool
    reads_enabled: bool
    writes_enabled: bool
    shadow_writes_enabled: bool

    @property
    def spreadsheet_reads_active(self) -> bool:
        return not self.reads_enabled

    @property
    def spreadsheet_writes_active(self) -> bool:
        return self.shadow_writes_enabled or not self.writes_enabled

    @property
    def fully_cut_over(self) -> bool:
        return self.reads_enabled and self.writes_enabled


def load_proposal_tracking_cutover_flags(
    environment: Mapping[str, str] | None = None,
) -> ProposalTrackingCutoverFlags:
    if environment is not None:
        values = environment
    else:
        from pcs_local_settings import proposal_tracking_cutover_environment

        values = proposal_tracking_cutover_environment()
        values.update({
            name: os.environ[name]
            for name in (MASTER_FLAG, READ_FLAG, WRITE_FLAG, SHADOW_WRITE_FLAG)
            if name in os.environ
        })
        values.update({
            name: os.environ[name]
            for name in ("PCS_PROPOSAL_STORAGE_MODE", "PCS_SUPABASE_ONLY")
            if name in os.environ
        })
    if "PCS_PROPOSAL_STORAGE_MODE" in values or "PCS_SUPABASE_ONLY" in values:
        from pcs_runtime_config import proposal_storage_environment, proposal_storage_mode

        normalized = dict(values)
        normalized.update(proposal_storage_environment(proposal_storage_mode(values)))
        values = normalized
    master = _enabled(values, MASTER_FLAG)
    return ProposalTrackingCutoverFlags(
        master_enabled=master,
        reads_enabled=master and _enabled(values, READ_FLAG),
        writes_enabled=master and _enabled(values, WRITE_FLAG),
        shadow_writes_enabled=master and _enabled(values, SHADOW_WRITE_FLAG),
    )


__all__ = [
    "MASTER_FLAG",
    "READ_FLAG",
    "SHADOW_WRITE_FLAG",
    "WRITE_FLAG",
    "ProposalTrackingCutoverFlags",
    "load_proposal_tracking_cutover_flags",
]
