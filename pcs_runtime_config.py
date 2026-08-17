"""Explicit runtime feature configuration shared by every PCS build."""

from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Mapping, MutableMapping


TRUE_VALUES = {"1", "true", "yes", "on"}
PROPOSAL_STORAGE_MODES = {"spreadsheet", "shadow", "supabase"}


def environment_flag(
    environment: Mapping[str, str], name: str, *, default: bool = False
) -> bool:
    raw_value = environment.get(name)
    if raw_value is None:
        return default
    return str(raw_value).strip().casefold() in TRUE_VALUES


def proposal_storage_mode(environment: Mapping[str, str] | None = None) -> str:
    """Return the authoritative proposal-tracking persistence mode.

    Explicit mode settings take precedence. Legacy cutover flags remain supported
    so an existing production installation does not change behavior during the
    integration release.
    """
    values = environment if environment is not None else os.environ
    configured = str(values.get("PCS_PROPOSAL_STORAGE_MODE") or "").strip().casefold()
    if configured:
        if configured not in PROPOSAL_STORAGE_MODES:
            choices = ", ".join(sorted(PROPOSAL_STORAGE_MODES))
            raise ValueError(
                f"PCS_PROPOSAL_STORAGE_MODE must be one of: {choices}."
            )
        return configured
    if environment_flag(values, "PCS_SUPABASE_ONLY"):
        return "supabase"
    master = environment_flag(values, "PROPOSAL_TRACKING_SUPABASE_ENABLED")
    reads = master and environment_flag(
        values, "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED"
    )
    writes = master and environment_flag(
        values, "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED"
    )
    shadow = master and environment_flag(
        values, "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED"
    )
    if reads and writes:
        return "supabase"
    if writes and shadow:
        return "shadow"
    return "spreadsheet"


def proposal_storage_environment(mode: str) -> dict[str, str]:
    normalized = str(mode or "").strip().casefold()
    if normalized not in PROPOSAL_STORAGE_MODES:
        choices = ", ".join(sorted(PROPOSAL_STORAGE_MODES))
        raise ValueError(f"Proposal storage mode must be one of: {choices}.")
    enabled = normalized != "spreadsheet"
    return {
        "PCS_PROPOSAL_STORAGE_MODE": normalized,
        "PCS_SUPABASE_ONLY": "1" if normalized == "supabase" else "0",
        "PROPOSAL_TRACKING_SUPABASE_ENABLED": "1" if enabled else "0",
        "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED": (
            "1" if normalized == "supabase" else "0"
        ),
        "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED": (
            "1" if enabled else "0"
        ),
        "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED": (
            "1" if normalized == "shadow" else "0"
        ),
    }


def apply_proposal_storage_mode(
    environment: MutableMapping[str, str], mode: str
) -> None:
    environment.update(proposal_storage_environment(mode))


@dataclass(frozen=True)
class RuntimeConfiguration:
    app_variant: str
    multi_tenant_enabled: bool
    proposal_storage_mode: str

    @property
    def proposal_database_source_enabled(self) -> bool:
        return self.proposal_storage_mode == "supabase"


def load_runtime_configuration(
    environment: Mapping[str, str] | None = None,
) -> RuntimeConfiguration:
    values = environment if environment is not None else os.environ
    return RuntimeConfiguration(
        app_variant=(
            str(values.get("PCS_APP_ENV") or "production").strip().casefold()
            or "production"
        ),
        multi_tenant_enabled=environment_flag(
            values, "PCS_MULTI_TENANT_ENABLED", default=False
        ),
        proposal_storage_mode=proposal_storage_mode(values),
    )


__all__ = [
    "PROPOSAL_STORAGE_MODES",
    "RuntimeConfiguration",
    "apply_proposal_storage_mode",
    "environment_flag",
    "load_runtime_configuration",
    "proposal_storage_environment",
    "proposal_storage_mode",
]
