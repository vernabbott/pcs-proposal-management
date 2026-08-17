"""Isolated runtime configuration for the PCS production integration build."""

from __future__ import annotations

import os
from pathlib import Path
from typing import MutableMapping

from pcs_runtime_config import apply_proposal_storage_mode


INTEGRATION_PILOTPOINT_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/"
    "PilotPoint IQ Roof Intelligence Report Beta"
)


def integration_environment(
    *,
    home: Path | None = None,
    pilotpoint_project_dir: Path | None = None,
) -> dict[str, str]:
    home = home or Path.home()
    data_dir = home / "Library" / "Application Support" / "PCS Proposal Integration"
    workspace = data_dir / "Workspace"
    accounts = workspace / "Accounts"
    pilotpoint_dir = pilotpoint_project_dir or INTEGRATION_PILOTPOINT_PROJECT_DIR
    values = {
        "PCS_APP_ENV": "integration",
        "PCS_APP_DISPLAY_NAME": "PCS Proposal Integration",
        "PCS_APP_STATE_DIR": "/tmp/pcs_proposal_integration_app",
        "PCS_DEFAULT_PORT": "5052",
        "PCS_SERVER_STARTUP_TIMEOUT": "90",
        "PCS_DATA_DIR": str(data_dir),
        "PCS_SETTINGS_PATH": str(data_dir / "settings.json"),
        "PCS_XLWINGS_LOG_PATH": str(data_dir / "pcs_xlwings.log"),
        "PCS_MULTI_TENANT_ENABLED": "1",
        "PCS_ALLOW_LOCAL_SUPABASE": "1",
        "ROOF_INTELLIGENCE_PROJECT_DIR": str(pilotpoint_dir),
        "ROOF_INTELLIGENCE_USER_KEY": "integration-local-user",
        "PCS_PROPOSAL_TEMP_DIR": str(workspace / "1. Open Proposals"),
        "PCS_CONTRACTS_DIR": str(workspace / "2. Signed Contracts"),
        "PCS_COMPLETED_DIR": str(workspace / "3. Finished Jobs"),
        "PCS_DEADFILE_DIR": str(workspace / "4. Dead Proposals"),
        "PCS_TEMPLATE_DIR": str(workspace / "Job Jacket Template"),
        "PCS_PROPOSALS_DIR": str(accounts / "PCS" / "1 - Open Proposals"),
        "PCS_DAVIDS_PROPOSALS_DIR": str(accounts / "David" / "1 - Open Proposals"),
        "PCS_LYDIAS_PROPOSALS_DIR": str(accounts / "Lydia" / "1 - Open Proposals"),
        "PCS_RANDYS_PROPOSALS_DIR": str(accounts / "Randy" / "1 - Open Proposals"),
        "PCS_EMAIL_TEMPLATE_DIR": str(workspace / "Marketing" / "Email Templates"),
        "PCS_EMAIL_LIST_DIR": str(workspace / "Marketing" / "Email Lists"),
        "PCS_PROPOSAL_TRACKER": str(workspace / "Proposal Tracking Integration.xlsx"),
        "PCS_PROPOSALS_WEB_URL": "",
        "DAVIDS_PROPOSALS_WEB_URL": "",
        "LYDIAS_PROPOSALS_WEB_URL": "",
        "RANDYS_PROPOSALS_WEB_URL": "",
    }
    apply_proposal_storage_mode(values, "supabase")
    return values


def apply_integration_environment(
    environ: MutableMapping[str, str] | None = None,
) -> dict[str, str]:
    environ = environ if environ is not None else os.environ
    pilotpoint_override = environ.get(
        "PCS_INTEGRATION_ROOF_INTELLIGENCE_PROJECT_DIR", ""
    ).strip()
    values = integration_environment(
        pilotpoint_project_dir=(
            Path(pilotpoint_override) if pilotpoint_override else None
        )
    )
    environ.update(values)

    integration_url = environ.get("PCS_INTEGRATION_SUPABASE_URL", "").strip()
    integration_key = environ.get(
        "PCS_INTEGRATION_SUPABASE_PUBLISHABLE_KEY", ""
    ).strip()
    if integration_url and integration_key:
        environ["PCS_SUPABASE_URL"] = integration_url
        environ["PCS_SUPABASE_PUBLISHABLE_KEY"] = integration_key
    else:
        # The integration build intentionally targets the existing local beta
        # Supabase stack unless explicit integration credentials are supplied.
        environ["PCS_SUPABASE_URL"] = "http://127.0.0.1:54321"
        environ.setdefault(
            "PCS_SUPABASE_PUBLISHABLE_KEY",
            "sb_publishable_ACJWlzQHlZjBrEguHvfOxg_3BJgxAaH",
        )
    # A desktop build must never inherit a privileged server credential.
    environ.pop("PCS_SUPABASE_SERVICE_ROLE_KEY", None)
    return values
