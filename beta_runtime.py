"""Isolated runtime configuration for the PCS beta desktop application."""

from __future__ import annotations

import os
from pathlib import Path
from typing import MutableMapping


BETA_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/"
    "PilotPoint IQ Roof Intelligence Report Beta"
)


def beta_environment(
    *,
    home: Path | None = None,
    pilotpoint_project_dir: Path | None = None,
) -> dict[str, str]:
    home = home or Path.home()
    data_dir = home / "Library" / "Application Support" / "PCS Proposal Management Beta"
    workspace = data_dir / "Workspace"
    accounts = workspace / "Accounts"
    pilotpoint_dir = pilotpoint_project_dir or BETA_PROJECT_DIR
    return {
        "PCS_APP_ENV": "beta",
        "PCS_APP_DISPLAY_NAME": "PCS Proposal Beta",
        "PCS_APP_STATE_DIR": "/tmp/pcs_proposal_beta_app",
        "PCS_DEFAULT_PORT": "5051",
        "PCS_SERVER_STARTUP_TIMEOUT": "90",
        "PCS_DATA_DIR": str(data_dir),
        "PCS_SETTINGS_PATH": str(data_dir / "settings.json"),
        "PCS_XLWINGS_LOG_PATH": str(data_dir / "pcs_xlwings.log"),
        "ROOF_INTELLIGENCE_PROJECT_DIR": str(pilotpoint_dir),
        "ROOF_INTELLIGENCE_USER_KEY": "beta-local-user",
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
        "PCS_PROPOSAL_TRACKER": str(workspace / "Proposal Tracking Beta.xlsx"),
        "PCS_PROPOSALS_WEB_URL": "",
        "DAVIDS_PROPOSALS_WEB_URL": "",
        "LYDIAS_PROPOSALS_WEB_URL": "",
        "RANDYS_PROPOSALS_WEB_URL": "",
        # Beta is fully cut over to tenant-scoped Supabase proposal tracking.
        # The beta workbook remains an inert rollback artifact and is neither
        # read nor written by the application.
        "PROPOSAL_TRACKING_SUPABASE_ENABLED": "1",
        "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED": "1",
        "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED": "1",
        "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED": "0",
    }


def apply_beta_environment(environ: MutableMapping[str, str] | None = None) -> dict[str, str]:
    environ = environ if environ is not None else os.environ
    pilotpoint_override = environ.get("PCS_BETA_ROOF_INTELLIGENCE_PROJECT_DIR", "").strip()
    values = beta_environment(
        pilotpoint_project_dir=Path(pilotpoint_override) if pilotpoint_override else None
    )
    for key, value in values.items():
        environ[key] = value

    beta_url = environ.get("PCS_BETA_SUPABASE_URL", "").strip()
    beta_key = environ.get("PCS_BETA_SUPABASE_PUBLISHABLE_KEY", "").strip()
    if beta_url and beta_key:
        environ["PCS_SUPABASE_URL"] = beta_url
        environ["PCS_SUPABASE_PUBLISHABLE_KEY"] = beta_key
    else:
        environ.pop("PCS_SUPABASE_URL", None)
        environ.pop("PCS_SUPABASE_PUBLISHABLE_KEY", None)
    # A packaged beta client must never inherit a worker/server credential.
    environ.pop("PCS_SUPABASE_SERVICE_ROLE_KEY", None)
    return values
