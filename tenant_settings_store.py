"""Tenant-owned settings and logical report-folder management."""

from __future__ import annotations

from contact_store import ContactStore, ContactStoreError
from pcs_local_settings import supabase_configuration
from tenant_context import current_tenant_context


class TenantSettingsStore(ContactStore):
    @classmethod
    def from_current_session(cls) -> "TenantSettingsStore":
        project_url, api_key = supabase_configuration()
        context = current_tenant_context()
        return cls(project_url, api_key, context.access_token, context.tenant_id)

    def list_report_folders(self) -> list[dict]:
        rows = self._request(
            "report_folder",
            params={"select": "id,name,is_archived", "order": "name.asc"},
        )
        return rows if isinstance(rows, list) else []

    def get_settings(self) -> dict:
        rows = self._request(
            "tenant_settings",
            params={"select": "tenant_id,default_report_folder_id,company_configuration", "limit": "1"},
        )
        return rows[0] if rows else {}

    def create_report_folder(self, name: object) -> dict:
        clean_name = " ".join(str(name or "").split())
        if not clean_name:
            raise ContactStoreError("Enter a report folder name.")
        rows = self._request(
            "report_folder",
            method="POST",
            payload={"name": clean_name},
            return_rows=True,
        )
        if not rows:
            raise ContactStoreError("The report folder was not created.")
        return rows[0]

    def set_default_report_folder(self, folder_id: object) -> None:
        clean_id = str(folder_id or "").strip()
        folders = {str(item["id"]) for item in self.list_report_folders() if not item.get("is_archived")}
        if clean_id not in folders:
            raise ContactStoreError("Choose an active report folder for this company.")
        self._request(
            "tenant_settings",
            method="PATCH",
            payload={"default_report_folder_id": clean_id},
        )

    def storage_prefix(self, report_id: str, revision_number: int) -> str:
        settings = self.get_settings()
        folder_id = settings.get("default_report_folder_id")
        if not folder_id:
            raise ContactStoreError("Choose a default report folder in Settings.")
        return (
            f"{self.tenant_id}/folders/{folder_id}/reports/{report_id}/"
            f"revisions/{int(revision_number)}"
        )
