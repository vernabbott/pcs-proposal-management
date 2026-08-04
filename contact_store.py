"""Server-side Supabase data access for PCS contact management."""

from __future__ import annotations

import json
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from pcs_local_settings import supabase_configuration


SHARED_UNKNOWN_EMAIL_DOMAINS = frozenset({
    "comcast.net",
    "gmail.com",
    "m.knck.io",
    "yahoo.com",
})


class ContactStoreError(RuntimeError):
    pass


class ContactConfigurationError(ContactStoreError):
    pass


class ContactStore:
    def __init__(self, project_url: str, service_key: str):
        if not project_url or not service_key:
            raise ContactConfigurationError(
                "Supabase is not configured. Open Local Settings and save the server key."
            )
        self.base_url = f"{project_url.rstrip('/')}/rest/v1"
        self.service_key = service_key

    @classmethod
    def from_local_settings(cls) -> "ContactStore":
        return cls(*supabase_configuration())

    def _request(self, table: str, *, method: str = "GET", params=None, payload=None, return_rows=False):
        query = f"?{urlencode(params or {}, doseq=True)}" if params else ""
        body = json.dumps(payload).encode("utf-8") if payload is not None else None
        headers = {
            "apikey": self.service_key,
            "Authorization": f"Bearer {self.service_key}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }
        if return_rows:
            headers["Prefer"] = "return=representation"
        request = Request(f"{self.base_url}/{table}{query}", data=body, headers=headers, method=method)
        try:
            with urlopen(request, timeout=30) as response:
                raw = response.read()
        except HTTPError as exc:
            try:
                detail = json.loads(exc.read().decode("utf-8")).get("message", "")
            except Exception:
                detail = ""
            raise ContactStoreError(detail or f"Supabase rejected the request ({exc.code}).") from exc
        except (URLError, TimeoutError, OSError) as exc:
            raise ContactStoreError("PCS could not connect to Supabase. Check the connection and try again.") from exc
        return json.loads(raw.decode("utf-8")) if raw else []

    def test_connection(self) -> None:
        self._request("organization", params={"select": "id", "limit": "1"})

    def list_organizations(self) -> list[dict]:
        rows = self._request(
            "organization",
            params={
                "select": (
                    "id,name,organization_type,main_office_address_line_1,"
                    "main_office_address_line_2,main_office_city,main_office_state,"
                    "main_office_zip_code"
                ),
                "is_active": "eq.true",
                "limit": "5000",
            },
        )
        return sorted(rows, key=lambda item: (item.get("name") or "").casefold())

    def list_contacts(self, *, search: str = "", status: str = "active") -> list[dict]:
        rows = self._request(
            "organization_contact",
            params={
                "select": (
                    "id,title,business_email,business_phone,mobile_phone,"
                    "branch_address_line_1,branch_address_line_2,branch_city,branch_state,"
                    "branch_zip_code,is_current,created_at,"
                    "do_not_contact,contact:contact("
                    "id,full_name,first_name,last_name,linkedin_url,notes),"
                    "organization:organization(id,name,organization_type,"
                    "main_office_address_line_1,main_office_address_line_2,"
                    "main_office_city,main_office_state,main_office_zip_code)"
                ),
                "limit": "5000",
            },
        )
        latest_by_contact: dict[str, dict] = {}
        for row in rows:
            contact = row.get("contact") or {}
            contact_id = contact.get("id")
            if not contact_id:
                continue
            existing = latest_by_contact.get(contact_id)
            if existing is None or row.get("is_current") or (
                not existing.get("is_current")
                and (row.get("created_at") or "") > (existing.get("created_at") or "")
            ):
                latest_by_contact[contact_id] = row

        result = list(latest_by_contact.values())
        if status == "active":
            result = [row for row in result if row.get("is_current")]
        elif status == "archived":
            result = [row for row in result if not row.get("is_current")]

        query = search.strip().casefold()
        if query:
            def searchable(row):
                contact = row.get("contact") or {}
                organization = row.get("organization") or {}
                values = (
                    contact.get("full_name"), organization.get("name"),
                    organization.get("organization_type"), row.get("title"),
                    row.get("business_email"), row.get("business_phone"), row.get("mobile_phone"),
                )
                return any(query in str(value or "").casefold() for value in values)
            result = [row for row in result if searchable(row)]

        return sorted(
            result,
            key=lambda row: ((row.get("contact") or {}).get("full_name") or "").casefold(),
        )

    def get_contact(self, contact_id: str) -> dict | None:
        return next((row for row in self.list_contacts(status="all") if row["contact"]["id"] == contact_id), None)

    def find_duplicate_contacts(self, values: dict) -> list[dict]:
        email = self._duplicate_value(values.get("business_email"))
        first_name = self._duplicate_value(values.get("first_name"))
        last_name = self._duplicate_value(values.get("last_name"))
        match_full_name = bool(first_name and last_name)
        matches = []
        for row in self.list_contacts(status="all"):
            person = row.get("contact") or {}
            reasons = []
            if email and self._duplicate_value(row.get("business_email")) == email:
                reasons.append("Same business email")
            if match_full_name and (
                self._duplicate_value(person.get("first_name")) == first_name
                and self._duplicate_value(person.get("last_name")) == last_name
            ):
                reasons.append("Same first and last name")
            if reasons:
                match = dict(row)
                match["duplicate_reasons"] = reasons
                matches.append(match)
        return matches

    def create_contact(self, values: dict) -> str:
        contact_payload = self._contact_payload(values)
        contacts = self._request("contact", method="POST", payload=contact_payload, return_rows=True)
        contact_id = contacts[0]["id"]
        try:
            self._request(
                "organization_contact",
                method="POST",
                payload=self._relationship_payload(contact_id, values),
            )
        except Exception:
            self._request("contact", method="DELETE", params={"id": f"eq.{contact_id}"})
            raise
        return contact_id

    def update_contact(self, contact_id: str, values: dict) -> None:
        existing = self.get_contact(contact_id)
        if not existing:
            raise ContactStoreError("That contact no longer exists.")
        self._request(
            "contact", method="PATCH", params={"id": f"eq.{contact_id}"},
            payload=self._contact_payload(values),
        )
        old_relationship_id = existing["id"]
        old_organization_id = (existing.get("organization") or {}).get("id")
        if existing.get("is_current") and old_organization_id == values["organization_id"]:
            payload = self._relationship_payload(contact_id, values)
            payload.pop("contact_id")
            payload.pop("organization_id")
            self._request(
                "organization_contact", method="PATCH", params={"id": f"eq.{old_relationship_id}"},
                payload=payload,
            )
            return

        if existing.get("is_current"):
            self._request(
                "organization_contact", method="PATCH", params={"id": f"eq.{old_relationship_id}"},
                payload={"is_current": False},
            )
        try:
            self._request(
                "organization_contact", method="POST",
                payload=self._relationship_payload(contact_id, values),
            )
        except Exception:
            if existing.get("is_current"):
                self._request(
                    "organization_contact", method="PATCH", params={"id": f"eq.{old_relationship_id}"},
                    payload={"is_current": True},
                )
            raise

    def archive_contact(self, contact_id: str) -> None:
        self._request(
            "organization_contact", method="PATCH",
            params={"contact_id": f"eq.{contact_id}", "is_current": "eq.true"},
            payload={"is_current": False},
        )

    def create_organization(self, name: str, organization_type: str, values: dict | None = None) -> str:
        payload = {
            "name": name.strip(),
            "organization_type": organization_type.strip() or "Other",
        }
        payload.update(self._main_office_payload(values or {}))
        rows = self._request(
            "organization", method="POST",
            payload=payload,
            return_rows=True,
        )
        return rows[0]["id"]

    def find_organization_by_name(self, name: str) -> dict | None:
        normalized_name = name.strip().casefold()
        if not normalized_name:
            return None
        return next(
            (
                organization
                for organization in self.list_organizations()
                if (organization.get("name") or "").strip().casefold() == normalized_name
            ),
            None,
        )

    def update_organization(
        self,
        organization_id: str,
        name: str,
        organization_type: str,
        values: dict,
    ) -> None:
        payload = {
            "name": name.strip(),
            "organization_type": organization_type.strip() or "Other",
        }
        payload.update(self._main_office_payload(values))
        self._request(
            "organization",
            method="PATCH",
            params={"id": f"eq.{organization_id}"},
            payload=payload,
        )

    def resolve_organization_from_email(self, email: str) -> str:
        domain = self._email_domain(email)
        if not domain:
            raise ContactStoreError(
                "Enter a business email address or select an organization."
            )
        if domain in SHARED_UNKNOWN_EMAIL_DOMAINS:
            return self._resolve_shared_unknown_organization()

        def find_domain_organization():
            rows = self._request(
                "organization",
                params={
                    "select": "id,normalized_name,email_domain,source_name",
                    "or": f"(normalized_name.eq.{domain},email_domain.ilike.{domain})",
                    "limit": "20",
                },
            )
            if not rows:
                return None
            for row in rows:
                if row.get("normalized_name") == domain:
                    return row
            for row in rows:
                if row.get("source_name") == "Contact email domain":
                    return row
            return rows[0] if len(rows) == 1 else None

        organization = find_domain_organization()
        if organization:
            if (organization.get("email_domain") or "").casefold() != domain:
                self._request(
                    "organization",
                    method="PATCH",
                    params={"id": f"eq.{organization['id']}"},
                    payload={"email_domain": domain},
                )
            return organization["id"]

        try:
            rows = self._request(
                "organization",
                method="POST",
                payload={
                    "name": domain,
                    "organization_type": "Unknown",
                    "email_domain": domain,
                },
                return_rows=True,
            )
            return rows[0]["id"]
        except ContactStoreError:
            # A simultaneous request may have created the same domain organization.
            organization = find_domain_organization()
            if organization:
                return organization["id"]
            raise

    def _resolve_shared_unknown_organization(self) -> str:
        def find_unknown_organization():
            rows = self._request(
                "organization",
                params={
                    "select": "id",
                    "normalized_name": "eq.unknown",
                    "limit": "1",
                },
            )
            return rows[0] if rows else None

        organization = find_unknown_organization()
        if organization:
            return organization["id"]

        try:
            rows = self._request(
                "organization",
                method="POST",
                payload={
                    "name": "Unknown",
                    "organization_type": "Unknown",
                    "source_name": "Shared unknown email domain",
                },
                return_rows=True,
            )
            return rows[0]["id"]
        except ContactStoreError:
            organization = find_unknown_organization()
            if organization:
                return organization["id"]
            raise

    @staticmethod
    def _contact_payload(values: dict) -> dict:
        first = values.get("first_name", "").strip()
        last = values.get("last_name", "").strip()
        return {
            "full_name": " ".join(part for part in (first, last) if part),
            "first_name": first or None,
            "last_name": last or None,
            "linkedin_url": values.get("linkedin_url", "").strip() or None,
            "notes": values.get("contact_notes", "").strip() or None,
        }

    @staticmethod
    def _relationship_payload(contact_id: str, values: dict) -> dict:
        return {
            "contact_id": contact_id,
            "organization_id": values["organization_id"],
            "title": values.get("title", "").strip() or None,
            "business_email": values.get("business_email", "").strip() or None,
            "business_phone": values.get("business_phone", "").strip() or None,
            "mobile_phone": values.get("mobile_phone", "").strip() or None,
            "branch_address_line_1": values.get("branch_address_line_1", "").strip() or None,
            "branch_address_line_2": values.get("branch_address_line_2", "").strip() or None,
            "branch_city": values.get("branch_city", "").strip() or None,
            "branch_state": values.get("branch_state", "").strip().upper() or None,
            "branch_zip_code": values.get("branch_zip_code", "").strip() or None,
            "do_not_contact": bool(values.get("do_not_contact")),
            "notes": values.get("relationship_notes", "").strip() or None,
        }

    @staticmethod
    def _main_office_payload(values: dict) -> dict:
        return {
            "main_office_address_line_1": values.get("main_office_address_line_1", "").strip() or None,
            "main_office_address_line_2": values.get("main_office_address_line_2", "").strip() or None,
            "main_office_city": values.get("main_office_city", "").strip() or None,
            "main_office_state": values.get("main_office_state", "").strip().upper() or None,
            "main_office_zip_code": values.get("main_office_zip_code", "").strip() or None,
        }

    @staticmethod
    def _duplicate_value(value) -> str:
        return " ".join(str(value or "").split()).casefold()

    @staticmethod
    def _email_domain(email: str) -> str:
        value = (email or "").strip()
        if "@" not in value:
            return ""
        return value.rsplit("@", 1)[1].strip().rstrip(".").casefold()


def get_contact_store() -> ContactStore:
    return ContactStore.from_local_settings()
