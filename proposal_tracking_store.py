"""Server-side Supabase access for PCS proposal tracking."""

from __future__ import annotations

import datetime
import re

from contact_store import ContactStore, ContactStoreError
from pcs_local_settings import supabase_configuration
from tenant_context import current_tenant_context


_EMAIL_PATTERN = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
_PROPOSAL_STATUSES = {"draft", "sent", "under_contract", "dead"}
_LEGACY_STATUS_MAP = {
    "follow_up": "sent",
    "won": "under_contract",
    "lost": "dead",
    "withdrawn": "dead",
    "archived": "dead",
}


class ProposalTrackingStore(ContactStore):
    PROPOSAL_SELECT = (
        "proposal_id,lead_source,submitted_by,estimated_by,estimate_completed_date,"
        "proposal_sent_date,follow_up_date,response_notes,status,created_at,updated_at,"
        "proposal:proposal(id,customer_name,project_street_address,project_address_line_2,"
        "project_city,project_state,project_zip_code,display_name,proposal_folder_name,"
        "created_at,updated_at,proposal_contact(proposal_id,organization_contact_id,"
        "contact_role,is_primary,organization_contact:organization_contact("
        "id,business_email,is_current,contact:contact("
        "id,full_name,first_name,last_name))))"
    )

    @classmethod
    def from_local_settings(cls) -> "ProposalTrackingStore":
        project_url, api_key = supabase_configuration()
        context = current_tenant_context()
        return cls(project_url, api_key, context.access_token, context.tenant_id)

    def test_connection(self) -> None:
        self._request("proposal_tracking", params={"select": "proposal_id", "limit": "1"})

    @staticmethod
    def _iso_date(value) -> str | None:
        if isinstance(value, datetime.datetime):
            return value.date().isoformat()
        if isinstance(value, datetime.date):
            return value.isoformat()
        text = str(value or "").strip()
        if not text or text == "-":
            return None
        for format_string in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.datetime.strptime(text, format_string).date().isoformat()
            except ValueError:
                pass
        raise ProposalTrackingStoreError(f"Enter a valid date for {text!r}.")

    @staticmethod
    def _extract_emails(value) -> list[str]:
        emails = []
        for match in _EMAIL_PATTERN.findall(str(value or "")):
            email = match.casefold()
            if email not in emails:
                emails.append(email)
        return emails

    @classmethod
    def _display_date(cls, value) -> str:
        iso_value = cls._iso_date(value)
        if not iso_value:
            return ""
        parsed = datetime.date.fromisoformat(iso_value)
        return f"{parsed.month}/{parsed.day}/{parsed.year}"

    @staticmethod
    def _split_name(value) -> tuple[str, str]:
        parts = " ".join(str(value or "").split()).split(" ")
        parts = [part for part in parts if part and part != "-"]
        if not parts:
            return "Unknown", ""
        return parts[0], " ".join(parts[1:])

    @staticmethod
    def _contact_values(row: dict) -> list[dict]:
        values = []
        for link in row.get("proposal_contact") or []:
            relationship = link.get("organization_contact") or {}
            person = relationship.get("contact") or {}
            if not relationship:
                continue
            values.append({
                "name": str(person.get("full_name") or "").strip(),
                "email": str(relationship.get("business_email") or "").strip(),
                "is_primary": bool(link.get("is_primary")),
                "is_current": bool(relationship.get("is_current")),
                "organization_contact_id": relationship.get("id"),
            })
        return sorted(
            values,
            key=lambda item: (
                not item["is_primary"],
                not item["is_current"],
                item["name"].casefold(),
            ),
        )

    @classmethod
    def _screen_entry(cls, row: dict) -> dict:
        proposal = row.get("proposal") or {}
        contacts = cls._contact_values(proposal)
        return {
            "row_number": row["proposal_id"],
            "customer": str(proposal.get("display_name") or "").strip(),
            "customer_name": str(proposal.get("customer_name") or "").strip(),
            "project_street_address": str(
                proposal.get("project_street_address") or ""
            ).strip(),
            "contact": ", ".join(item["name"] for item in contacts if item["name"]),
            "email_address": "; ".join(item["email"] for item in contacts if item["email"]),
            "lead_source": str(row.get("lead_source") or "").strip(),
            "submitted_by": str(row.get("submitted_by") or "").strip(),
            "estimated_by": str(row.get("estimated_by") or "").strip(),
            "estimate_date_input": cls._display_date(row.get("estimate_completed_date")),
            "proposal_date_input": cls._display_date(row.get("proposal_sent_date")),
            "follow_up_date_input": cls._display_date(row.get("follow_up_date")),
            "response_notes": str(row.get("response_notes") or "").strip(),
            "status": str(row.get("status") or "draft").strip(),
            "proposal_date": row.get("proposal_sent_date"),
        }

    def list_proposals(self) -> list[dict]:
        rows = self._request(
            "proposal_tracking",
            params={
                "select": self.PROPOSAL_SELECT,
                "limit": "5000",
            },
        )
        return sorted(
            rows,
            key=lambda row: (
                str((row.get("proposal") or {}).get("display_name") or "").casefold(),
                str(row.get("proposal_sent_date") or ""),
                str(row.get("created_at") or ""),
            ),
        )

    def list_missing_entries(self) -> list[dict]:
        entries = []
        for row in self.list_proposals():
            entry = self._screen_entry(row)
            if any((
                not entry["contact"],
                not entry["email_address"],
                not entry["proposal_date_input"],
                not entry["estimate_date_input"],
            )):
                entries.append(entry)
        return entries

    def list_weekly_follow_ups(self, cutoff_date: datetime.date) -> list[dict]:
        rows = self._request(
            "proposal_tracking",
            params={
                "select": self.PROPOSAL_SELECT,
                "proposal_sent_date": f"lte.{cutoff_date.isoformat()}",
                "follow_up_date": "is.null",
                "status": "eq.sent",
                "order": "submitted_by.asc,proposal_sent_date.asc",
                "limit": "5000",
            },
        )
        entries = []
        for row in rows:
            entry = self._screen_entry(row)
            proposal_date = self._iso_date(entry["proposal_date"])
            if not proposal_date:
                continue
            entry["proposal_date"] = datetime.date.fromisoformat(proposal_date)
            entry["proposal_date_display"] = entry["proposal_date"].strftime("%-m/%-d/%Y")
            entries.append(entry)
        return entries

    def update_entries(self, entries: list[dict]) -> int:
        updated = 0
        for entry in entries:
            if entry.get("is_new"):
                proposal_id = self._create_proposal(entry)
            else:
                proposal_id = self._resolve_proposal_id(entry.get("row_number"))
                if not proposal_id:
                    continue
                self._request(
                    "proposal_tracking",
                    method="PATCH",
                    params={"proposal_id": f"eq.{proposal_id}"},
                    payload=self._editable_payload(entry),
                )
            self._link_contacts(
                proposal_id,
                entry.get("contact"),
                entry.get("email_address"),
            )
            updated += 1
        return updated

    def _resolve_proposal_id(self, value) -> str:
        identifier = str(value or "").strip()
        if not identifier:
            return ""
        if identifier.isdigit():
            rows = self._request(
                "proposal_tracking",
                params={
                    "select": "proposal_id",
                    "source_name": "eq.Proposal Tracking.xlsx",
                    "source_row_number": f"eq.{identifier}",
                    "limit": "1",
                },
            )
            return rows[0]["proposal_id"] if rows else ""
        return identifier

    def _create_proposal(self, entry: dict) -> str:
        customer_name = " ".join(str(entry.get("customer_name") or "").split())
        project_street_address = " ".join(
            str(entry.get("project_street_address") or "").split()
        )
        if not customer_name:
            combined = " ".join(str(entry.get("customer") or "").split())
            if " - " in combined:
                customer_name, project_street_address = combined.rsplit(" - ", 1)
            else:
                customer_name = combined
        if not customer_name:
            raise ProposalTrackingStoreError("Customer name is required.")
        proposal_payload = {
            "customer_name": customer_name,
            "project_street_address": project_street_address or None,
            "proposal_folder_name": (
                f"{customer_name} - {project_street_address}"
                if project_street_address else customer_name
            ),
        }
        rows = self._request(
            "proposal", method="POST", payload=proposal_payload, return_rows=True
        )
        proposal_id = rows[0]["id"]
        tracking_payload = {
            "proposal_id": proposal_id,
            "source_name": "PCS application",
            **self._editable_payload(entry, infer_status=True),
        }
        self._request(
            "proposal_tracking",
            method="POST",
            payload=tracking_payload,
        )
        return proposal_id

    def _editable_payload(self, entry: dict, *, infer_status: bool = False) -> dict:
        proposal_date = self._iso_date(entry.get("proposal_date"))
        follow_up_date = self._iso_date(entry.get("follow_up_date"))
        estimate_date = self._iso_date(entry.get("estimate_date"))
        payload = {
            "lead_source": str(entry.get("lead_source") or "").strip() or None,
            "submitted_by": str(entry.get("submitted_by") or "").strip() or None,
            "estimated_by": str(entry.get("estimated_by") or "").strip() or None,
            "estimate_completed_date": estimate_date,
            "proposal_sent_date": proposal_date,
            "follow_up_date": follow_up_date,
        }
        requested_status = str(entry.get("status") or "").strip().casefold()
        requested_status = _LEGACY_STATUS_MAP.get(requested_status, requested_status)
        if requested_status:
            if requested_status not in _PROPOSAL_STATUSES:
                raise ProposalTrackingStoreError(
                    "Proposal status must be draft, sent, under contract, or dead."
                )
            payload["status"] = requested_status
        elif infer_status:
            payload["status"] = "sent" if proposal_date else "draft"
        return payload

    def _find_or_create_organization_contact(self, name: str, email: str) -> str:
        rows = self._request(
            "organization_contact",
            params={
                "select": "id,is_current,updated_at",
                "normalized_email": f"eq.{email.casefold()}",
                "order": "is_current.desc,updated_at.desc",
                "limit": "2",
            },
        )
        if rows:
            return rows[0]["id"]

        organization_id = self.resolve_organization_from_email(email)
        first_name, last_name = self._split_name(name)
        contact_id = self.create_contact({
            "first_name": first_name,
            "last_name": last_name,
            "organization_id": organization_id,
            "business_email": email,
        })
        relationships = self._request(
            "organization_contact",
            params={
                "select": "id",
                "contact_id": f"eq.{contact_id}",
                "is_current": "eq.true",
                "limit": "1",
            },
        )
        if not relationships:
            raise ProposalTrackingStoreError("The proposal contact relationship was not created.")
        return relationships[0]["id"]

    def _link_contacts(self, proposal_id: str, contact_name, email_value) -> None:
        emails = self._extract_emails(email_value)
        if not emails:
            return
        names = [part.strip() for part in re.split(r"\s*(?:,|&|;)\s*", str(contact_name or "")) if part.strip()]
        relationship_ids = [
            self._find_or_create_organization_contact(
                names[index] if index < len(names) else names[0] if names else "Unknown",
                email,
            )
            for index, email in enumerate(emails)
        ]
        self._request(
            "proposal_contact",
            method="PATCH",
            params={"proposal_id": f"eq.{proposal_id}", "is_primary": "eq.true"},
            payload={"is_primary": False, "contact_role": "additional"},
        )
        for index, relationship_id in enumerate(relationship_ids):
            payload = {
                "proposal_id": proposal_id,
                "organization_contact_id": relationship_id,
                "contact_role": "primary" if index == 0 else "additional",
                "is_primary": index == 0,
            }
            existing = self._request(
                "proposal_contact",
                params={
                    "select": "proposal_id",
                    "proposal_id": f"eq.{proposal_id}",
                    "organization_contact_id": f"eq.{relationship_id}",
                    "limit": "1",
                },
            )
            if existing:
                self._request(
                    "proposal_contact",
                    method="PATCH",
                    params={
                        "proposal_id": f"eq.{proposal_id}",
                        "organization_contact_id": f"eq.{relationship_id}",
                    },
                    payload={
                        "contact_role": payload["contact_role"],
                        "is_primary": payload["is_primary"],
                    },
                )
            else:
                self._request("proposal_contact", method="POST", payload=payload)

    def mark_follow_ups(self, proposal_ids, follow_up_date: datetime.date) -> int:
        ids = [
            self._resolve_proposal_id(value)
            for value in proposal_ids
            if str(value).strip()
        ]
        ids = [value for value in ids if value]
        if not ids:
            return 0
        rows = self._request(
            "proposal_tracking",
            method="PATCH",
            params={"proposal_id": f"in.({','.join(ids)})"},
            payload={"follow_up_date": follow_up_date.isoformat(), "status": "sent"},
            return_rows=True,
        )
        return len(rows)

    def upsert_from_proposal_save(
        self,
        *,
        created_date,
        customer_name,
        street_address,
        city,
        state,
        zip_code,
        submitted_by,
        folder_name,
        lead_value="",
        estimated_by="Vern",
    ) -> str:
        folder_name = str(folder_name or "").strip()
        rows = self._request(
            "proposal",
            params={
                "select": "id",
                "proposal_folder_name": f"eq.{folder_name}",
                "order": "updated_at.desc",
                "limit": "1",
            },
        )
        if not rows:
            rows = self._request(
                "proposal",
                params={
                    "select": "id",
                    "display_name": f"eq.{folder_name}",
                    "order": "updated_at.desc",
                    "limit": "1",
                },
            )
        proposal_payload = {
            "customer_name": " ".join(str(customer_name or "").split()),
            "project_street_address": " ".join(str(street_address or "").split()) or None,
            "project_city": " ".join(str(city or "").split()) or None,
            "project_state": str(state or "").strip().upper() or None,
            "project_zip_code": str(zip_code or "").strip() or None,
            "proposal_folder_name": folder_name or None,
        }
        tracking_payload = {
            "lead_source": str(lead_value or "").strip() or None,
            "submitted_by": str(submitted_by or "").strip() or None,
            "estimated_by": str(estimated_by or "").strip() or None,
        }
        if not rows:
            created = self._request(
                "proposal", method="POST", payload=proposal_payload, return_rows=True
            )
            proposal_id = created[0]["id"]
            tracking_payload.update({
                "proposal_id": proposal_id,
                "estimate_completed_date": self._iso_date(created_date),
                "status": "draft",
                "source_name": "PCS application",
            })
            self._request(
                "proposal_tracking", method="POST", payload=tracking_payload
            )
            return proposal_id
        proposal_id = rows[0]["id"]
        self._request(
            "proposal",
            method="PATCH",
            params={"id": f"eq.{proposal_id}"},
            payload=proposal_payload,
        )
        tracking_rows = self._request(
            "proposal_tracking",
            params={"select": "proposal_id", "proposal_id": f"eq.{proposal_id}", "limit": "1"},
        )
        if tracking_rows:
            self._request(
                "proposal_tracking",
                method="PATCH",
                params={"proposal_id": f"eq.{proposal_id}"},
                payload=tracking_payload,
            )
        else:
            tracking_payload.update({
                "proposal_id": proposal_id,
                "estimate_completed_date": self._iso_date(created_date),
                "status": "draft",
                "source_name": "PCS application",
            })
            self._request(
                "proposal_tracking", method="POST", payload=tracking_payload
            )
        return proposal_id


class ProposalTrackingStoreError(ContactStoreError):
    pass


def get_proposal_tracking_store() -> ProposalTrackingStore:
    return ProposalTrackingStore.from_local_settings()
