"""Server-side Supabase access for PCS proposal tracking."""

from __future__ import annotations

import datetime
import re
import uuid

from contact_store import ContactStore, ContactStoreError
from pcs_local_settings import supabase_configuration
from tenant_context import current_tenant_context


_EMAIL_PATTERN = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
_PROPOSAL_STATUSES = {"draft", "sent", "under_contract", "finished", "dead"}
_LEGACY_STATUS_MAP = {
    "follow_up": "sent",
    "won": "under_contract",
    "lost": "dead",
    "withdrawn": "dead",
    "archived": "dead",
    "under contract": "under_contract",
    "under-contract": "under_contract",
}


class ProposalTrackingStore(ContactStore):
    PROPOSAL_SELECT = (
        "proposal_id,lead_source,submitted_by,estimated_by,estimate_completed_date,"
        "proposal_sent_date,follow_up_date,follow_up_required,response_notes,status,created_at,updated_at,"
        "proposal:proposal(id,customer_name,project_street_address,project_address_line_2,"
        "project_city,project_state,project_zip_code,display_name,proposal_folder_name,"
        "created_at,updated_at,proposal_contact(proposal_id,organization_contact_id,"
        "contact_role,is_primary,organization_contact:organization_contact("
        "id,business_email,is_current,contact:contact("
        "id,full_name,first_name,last_name))))"
    )
    MANAGEMENT_SELECT = (
        "id,customer_name,project_street_address,project_address_line_2,"
        "project_city,project_state,project_zip_code,display_name,proposal_folder_name,draft_detail,"
        "created_at,updated_at,proposal_tracking!inner(status,submitted_by,"
        "estimated_by,lead_source,response_notes,estimate_completed_date,"
        "proposal_sent_date,follow_up_date,follow_up_required,created_at,updated_at),"
        "proposal_contact(organization_contact_id,contact_role,is_primary,"
        "organization_contact:organization_contact(id,business_email,is_current,"
        "contact:contact(id,full_name),organization:organization(id,name)))"
    )
    CONTACT_OPTION_SELECT = (
        "id,business_email,is_current,contact:contact(id,full_name),"
        "organization:organization(id,name)"
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
            "follow_up_required": bool(row.get("follow_up_required", True)),
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

    @staticmethod
    def _management_timestamp(value) -> datetime.datetime | None:
        text = str(value or "").strip()
        if not text:
            return None
        try:
            parsed = datetime.datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            return None
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone().replace(tzinfo=None)
        return parsed

    @classmethod
    def _management_entry(cls, row: dict) -> dict:
        tracking = row.get("proposal_tracking") or {}
        if isinstance(tracking, list):
            tracking = tracking[0] if tracking else {}
        customer_name = " ".join(str(row.get("customer_name") or "").split())
        street_address = " ".join(
            str(row.get("project_street_address") or "").split()
        )
        display_name = " ".join(str(row.get("display_name") or "").split())
        if not display_name:
            display_name = (
                f"{customer_name} - {street_address}" if street_address else customer_name
            )
        folder_name = " ".join(
            str(row.get("proposal_folder_name") or display_name).split()
        )
        timestamps = [
            cls._management_timestamp(row.get("updated_at") or row.get("created_at")),
            cls._management_timestamp(
                tracking.get("updated_at") or tracking.get("created_at")
            ),
        ]
        last_modified = max((value for value in timestamps if value), default=None)
        contact_links = sorted(
            row.get("proposal_contact") or [],
            key=lambda link: (
                not bool(link.get("is_primary")),
                str(link.get("contact_role") or "") != "primary",
            ),
        )
        primary_relationship = (
            (contact_links[0].get("organization_contact") or {})
            if contact_links else {}
        )
        primary_person = primary_relationship.get("contact") or {}
        primary_organization = primary_relationship.get("organization") or {}
        return {
            "id": str(row.get("id") or ""),
            "name": display_name,
            "folder_name": folder_name,
            "customer_name": customer_name,
            "project_street_address": street_address,
            "project_address_line_2": str(
                row.get("project_address_line_2") or ""
            ).strip(),
            "project_city": str(row.get("project_city") or "").strip(),
            "project_state": str(row.get("project_state") or "").strip(),
            "project_zip_code": str(row.get("project_zip_code") or "").strip(),
            "draft_detail": (
                row.get("draft_detail")
                if isinstance(row.get("draft_detail"), dict)
                else {}
            ),
            "status": str(tracking.get("status") or "draft").strip(),
            "submitted_by": str(tracking.get("submitted_by") or "").strip(),
            "estimated_by": str(tracking.get("estimated_by") or "").strip(),
            "lead_source": str(tracking.get("lead_source") or "").strip(),
            "response_notes": str(tracking.get("response_notes") or "").strip(),
            "estimate_completed_date": cls._iso_date(
                tracking.get("estimate_completed_date")
            ),
            "estimate_completed_date_display": cls._display_date(
                tracking.get("estimate_completed_date")
            ),
            "proposal_sent_date": cls._iso_date(tracking.get("proposal_sent_date")),
            "proposal_sent_date_display": cls._display_date(
                tracking.get("proposal_sent_date")
            ),
            "follow_up_date": cls._iso_date(tracking.get("follow_up_date")),
            "follow_up_date_display": cls._display_date(
                tracking.get("follow_up_date")
            ),
            "follow_up_required": bool(tracking.get("follow_up_required", True)),
            "organization_contact_id": str(
                primary_relationship.get("id") or ""
            ),
            "contact_id": str(primary_person.get("id") or ""),
            "contact_name": str(primary_person.get("full_name") or "").strip(),
            "contact_email": str(
                primary_relationship.get("business_email") or ""
            ).strip(),
            "organization_name": str(
                primary_organization.get("name") or ""
            ).strip(),
            "last_modified": last_modified,
            "last_modified_display": (
                last_modified.strftime("%m/%d/%Y") if last_modified else "Unavailable"
            ),
        }

    def list_management_proposals(self, statuses) -> list[dict]:
        requested_statuses = {
            str(status or "").strip().lower() for status in statuses or ()
        }
        requested_statuses.discard("")
        unsupported = requested_statuses - _PROPOSAL_STATUSES
        if unsupported:
            raise ProposalTrackingStoreError(
                f"Unsupported proposal status: {sorted(unsupported)[0]}"
            )
        if not requested_statuses:
            return []
        rows = self._request(
            "proposal",
            params={
                "select": self.MANAGEMENT_SELECT,
                "proposal_tracking.status": (
                    f"in.({','.join(sorted(requested_statuses))})"
                ),
                "limit": "5000",
            },
        )
        entries = [self._management_entry(row) for row in rows]
        return sorted(entries, key=lambda entry: entry["name"].casefold())

    def get_management_proposal(self, proposal_id: str) -> dict | None:
        identifier = str(proposal_id or "").strip()
        if not identifier:
            return None
        rows = self._request(
            "proposal",
            params={
                "select": self.MANAGEMENT_SELECT,
                "id": f"eq.{identifier}",
                "limit": "1",
            },
        )
        return self._management_entry(rows[0]) if rows else None

    def get_management_proposal_by_folder(self, folder_name: str) -> dict | None:
        name = " ".join(str(folder_name or "").split())
        if not name:
            return None
        rows = self._request(
            "proposal",
            params={
                "select": self.MANAGEMENT_SELECT,
                "proposal_folder_name": f"eq.{name}",
                "order": "updated_at.desc",
                "limit": "1",
            },
        )
        if not rows:
            rows = self._request(
                "proposal",
                params={
                    "select": self.MANAGEMENT_SELECT,
                    "display_name": f"eq.{name}",
                    "order": "updated_at.desc",
                    "limit": "1",
                },
            )
        return self._management_entry(rows[0]) if rows else None

    @staticmethod
    def _contact_option(row: dict) -> dict:
        person = row.get("contact") or {}
        organization = row.get("organization") or {}
        return {
            "id": str(row.get("id") or ""),
            "name": str(person.get("full_name") or "").strip(),
            "email": str(row.get("business_email") or "").strip(),
            "organization": str(organization.get("name") or "").strip(),
        }

    def list_management_contact_options(self) -> list[dict]:
        rows = []
        last_id = ""
        while True:
            params = {
                "select": self.CONTACT_OPTION_SELECT,
                "is_current": "eq.true",
                "order": "id.asc",
                "limit": "1000",
            }
            if last_id:
                params["id"] = f"gt.{last_id}"
            page = self._request("organization_contact", params=params)
            rows.extend(page)
            if len(page) < 1000:
                break
            last_id = str(page[-1].get("id") or "")
            if not last_id:
                break
        options = [
            self._contact_option(row) for row in rows
            if (row.get("contact") or {}).get("full_name") or row.get("business_email")
        ]
        return sorted(
            options,
            key=lambda option: (
                option["name"].casefold(),
                option["email"].casefold(),
            ),
        )

    def _get_organization_contact_option(self, relationship_id: str) -> dict:
        rows = self._request(
            "organization_contact",
            params={
                "select": self.CONTACT_OPTION_SELECT,
                "id": f"eq.{relationship_id}",
                "is_current": "eq.true",
                "limit": "1",
            },
        )
        if not rows:
            raise ProposalTrackingStoreError("That contact is no longer available.")
        return self._contact_option(rows[0])

    def _set_primary_contact(self, proposal_id: str, relationship_id: str) -> None:
        proposals = self._request(
            "proposal",
            params={"select": "id", "id": f"eq.{proposal_id}", "limit": "1"},
        )
        if not proposals:
            raise ProposalTrackingStoreError("That proposal is no longer available.")
        existing_links = self._request(
            "proposal_contact",
            params={
                "select": "organization_contact_id,is_primary",
                "proposal_id": f"eq.{proposal_id}",
                "limit": "5000",
            },
        )
        selected_link = next(
            (
                link for link in existing_links
                if str(link.get("organization_contact_id")) == relationship_id
            ),
            None,
        )
        if selected_link and selected_link.get("is_primary"):
            return
        previous_primary_ids = [
            str(link.get("organization_contact_id"))
            for link in existing_links if link.get("is_primary")
        ]
        try:
            if previous_primary_ids:
                self._request(
                    "proposal_contact",
                    method="PATCH",
                    params={
                        "proposal_id": f"eq.{proposal_id}",
                        "is_primary": "eq.true",
                    },
                    payload={"is_primary": False, "contact_role": "additional"},
                )
            if selected_link:
                self._request(
                    "proposal_contact",
                    method="PATCH",
                    params={
                        "proposal_id": f"eq.{proposal_id}",
                        "organization_contact_id": f"eq.{relationship_id}",
                    },
                    payload={"is_primary": True, "contact_role": "primary"},
                )
            else:
                self._request(
                    "proposal_contact",
                    method="POST",
                    payload={
                        "proposal_id": proposal_id,
                        "organization_contact_id": relationship_id,
                        "is_primary": True,
                        "contact_role": "primary",
                    },
                )
        except Exception:
            for previous_id in previous_primary_ids:
                try:
                    self._request(
                        "proposal_contact",
                        method="PATCH",
                        params={
                            "proposal_id": f"eq.{proposal_id}",
                            "organization_contact_id": f"eq.{previous_id}",
                        },
                        payload={"is_primary": True, "contact_role": "primary"},
                    )
                except Exception:
                    pass
            raise

    def assign_or_create_primary_contact(
        self,
        proposal_id: str,
        *,
        organization_contact_id: str = "",
        contact_name: str = "",
        email: str = "",
        organization_name: str = "",
    ) -> dict:
        relationship_id = str(organization_contact_id or "").strip()
        created = False
        if relationship_id:
            option = self._get_organization_contact_option(relationship_id)
        else:
            clean_name = " ".join(str(contact_name or "").split())
            clean_email = str(email or "").strip().casefold()
            if not clean_name:
                raise ProposalTrackingStoreError("Enter the contact name.")
            if not clean_email or _EMAIL_PATTERN.fullmatch(clean_email) is None:
                raise ProposalTrackingStoreError("Enter a valid email address.")
            matches = self._request(
                "organization_contact",
                params={
                    "select": self.CONTACT_OPTION_SELECT,
                    "normalized_email": f"eq.{clean_email}",
                    "is_current": "eq.true",
                    "order": "updated_at.desc",
                    "limit": "2",
                },
            )
            if matches:
                option = self._contact_option(matches[0])
                relationship_id = option["id"]
            else:
                organization = self.find_organization_for_email(clean_email)
                if organization:
                    organization_id = organization["id"]
                elif not str(organization_name or "").strip():
                    raise ProposalContactOrganizationRequired(
                        self._email_domain(clean_email)
                    )
                else:
                    organization_id = self.resolve_named_organization_for_email(
                        organization_name,
                        clean_email,
                    )
                first_name, last_name = self._split_name(clean_name)
                contact_id = self.create_contact({
                    "first_name": first_name,
                    "last_name": last_name,
                    "organization_id": organization_id,
                    "business_email": clean_email,
                })
                relationships = self._request(
                    "organization_contact",
                    params={
                        "select": self.CONTACT_OPTION_SELECT,
                        "contact_id": f"eq.{contact_id}",
                        "is_current": "eq.true",
                        "limit": "1",
                    },
                )
                if not relationships:
                    raise ProposalTrackingStoreError(
                        "The new contact relationship was not created."
                    )
                option = self._contact_option(relationships[0])
                relationship_id = option["id"]
                created = True
        self._set_primary_contact(str(proposal_id), relationship_id)
        return {**option, "created": created}

    def update_proposal_customer_name(
        self,
        proposal_id: str,
        customer_name: str,
    ) -> None:
        identifier = str(proposal_id or "").strip()
        clean_name = " ".join(str(customer_name or "").split())
        if not identifier:
            raise ProposalTrackingStoreError("That proposal could not be selected.")
        if not clean_name:
            return
        rows = self._request(
            "proposal",
            method="PATCH",
            params={"id": f"eq.{identifier}"},
            payload={"customer_name": clean_name},
            return_rows=True,
        )
        if not rows:
            raise ProposalTrackingStoreError("That proposal is no longer available.")

    def list_missing_entries(self) -> list[dict]:
        entries = []
        for row in self.list_proposals():
            entry = self._screen_entry(row)
            if entry["status"] == "dead":
                continue
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
                "follow_up_required": "eq.true",
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
        try:
            return str(uuid.UUID(identifier))
        except (ValueError, AttributeError):
            return ""

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

    def save_proposal_draft_detail(
        self, proposal_id: str, draft_detail: dict
    ) -> None:
        """Persist the unsaved proposal-detail form during a contact detour."""
        identifier = str(proposal_id or "").strip()
        if not identifier:
            raise ProposalTrackingStoreError("A proposal draft is required.")
        if not isinstance(draft_detail, dict):
            raise ProposalTrackingStoreError("Proposal draft detail must be an object.")
        rows = self._request(
            "proposal",
            method="PATCH",
            params={"id": f"eq.{identifier}"},
            payload={"draft_detail": draft_detail},
            return_rows=True,
        )
        if not rows:
            raise ProposalTrackingStoreError(
                "That proposal draft is no longer available."
            )

    def clear_proposal_draft_detail(self, proposal_id: str) -> None:
        self.save_proposal_draft_detail(proposal_id, {})

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
                    "Proposal status must be draft, sent, under contract, finished, or dead."
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
        proposal_id="",
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
        requested_proposal_id = str(proposal_id or "").strip()
        if requested_proposal_id:
            rows = self._request(
                "proposal",
                params={
                    "select": "id",
                    "id": f"eq.{requested_proposal_id}",
                    "limit": "1",
                },
            )
            if not rows:
                raise ProposalTrackingStoreError(
                    "That proposal draft is no longer available."
                )
        else:
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
        if requested_proposal_id and created_date:
            tracking_payload["estimate_completed_date"] = self._iso_date(
                created_date
            )
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
            try:
                self._request(
                    "proposal_tracking", method="POST", payload=tracking_payload
                )
            except Exception:
                try:
                    self._request(
                        "proposal",
                        method="DELETE",
                        params={"id": f"eq.{proposal_id}"},
                    )
                except Exception:
                    pass
                raise
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


class ProposalContactOrganizationRequired(ContactStoreError):
    def __init__(self, domain: str):
        self.domain = str(domain or "").strip().casefold()
        super().__init__(
            f"Enter the organization name for the {self.domain or 'email'} contact."
        )


class ProposalTrackingStoreError(ContactStoreError):
    pass


def get_proposal_tracking_store() -> ProposalTrackingStore:
    return ProposalTrackingStore.from_local_settings()
