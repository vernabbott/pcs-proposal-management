#!/usr/bin/env python3
"""Import Colorado IREM AMO firms and independently verified contacts.

IREM's public AMO firm profiles are used for company identity and office data.
Individual IREM member-detail pages are intentionally not collected because
IREM disallows automated access to those pages in robots.txt. Named contacts
in this import come only from the companies' own public team/office pages.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
import html
import json
import os
from pathlib import Path
import re
from typing import Iterable
from urllib.request import Request, urlopen


SOURCE_NAME = "IREM AMO Directory"
IREM_PROFILE_BASE = "https://www.irem.org/about-irem/directory/amo-firm-details?id="


@dataclass(frozen=True)
class AmoOffice:
    canonical_name: str
    irem_id: str
    membership: str
    aliases: tuple[str, ...] = ()


@dataclass(frozen=True)
class VerifiedContact:
    company_name: str
    full_name: str
    title: str
    source_url: str


AMO_OFFICES = (
    AmoOffice("Camden Property Trust", "MDAwODkzNzk=", "AMO Branch"),
    AmoOffice("CBRE", "MDQ2MjEyMDI=", "AMO Branch", ("CBRE, Inc.",)),
    AmoOffice(
        "Cushman & Wakefield",
        "MDAwODk0NjA=",
        "AMO Branch",
        ("Cushman & Wakefield of Colorado, Inc.",),
    ),
    AmoOffice("Cushman & Wakefield", "MDYwMzI0OTY=", "AMO Branch", ("Cushman & Wakefield, Inc.",)),
    AmoOffice("Denver Housing Authority", "MDU5MTEwMjA=", "AMO Headquarters"),
    AmoOffice("Echelon Property Group", "MDY0MTM4NTg=", "AMO Branch", ("Echelon Property Group, LLC",)),
    AmoOffice("Greystar", "MDYyMTA2Njc=", "AMO Branch", ("Greystar Management",)),
    AmoOffice("Griffis/Blessing", "MDQyNTc0ODc=", "AMO Headquarters", ("Griffis Blessing, Inc.",)),
    AmoOffice("Mission Rock Residential", "MDYxNDU0Njk=", "AMO Headquarters", ("Mission Rock Residential LLC",)),
    AmoOffice("Mountain-n-Plains", "MDQwNjk0NzY=", "AMO Headquarters", ("Mountain-n-Plains, Inc.",)),
    AmoOffice("Physicians Realty Trust", "MDYxNzc3NTg=", "AMO Branch"),
    AmoOffice("Sentry Management", "MDYyNTM5MTU=", "AMO Branch", ("Sentry Management, Inc.",)),
    AmoOffice("Simpson Property Group", "MDAxMTUwMzc=", "AMO Branch", ("Simpson Property Group, L.P.",)),
    AmoOffice(
        "Tiarna Real Estate Services",
        "MDAyMTUyNjE=",
        "AMO Branch",
        ("Tiarna Real Estate Services, Inc. Colorado Regional Office", "Tiarna Real Estate Services, Inc."),
    ),
)


VERIFIED_CONTACTS = (
    VerifiedContact(
        "Mission Rock Residential",
        "Meredith Wright",
        "Chief Executive Officer",
        "https://www.missionrockresidential.com/our-team",
    ),
    VerifiedContact(
        "Mission Rock Residential",
        "Janelle French",
        "Executive Vice President, Operations",
        "https://www.missionrockresidential.com/our-team",
    ),
    VerifiedContact(
        "Mission Rock Residential",
        "Erin Mangers",
        "Vice President, Operations",
        "https://www.missionrockresidential.com/our-team",
    ),
    VerifiedContact(
        "Mountain-n-Plains",
        "Julia Crawmer",
        "Broker, Asset Management",
        "https://mountain-n-plains.com/contact-us/",
    ),
    VerifiedContact(
        "Mountain-n-Plains",
        "Donna Knopp",
        "Associate Broker, Commercial Property Manager",
        "https://mountain-n-plains.com/contact-us/",
    ),
    VerifiedContact(
        "Mountain-n-Plains",
        "Lacey Fleming",
        "Employing Broker, Residential Manager",
        "https://mountain-n-plains.com/contact-us/",
    ),
    VerifiedContact(
        "Sentry Management",
        "Susan Horton",
        "Division President, Denver-Boulder",
        "https://www.sentrymgt.com/offices/boulder/",
    ),
    VerifiedContact(
        "Tiarna Real Estate Services",
        "Thomas A. McAndrews",
        "Chief Executive Officer",
        "https://www.tiarna.com/team/",
    ),
    VerifiedContact(
        "Tiarna Real Estate Services",
        "Robert S. Alleborn",
        "Executive Vice President",
        "https://www.tiarna.com/team/",
    ),
    VerifiedContact(
        "Tiarna Real Estate Services",
        "Michael T. McAndrews",
        "Executive Vice President",
        "https://www.tiarna.com/team/",
    ),
)


def clean(value: object) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").split())


def normalized_name(value: str) -> str:
    return clean(value).casefold()


def split_person_name(value: str) -> tuple[str | None, str | None]:
    parts = clean(value).split()
    if not parts:
        return None, None
    if len(parts) == 1:
        return parts[0], None
    return parts[0], parts[-1]


def span_value(page: str, suffix: str) -> str | None:
    pattern = rf'<span[^>]+id="[^"]*{re.escape(suffix)}"[^>]*>(.*?)</span>'
    match = re.search(pattern, page, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        return None
    value = re.sub(r"<[^>]+>", "", match.group(1))
    value = clean(html.unescape(value))
    return value or None


def fetch_profile(office: AmoOffice) -> dict[str, object]:
    source_url = IREM_PROFILE_BASE + office.irem_id
    request = Request(source_url, headers={"User-Agent": "PCS-Property-Management-Research/1.0"})
    with urlopen(request, timeout=30) as response:
        page = response.read().decode("utf-8", errors="replace")

    city_state = span_value(page, "_CityState") or ""
    city = state = zip_code = None
    match = re.match(r"^(.*?)\s+([A-Z]{2})\s+([0-9]{5}(?:-[0-9]{4})?)$", city_state)
    if match:
        city, state, zip_code = (clean(part) or None for part in match.groups())

    website = span_value(page, "_Web")
    if website and not website.startswith(("http://", "https://")):
        website = "https://" + website

    return {
        "canonical_name": office.canonical_name,
        "profile_name": span_value(page, "_Label_Name"),
        "membership": office.membership,
        "aliases": office.aliases,
        "address_line_1": span_value(page, "_Addrs1"),
        "address_line_2": span_value(page, "_Addrs2"),
        "city": city,
        "state": state,
        "zip_code": zip_code,
        "phone": span_value(page, "_phone"),
        "website": website,
        "source_url": source_url,
    }


def merge_profiles(profiles: Iterable[dict[str, object]]) -> dict[str, dict[str, object]]:
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for profile in profiles:
        grouped[str(profile["canonical_name"])].append(profile)

    merged: dict[str, dict[str, object]] = {}
    for canonical_name, items in grouped.items():
        record: dict[str, object] = {
            "canonical_name": canonical_name,
            "aliases": [],
            "profile_names": [],
            "memberships": [],
            "source_urls": [],
        }
        for item in items:
            for field in ("address_line_1", "address_line_2", "city", "state", "zip_code", "phone", "website"):
                if not record.get(field) and item.get(field):
                    record[field] = item[field]
            record["aliases"].extend(item.get("aliases") or [])  # type: ignore[union-attr]
            if item.get("profile_name"):
                record["profile_names"].append(item["profile_name"])  # type: ignore[union-attr]
            record["memberships"].append(item["membership"])  # type: ignore[union-attr]
            record["source_urls"].append(item["source_url"])  # type: ignore[union-attr]
        merged[canonical_name] = record
    return merged


def database_url(env_file: Path) -> str:
    from dotenv import load_dotenv

    load_dotenv(env_file)
    value = os.getenv("DATABASE_URL")
    if not value:
        raise RuntimeError(f"DATABASE_URL is not configured in {env_file}")
    if value.startswith("postgres://"):
        return value.replace("postgres://", "postgresql+psycopg2://", 1)
    if value.startswith("postgresql://"):
        return value.replace("postgresql://", "postgresql+psycopg2://", 1)
    return value


def load_database(
    companies: dict[str, dict[str, object]],
    contacts: Iterable[VerifiedContact],
    env_file: Path,
    commit: bool,
) -> dict[str, int]:
    from sqlalchemy import create_engine, text

    engine = create_engine(database_url(env_file), pool_pre_ping=True, connect_args={"sslmode": "require"})
    stats = {
        "irem_offices_fetched": len(AMO_OFFICES),
        "companies_eligible": len(companies),
        "companies_inserted": 0,
        "companies_updated": 0,
        "contacts_eligible": len(tuple(contacts)),
        "contacts_inserted": 0,
        "contacts_updated": 0,
        "committed": int(commit),
    }

    company_find = text("""
        select id, name
        from public.property_management_companies
        where normalized_name = any(:candidate_names)
        order by case when normalized_name = :canonical_name then 0 else 1 end, created_at
        limit 1
    """)
    company_insert = text("""
        insert into public.property_management_companies
            (name, website, main_phone, address_line_1, address_line_2, city, state, zip_code,
             source_name, source_url, verified_at, notes)
        values
            (:name, :website, :phone, :address_line_1, :address_line_2, :city, :state, :zip_code,
             :source_name, :source_url, now(), :notes)
        returning id
    """)
    company_update = text("""
        update public.property_management_companies
        set website = coalesce(website, :website),
            main_phone = coalesce(main_phone, :phone),
            address_line_1 = coalesce(address_line_1, :address_line_1),
            address_line_2 = coalesce(address_line_2, :address_line_2),
            city = coalesce(city, :city),
            state = coalesce(state, :state),
            zip_code = coalesce(zip_code, :zip_code),
            source_name = coalesce(source_name, :source_name),
            source_url = coalesce(source_url, :source_url),
            verified_at = now(),
            is_active = true,
            notes = case
                when notes is null then :notes
                when position(:note_marker in notes) = 0 then notes || E'\\n' || :notes
                else notes
            end
        where id = :id
    """)
    contact_find = text("""
        select id
        from public.property_management_contacts
        where company_id = :company_id and normalized_name = :normalized_name
        order by created_at
        limit 1
    """)
    contact_insert = text("""
        insert into public.property_management_contacts
            (company_id, full_name, first_name, last_name, title, source_name, source_url,
             verified_at, notes)
        values
            (:company_id, :full_name, :first_name, :last_name, :title, :source_name, :source_url,
             now(), :notes)
        returning id
    """)
    contact_update = text("""
        update public.property_management_contacts
        set title = coalesce(:title, title),
            source_name = coalesce(source_name, :source_name),
            source_url = coalesce(source_url, :source_url),
            verified_at = now(),
            is_current = true,
            notes = case
                when notes is null then :notes
                when position(:note_marker in notes) = 0 then notes || E'\\n' || :notes
                else notes
            end
        where id = :id
    """)

    company_ids: dict[str, object] = {}
    contacts = tuple(contacts)
    with engine.connect() as connection:
        transaction = connection.begin()
        try:
            for canonical_name, company in sorted(companies.items()):
                aliases = [canonical_name, *(company.get("aliases") or []), *(company.get("profile_names") or [])]
                candidate_names = sorted({normalized_name(value) for value in aliases if clean(value)})
                existing = connection.execute(
                    company_find,
                    {"candidate_names": candidate_names, "canonical_name": normalized_name(canonical_name)},
                ).mappings().one_or_none()
                source_urls = list(company["source_urls"])
                note_marker = "IREM AMO verification:"
                memberships = ", ".join(sorted(set(company["memberships"])))
                notes = (
                    f"{note_marker} active Colorado listing(s) verified in the public AMO directory "
                    f"({memberships}); profiles: {'; '.join(source_urls)}"
                )
                values = {
                    "name": canonical_name,
                    "website": company.get("website"),
                    "phone": company.get("phone"),
                    "address_line_1": company.get("address_line_1"),
                    "address_line_2": company.get("address_line_2"),
                    "city": company.get("city"),
                    "state": company.get("state"),
                    "zip_code": company.get("zip_code"),
                    "source_name": SOURCE_NAME,
                    "source_url": source_urls[0],
                    "notes": notes,
                    "note_marker": note_marker,
                }
                if existing is None:
                    company_id = connection.execute(company_insert, values).scalar_one()
                    stats["companies_inserted"] += 1
                else:
                    company_id = existing["id"]
                    connection.execute(company_update, {**values, "id": company_id})
                    stats["companies_updated"] += 1
                company_ids[canonical_name] = company_id

            for contact in contacts:
                company_id = company_ids[contact.company_name]
                first_name, last_name = split_person_name(contact.full_name)
                note_marker = "Official-company contact verification:"
                notes = (
                    f"{note_marker} company relationship and title verified on the public company page; "
                    "direct email and phone were not published."
                )
                values = {
                    "company_id": company_id,
                    "normalized_name": normalized_name(contact.full_name),
                    "full_name": contact.full_name,
                    "first_name": first_name,
                    "last_name": last_name,
                    "title": contact.title,
                    "source_name": "Official company website",
                    "source_url": contact.source_url,
                    "notes": notes,
                    "note_marker": note_marker,
                }
                contact_id = connection.execute(contact_find, values).scalar_one_or_none()
                if contact_id is None:
                    connection.execute(contact_insert, values)
                    stats["contacts_inserted"] += 1
                else:
                    connection.execute(contact_update, {**values, "id": contact_id})
                    stats["contacts_updated"] += 1

            if commit:
                transaction.commit()
            else:
                transaction.rollback()
        except Exception:
            transaction.rollback()
            raise
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--env-file", type=Path, required=True, help="Environment file containing DATABASE_URL")
    parser.add_argument("--commit", action="store_true", help="Commit changes; otherwise run a rollback-only preview")
    args = parser.parse_args()

    profiles = [fetch_profile(office) for office in AMO_OFFICES]
    companies = merge_profiles(profiles)
    stats = load_database(companies, VERIFIED_CONTACTS, args.env_file, args.commit)
    print(json.dumps(stats, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
