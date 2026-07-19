#!/usr/bin/env python3
"""Import Denver Metro BOMA property managers and contacts into Supabase.

The public BOMA principal directory mixes buildings, organizations, and people.
This importer only writes the phase-one entities that can be supported by an
explicit property-management field, an organization profile, or a building
name whose final segment names the manager and whose page lists contacts.
"""

from __future__ import annotations

import argparse
import html
from html.parser import HTMLParser
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen
from urllib.robotparser import RobotFileParser


BASE_URL = "https://members.bomadenver.org"
DIRECTORY_URL = BASE_URL + "/principal-member-directory/FindStartsWith?term=%23%21"
ROBOTS_URL = BASE_URL + "/robots.txt"
USER_AGENT = "PCS-Proposal-Management/1.0 (+public BOMA directory import)"
SOURCE_NAME = "Denver Metro BOMA Principal Member Directory"

MANAGEMENT_TITLE_WORDS = (
    "property manager",
    "general manager",
    "facility manager",
    "facilities manager",
    "operations",
    "asset manager",
)

COMPANY_ALIASES = {
    "americas capital partners": "America's Capital Partners",
    "cole taylor": "",
    "cushman and wakefield": "Cushman & Wakefield",
    "elevate real estate": "Elevate Real Estate Services",
    "industrial": "",
    "lba realty, llc": "LBA Realty",
    "nexcore group": "NexCore Group",
    "patrinely group": "Patrinely",
    "prime west": "Prime West Real Estate Services",
    "urban rennaisance group": "Urban Renaissance Group",
}


@dataclass
class Address:
    line_1: str | None = None
    line_2: str | None = None
    city: str | None = None
    state: str | None = None
    zip_code: str | None = None


@dataclass
class Company:
    name: str
    source_url: str
    website: str | None = None
    phone: str | None = None
    address: Address | None = None


@dataclass
class Contact:
    company_name: str
    full_name: str
    title: str | None
    source_url: str
    phone: str | None = None
    address: Address | None = None


class DetailLinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: set[str] = set()

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        href = dict(attrs).get("href") or ""
        if "/principal-member-directory/Details/" in href:
            self.links.add(urljoin(BASE_URL, href))


class TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def text(self) -> str:
        return clean_text(" ".join(self.parts))


def clean_text(value: object) -> str:
    return " ".join(html.unescape(str(value or "")).split())


def strip_html(value: str) -> str:
    parser = TextExtractor()
    parser.feed(value)
    return parser.text()


def normalized_key(value: str) -> str:
    return clean_text(value).casefold()


def canonical_company_name(value: str) -> str:
    name = clean_text(value).strip(" -:")
    return COMPANY_ALIASES.get(normalized_key(name), name)


def split_person_name(full_name: str) -> tuple[str | None, str | None]:
    parts = clean_text(full_name).split()
    if not parts:
        return None, None
    if len(parts) == 1:
        return parts[0], None
    return parts[0], " ".join(parts[1:])


def looks_like_person(value: str) -> bool:
    name = clean_text(value)
    parts = name.split()
    if not 2 <= len(parts) <= 5 or any(char.isdigit() for char in name):
        return False
    blocked = ("building", "center", "tower", "campus", "plaza", "library", "properties", "management")
    lowered = name.casefold()
    return not any(word in lowered for word in blocked)


def fetch(url: str, attempts: int = 3) -> str:
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        request = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "text/html,*/*"})
        try:
            with urlopen(request, timeout=45) as response:
                return response.read().decode("utf-8", "replace")
        except HTTPError as exc:
            last_error = exc
            if exc.code < 500 or attempt == attempts:
                raise
        except (TimeoutError, URLError, OSError) as exc:
            last_error = exc
            if attempt == attempts:
                raise
        time.sleep(attempt * 1.5)
    raise RuntimeError(f"Unable to fetch {url}: {last_error}")


def check_robots() -> None:
    parser = RobotFileParser()
    parser.set_url(ROBOTS_URL)
    parser.read()
    if not parser.can_fetch(USER_AGENT, DIRECTORY_URL):
        raise RuntimeError("Denver Metro BOMA robots.txt does not permit this directory import")


def match_text(pattern: str, page: str, flags: int = re.IGNORECASE | re.DOTALL) -> str | None:
    match = re.search(pattern, page, flags)
    return strip_html(match.group(1)) if match else None


def page_address(page: str) -> Address:
    street = match_text(r'<span[^>]+itemprop="streetAddress"[^>]*>(.*?)</span>', page)
    city = match_text(r'<span[^>]+itemprop="addressLocality"[^>]*>(.*?)</span>', page)
    state = match_text(r'<span[^>]+itemprop="addressRegion"[^>]*>(.*?)</span>', page)
    zip_code = match_text(r'<span[^>]+itemprop="postalCode"[^>]*>(.*?)</span>', page)
    line_1 = street
    line_2 = None
    if street:
        suite_match = re.search(r"^(.*?)(?:,?\s+(Suite|Ste\.?|#)\s*([^,]+))$", street, re.IGNORECASE)
        if suite_match:
            line_1 = clean_text(suite_match.group(1)).rstrip(",")
            line_2 = clean_text(f"{suite_match.group(2)} {suite_match.group(3)}")
    return Address(line_1=line_1, line_2=line_2, city=city, state=state, zip_code=zip_code)


def page_phone(page: str) -> str | None:
    return match_text(r'<span[^>]+itemprop="telephone"[^>]*>(.*?)</span>', page)


def external_website(page: str) -> str | None:
    candidates = re.findall(r'<a[^>]+href="(https?://[^"]+)"[^>]*>\s*(?:Visit Website)?\s*</a>', page, re.I)
    blocked = ("bomadenver.org", "google.com", "facebook.com", "linkedin.com", "twitter.com", "instagram.com")
    for candidate in candidates:
        if not any(domain in candidate.casefold() for domain in blocked):
            return html.unescape(candidate)
    return None


def contact_cards(page: str) -> list[tuple[str, str | None, str | None]]:
    contacts: list[tuple[str, str | None, str | None]] = []
    starts = [match.start() for match in re.finditer(r'<div class="card gz-contact-card">', page)]
    for index, start in enumerate(starts):
        end = starts[index + 1] if index + 1 < len(starts) else page.find("<!--", start)
        chunk = page[start:end if end > start else len(page)]
        name = None
        for anchor_text in re.findall(
            r'<a[^>]+href="[^"]*/principal-member-directory/Details/[^"]+"[^>]*>(.*?)</a>',
            chunk,
            re.I | re.S,
        ):
            candidate = strip_html(anchor_text)
            if candidate:
                name = candidate
                break
        if not name:
            continue
        title = match_text(r'<div class="gz-member-reptitle">(.*?)</div>', chunk)
        phone_match = re.search(r'href="tel:([^"]+)"', chunk, re.I)
        phone = clean_text(phone_match.group(1)) if phone_match else None
        contacts.append((name, title, phone))
    return contacts


def inferred_company_from_title(title: str, contacts: list[tuple[str, str | None, str | None]]) -> str | None:
    if not contacts or not any(
        contact_title and any(word in contact_title.casefold() for word in MANAGEMENT_TITLE_WORDS)
        for _, contact_title, _ in contacts
    ):
        return None
    parts = [clean_text(part) for part in re.split(r"\s+-\s+|\s*-\s*", title) if clean_text(part)]
    if len(parts) < 2:
        return None
    candidate = parts[-1]
    if any(char.isdigit() for char in candidate) or len(candidate) < 3:
        return None
    return canonical_company_name(candidate)


def parse_detail(url: str, page: str) -> tuple[Company | None, list[Contact], str]:
    title = match_text(r'<h1 class="gz-pagetitle"[^>]*>(.*?)</h1>', page)
    if not title:
        return None, [], "missing_title"
    org = match_text(r'<div class="gz-details-org">(.*?)</div>', page)
    manager = match_text(r"Property Management\s*:\s*(.*?)</p>", page)
    cards = contact_cards(page)
    address = page_address(page)
    phone = page_phone(page)

    company_name: str | None = None
    reason = "skipped"
    company_address: Address | None = None
    company_phone: str | None = None
    website: str | None = None

    if manager:
        company_name = canonical_company_name(manager)
        reason = "explicit_manager"
    elif org and looks_like_person(title):
        company_name = canonical_company_name(org)
        company_address = address
        company_phone = phone
        website = external_website(page)
        reason = "organization_profile"
    else:
        company_name = inferred_company_from_title(title, cards)
        if company_name:
            reason = "title_inference"

    if not company_name:
        return None, [], reason

    company = Company(
        name=company_name,
        source_url=url,
        website=website,
        phone=company_phone,
        address=company_address,
    )
    contacts: list[Contact] = []
    for name, contact_title, contact_phone in cards:
        contacts.append(
            Contact(
                company_name=company_name,
                full_name=name,
                title=contact_title,
                phone=contact_phone,
                address=address,
                source_url=url,
            )
        )

    if reason == "organization_profile" and not cards:
        contacts.append(
            Contact(
                company_name=company_name,
                full_name=title,
                title=None,
                phone=phone,
                address=address,
                source_url=url,
            )
        )
    return company, contacts, reason


def merge_company(existing: Company, incoming: Company) -> Company:
    return Company(
        name=existing.name,
        source_url=existing.source_url,
        website=existing.website or incoming.website,
        phone=existing.phone or incoming.phone,
        address=existing.address or incoming.address,
    )


def collect_directory(delay: float) -> tuple[dict[str, Company], dict[tuple[str, str], Contact], dict[str, int]]:
    check_robots()
    directory_page = fetch(DIRECTORY_URL)
    parser = DetailLinkParser()
    parser.feed(directory_page)
    urls = sorted(parser.links)
    if len(urls) < 300:
        raise RuntimeError(f"Expected at least 300 BOMA detail pages; found {len(urls)}")

    companies: dict[str, Company] = {}
    contacts: dict[tuple[str, str], Contact] = {}
    organization_profiles: list[tuple[Company, list[Contact]]] = []
    counts = {
        "directory_links": len(urls),
        "fetched": 0,
        "fetch_errors": 0,
        "explicit_manager": 0,
        "organization_profile": 0,
        "organization_profile_matched": 0,
        "organization_profile_unmatched": 0,
        "title_inference": 0,
        "skipped": 0,
        "missing_title": 0,
    }
    for index, url in enumerate(urls, start=1):
        try:
            page = fetch(url)
            counts["fetched"] += 1
            company, page_contacts, reason = parse_detail(url, page)
            counts[reason] = counts.get(reason, 0) + 1
            if company:
                if reason == "organization_profile":
                    organization_profiles.append((company, page_contacts))
                else:
                    company_key = normalized_key(company.name)
                    companies[company_key] = merge_company(companies[company_key], company) if company_key in companies else company
                    for contact in page_contacts:
                        contact_key = (company_key, normalized_key(contact.full_name))
                        current = contacts.get(contact_key)
                        if current is None or (not current.title and contact.title):
                            contacts[contact_key] = contact
        except Exception as exc:
            counts["fetch_errors"] += 1
            print(f"warning: unable to process {url}: {exc}", file=sys.stderr, flush=True)
        if index % 25 == 0 or index == len(urls):
            print(
                f"processed={index}/{len(urls)} companies={len(companies)} contacts={len(contacts)} errors={counts['fetch_errors']}",
                flush=True,
            )
        if delay:
            time.sleep(delay)

    # An individual profile's organization can be a building instead of the
    # person's employer. Use these profiles only to enrich companies that were
    # independently established by an explicit management field or a
    # management-labeled building listing.
    for company, page_contacts in organization_profiles:
        company_key = normalized_key(company.name)
        if company_key not in companies:
            counts["organization_profile_unmatched"] += 1
            continue
        counts["organization_profile_matched"] += 1
        companies[company_key] = merge_company(companies[company_key], company)
        for contact in page_contacts:
            contact_key = (company_key, normalized_key(contact.full_name))
            current = contacts.get(contact_key)
            if current is None or (not current.title and contact.title):
                contacts[contact_key] = contact
    return companies, contacts, counts


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


def address_values(address: Address | None) -> dict[str, str | None]:
    address = address or Address()
    return {
        "address_line_1": address.line_1,
        "address_line_2": address.line_2,
        "city": address.city,
        "state": address.state,
        "zip_code": address.zip_code,
    }


def load_database(companies: dict[str, Company], contacts: dict[tuple[str, str], Contact], env_file: Path) -> dict[str, int]:
    from sqlalchemy import create_engine, text

    engine = create_engine(database_url(env_file), pool_pre_ping=True, connect_args={"sslmode": "require"})
    stats = {"companies_inserted": 0, "companies_updated": 0, "contacts_inserted": 0, "contacts_updated": 0}
    company_ids: dict[str, object] = {}

    company_insert = text("""
        insert into public.property_management_companies
            (name, website, main_phone, address_line_1, address_line_2, city, state, zip_code,
             source_name, source_url, verified_at)
        values
            (:name, :website, :phone, :address_line_1, :address_line_2, :city, :state, :zip_code,
             :source_name, :source_url, now())
        on conflict on constraint property_management_companies_normalized_name_key do nothing
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
            is_active = true
        where normalized_name = :normalized_name
        returning id
    """)
    contact_find = text("""
        select id
        from public.property_management_contacts
        where company_id = :company_id and normalized_name = :normalized_name
        order by normalized_email is not null desc, created_at
        limit 1
    """)
    contact_insert = text("""
        insert into public.property_management_contacts
            (company_id, full_name, first_name, last_name, title, direct_phone,
             address_line_1, address_line_2, city, state, zip_code,
             source_name, source_url, verified_at)
        values
            (:company_id, :full_name, :first_name, :last_name, :title, :phone,
             :address_line_1, :address_line_2, :city, :state, :zip_code,
             :source_name, :source_url, now())
        returning id
    """)
    contact_update = text("""
        update public.property_management_contacts
        set title = coalesce(:title, title),
            direct_phone = coalesce(direct_phone, :phone),
            address_line_1 = coalesce(address_line_1, :address_line_1),
            address_line_2 = coalesce(address_line_2, :address_line_2),
            city = coalesce(city, :city),
            state = coalesce(state, :state),
            zip_code = coalesce(zip_code, :zip_code),
            source_name = coalesce(source_name, :source_name),
            source_url = coalesce(source_url, :source_url),
            verified_at = now(),
            is_current = true
        where id = :id
    """)

    try:
        with engine.begin() as connection:
            for key, company in sorted(companies.items()):
                values = {
                    "name": company.name,
                    "normalized_name": key,
                    "website": company.website,
                    "phone": company.phone,
                    "source_name": SOURCE_NAME,
                    "source_url": company.source_url,
                    **address_values(company.address),
                }
                company_id = connection.execute(company_insert, values).scalar_one_or_none()
                if company_id is None:
                    company_id = connection.execute(company_update, values).scalar_one()
                    stats["companies_updated"] += 1
                else:
                    stats["companies_inserted"] += 1
                company_ids[key] = company_id

            for (company_key, contact_key), contact in sorted(contacts.items()):
                company_id = company_ids[company_key]
                first_name, last_name = split_person_name(contact.full_name)
                values = {
                    "company_id": company_id,
                    "normalized_name": contact_key,
                    "full_name": contact.full_name,
                    "first_name": first_name,
                    "last_name": last_name,
                    "title": contact.title,
                    "phone": contact.phone,
                    "source_name": SOURCE_NAME,
                    "source_url": contact.source_url,
                    **address_values(contact.address),
                }
                contact_id = connection.execute(contact_find, values).scalar_one_or_none()
                if contact_id is None:
                    connection.execute(contact_insert, values)
                    stats["contacts_inserted"] += 1
                else:
                    connection.execute(contact_update, {**values, "id": contact_id})
                    stats["contacts_updated"] += 1
    finally:
        engine.dispose()
    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import the public Denver Metro BOMA principal directory")
    parser.add_argument("--commit", action="store_true", help="Write the normalized companies and contacts to Supabase")
    parser.add_argument("--delay", type=float, default=0.2, help="Delay between detail-page requests (default: 0.2 seconds)")
    parser.add_argument("--env-file", type=Path, required=True, help="Environment file containing DATABASE_URL")
    parser.add_argument("--show-companies", action="store_true", help="Print normalized company names for review")
    args = parser.parse_args()
    if args.delay < 0:
        parser.error("--delay must not be negative")
    return args


def main() -> int:
    args = parse_args()
    companies, contacts, collection_stats = collect_directory(args.delay)
    print("collection_stats=" + repr(collection_stats), flush=True)
    print(f"normalized_companies={len(companies)} normalized_contacts={len(contacts)}", flush=True)
    if args.show_companies:
        for company in sorted(companies.values(), key=lambda item: item.name.casefold()):
            print(f"company={company.name}", flush=True)
    if not args.commit:
        print("dry_run=true database_changes=0", flush=True)
        return 0
    load_stats = load_database(companies, contacts, args.env_file)
    print("load_stats=" + repr(load_stats), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
