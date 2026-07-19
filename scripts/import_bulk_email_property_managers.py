#!/usr/bin/env python3
"""Import high-confidence property-management domains from a reviewed email list.

The review JSON is produced from the bulk-email audit. Only rows explicitly
classified as high-confidence ``Yes`` matches are eligible. Email validation
checks syntax and the domain's MX (or legacy address) records; it does not claim
that an individual mailbox accepted an SMTP delivery.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
import json
import os
from pathlib import Path
import re
import subprocess
from typing import Any


SOURCE_NAME = "PCS Bulk Email List / verified property-management domain"
UNKNOWN_CONTACT_NAME = "Unknown"
EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]{0,251}[A-Za-z0-9])?$", re.ASCII)
GENERIC_MAILBOXES = {
    "admin",
    "corporate",
    "customercare",
    "customerservice",
    "hello",
    "info",
    "leasing",
    "office",
    "service",
}


def clean(value: object) -> str:
    return " ".join(str(value or "").split())


def normalized_name(value: str) -> str:
    return clean(value).casefold()


def validated_rows(review_json: Path) -> tuple[list[dict[str, str]], dict[str, int]]:
    payload = json.loads(review_json.read_text(encoding="utf-8"))
    rows: list[dict[str, str]] = []
    stats = {
        "review_rows": len(payload.get("emails", [])),
        "eligible_high_confidence_yes": 0,
        "invalid_syntax": 0,
        "missing_company": 0,
        "duplicate_email": 0,
    }
    seen: set[str] = set()
    for raw in payload.get("emails", []):
        if raw.get("classification") != "Yes" or raw.get("confidence") != "High":
            continue
        stats["eligible_high_confidence_yes"] += 1
        email = clean(raw.get("email")).casefold()
        company = clean(raw.get("company"))
        domain = clean(raw.get("domain")).casefold()
        source_url = clean(raw.get("source")) or f"https://{domain}"
        if not company or not domain:
            stats["missing_company"] += 1
            continue
        if not EMAIL_RE.fullmatch(email) or email.rsplit("@", 1)[1] != domain:
            stats["invalid_syntax"] += 1
            continue
        if email in seen:
            stats["duplicate_email"] += 1
            continue
        seen.add(email)
        rows.append({"email": email, "domain": domain, "company": company, "source_url": source_url})
    return rows, stats


def dig(domain: str, record_type: str) -> list[str]:
    result = subprocess.run(
        ["/usr/bin/dig", "+short", record_type, domain],
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )
    if result.returncode != 0:
        return []
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def mail_domain_status(domain: str) -> tuple[str, str]:
    mx = dig(domain, "MX")
    if mx:
        if any(line.split()[0] == "0" and line.rstrip().endswith(".") and len(line.split()) == 2 and line.split()[1] == "." for line in mx):
            return "no_mail", "Null MX"
        return "mx", "; ".join(mx[:3])
    address = dig(domain, "A") or dig(domain, "AAAA")
    if address:
        return "address_fallback", "; ".join(address[:3])
    return "no_mail", "No MX, A, or AAAA record"


def verify_domains(rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], dict[str, dict[str, str]]]:
    domains = sorted({row["domain"] for row in rows})
    with ThreadPoolExecutor(max_workers=8) as executor:
        outcomes = list(executor.map(mail_domain_status, domains))
    status = {domain: {"status": result[0], "evidence": result[1]} for domain, result in zip(domains, outcomes)}
    accepted = [row for row in rows if status[row["domain"]]["status"] != "no_mail"]
    return accepted, status


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


def grouped_companies(rows: list[dict[str, str]]) -> dict[str, dict[str, Any]]:
    grouped: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        grouped[normalized_name(row["company"])].append(row)
    companies: dict[str, dict[str, Any]] = {}
    for key, company_rows in grouped.items():
        # Prefer the shortest domain as the canonical corporate domain when a
        # company appears under both a corporate and property-specific domain.
        canonical = sorted(company_rows, key=lambda row: (len(row["domain"]), row["domain"]))[0]
        generic_emails = sorted(
            row["email"]
            for row in company_rows
            if row["email"].split("@", 1)[0] in GENERIC_MAILBOXES
        )
        companies[key] = {
            "name": canonical["company"],
            "email_domain": canonical["domain"],
            "website": canonical["source_url"],
            "main_email": generic_emails[0] if generic_emails else None,
            "rows": company_rows,
        }
    return companies


def reconcile_database(
    rows: list[dict[str, str]],
    env_file: Path,
    commit: bool,
    source_name: str = SOURCE_NAME,
    verification_label: str = "Bulk-email",
) -> dict[str, int]:
    from sqlalchemy import create_engine, text

    engine = create_engine(database_url(env_file), pool_pre_ping=True, connect_args={"sslmode": "require"})
    companies = grouped_companies(rows)
    stats = {
        "companies_eligible": len(companies),
        "companies_inserted": 0,
        "companies_updated": 0,
        "contacts_eligible": len(rows),
        "contacts_inserted": 0,
        "contacts_updated": 0,
        "email_company_conflicts": 0,
    }
    company_find = text("""
        select id
        from public.property_management_companies
        where normalized_name = :normalized_name
    """)
    company_insert = text("""
        insert into public.property_management_companies
            (name, website, email_domain, main_email, source_name, source_url,
             verified_at, notes)
        values
            (:name, :website, :email_domain, :main_email, :source_name, :source_url,
             now(), :notes)
        returning id
    """)
    company_update = text("""
        update public.property_management_companies
        set website = coalesce(website, :website),
            email_domain = coalesce(email_domain, :email_domain),
            main_email = coalesce(main_email, :main_email),
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
        select id, company_id
        from public.property_management_contacts
        where normalized_email = :email
        order by created_at
        limit 1
    """)
    contact_insert = text("""
        insert into public.property_management_contacts
            (company_id, full_name, business_email, source_name, source_url,
             verified_at, notes)
        values
            (:company_id, :full_name, :email, :source_name, :source_url,
             now(), :notes)
    """)
    contact_update = text("""
        update public.property_management_contacts
        set verified_at = now(),
            is_current = true,
            source_name = coalesce(source_name, :source_name),
            source_url = coalesce(source_url, :source_url),
            notes = case
                when notes is null then :notes
                when position(:note_marker in notes) = 0 then notes || E'\\n' || :notes
                else notes
            end
        where id = :id
    """)
    company_note_marker = f"{verification_label} domain verification:"
    company_notes = (
        f"{company_note_marker} company identity and mail-capable domain confirmed; "
        "address and phone are unknown."
    )
    contact_note_marker = f"{verification_label} verification:"
    contact_notes = (
        f"{contact_note_marker} syntax and employer domain mail routing confirmed; "
        "mailbox deliverability, person name, title, address, and phone are unknown."
    )
    try:
        with engine.connect() as connection:
            transaction = connection.begin()
            try:
                company_ids: dict[str, object] = {}
                for key, company in sorted(companies.items()):
                    company_id = connection.execute(company_find, {"normalized_name": key}).scalar_one_or_none()
                    values = {
                        "name": company["name"],
                        "website": company["website"],
                        "email_domain": company["email_domain"],
                        "main_email": company["main_email"],
                        "source_name": source_name,
                        "source_url": company["website"],
                        "notes": company_notes,
                        "note_marker": company_note_marker,
                    }
                    if company_id is None:
                        company_id = connection.execute(company_insert, values).scalar_one()
                        stats["companies_inserted"] += 1
                    else:
                        connection.execute(company_update, {**values, "id": company_id})
                        stats["companies_updated"] += 1
                    company_ids[key] = company_id

                for row in sorted(rows, key=lambda item: item["email"]):
                    company_id = company_ids[normalized_name(row["company"])]
                    existing = connection.execute(contact_find, {"email": row["email"]}).first()
                    values = {
                        "company_id": company_id,
                        "full_name": UNKNOWN_CONTACT_NAME,
                        "email": row["email"],
                        "source_name": source_name,
                        "source_url": row["source_url"],
                        "notes": contact_notes,
                        "note_marker": contact_note_marker,
                    }
                    if existing is None:
                        connection.execute(contact_insert, values)
                        stats["contacts_inserted"] += 1
                    elif existing.company_id == company_id:
                        connection.execute(contact_update, {**values, "id": existing.id})
                        stats["contacts_updated"] += 1
                    else:
                        stats["email_company_conflicts"] += 1
                if commit:
                    transaction.commit()
                else:
                    transaction.rollback()
            except Exception:
                transaction.rollback()
                raise
    finally:
        engine.dispose()
    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import reviewed property-management companies and emails")
    parser.add_argument("--review-json", type=Path, required=True, help="Bulk-email review JSON")
    parser.add_argument("--env-file", type=Path, required=True, help="Environment file containing DATABASE_URL")
    parser.add_argument("--source-name", default=SOURCE_NAME, help="Source label stored with imported rows")
    parser.add_argument(
        "--verification-label",
        default="Bulk-email",
        help="Short label used in verification notes",
    )
    parser.add_argument("--commit", action="store_true", help="Commit the transaction; otherwise roll it back")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows, validation_stats = validated_rows(args.review_json)
    accepted_rows, domain_status = verify_domains(rows)
    rejected_domains = sorted(domain for domain, result in domain_status.items() if result["status"] == "no_mail")
    dns_stats = {
        "domains_checked": len(domain_status),
        "domains_with_mx": sum(result["status"] == "mx" for result in domain_status.values()),
        "domains_with_address_fallback": sum(result["status"] == "address_fallback" for result in domain_status.values()),
        "domains_rejected": len(rejected_domains),
        "emails_rejected_for_domain": len(rows) - len(accepted_rows),
    }
    print("validation_stats=" + repr(validation_stats), flush=True)
    print("dns_stats=" + repr(dns_stats), flush=True)
    if rejected_domains:
        print("rejected_domains=" + repr(rejected_domains), flush=True)
    database_stats = reconcile_database(
        accepted_rows,
        args.env_file,
        args.commit,
        source_name=args.source_name,
        verification_label=args.verification_label,
    )
    print("database_stats=" + repr(database_stats), flush=True)
    print(f"commit={args.commit}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
