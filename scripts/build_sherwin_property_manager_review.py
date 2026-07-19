#!/usr/bin/env python3
"""Extract and conservatively classify the Sherwin July 2025 email DOCX."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re

from docx import Document


# These domains were already verified and reconciled to companies in Supabase.
EXISTING: dict[str, str] = {
    "assetliving.com": "Asset Living",
    "brcrealestate.com": "BRC Multifamily",
    "cardinalgroup.com": "Cardinal Group Management",
    "centralmgt.com": "Central Management",
    "comcapmgmt.com": "ComCap Management",
    "coughlinmgt.com": "Coughlin Management Company",
    "cowboyproperties.com": "Cowboy Properties",
    "cruisemgmt.com": "Cruise Management",
    "dunmire.net": "Dunmire Property Management",
    "firmusmgmt.com": "Firmus Management",
    "gaddismanagement.com": "Gaddis Management",
    "greystar.com": "Greystar",
    "habitatmgmt.com": "Habitat Management",
    "highmarkres.com": "Highmark Residential",
    "incopm.com": "INCO Property Management",
    "ipmcolorado.com": "IPM Colorado",
    "ipmresidentialpm.com": "IPM Residential Property Management",
    "jp-co.com": "Jordon Perlmutter & Co.",
    "legacypartners.com": "Legacy Partners",
    "mbpropertymanagement.com": "MB Property Management",
    "mercyhousing.org": "Mercy Housing Management Group",
    "moorepm.com": "Moore Property Management Services",
    "mosaic1mgt.com": "Mosaic Management",
    "pierpropertyservices.com": "Pier Property Services",
    "placesmanagement.com": "Places Management",
    "platinum-properties.net": "Platinum Properties",
    "polarstarproperties.com": "Polar Star Properties",
    "rentbouldernow.com": "Boulder Property Management",
    "renthrg.com": "HRG Property Management",
    "residentialniche.com": "Residential Niche",
    "rio-realestate.com": "RIO Real Estate",
    "rosenthalproperties.com": "Rosenthal Properties",
    "rpmpropmgt.com": "Real Property Management",
    "skm-management.com": "SKM Management",
    "smpmanagers.com": "SMP Management",
    "thesitusgroup.com": "The Situs Group",
    "zealmgmt.com": "Zeal Property Management",
}

# New employers confirmed from organization or authoritative portfolio sources.
NEW: dict[str, tuple[str, str]] = {
    "auroraha.org": ("Housing Authority of the City of Aurora", "https://www.aurorahousing.org/about"),
    "aurorahousing.org": ("Housing Authority of the City of Aurora", "https://www.aurorahousing.org/about"),
    "availcolorado.com": ("Avail Property Management", "https://availcolorado.com/who-we-are/"),
    "boulderhousing.org": ("Boulder Housing Partners", "https://boulderhousing.org/about/"),
    "brothersredevelopment.org": ("Brothers Property Management", "https://brothersredevelopment.org/property-management/"),
    "cornerstonec21.com": ("Century 21 Cornerstone", "https://www.cornerstonec21.com/management-services"),
    "corumrealestate.com": ("Corum Real Estate", "https://www.corumrealestate.com/services"),
    "deltahousingauthority.org": ("Delta Housing Authority", "https://deltahousingauthority.org/rural-development/"),
    "greeley-weldha.org": ("Greeley Housing Authority", "https://greeleyhousing.org/properties"),
    "greccio.org": ("Greccio Housing", "https://www.greccio.org/greccio-commercial"),
    "hopecommunities.org": ("Hope Communities", "https://hopecommunities.org/affordable-housing/carolton-arms-apartments/"),
    "interurbancompanies.com": ("Interurban Companies", "https://www.interurbancompanies.com/interurban-properties/"),
    "lcmpm.com": ("LCM Property Management", "https://www.lcmpropertymanagement.com/"),
    "mwhs.org": ("Metro West Housing Solutions", "https://mwhs.org/"),
    "oakridgepropertiesllc.com": ("Oakridge Properties", "https://oakridgepropertiesllc.com/"),
    "spectrumcre.com": ("Spectrum Commercial Real Estate", "https://spectrumcre.com/"),
    "terramanagementgroupllc.com": ("Terra Management Group", "https://www.brightonco.gov/AgendaCenter/ViewFile/Agenda/_05022024-1784"),
    "thistle.us": ("Thistle Community Housing", "https://www.thistlecommunityhousing.org/about-us"),
    "urbanphenix.com": ("Urban Phenix", "https://urbanphenix.com/third-party-property-management/"),
    "waypointre.com": ("Waypoint Real Estate", "https://www.waypointre.com/"),
}

PERSONAL_DOMAINS = {"aol.com", "comcast.net", "cox.net", "gmail.com", "hotmail.com", "icloud.com", "msn.com", "rmi.net"}
EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+$")


def extract_emails(path: Path) -> list[str]:
    text = "\n".join(paragraph.text for paragraph in Document(path).paragraphs)
    tokens = [token.strip(" ,;\t\r\n").casefold() for token in re.split(r"[;\s]+", text)]
    return [email for email in tokens if EMAIL_RE.fullmatch(email)]


def build_review(path: Path) -> dict[str, object]:
    extracted = extract_emails(path)
    verified = {domain: (company, f"https://{domain}/") for domain, company in EXISTING.items()}
    verified.update(NEW)
    seen: set[str] = set()
    rows: list[dict[str, str]] = []
    duplicates = 0
    for email in extracted:
        if email in seen:
            duplicates += 1
            continue
        seen.add(email)
        domain = email.rsplit("@", 1)[1]
        match = verified.get(domain)
        if match:
            company, source = match
            rows.append({
                "email": email,
                "domain": domain,
                "company": company,
                "classification": "Yes",
                "confidence": "High",
                "source": source,
                "reason": "Employer is verified as providing property management or directly operating a managed real-estate portfolio.",
            })
        else:
            personal = domain in PERSONAL_DOMAINS
            rows.append({
                "email": email,
                "domain": domain,
                "company": "",
                "classification": "No",
                "confidence": "High" if personal else "Medium",
                "source": "",
                "reason": (
                    "Personal mailbox cannot be reliably assigned to a property-management company."
                    if personal
                    else "No sufficiently strong evidence that the employer is a property-management company; excluded conservatively."
                ),
            })
    yes = sum(row["classification"] == "Yes" for row in rows)
    return {
        "source_file": str(path),
        "review_method": "Conservative employer-domain verification against current company sources; universities, governments, vendors, personal mailboxes, and ambiguous employers excluded.",
        "summary": {
            "emails_extracted": len(extracted),
            "unique_emails": len(rows),
            "duplicate_emails": duplicates,
            "high_confidence_yes": yes,
            "excluded": len(rows) - yes,
            "verified_companies": len({row["company"] for row in rows if row["classification"] == "Yes"}),
        },
        "emails": rows,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    payload = build_review(args.docx)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload["summary"], indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
