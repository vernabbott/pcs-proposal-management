#!/usr/bin/env python3
"""Extract and conservatively classify the CoStar property-manager email DOCX."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re

from docx import Document


VERIFIED: dict[str, tuple[str, str]] = {
    "amli.com": ("AMLI Residential", "https://www.amli.com/"),
    "apexreco.com": ("Apex Real Estate Advisors", "https://apexreco.com/"),
    "bartellre.com": ("Bartell & Company Real Estate", "https://bartellre.com/"),
    "bearpawrentals.com": ("Bear Paw Property Management", "https://www.bearpawrentals.com/"),
    "bluesailpm.com": ("Blue Sail Property Management", "https://bluesailpm.com/"),
    "brookfieldpropertiesretail.com": ("Brookfield Properties", "https://www.brookfieldproperties.com/"),
    "cbre.com": ("CBRE", "https://www.cbre.com/services/manage-properties-and-portfolios"),
    "cis.cushwake.com": ("Cushman & Wakefield", "https://www.cushmanwakefield.com/en/services/property-management"),
    "colliers.com": ("Colliers", "https://www.colliers.com/en/services/real-estate-management-services"),
    "colliersb-k.com": ("Colliers", "https://www.colliers.com/en/services/real-estate-management-services"),
    "coloradohtc.com": ("Colorado Health & Tech Centers", "https://www.coloradohtc.com/"),
    "coloradorpm.com": ("Colorado RPM", "https://coloradorpm.com/"),
    "continuumllc.com": ("Continuum Partners", "https://continuumpartners.com/"),
    "continuumpartners.com": ("Continuum Partners", "https://continuumpartners.com/"),
    "creginc.com": ("Crosbie Real Estate Group / Crosbie Management Services", "https://creginc.com/"),
    "cushwake.com": ("Cushman & Wakefield", "https://www.cushmanwakefield.com/en/services/property-management"),
    "dunton-commercial.com": ("Dunton Commercial", "https://dunton-commercial.com/"),
    "firt.com": ("First Industrial Realty Trust", "https://www.firstindustrial.com/"),
    "gb85.com": ("Griffis/Blessing", "https://www.gb85.com/"),
    "graniteprop.com": ("Granite Properties", "https://graniteprop.com/about-granite/operations/"),
    "greystar.com": ("Greystar", "https://www.greystar.com/"),
    "hospitality-work.com": ("Hospitality at Work", "https://hospitality-work.com/"),
    "kewrealty.com": ("KEW Realty", "https://kewrealty.com/about/"),
    "kimcorealty.com": ("Kimco Realty", "https://www.kimcorealty.com/"),
    "lafayettecompany.com": ("Lafayette Property Company", "https://lafayettecompany.com/"),
    "lcpmanagement.net": ("LCP Management", "https://lcpmanagement.net/"),
    "legacyproperties-pm.com": ("Legacy Properties-PM", "https://legacyproperties-pm.com/"),
    "lillibridge.com": ("Lillibridge Healthcare Services", "https://www.lillibridge.com/about"),
    "makersre.com": ("Makers Commercial", "https://www.makersre.com/"),
    "mercyhousing.org": ("Mercy Housing Management Group", "https://www.mercyhousing.org/partner-with-us/property-management/"),
    "newcastle83.com": ("Newcastle Properties", "https://www.newcastle83.com/"),
    "nmrk.com": ("Newmark", "https://www.nmrk.com/services/property-management"),
    "northstarcp.com": ("Northstar Commercial Partners", "https://www.northstarcommercialpartners.com/about/"),
    "oldvine.net": ("Old Vine Property Group", "https://oldvine.net/"),
    "paramountpropertyco.com": ("Paramount Property Company", "https://paramountpropertycompany.com/"),
    "parkrpm.com": ("Park Realty & Property Management", "https://parkrpm.com/"),
    "pmimilehigh.com": ("Copper Vine Property Management (formerly PMI Mile High)", "https://www.coppervinepropertymanagement.com/"),
    "ppmdenver.com": ("Performance Property Management", "https://ppmdenver.com/"),
    "proequityam.com": ("ProEquity Asset Management", "https://proequity.am/"),
    "prologis.com": ("Prologis", "https://www.prologis.com/"),
    "propertysensere.com": ("PropertySense Real Estate", "https://www.propertysensere.com/"),
    "propprealty.com": ("Propp Realty Management", "https://propprealtymanagement.com/"),
    "realcapitalsolutions.com": ("Real Capital Solutions", "https://www.realcapitalsolutions.com/"),
    "regencycenters.com": ("Regency Centers", "https://www.regencycenters.com/"),
    "rpmpropmgt.com": ("Real Property Management", "https://www.rpmpropmgt.com/"),
    "sessionsllc.com": ("Sessions Group", "https://sessionsllc.com/"),
    "shamesmakovsky.com": ("NAI Shames Makovsky", "https://shamesmakovsky.com/"),
    "skbcos.com": ("ScanlanKemperBard Companies", "https://www.skbcos.com/"),
    "streamrealty.com": ("Stream Realty Partners", "https://streamrealty.com/services/property-management/"),
    "theshermanagencyinc.com": ("The Sherman Agency", "https://www.theshermanagencyinc.com/management-services/"),
    "thesitusgroup.com": ("The Situs Group", "https://www.thesitusgroup.com/"),
    "thompsonthrift.com": ("Thompson Thrift", "https://www.thompsonthrift.com/residential/our-services"),
    "transwestern.com": ("Transwestern", "https://transwestern.com/services/asset-services"),
    "trevey.com": ("Trevey Commercial Real Estate", "https://trevey.com/our-services/"),
    "westardenver.com": ("Westar Real Property Services", "https://www.westardenver.com/"),
    "westerncenters.com": ("Western Centers", "https://www.westerncenters.com/about/"),
}

PERSONAL_DOMAINS = {"aol.com", "comcast.net", "gmail.com", "hotmail.com", "icloud.com", "msn.com"}
EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+$")


def extract_emails(path: Path) -> list[str]:
    text = "\n".join(paragraph.text for paragraph in Document(path).paragraphs)
    emails = [token.strip(" ,;\t\r\n").casefold() for token in re.split(r"[;\s]+", text)]
    return [email for email in emails if EMAIL_RE.fullmatch(email)]


def build_review(path: Path) -> dict[str, object]:
    extracted = extract_emails(path)
    seen: set[str] = set()
    rows: list[dict[str, str]] = []
    duplicates = 0
    for email in extracted:
        if email in seen:
            duplicates += 1
            continue
        seen.add(email)
        domain = email.rsplit("@", 1)[1]
        verified = VERIFIED.get(domain)
        if verified:
            company, source = verified
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
        "review_method": "Conservative employer-domain verification against current company sources; personal and ambiguous domains excluded.",
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
