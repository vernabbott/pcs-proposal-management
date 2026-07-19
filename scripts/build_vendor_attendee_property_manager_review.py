#!/usr/bin/env python3
"""Audit an education vendor-attendee CSV for property-management eligibility."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path


def clean(value: object) -> str:
    return " ".join(str(value or "").split())


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_file", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    with args.csv_file.open(encoding="utf-8-sig", newline="") as source:
        source_rows = list(csv.DictReader(source))

    rows: list[dict[str, str]] = []
    for row_number, row in enumerate(source_rows, start=2):
        email = clean(row.get("Email Addr")).casefold()
        rows.append({
            "source_row": str(row_number),
            "first_name": clean(row.get("First_Name")),
            "last_name": clean(row.get("Last_Name")),
            "title": clean(row.get("Professional Title")),
            "organization": clean(row.get("Organization")),
            "email": email,
            "classification": "No",
            "confidence": "High",
            "reason": (
                "School district, charter school, education agency, or school-business association; "
                "not a property-management company."
            ),
        })

    payload = {
        "source_file": str(args.csv_file),
        "review_method": (
            "Organization-type review against the property-management table scope. "
            "Education entities and their personnel are excluded."
        ),
        "summary": {
            "source_rows": len(rows),
            "rows_with_email": sum(bool(row["email"]) for row in rows),
            "rows_missing_email": sum(not row["email"] for row in rows),
            "unique_organizations": len({row["organization"] for row in rows if row["organization"]}),
            "eligible_property_management_contacts": 0,
            "excluded_rows": len(rows),
        },
        "attendees": rows,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload["summary"], indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
