"""Readable, filesystem-safe filenames for Roof Intelligence PDFs."""

from __future__ import annotations

import re


def roof_report_pdf_filename(
    address: object,
    city: object,
    *,
    revision_index: int | None = None,
) -> str:
    """Name every customer PDF from only its street address and city."""
    street = str(address or "").split(",", 1)[0].strip()
    street = re.sub(
        r"\s+(?:APT|APARTMENT|UNIT|STE|SUITE|#)\b.*$",
        "",
        street,
        flags=re.IGNORECASE,
    ).strip()
    city_name = str(city or "").strip()

    def readable(value: str) -> str:
        words = []
        for word in value.split():
            upper = word.upper()
            if upper in {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}:
                words.append(upper)
            elif re.fullmatch(r"\d+(?:ST|ND|RD|TH)", upper):
                words.append(upper[:-2] + upper[-2:].lower())
            else:
                words.append(word.title())
        return " ".join(words)

    base = " ".join(part for part in (readable(street), readable(city_name)) if part)
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", base)
    base = " ".join(base.split()).strip(" .") or "Roof Intelligence Report"
    base = base[:180].rstrip(" .")

    if revision_index is not None:
        try:
            revision_value = int(revision_index)
        except (TypeError, ValueError) as exc:
            raise ValueError("revision_index must be a positive integer") from exc
        if revision_value < 1:
            raise ValueError("revision_index must be a positive integer")
    return f"{base}.pdf"
