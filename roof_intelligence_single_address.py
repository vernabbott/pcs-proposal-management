#!/usr/bin/env python3
"""Generate a Roof Intelligence report for one address using the GIS report project."""

from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import sys
import time
from pathlib import Path


DEFAULT_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/PCS Roof Intelligence Report"
)


STREET_ALIASES = {
    "AVENUE": "AVE",
    "BOULEVARD": "BLVD",
    "CIRCLE": "CIR",
    "COURT": "CT",
    "DRIVE": "DR",
    "LANE": "LN",
    "PARKWAY": "PKWY",
    "PLACE": "PL",
    "ROAD": "RD",
    "STREET": "ST",
    "TERRACE": "TER",
    "NORTH": "N",
    "SOUTH": "S",
    "EAST": "E",
    "WEST": "W",
}


def normalize_address(value: object) -> str:
    tokens = re.findall(r"[A-Z0-9]+", str(value or "").upper())
    normalized_tokens = []
    for token in tokens:
        if token in {"DENVER", "CO", "COLORADO"}:
            continue
        if len(token) == 5 and token.isdigit():
            continue
        normalized_tokens.append(STREET_ALIASES.get(token, token))
    return " ".join(normalized_tokens)


def address_zip(value: object) -> str:
    match = re.search(r"\b(\d{5})(?:-\d{4})?\b", str(value or ""))
    return match.group(1) if match else ""


def score_address(query: str, candidate: str) -> float:
    if not query or not candidate:
        return 0.0
    if query == candidate:
        return 1.0
    if query.startswith(candidate) or candidate.startswith(query):
        return 0.96
    query_parts = set(query.split())
    candidate_parts = set(candidate.split())
    if query_parts and query_parts.issubset(candidate_parts):
        return 0.93
    return difflib.SequenceMatcher(None, query, candidate).ratio()


def sql_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def live_address_where_clauses(address: str, collector) -> list[str]:
    normalized = normalize_address(address)
    tokens = normalized.split()
    if not tokens:
        return []

    number = tokens[0]
    street_tokens = [token for token in tokens[1:] if not token.isdigit()]
    zip_code = address_zip(address)
    clauses = []

    if number.isdigit() and street_tokens:
        street = street_tokens[0]
        primary_clause = f"SITUS_ADDRESS_LINE1 LIKE {sql_literal(number + '%' + street + '%')}"
        if zip_code and "SITUS_ZIP" in collector.collect_parcel_fields():
            clauses.append(f"{primary_clause} AND SITUS_ZIP LIKE {sql_literal(zip_code + '%')}")
        clauses.append(primary_clause)
    if number.isdigit():
        clauses.append(f"SITUS_ADDRESS_LINE1 LIKE {sql_literal(number + '%')}")
    if street_tokens:
        clauses.append(f"SITUS_ADDRESS_LINE1 LIKE {sql_literal('%' + street_tokens[0] + '%')}")

    return list(dict.fromkeys(clauses))


def collect_live_parcels_for_address(address: str, collector) -> list[dict]:
    parcels: list[dict] = []
    seen: set[str] = set()
    for where in live_address_where_clauses(address, collector):
        page = collector.fetch_page(
            collector.PARCELS_URL,
            where,
            0,
            collector.collect_parcel_fields(),
            return_geometry=True,
        )
        for feature in page.get("features", []):
            attrs = feature.get("attributes", {}) or {}
            attrs["parcel_shape_area"] = attrs.get("Shape__Area")
            attrs["parcel_geometry"] = collector.geometry_to_wkt(feature.get("geometry"))
            attrs["full_parcel_number"] = collector.parcel_join_key(attrs)
            key = attrs.get("full_parcel_number") or json.dumps(attrs, sort_keys=True)
            if key in seen:
                continue
            seen.add(key)
            parcels.append(attrs)
        if parcels:
            break
    return parcels


def find_parcel_for_address(address: str, parcels: list[dict], collector) -> tuple[dict, float, str]:
    query = normalize_address(address)
    zip_code = address_zip(address)
    candidates = parcels
    if zip_code:
        zip_matches = [parcel for parcel in parcels if collector.parcel_zip(parcel) == zip_code]
        if zip_matches:
            candidates = zip_matches

    best_parcel: dict | None = None
    best_score = 0.0
    best_address = ""
    for parcel in candidates:
        for raw_candidate in (
            parcel.get("SITUS_ADDRESS_LINE1"),
            collector.address_from_record(parcel),
        ):
            candidate = normalize_address(raw_candidate)
            current_score = score_address(query, candidate)
            if current_score > best_score:
                best_score = current_score
                best_parcel = parcel
                best_address = str(raw_candidate or "")

    if not best_parcel or best_score < 0.74:
        raise RuntimeError(f"No parcel match found for address: {address}")
    return best_parcel, best_score, best_address


def find_parcel_live_or_cached(address: str, parcel_cache: Path, collector) -> tuple[dict, float, str, str]:
    try:
        live_parcels = collect_live_parcels_for_address(address, collector)
        if live_parcels:
            parcel, score, matched_address = find_parcel_for_address(address, live_parcels, collector)
            return parcel, score, matched_address, "Live Denver parcel service"
    except Exception as exc:
        print(f"Warning: live parcel lookup failed: {exc}", file=sys.stderr)

    if not parcel_cache.exists():
        raise RuntimeError(
            f"No live parcel match found for {address}, and fallback parcel cache was not found: {parcel_cache}"
        )

    parcels = collector.load_or_collect_parcels(str(parcel_cache))
    parcel, score, matched_address = find_parcel_for_address(address, parcels, collector)
    return parcel, score, matched_address, "Local parcel cache fallback"


def report_row_from_record(record: dict, collector) -> dict:
    collector.add_output_fields(record)
    return {label: record.get(field, "") for field, label in collector.OUTPUT_FIELDS}


def safe_report_name(parcel: str, address: str) -> str:
    safe_address = "".join(ch if ch.isalnum() else "-" for ch in address.lower()).strip("-")
    return f"{parcel}-{safe_address or 'roof-report'}.pdf"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate one Roof Intelligence report by address.")
    parser.add_argument("--address", required=True, help="Property address")
    parser.add_argument("--project-dir", default=str(DEFAULT_PROJECT_DIR), help="Roof Intelligence project directory")
    parser.add_argument("--parcel-cache", default="colorado_parcel_data.csv", help="Optional fallback parcel CSV path")
    parser.add_argument("--output-dir", default="roof_intelligence_reports", help="Report output directory")
    parser.add_argument("--image-dir", default="aerial_images_single_address", help="Aerial image output directory")
    parser.add_argument("--analysis-cache-dir", default="roof_ai_analysis_cache", help="AI analysis cache directory")
    parser.add_argument("--use-ai", action="store_true", help="Run AI vision analysis")
    parser.add_argument("--ai-provider", choices=("openai", "gemini"), default="openai", help="AI provider")
    parser.add_argument("--ai-model", default=None, help="AI model override")
    parser.add_argument("--allow-ai-fallback", action="store_true", help="Generate fallback analysis if AI fails")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_dir = Path(args.project_dir).expanduser().resolve()
    if not project_dir.exists():
        raise RuntimeError(f"Roof Intelligence project directory does not exist: {project_dir}")

    sys.path.insert(0, str(project_dir))
    import collect_denver_buildings_with_parcels as collector
    import generate_roof_intelligence_reports as reports

    reports.load_env_file(project_dir / ".env")

    parcel_cache = Path(args.parcel_cache)
    if not parcel_cache.is_absolute():
        parcel_cache = project_dir / parcel_cache

    collector.init_crs_transformers(collector.DENVER_BUILDINGS_URL, collector.PARCELS_URL)
    parcel, match_score, matched_address, lookup_source = find_parcel_live_or_cached(args.address, parcel_cache, collector)

    parcel_geometry = collector.get_parcel_bounds_in_building_crs([parcel])
    if not parcel_geometry:
        raise RuntimeError(f"Parcel geometry was not available for {matched_address}")

    buildings = collector.collect_buildings(None, parcel_geometry)
    combined = collector.combine_data(buildings, [parcel])
    parcel_key = collector.parcel_join_key(parcel)
    matched_records = [
        record for record in combined if collector.parcel_join_key(record) == parcel_key
    ]
    if not matched_records:
        raise RuntimeError(f"No building footprint matched parcel {parcel_key} for {matched_address}")

    matched_records.sort(key=lambda record: float(record.get("roof_squares", 0) or 0), reverse=True)
    record = matched_records[0]
    collector.add_aerial_image_fields(record)

    image_dir = Path(args.image_dir)
    if not image_dir.is_absolute():
        image_dir = project_dir / image_dir
    collector.download_aerial_images(record, str(image_dir))

    row = report_row_from_record(record, collector)
    denver_path = reports.resolve_path(project_dir, reports.normalize_text(row.get("Denver GIS Aerial Image File")))
    ai_model = args.ai_model or reports.default_ai_model(args.ai_provider)
    cache_dir = Path(args.analysis_cache_dir)
    if not cache_dir.is_absolute():
        cache_dir = project_dir / cache_dir
    analysis = reports.load_or_create_analysis(
        row,
        denver_path,
        None,
        cache_dir,
        args.use_ai,
        args.ai_provider,
        ai_model,
        args.allow_ai_fallback,
    )
    analysis = reports.apply_aerial_age_adjustment(row, analysis)

    output_dir = Path(args.output_dir)
    if not output_dir.is_absolute():
        output_dir = project_dir / output_dir
    output_path = output_dir / safe_report_name(reports.normalize_text(row.get("Parcel Number")) or parcel_key, reports.normalize_text(row.get("Address")))
    reports.render_report(row, analysis, denver_path, None, output_path)

    result = {
        "address": row.get("Address") or matched_address,
        "city": row.get("Building City"),
        "state": row.get("Building State"),
        "zip": row.get("Building ZIP"),
        "parcel": row.get("Parcel Number") or parcel_key,
        "match_score": round(match_score, 3),
        "lookup_source": lookup_source,
        "roof_squares": reports.roof_squares(row),
        "building_footprint_sqft": row.get("Building Footprint Sq Ft"),
        "aerial_image_file": str(denver_path) if denver_path else "",
        "analysis_source": reports.analysis_source_label(analysis),
        "report_path": str(output_path),
    }
    print(json.dumps(result))
    time.sleep(0.1)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(json.dumps({"error": str(exc)}))
        raise
