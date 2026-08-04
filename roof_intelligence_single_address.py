#!/usr/bin/env python3
"""Generate a Roof Intelligence report for one address using the GIS report project."""

from __future__ import annotations

import argparse
import difflib
from functools import lru_cache
import json
import os
import re
import sys
import time
import uuid
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from roof_intelligence_area_batch import (
    COUNTY_BOUNDS_PADDING_DEGREES,
    COUNTY_WGS84_BOUNDS,
)
from roof_report_naming import roof_report_pdf_filename


DEFAULT_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/PilotPoint IQ Roof Intelligence Report"
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

STREET_SEARCH_IGNORED = {
    "N", "S", "E", "W", "NE", "NW", "SE", "SW",
    "AVE", "BLVD", "CIR", "CT", "DR", "LN", "PKWY", "PL", "RD", "ST", "TER",
}


def normalize_address(value: object) -> str:
    # The parcel services generally store only the street line. Prefer the
    # street portion of a full mailing address so the city name does not lower
    # the match score for counties outside Denver.
    street_line = str(value or "").split(",", 1)[0]
    tokens = re.findall(r"[A-Z0-9]+", street_line.upper())
    normalized_tokens = []
    for token in tokens:
        if token in {"DENVER", "CO", "COLORADO"}:
            continue
        normalized_tokens.append(STREET_ALIASES.get(token, token))
    return " ".join(normalized_tokens)


def address_zip(value: object) -> str:
    matches = re.findall(r"\b(\d{5})(?:-\d{4})?\b", str(value or ""))
    return matches[-1] if matches else ""


def score_address(query: str, candidate: str) -> float:
    if not query or not candidate:
        return 0.0
    query_tokens = query.split()
    candidate_tokens = candidate.split()
    if query_tokens[0].isdigit() and candidate_tokens[0].isdigit():
        if query_tokens[0] != candidate_tokens[0]:
            return 0.0
    query_street_tokens = {
        token for token in query_tokens[1:]
        if token not in STREET_SEARCH_IGNORED and not token.isdigit()
    }
    candidate_street_tokens = {
        token for token in candidate_tokens[1:]
        if token not in STREET_SEARCH_IGNORED and not token.isdigit()
    }
    # Parcel queries use SQL LIKE expressions such as "%HIGH%", which also
    # return unrelated streets such as HIGGINS. A fuzzy whole-string score is
    # not sufficient evidence of a property match unless an actual street-name
    # token agrees.
    if (
        query_street_tokens
        and candidate_street_tokens
        and query_street_tokens.isdisjoint(candidate_street_tokens)
    ):
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
    street_tokens = [
        token for token in tokens[1:]
        if not token.isdigit() and STREET_ALIASES.get(token, token) not in STREET_SEARCH_IGNORED
    ]
    zip_code = address_zip(address)
    clauses = []

    fields = set(collector.collect_parcel_fields())
    address_fields = [
        field
        for field in (
            "SITUS_ADDRESS_LINE1", "Situs_Address", "PRPADDRESS", "concataddr1",
            "PropertyAddress", "SITUS_FULL_ADDRESS", "LOCADDRESS", "SITUS",
        )
        if field in fields
    ]
    if number.isdigit() and street_tokens:
        street = street_tokens[0]
        zip_clause = collector.parcel_zip_where({zip_code}) if zip_code else "1=1"
        for field in address_fields:
            primary_clause = f"{field} LIKE {sql_literal(number + '%' + street + '%')}"
            if zip_clause != "1=1":
                clauses.append(f"{primary_clause} AND {zip_clause}")
            clauses.append(primary_clause)
    for field in address_fields:
        if number.isdigit():
            clauses.append(f"{field} LIKE {sql_literal(number + '%')}")
        if street_tokens:
            clauses.append(f"{field} LIKE {sql_literal('%' + street_tokens[0] + '%')}")

    return list(dict.fromkeys(clauses))


@lru_cache(maxsize=64)
def geocode_address_location(address: str) -> dict:
    params = urlencode(
        {
            "SingleLine": address,
            "outFields": "Match_addr,Addr_type,Subregion,Postal,City,RegionAbbr",
            "maxLocations": 1,
            "f": "json",
        }
    )
    request = Request(
        "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/"
        "findAddressCandidates?" + params,
        headers={"User-Agent": "PCS-Roof-Intelligence/1.0"},
    )
    with urlopen(request, timeout=30) as response:
        payload = json.load(response)
    candidates = payload.get("candidates") or []
    if not candidates:
        raise RuntimeError(f"Unable to geocode address: {address}")
    candidate = candidates[0]
    location = candidate.get("location") or {}
    attributes = candidate.get("attributes") or {}
    return {
        "longitude": float(location["x"]),
        "latitude": float(location["y"]),
        "county": str(attributes.get("Subregion") or "").strip(),
        "postal_code": str(attributes.get("Postal") or "").strip()[:5],
        "city": str(attributes.get("City") or "").strip(),
        "state": str(attributes.get("RegionAbbr") or "").strip(),
    }


def geocode_address_point(address: str) -> tuple[float, float]:
    location = geocode_address_location(address)
    return location["longitude"], location["latitude"]


def normalized_county_key(value: object) -> str:
    normalized = re.sub(r"\bCOUNTY\b", "", str(value or "").upper())
    return "_".join(re.findall(r"[A-Z0-9]+", normalized)).lower()


def profile_contains_point(profile, longitude: float, latitude: float) -> bool:
    county_bounds = COUNTY_WGS84_BOUNDS.get(str(getattr(profile, "key", "")).strip().lower())
    if county_bounds is None:
        return False
    west, south, east, north = county_bounds
    padding = COUNTY_BOUNDS_PADDING_DEGREES
    return (
        west - padding <= longitude <= east + padding
        and south - padding <= latitude <= north + padding
    )


def shortlist_county_profiles(address: str, profiles: dict) -> list:
    """Prioritize counties supported by the full-address geocode and requested ZIP."""
    requested_zip = address_zip(address)
    if not requested_zip:
        return []
    try:
        location = geocode_address_location(address)
    except Exception:
        return []

    geocoded_zip = str(location.get("postal_code") or "")[:5]
    if geocoded_zip and geocoded_zip != requested_zip:
        return []

    prioritized_keys: list[str] = []
    geocoded_county = normalized_county_key(location.get("county"))
    if geocoded_county in profiles:
        prioritized_keys.append(geocoded_county)

    longitude = float(location["longitude"])
    latitude = float(location["latitude"])
    for key, profile in profiles.items():
        if key not in prioritized_keys and profile_contains_point(profile, longitude, latitude):
            prioritized_keys.append(key)
    return [profiles[key] for key in prioritized_keys]


def collect_spatial_parcel_for_address(address: str, collector) -> list[dict]:
    longitude, latitude = geocode_address_point(address)
    params = {
        "where": "1=1",
        "outFields": ",".join(collector.collect_parcel_fields()),
        "returnGeometry": "true",
        "geometry": f"{longitude},{latitude}",
        "geometryType": "esriGeometryPoint",
        "spatialRel": "esriSpatialRelIntersects",
        "inSR": "4326",
        "outSR": str(collector.PARCEL_CRS),
        "resultRecordCount": "10",
        "f": "json",
    }
    payload = collector.fetch_arcgis_json(collector.PARCELS_URL + "?" + urlencode(params))
    return list(payload.get("features") or [])


def parcel_records_from_features(features: list[dict], collector) -> list[dict]:
    parcels: list[dict] = []
    seen: set[str] = set()
    for feature in features:
        attrs = feature.get("attributes", {}) or {}
        attrs["parcel_shape_area"] = attrs.get("Shape__Area") or attrs.get("SHAPE__Area")
        attrs["parcel_geometry"] = collector.geometry_to_wkt(feature.get("geometry"))
        attrs["full_parcel_number"] = collector.parcel_join_key(attrs)
        key = attrs.get("full_parcel_number") or json.dumps(attrs, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        parcels.append(attrs)
    return parcels


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
    if not parcels:
        parcels.extend(
            parcel_records_from_features(
                collect_spatial_parcel_for_address(address, collector),
                collector,
            )
        )
    return parcels


def collect_live_parcel_by_id(parcel_id: str, collector) -> dict | None:
    """Load the exact parcel selected during map-area discovery."""
    requested = re.sub(r"[^A-Za-z0-9]", "", str(parcel_id or "")).upper()
    if not requested:
        return None
    fields = collector.collect_parcel_fields()
    identifier_fields = (
        "SCHEDNUM", "PARID", "ParcelNo", "PARCELNUMBER", "PARCELNUM", "PARCEL_SPN",
        "PARCEL_ID", "PARCELID", "PARCELNB", "PARCEL", "PIN", "SPN", "AIN", "Folio",
    )
    for field in identifier_fields:
        if field not in fields:
            continue
        try:
            page = collector.fetch_page(
                collector.PARCELS_URL,
                f"{field} = {sql_literal(parcel_id)}",
                0,
                fields,
                return_geometry=True,
            )
        except Exception:
            continue
        for feature in page.get("features", []):
            attrs = feature.get("attributes", {}) or {}
            actual = re.sub(r"[^A-Za-z0-9]", "", str(collector.parcel_join_key(attrs) or "")).upper()
            if actual != requested:
                continue
            attrs["parcel_shape_area"] = attrs.get("Shape__Area") or attrs.get("SHAPE__Area")
            attrs["parcel_geometry"] = collector.geometry_to_wkt(feature.get("geometry"))
            attrs["full_parcel_number"] = collector.parcel_join_key(attrs)
            return attrs
    return None


def find_parcel_for_address(
    address: str,
    parcels: list[dict],
    collector,
    *,
    allow_single_fallback: bool = True,
) -> tuple[dict, float, str]:
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

    if (
        allow_single_fallback
        and (not best_parcel or best_score < 0.74)
        and len(candidates) == 1
    ):
        return candidates[0], 0.9, str(address).split(",", 1)[0].strip()
    if not best_parcel or best_score < 0.74:
        raise RuntimeError(f"No parcel match found for address: {address}")
    return best_parcel, best_score, best_address


def find_parcel_live(address: str, collector, county_name: str) -> tuple[dict, float, str, str]:
    """Resolve an address only through the selected county's live parcel service."""
    live_parcels = collect_live_parcels_for_address(address, collector)
    if not live_parcels:
        raise RuntimeError(f"No live {county_name} parcel match found for address: {address}")
    try:
        parcel, score, matched_address = find_parcel_for_address(
            address,
            live_parcels,
            collector,
            allow_single_fallback=False,
        )
    except RuntimeError as text_match_error:
        spatial_parcels = parcel_records_from_features(
            collect_spatial_parcel_for_address(address, collector),
            collector,
        )
        if not spatial_parcels:
            raise text_match_error
        parcel, score, matched_address = find_parcel_for_address(address, spatial_parcels, collector)
    return parcel, score, matched_address, f"Live {county_name} parcel service"


def configure_collector_for_county(collector, profile) -> None:
    collector.BUILDINGS_URL = profile.building_url
    collector.PARCELS_URL = profile.parcel_url
    collector.IMAGERY_SOURCES = list(profile.imagery_sources)
    collector.BUILDING_SOURCE_KIND = getattr(profile, "building_source", "arcgis")
    collector.ACTIVE_COUNTY_NAME = profile.display_name.replace(" County", "")
    collector.ACTIVE_STATE = "CO"
    collector._COLLECT_PARCEL_FIELDS = None
    collector._COLLECT_BUILDING_FIELDS = None
    building_crs = getattr(profile, "building_crs", None)
    if building_crs is None:
        collector.init_crs_transformers(collector.BUILDINGS_URL, collector.PARCELS_URL)
    else:
        collector.init_crs_transformers(
            collector.BUILDINGS_URL,
            collector.PARCELS_URL,
            building_crs,
        )


def resolve_county_and_parcel(address: str, collector, profiles: dict) -> tuple[object, dict, float, str, str]:
    """Resolve the county by finding the address/ZIP in configured parcel services."""
    zip_code = address_zip(address)
    if not zip_code:
        raise RuntimeError("A five-digit ZIP code is required to determine the property county.")

    matches = []
    failures = []

    def query_profiles(selected_profiles) -> None:
        for profile in selected_profiles:
            try:
                configure_collector_for_county(collector, profile)
                parcels = collect_live_parcels_for_address(address, collector)
                if not parcels:
                    continue
                try:
                    parcel, score, matched_address = find_parcel_for_address(
                        address,
                        parcels,
                        collector,
                        allow_single_fallback=False,
                    )
                except RuntimeError as text_match_error:
                    spatial_parcels = parcel_records_from_features(
                        collect_spatial_parcel_for_address(address, collector),
                        collector,
                    )
                    if not spatial_parcels:
                        raise text_match_error
                    parcel, score, matched_address = find_parcel_for_address(
                        address,
                        spatial_parcels,
                        collector,
                    )
                parcel_zip = collector.parcel_zip(parcel)
                # Some supported county layers (notably Adams) leave the parcel
                # ZIP blank even for an exact situs-address match.
                if not parcel_zip or parcel_zip == zip_code:
                    matches.append((score, profile, parcel, matched_address))
            except Exception as exc:
                failures.append(f"{profile.display_name}: {exc}")

    shortlist = shortlist_county_profiles(address, profiles)
    query_profiles(shortlist)
    if not matches:
        shortlisted_keys = {profile.key for profile in shortlist}
        query_profiles(
            profile for profile in profiles.values()
            if profile.key not in shortlisted_keys
        )

    if not matches:
        detail = "; ".join(failures[-2:])
        suffix = f" ({detail})" if detail else ""
        raise RuntimeError(
            f"No supported county parcel match was found for ZIP {zip_code}. "
            "Supported counties are "
            + ", ".join(profile.display_name for profile in profiles.values())
            + f".{suffix}"
        )

    matches.sort(key=lambda item: item[0], reverse=True)
    score, profile, parcel, matched_address = matches[0]
    configure_collector_for_county(collector, profile)
    return profile, parcel, score, matched_address, f"Live {profile.display_name} parcel service"


def report_row_from_record(record: dict, collector) -> dict:
    collector.add_output_fields(record)
    return {label: record.get(field, "") for field, label in collector.OUTPUT_FIELDS}


def select_building_for_address(records: list[dict], address: str, collector) -> dict:
    if len(records) == 1:
        return records[0]
    longitude, latitude = geocode_address_point(address)
    point = collector.shape({"type": "Point", "coordinates": [longitude, latitude]})
    if collector.PARCEL_CRS != 4326:
        transformer = collector.Transformer.from_crs(4326, collector.PARCEL_CRS, always_xy=True)
        point = collector.transform(transformer.transform, point)

    ranked = []
    for record in records:
        polygon = collector.get_building_polygon(record)
        if polygon is None:
            continue
        ranked.append((0 if polygon.covers(point) else 1, polygon.distance(point), record))
    if not ranked:
        return max(records, key=lambda record: float(record.get("roof_squares", 0) or 0))
    ranked.sort(key=lambda item: (item[0], item[1]))
    return ranked[0][2]


def safe_report_name(address: str, city: str) -> str:
    return roof_report_pdf_filename(address, city)


def generate_roof_analysis(reports, row, aerial_path, cache_dir, args, ai_model):
    """Run PilotPoint's production single-call roof-reference workflow for PCS orders."""
    return reports.load_or_create_analysis(
        row,
        aerial_path,
        None,
        cache_dir,
        args.use_ai,
        args.ai_provider,
        ai_model,
        args.allow_ai_fallback,
        use_roof_references=True,
    )


def should_pause_for_footprint_review(
    validation: dict,
    footprint_source: str,
    allow_pending_review: bool,
) -> bool:
    return (
        validation.get("status") == "discrepancy"
        and footprint_source == "auto"
        and not allow_pending_review
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate one Roof Intelligence report by address.")
    parser.add_argument("--address", required=True, help="Property address")
    parser.add_argument(
        "--parcel-id",
        default="",
        help="Parcel selected by PCS; verifies address resolution and scopes assessor enrichment",
    )
    parser.add_argument(
        "--county",
        default="auto",
        help="County profile key, or 'auto' to resolve it from the address ZIP and parcel services",
    )
    parser.add_argument("--project-dir", default=str(DEFAULT_PROJECT_DIR), help="Roof Intelligence project directory")
    parser.add_argument("--output-dir", default="roof_intelligence_reports", help="Report output directory")
    parser.add_argument("--image-dir", default="aerial_images_single_address", help="Aerial image output directory")
    parser.add_argument("--analysis-cache-dir", default="roof_ai_analysis_cache", help="AI analysis cache directory")
    parser.add_argument("--use-ai", action="store_true", help="Run AI vision analysis")
    parser.add_argument("--ai-provider", choices=("openai", "gemini"), default="openai", help="AI provider")
    parser.add_argument("--ai-model", default=None, help="AI model override")
    parser.add_argument("--allow-ai-fallback", action="store_true", help="Generate fallback analysis if AI fails")
    parser.add_argument(
        "--footprint-source",
        choices=("auto", "supabase", "county"),
        default="auto",
        help="Audited source resolution for a previously reported footprint discrepancy",
    )
    parser.add_argument("--footprint-override-reason", default="")
    parser.add_argument(
        "--roof-area-override",
        type=float,
        default=None,
        help="Audited property-level square footage selected for future reports",
    )
    parser.add_argument(
        "--allow-pending-footprint-review",
        action="store_true",
        help="Generate a batch report with the primary footprint while preserving its review item",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_dir = Path(args.project_dir).expanduser().resolve()
    if not project_dir.exists():
        raise RuntimeError(f"Roof Intelligence project directory does not exist: {project_dir}")

    sys.path.insert(0, str(project_dir))
    import collect_county_buildings_with_parcels as collector
    import generate_roof_intelligence_reports as reports
    from assessor_detail import (
        enrich_report_row,
        fetch_assessor_details,
        normalize_identifier,
        validate_assessor_footprint,
    )
    from county_config import COUNTY_PROFILES, county_profile
    from building_footprint_store import mark_canonical_pending_review, save_canonical_footprint

    reports.load_env_file(project_dir / ".env")

    if args.county.strip().lower() == "auto":
        profile, parcel, match_score, matched_address, lookup_source = resolve_county_and_parcel(
            args.address,
            collector,
            COUNTY_PROFILES,
        )
    else:
        profile = county_profile(args.county)
        configure_collector_for_county(collector, profile)
        parcel = collect_live_parcel_by_id(args.parcel_id, collector) if args.parcel_id else None
        if parcel is not None:
            match_score = 1.0
            matched_address = collector.address_from_record(parcel) or args.address
            lookup_source = f"Live {profile.display_name} parcel service (PCS-selected parcel)"
        else:
            parcel, match_score, matched_address, lookup_source = find_parcel_live(
                args.address,
                collector,
                profile.display_name,
            )

    parcel_geometry = collector.get_parcel_bounds_in_building_crs([parcel])
    if not parcel_geometry:
        raise RuntimeError(f"Parcel geometry was not available for {matched_address}")

    primary_error = ""
    secondary_error = ""
    try:
        buildings = collector.collect_buildings(None, parcel_geometry)
    except Exception as exc:
        buildings = []
        primary_error = " ".join(str(exc).split())[:300]
    combined = collector.combine_data(buildings, [parcel]) if buildings else []
    parcel_key = collector.parcel_join_key(parcel)
    if args.parcel_id and normalize_identifier(args.parcel_id) != normalize_identifier(parcel_key):
        raise RuntimeError(
            f"PCS selected parcel {args.parcel_id}, but the address resolved to parcel {parcel_key}."
        )
    matched_records = [
        record for record in combined if collector.parcel_join_key(record) == parcel_key
    ]
    preliminary_record = (
        select_building_for_address(matched_records, args.address, collector)
        if matched_records else None
    )
    canonical_status = str((preliminary_record or {}).get("canonical_status") or "")
    canonical_is_final = canonical_status in {"validated", "single_source", "manually_resolved"}
    if canonical_is_final:
        secondary_buildings = []
    else:
        try:
            secondary_buildings = collector.collect_secondary_buildings(None, parcel_geometry)
        except Exception as exc:
            secondary_buildings = []
            secondary_error = " ".join(str(exc).split())[:300]
    secondary_combined = collector.combine_data(secondary_buildings, [parcel]) if secondary_buildings else []
    secondary_matches = [
        record for record in secondary_combined if collector.parcel_join_key(record) == parcel_key
    ]
    footprint_warnings: list[str] = []
    footprint_validation: dict = {}
    if canonical_is_final:
        record = preliminary_record
        footprint_validation = record.get("canonical_validation") or {
            "status": canonical_status,
            "difference_pct": record.get("difference_pct"),
        }
        footprint_warnings.append(
            f"The approved canonical footprint was used ({canonical_status.replace('_', ' ')})."
        )
    elif not matched_records and secondary_matches:
        record = select_building_for_address(secondary_matches, args.address, collector)
        footprint_validation = {
            "status": "county_only",
            "secondary_sqft": record.get("building_footprint_sqft"),
        }
        footprint_warnings.append(
            "Building footprint was available only from the county GIS footprint layer; "
            + (f"the Supabase lookup failed ({primary_error})." if primary_error else "the Supabase Microsoft footprint table had no matching structure.")
        )
    elif not matched_records:
        failure_details = " | ".join(value for value in (primary_error, secondary_error) if value)
        raise RuntimeError(
            f"No building footprint matched parcel {parcel_key} for {matched_address}"
            + (f": {failure_details}" if failure_details else "")
        )
    else:
        record = preliminary_record
        footprint_validation = collector.validate_building_footprint_sources(
            record, secondary_buildings
        )
        if footprint_validation.get("status") == "primary_only":
            footprint_warnings.append(
                "Building footprint was available only from the Supabase Microsoft footprint table; "
                + (f"the county lookup failed ({secondary_error})." if secondary_error else "the county GIS footprint layer had no overlapping structure.")
            )
    override_reason = " ".join(args.footprint_override_reason.split())
    if args.footprint_source != "auto" and not canonical_is_final:
        if len(override_reason) < 10:
            raise RuntimeError("A footprint discrepancy override requires a reason of at least 10 characters.")
        if args.footprint_source == "county":
            if not secondary_matches:
                raise RuntimeError("The approved county footprint is not available for this property.")
            record = select_building_for_address(secondary_matches, args.address, collector)
        elif not matched_records:
            raise RuntimeError("The approved Supabase footprint is not available for this property.")
        footprint_warnings.append(
            f"Footprint discrepancy was resolved using the {args.footprint_source} source. Reason: {override_reason}"
        )
    if not canonical_is_final:
        secondary_record = (
            select_building_for_address(secondary_matches, args.address, collector)
            if secondary_matches else None
        )
        canonical_record = save_canonical_footprint(
            profile.display_name,
            parcel_key,
            preliminary_record,
            secondary_record,
            footprint_validation,
            address=matched_address or args.address,
            selected_source=args.footprint_source,
            reason=override_reason,
            resolved_by="PCS local user" if args.footprint_source != "auto" else "system",
        )
        record.update({
            "canonical_id": canonical_record.get("canonical_id"),
            "canonical_status": canonical_record.get("canonical_status"),
            "canonical_validation": canonical_record.get("canonical_validation"),
        })
    if footprint_validation.get("status") == "discrepancy" and args.allow_pending_footprint_review:
        footprint_warnings.append(
            "The county GIS footprint exceeds the Microsoft footprint by "
            f"{footprint_validation['county_excess_pct']:.2f}%. The batch report uses the Microsoft "
            "footprint while the canonical footprint remains pending review."
        )
    elif should_pause_for_footprint_review(
        footprint_validation,
        args.footprint_source,
        args.allow_pending_footprint_review,
    ):
        raise RuntimeError(
            f"Building footprint discrepancy needs attention for {profile.display_name} parcel {parcel_key}: "
            "Supabase Microsoft footprint "
            f"{footprint_validation['primary_sqft']:.0f} sq ft versus county GIS footprint "
            f"{footprint_validation['secondary_sqft']:.0f} sq ft "
            f"(county is {footprint_validation['county_excess_pct']:.2f}% larger; 5% allowed). "
            f"Canonical footprint {record.get('canonical_id')} is pending review."
        )
    collector.add_aerial_image_fields(record)

    image_dir = Path(args.image_dir)
    if not image_dir.is_absolute():
        image_dir = project_dir / image_dir
    collector.download_aerial_images(record, str(image_dir))

    row = report_row_from_record(record, collector)
    # A parcel can contain multiple addressed buildings and publish only its
    # primary assessor situs (1704 High St is on the parcel whose assessor
    # label is 1720 N High St). Individual reports must retain the requested
    # building address after spatial parcel resolution.
    row["Address"] = str(args.address).split(",", 1)[0].strip()
    if not str(row.get("Building ZIP") or "").strip():
        row["Building ZIP"] = address_zip(args.address)
    assessor_result = None
    assessor_warnings: list[str] = list(footprint_warnings)
    assessor_footprint_validation: dict = {}
    try:
        assessor_result = fetch_assessor_details(profile.key, [parcel_key])
        enrich_report_row(row, assessor_result)
        assessor_warnings.extend(assessor_result.warnings)
        assessor_footprint_validation = validate_assessor_footprint(
            row.get("Building Footprint Sq Ft"), assessor_result.records
        )
        if should_pause_for_footprint_review(
            assessor_footprint_validation,
            args.footprint_source,
            args.allow_pending_footprint_review,
        ):
            if record.get("canonical_id"):
                mark_canonical_pending_review(record["canonical_id"], {
                    **assessor_footprint_validation,
                    "comparison": "county_assessor",
                })
            raise RuntimeError(
                f"Building footprint discrepancy needs attention for {profile.display_name} parcel {parcel_key}: selected footprint "
                f"{assessor_footprint_validation['primary_sqft']:.0f} sq ft versus explicit "
                f"county assessor footprint {assessor_footprint_validation['assessor_sqft']:.0f} sq ft "
                f"(county is {assessor_footprint_validation['county_excess_pct']:.2f}% larger; 5% allowed)."
            )
        if assessor_footprint_validation.get("status") == "discrepancy" and args.allow_pending_footprint_review:
            assessor_warnings.append(
                "The explicit county assessor area exceeds the Microsoft footprint by "
                f"{assessor_footprint_validation['county_excess_pct']:.2f}%. The batch report was generated "
                "while the footprint remains pending review."
            )
        if assessor_footprint_validation.get("status") == "discrepancy":
            assessor_warnings.append(
                f"The explicit assessor footprint discrepancy was approved using the {args.footprint_source} geometry. "
                f"Reason: {override_reason}"
            )
        if assessor_footprint_validation.get("status") == "not_comparable":
            assessor_warnings.append(assessor_footprint_validation["reason"])
        missing_sources = [
            key for key, count in assessor_result.source_counts.items() if count == 0
        ]
        if assessor_result.records and missing_sources:
            assessor_warnings.append(
                "Assessor information was available, but no matching record was returned by: "
                + ", ".join(missing_sources)
            )
    except ValueError as exc:
        # Denver retains its existing assessor path; the dedicated source map
        # currently covers the nine surrounding counties.
        if "Unsupported assessor county" not in str(exc):
            raise
    if args.roof_area_override is not None:
        if args.roof_area_override < 0:
            raise RuntimeError("The property square-footage override cannot be negative.")
        row["Building Footprint Sq Ft"] = float(args.roof_area_override)
        assessor_warnings.append(
            "The approved PCS property-level square-footage override was applied to this fresh report."
        )
    aerial_path = reports.resolve_path(
        project_dir,
        reports.normalize_text(row.get("Primary Aerial Image File")),
    )
    analysis_aerial_path = reports.resolve_path(
        project_dir,
        reports.aerial_analysis_image_file_value(row),
    )
    if not analysis_aerial_path or not analysis_aerial_path.is_file():
        raise RuntimeError(
            "The footprint-masked target roof image could not be created; "
            "roof analysis was stopped to prevent neighboring buildings from being included."
        )
    ai_model = args.ai_model or reports.default_ai_model(args.ai_provider)
    cache_dir = Path(args.analysis_cache_dir)
    if not cache_dir.is_absolute():
        cache_dir = project_dir / cache_dir
    analysis = generate_roof_analysis(
        reports,
        row,
        analysis_aerial_path,
        cache_dir,
        args,
        ai_model,
    )
    analysis = reports.apply_aerial_age_adjustment(row, analysis)

    output_dir = Path(args.output_dir)
    if not output_dir.is_absolute():
        output_dir = project_dir / output_dir
    output_path = output_dir / safe_report_name(
        reports.normalize_text(row.get("Address")),
        reports.normalize_text(row.get("Building City")),
    )
    reports.render_report(row, analysis, aerial_path, None, output_path)

    report_id = None
    report_snapshot = None
    from roof_intelligence_cutover_flags import load_cutover_flags

    if load_cutover_flags().editing_enabled:
        from roof_intelligence_snapshot import create_initial_snapshot

        report_id = str(uuid.uuid4())
        canonical_key = f"{profile.display_name.upper()}:{str(row.get('Parcel Number') or parcel_key).upper()}"
        report_snapshot = create_initial_snapshot(
            report_id=report_id,
            property_data={
                "canonical_key": canonical_key,
                "canonical_footprint_id": record.get("canonical_id"),
                "address": row.get("Address") or matched_address,
                "city": row.get("Building City"),
                "state": row.get("Building State"),
                "zip_code": row.get("Building ZIP"),
                "county": profile.display_name,
                "parcel_number": row.get("Parcel Number") or parcel_key,
                "latitude": record.get("latitude"),
                "longitude": record.get("longitude"),
            },
            report_fields=json.loads(json.dumps(row, default=str)) | {
                "roof_area_sqft": row.get("Building Footprint Sq Ft") or 0,
                "roof_squares": reports.roof_squares(row),
            },
            analysis=json.loads(json.dumps(analysis, default=str)),
            imagery={
                "source": str(row.get("Primary Aerial Source") or ""),
                "capture_date": row.get("Primary Aerial Photo Date") or None,
                "local_report_image_path": str(aerial_path) if aerial_path else "",
                "local_analysis_image_path": str(analysis_aerial_path),
                "target_mask_version": row.get("Primary Aerial Target Mask Version") or None,
                "target_mask_coverage": row.get("Primary Aerial Target Mask Coverage") or None,
                "limitations": list(assessor_warnings),
            },
            created_by=os.environ.get("ROOF_INTELLIGENCE_USER_KEY", "local-user"),
            persistent_square_footage_override=args.roof_area_override is not None,
        )

    result = {
        "report_id": report_id,
        "report_snapshot": report_snapshot,
        "address": row.get("Address") or matched_address,
        "city": row.get("Building City"),
        "state": row.get("Building State"),
        "zip": row.get("Building ZIP"),
        "county": profile.display_name,
        "county_profile": profile.key,
        "parcel": row.get("Parcel Number") or parcel_key,
        "match_score": round(match_score, 3),
        "lookup_source": lookup_source,
        "roof_squares": reports.roof_squares(row),
        "building_footprint_sqft": row.get("Building Footprint Sq Ft"),
        "square_footage_override_applied": args.roof_area_override is not None,
        "year_built": row.get("Year Built"),
        "effective_year_built": row.get("Effective Year Built"),
        "roof_type": analysis.get("roof_type"),
        "condition_score": analysis.get("condition_score") or analysis.get("overall_score"),
        "risk_level": analysis.get("risk_level"),
        "imagery_source": row.get("Primary Aerial Source"),
        "imagery_capture_date": row.get("Primary Aerial Photo Date"),
        "aerial_image_file": str(aerial_path) if aerial_path else "",
        "analysis_aerial_image_file": str(analysis_aerial_path),
        "analysis_source": reports.analysis_source_label(analysis),
        "assessor_record_count": len(assessor_result.records) if assessor_result else 0,
        "assessor_source_counts": assessor_result.source_counts if assessor_result else {},
        "assessor_detail_links": assessor_result.detail_links if assessor_result else [],
        "assessor_warnings": assessor_warnings,
        "footprint_validation": footprint_validation,
        "assessor_footprint_validation": assessor_footprint_validation,
        "footprint_resolution": {
            "selected_source": args.footprint_source,
            "reason": args.footprint_override_reason,
        } if args.footprint_source != "auto" else {},
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
