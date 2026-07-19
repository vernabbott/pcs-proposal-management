#!/usr/bin/env python3
"""Discover report candidates inside a WGS84 rectangle or radius across supported counties."""

from __future__ import annotations

import argparse
from datetime import date
import json
import math
from pathlib import Path
import sys
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


DEFAULT_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/PilotPoint IQ Roof Intelligence Report"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Find Roof Intelligence candidates inside a selected map area.")
    parser.add_argument("--project-dir", default=str(DEFAULT_PROJECT_DIR))
    parser.add_argument("--north", required=True, type=float)
    parser.add_argument("--south", required=True, type=float)
    parser.add_argument("--east", required=True, type=float)
    parser.add_argument("--west", required=True, type=float)
    parser.add_argument("--selection-type", choices=("rectangle", "radius"), default="rectangle")
    parser.add_argument("--center-lat", type=float)
    parser.add_argument("--center-lng", type=float)
    parser.add_argument("--radius-miles", type=float)
    parser.add_argument("--minimum-roof-size", type=float, default=10_000)
    parser.add_argument("--max-candidates", type=int, default=2_000)
    return parser.parse_args()


def _number(value: object) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number


def record_age(record: dict, current_year: int | None = None) -> int | None:
    value = record.get("effective_year_built") or record.get("year_built")
    year = _number(value)
    if year is None or year <= 0:
        return None
    return max(0, (current_year or date.today().year) - int(year))


def reverse_geocode_record(record: dict, collector) -> dict[str, str]:
    polygon = collector.get_parcel_polygon(record)
    if polygon is None:
        return {}
    point = polygon.centroid
    if collector.PARCEL_CRS != 4326:
        transformer = collector.Transformer.from_crs(collector.PARCEL_CRS, 4326, always_xy=True)
        longitude, latitude = transformer.transform(point.x, point.y)
    else:
        longitude, latitude = point.x, point.y
    params = urlencode(
        {
            "location": f"{longitude},{latitude}",
            "outSR": 4326,
            "featureTypes": "StreetAddress,PointAddress,Parcel",
            "f": "json",
        }
    )
    request = Request(
        "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/reverseGeocode?"
        + params,
        headers={"User-Agent": "PCS-Roof-Intelligence/1.0"},
    )
    try:
        with urlopen(request, timeout=30) as response:
            payload = json.load(response)
    except Exception:
        return {}
    address = payload.get("address") or {}
    return {
        "street": str(address.get("Address") or "").strip(),
        "city": str(address.get("City") or address.get("Subregion") or "").strip(),
        "state": str(address.get("RegionAbbr") or "CO").strip(),
        "zip": str(address.get("Postal") or "").strip()[:5],
    }


def candidate_from_record(record: dict, profile, collector, minimum_roof_size: float) -> dict | None:
    collector.add_output_fields(record)
    roof_area = _number(record.get("roof_area_est") or record.get("building_footprint_sqft")) or 0.0
    if roof_area < minimum_roof_size:
        return None
    age = record_age(record)
    street = str(record.get("property_address") or collector.address_from_record(record) or "").strip()
    zip_code = str(record.get("property_zip") or collector.parcel_zip(record) or "").strip()[:5]
    location = {}
    if not street or not zip_code:
        location = reverse_geocode_record(record, collector)
        street = street or location.get("street", "")
        zip_code = zip_code or location.get("zip", "")
    if not street or not zip_code:
        return None
    city = str(record.get("property_city") or location.get("city") or "").strip()
    state = str(record.get("property_state") or location.get("state") or "CO").strip() or "CO"
    locality = ", ".join(part for part in (city, state) if part)
    full_address = f"{street}, {locality} {zip_code}" if locality else f"{street}, {zip_code}"
    parcel = str(record.get("parcel_number") or collector.parcel_join_key(record) or "").strip()
    if not parcel:
        return None
    return {
        "candidate_key": f"{profile.key}:{parcel}",
        "address": full_address,
        "street_address": street,
        "city": city,
        "state": state,
        "zip": zip_code,
        "county": profile.display_name,
        "county_profile": profile.key,
        "parcel": parcel,
        "roof_area_sqft": round(roof_area, 1),
        "roof_squares": _number(record.get("roof_squares")),
        "year_built": record.get("year_built") or None,
        "effective_year_built": record.get("effective_year_built") or None,
        "age_estimate_years": age,
        "footprint_source": record.get("footprint_source") or "supabase",
        "footprint_validation": record.get("footprint_validation") or {},
    }


def circle_polygon(center: dict[str, float], radius_miles: float, point_count: int = 72) -> dict:
    angular_radius = radius_miles / 3958.7613
    latitude = math.radians(center["lat"])
    longitude = math.radians(center["lng"])
    ring = []
    for index in range(point_count):
        bearing = 2 * math.pi * index / point_count
        destination_latitude = math.asin(
            math.sin(latitude) * math.cos(angular_radius)
            + math.cos(latitude) * math.sin(angular_radius) * math.cos(bearing)
        )
        destination_longitude = longitude + math.atan2(
            math.sin(bearing) * math.sin(angular_radius) * math.cos(latitude),
            math.cos(angular_radius) - math.sin(latitude) * math.sin(destination_latitude),
        )
        ring.append([math.degrees(destination_longitude), math.degrees(destination_latitude)])
    ring.append(ring[0])
    return {"rings": [ring], "spatialReference": {"wkid": 4326}}


def spatial_query(bounds: dict[str, float], selection: dict | None = None) -> tuple[str, str]:
    if selection and selection.get("selection_type") == "radius":
        geometry = circle_polygon(selection["center"], float(selection["radius_miles"]))
        return json.dumps(geometry, separators=(",", ":")), "esriGeometryPolygon"
    envelope = f'{bounds["west"]},{bounds["south"]},{bounds["east"]},{bounds["north"]}'
    return envelope, "esriGeometryEnvelope"


def fetch_arcgis_post(collector, url: str, params: dict) -> dict:
    attempts = int(getattr(collector, "REQUEST_ATTEMPTS", 3))
    timeout = float(getattr(collector, "REQUEST_TIMEOUT", 60))
    request_data = urlencode(params).encode("utf-8")
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        request = Request(
            url,
            data=request_data,
            headers={
                "User-Agent": "Python Building Collector",
                "Content-Type": "application/x-www-form-urlencoded",
            },
            method="POST",
        )
        try:
            with urlopen(request, timeout=timeout) as response:
                payload = json.load(response)
            if payload.get("error"):
                details = payload["error"]
                message = details.get("message") if isinstance(details, dict) else str(details)
                raise RuntimeError(f"ArcGIS query error: {message or 'unknown error'}")
            return payload
        except HTTPError as exc:
            last_error = exc
            if exc.code < 500 or attempt == attempts:
                raise RuntimeError(f"HTTP error: {exc.code} {exc.reason}") from exc
        except (TimeoutError, URLError, OSError) as exc:
            last_error = exc
            if attempt == attempts:
                raise RuntimeError(f"URL error after {attempts} attempts: {exc}") from exc
        time.sleep(2 * attempt)
    raise RuntimeError(f"URL error: {last_error}")


def fetch_parcels_in_bounds(
    collector,
    bounds: dict[str, float],
    max_records: int = 10_000,
    selection: dict | None = None,
) -> list[dict]:
    fields = collector.collect_parcel_fields(collector.PARCELS_URL)
    metadata_fields = collector.available_layer_fields(collector.PARCELS_URL)
    object_id = next(
        (field.get("name") for field in metadata_fields if field.get("type") == "esriFieldTypeOID"),
        "OBJECTID",
    )
    offset = 0
    parcels: list[dict] = []
    geometry, geometry_type = spatial_query(bounds, selection)
    while len(parcels) < max_records:
        params = {
            "where": "1=1",
            "outFields": ",".join(fields),
            "returnGeometry": "true",
            "geometry": geometry,
            "geometryType": geometry_type,
            "spatialRel": "esriSpatialRelIntersects",
            "inSR": "4326",
            "outSR": str(collector.PARCEL_CRS),
            "f": "json",
            "resultOffset": offset,
            "resultRecordCount": collector.PAGE_SIZE,
            "orderByFields": f"{object_id} ASC",
        }
        if geometry_type == "esriGeometryPolygon":
            page = fetch_arcgis_post(collector, collector.PARCELS_URL, params)
        else:
            page = collector.fetch_arcgis_json(collector.PARCELS_URL + "?" + urlencode(params))
        features = page.get("features", [])
        if not features:
            break
        for feature in features:
            attrs = feature.get("attributes", {}) or {}
            attrs["parcel_shape_area"] = attrs.get("Shape__Area") or attrs.get("SHAPE__Area")
            attrs["parcel_geometry"] = collector.geometry_to_wkt(feature.get("geometry"))
            attrs["full_parcel_number"] = collector.parcel_join_key(attrs)
            parcels.append(attrs)
            if len(parcels) >= max_records:
                break
        offset += len(features)
        if len(features) < collector.PAGE_SIZE:
            break
    return parcels


def discover_candidates(
    project_dir: Path,
    bounds: dict[str, float],
    minimum_roof_size: float,
    max_candidates: int,
    selection: dict | None = None,
) -> tuple[list[dict], list[str]]:
    sys.path.insert(0, str(project_dir))
    import collect_denver_buildings_with_parcels as collector
    from county_config import COUNTY_PROFILES
    from roof_intelligence_single_address import configure_collector_for_county

    by_key: dict[str, dict] = {}
    warnings: list[str] = []
    discovered_parcel_count = 0
    for profile in COUNTY_PROFILES.values():
        try:
            configure_collector_for_county(collector, profile)
            parcels = fetch_parcels_in_bounds(collector, bounds, selection=selection)
            if not parcels:
                continue
            discovered_parcel_count += len(parcels)
            parcel_geometry = collector.get_parcel_bounds_in_building_crs(parcels)
            if not parcel_geometry:
                warnings.append(f"{profile.display_name}: parcel geometry was unavailable")
                continue
            primary_error = ""
            secondary_error = ""
            try:
                buildings = collector.collect_buildings(None, parcel_geometry)
            except Exception as exc:
                buildings = []
                primary_error = " ".join(str(exc).split())[:200]
            try:
                secondary_buildings = collector.collect_secondary_buildings(None, parcel_geometry)
            except Exception as exc:
                secondary_buildings = []
                secondary_error = " ".join(str(exc).split())[:200]
            parcel_keys = {collector.parcel_join_key(parcel) for parcel in parcels}
            records = collector.combine_data(buildings, parcels) if buildings else []
            has_primary_matches = any(
                collector.parcel_join_key(record) in parcel_keys for record in records
            )
            if not has_primary_matches and secondary_buildings:
                records = collector.combine_data(secondary_buildings, parcels)
                buildings = []
                warnings.append(
                    f"{profile.display_name}: using county footprints because Supabase "
                    + (f"failed ({primary_error})" if primary_error else "returned no selected footprints")
                )
            elif primary_error and not secondary_buildings:
                warnings.append(f"{profile.display_name}: Supabase failed ({primary_error})")
            if secondary_error:
                warnings.append(f"{profile.display_name}: county footprint lookup failed ({secondary_error})")
            for record in records:
                if collector.parcel_join_key(record) not in parcel_keys:
                    continue
                if buildings:
                    canonical_status = str(record.get("canonical_status") or "")
                    if canonical_status in {"validated", "single_source", "manually_resolved"}:
                        record["footprint_source"] = "canonical"
                        record["footprint_validation"] = record.get("canonical_validation") or {
                            "status": canonical_status
                        }
                    else:
                        record["footprint_source"] = "supabase"
                        record["footprint_validation"] = collector.validate_building_footprint_sources(
                            record, secondary_buildings
                        )
                else:
                    record["footprint_source"] = "county"
                    record["footprint_validation"] = {"status": "county_only"}
                candidate = candidate_from_record(record, profile, collector, minimum_roof_size)
                if not candidate:
                    continue
                previous = by_key.get(candidate["candidate_key"])
                if previous is None or candidate["roof_area_sqft"] > previous["roof_area_sqft"]:
                    by_key[candidate["candidate_key"]] = candidate
        except Exception as exc:
            warnings.append(f"{profile.display_name}: {' '.join(str(exc).split())[:240]}")

    if not by_key and not discovered_parcel_count and warnings:
        raise RuntimeError("County parcel discovery failed: " + " | ".join(warnings))

    candidates = sorted(
        by_key.values(),
        key=lambda candidate: (-float(candidate.get("roof_area_sqft") or 0), candidate.get("address") or ""),
    )[:max_candidates]
    return candidates, warnings


def main() -> int:
    args = parse_args()
    project_dir = Path(args.project_dir).expanduser().resolve()
    if not project_dir.is_dir():
        raise RuntimeError(f"Roof Intelligence project directory does not exist: {project_dir}")
    if not (args.south < args.north and args.west < args.east):
        raise ValueError("The selected map bounds are invalid.")
    selection = {"selection_type": args.selection_type}
    if args.selection_type == "radius":
        if args.center_lat is None or args.center_lng is None or not args.radius_miles or args.radius_miles <= 0:
            raise ValueError("A valid radius center and distance are required.")
        selection.update({
            "center": {"lat": args.center_lat, "lng": args.center_lng},
            "radius_miles": args.radius_miles,
        })
    candidates, warnings = discover_candidates(
        project_dir,
        {"north": args.north, "south": args.south, "east": args.east, "west": args.west},
        args.minimum_roof_size,
        args.max_candidates,
        selection,
    )
    print(json.dumps({"candidates": candidates, "warnings": warnings}))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(json.dumps({"error": " ".join(str(exc).split())[:500]}))
        raise
