#!/usr/bin/env python3
"""
Mega Pipes Quotation System
Module: tools/build_pin_master.py

Purpose:
  Ingests the raw All-India Post Office Directory CSV (165,627 rows)
  and builds a clean, authoritative 1-record-per-PIN normalized master table (19,586 PINs).

Approved Master CSV Columns:
  PIN, Latitude, Longitude, OfficeName, OfficeType, District, State,
  RecordCount, ValidCoordCount, QualityStatus, SpreadKM, SelectionRule

Representative-Coordinate Rules (3-Tier Hierarchy):
  1. Tier 1: Authoritative Postal Hub (HO/PO):
     If a Head Post Office (HO) or Sub-Office (PO) has valid coordinates within India
     (Lat 6.0N-38.0N, Lon 68.0E-98.0E), it is chosen as the primary representative coordinate.
     - Confirmed by branch offices within 25 km: HIGH confidence (HO_PO_HUB_CONFIRMED).
     - Natural delivery variation (25-50 km): MEDIUM confidence (HO_PO_HUB_WITH_VARIATION).
     - Single office records: MEDIUM confidence (SINGLE_OFFICE).
  2. Tier 2: Branch Office Cluster Medoid:
     If only Branch Offices (BO) have coordinates, extreme geocoding outliers (>100 km from median)
     are rejected, and the spatial medoid (point minimizing sum of distances to other inliers) is selected.
     - Inlier spread <= 15 km: HIGH confidence (BO_MEDOID_CONVERGENT).
     - Inlier spread 15-35 km: MEDIUM confidence (BO_MEDOID_NORMAL).
     - Outliers rejected or inlier spread > 35 km: LOW_FLAGGED confidence (BO_MEDOID_OUTLIERS_REJECTED).
  3. Tier 3: Unresolved:
     If all offices have 'NA' or missing coordinates, Latitude and Longitude are stored as null/empty.
     Coordinates are NEVER fabricated. Cataloged in DATA/unresolved_pins.csv for administrative review.
"""

import csv
import math
import os
import sys
import time
import shutil
from collections import defaultdict, Counter
from pathlib import Path
from typing import Tuple, List, Dict, Any

# Project paths
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "DATA"
DEFAULT_RAW_CSV = Path(r"C:\Users\Administrator\Downloads\5c2f62fe-5afa-4119-a499-fec9d604d5bd.csv")
OUTPUT_MASTER_CSV = DATA_DIR / "india_pincode_master.csv"
OUTPUT_UNRESOLVED_CSV = DATA_DIR / "unresolved_pins.csv"

# Geographic bounding box for India (Mainland + Islands)
LAT_MIN, LAT_MAX = 6.0, 38.0
LON_MIN, LON_MAX = 68.0, 98.0

# Threshold parameters (approved baseline)
HUB_SPREAD_MAX_KM = 50.0       # Maximum acceptable spread from hub to inlier BOs
HIGH_AGREEMENT_KM = 25.0       # Maximum spread for HIGH confidence hub confirmation
MEDOID_CONVERGENCE_KM = 15.0   # Maximum spread for HIGH confidence BO medoid
BO_NORMAL_SPREAD_KM = 35.0     # Maximum spread for MEDIUM confidence BO medoid
OUTLIER_THRESHOLD_KM = 100.0   # Distance from median to qualify as extreme geocoding outlier


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Calculate the great-circle distance between two points in kilometres."""
    if lat1 == lat2 and lon1 == lon2:
        return 0.0
    R = 6371.0  # Earth's mean radius in km
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2.0) ** 2 +
         math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2.0) ** 2)
    a = min(1.0, max(0.0, a))
    c = 2.0 * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))
    return round(R * c, 3)


def build_pin_master(raw_csv_path: Path = DEFAULT_RAW_CSV) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Ingest raw India Post CSV, normalize to 1 record per PIN,
    and atomically write output master and unresolved CSV files.
    """
    if not raw_csv_path.exists():
        raise FileNotFoundError(f"Raw India Post CSV not found at: {raw_csv_path}")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    
    print(f"Reading raw India Post CSV from: {raw_csv_path} ...", flush=True)
    pins = defaultdict(list)
    total_raw_rows = 0
    missing_coords_count = 0
    oob_coords_count = 0
    float_parse_errors = 0
    valid_coords_count = 0

    with open(raw_csv_path, mode="r", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            total_raw_rows += 1
            pin_raw = row["pincode"].strip()
            # Preserve strictly as 6-character string with leading zeros
            pin_clean = pin_raw.zfill(6) if pin_raw.isdigit() else pin_raw
            
            lat_raw = row.get("latitude", "").strip()
            lon_raw = row.get("longitude", "").strip()
            has_coord = False
            lat_val, lon_val = None, None
            
            if lat_raw.upper() in ("NA", "", "NULL", "NONE") or lon_raw.upper() in ("NA", "", "NULL", "NONE"):
                missing_coords_count += 1
            else:
                try:
                    lt = float(lat_raw)
                    ln = float(lon_raw)
                    if LAT_MIN <= lt <= LAT_MAX and LON_MIN <= ln <= LON_MAX:
                        lat_val, lon_val = lt, ln
                        has_coord = True
                        valid_coords_count += 1
                    else:
                        oob_coords_count += 1
                except ValueError:
                    float_parse_errors += 1

            pins[pin_clean].append({
                "office_name": row.get("officename", "").strip(),
                "office_type": row.get("officetype", "").strip().upper(),
                "delivery": row.get("delivery", "").strip(),
                "district": row.get("district", "").strip(),
                "state": row.get("statename", "").strip(),
                "lat": lat_val,
                "lon": lon_val,
                "has_coord": has_coord
            })

    unique_pins = len(pins)
    print(f"Loaded {total_raw_rows:,} records across {unique_pins:,} unique PINs in {time.time()-t0:.2f}s", flush=True)
    print(f"  Valid coordinates: {valid_coords_count:,}")
    print(f"  Missing (NA): {missing_coords_count:,}")
    print(f"  Out of bounds: {oob_coords_count:,}")
    print(f"  Float parse errors: {float_parse_errors}")

    normalized_records = []
    unresolved_records = []
    quality_counts = Counter()
    rule_counts = Counter()

    for pin in sorted(pins.keys()):
        offices = pins[pin]
        total_offices = len(offices)
        valid_offices = [o for o in offices if o["has_coord"]]
        valid_count = len(valid_offices)

        # Primary District and State (modal choice among offices)
        dist_counts = Counter(o["district"] for o in offices if o["district"])
        state_counts = Counter(o["state"] for o in offices if o["state"])
        pref_district = dist_counts.most_common(1)[0][0] if dist_counts else "UNKNOWN"
        pref_state = state_counts.most_common(1)[0][0] if state_counts else "UNKNOWN"

        if valid_count == 0:
            quality = "UNRESOLVED"
            rule = "NO_COORDINATES"
            rep = offices[0]
            quality_counts[quality] += 1
            rule_counts[rule] += 1
            
            rec = {
                "PIN": pin,
                "Latitude": "",
                "Longitude": "",
                "OfficeName": rep["office_name"],
                "OfficeType": rep["office_type"],
                "District": pref_district,
                "State": pref_state,
                "RecordCount": total_offices,
                "ValidCoordCount": 0,
                "QualityStatus": quality,
                "SpreadKM": "0.0",
                "SelectionRule": rule
            }
            normalized_records.append(rec)
            unresolved_records.append(rec)
            continue

        # Check if Head Office (HO) or Sub-Office (PO) has valid coordinates
        ho_offices = [o for o in valid_offices if o["office_type"] == "HO"]
        po_offices = [o for o in valid_offices if o["office_type"] == "PO"]
        ho_po_valid = ho_offices + po_offices

        if len(valid_offices) == 1:
            sel = valid_offices[0]
            quality = "MEDIUM"
            rule = "SINGLE_OFFICE"
            spread = 0.0
        elif ho_po_valid:
            sel = ho_po_valid[0]  # Primary authoritative hub (HO preferred, then PO)
            # Check spread of other valid offices to this hub
            diffs = [haversine(sel["lat"], sel["lon"], o["lat"], o["lon"]) for o in valid_offices]
            inliers = [d for d in diffs if d <= HUB_SPREAD_MAX_KM]
            max_inlier_d = max(inliers) if inliers else 0.0
            outliers = len(diffs) - len(inliers)
            spread = max_inlier_d

            if outliers > 0 or max_inlier_d > HIGH_AGREEMENT_KM:
                quality = "MEDIUM"
                rule = "HO_PO_HUB_WITH_VARIATION"
            else:
                quality = "HIGH"
                rule = "HO_PO_HUB_CONFIRMED"
        else:
            # Only Branch Offices (BO) have coordinates. Medoid selection with outlier rejection.
            lats = sorted([o["lat"] for o in valid_offices])
            lons = sorted([o["lon"] for o in valid_offices])
            med_lat = lats[len(lats) // 2]
            med_lon = lons[len(lons) // 2]

            inliers = [o for o in valid_offices if haversine(o["lat"], o["lon"], med_lat, med_lon) <= OUTLIER_THRESHOLD_KM]
            if not inliers:
                inliers = valid_offices

            best_sum = float("inf")
            best_o = inliers[0]
            for o1 in inliers:
                s = sum(haversine(o1["lat"], o1["lon"], o2["lat"], o2["lon"]) for o2 in inliers)
                if s < best_sum:
                    best_sum = s
                    best_o = o1
            sel = best_o

            max_d = 0.0
            for i in range(len(inliers)):
                for j in range(i + 1, len(inliers)):
                    d = haversine(inliers[i]["lat"], inliers[i]["lon"], inliers[j]["lat"], inliers[j]["lon"])
                    if d > max_d:
                        max_d = d
            spread = max_d
            outliers = len(valid_offices) - len(inliers)

            if outliers > 0 or max_d > BO_NORMAL_SPREAD_KM:
                quality = "LOW_FLAGGED"
                rule = "BO_MEDOID_OUTLIERS_REJECTED"
            elif max_d <= MEDOID_CONVERGENCE_KM:
                quality = "HIGH"
                rule = "BO_MEDOID_CONVERGENT"
            else:
                quality = "MEDIUM"
                rule = "BO_MEDOID_NORMAL"

        quality_counts[quality] += 1
        rule_counts[rule] += 1

        normalized_records.append({
            "PIN": pin,
            "Latitude": f"{sel['lat']:.6f}",
            "Longitude": f"{sel['lon']:.6f}",
            "OfficeName": sel["office_name"],
            "OfficeType": sel["office_type"],
            "District": pref_district,
            "State": pref_state,
            "RecordCount": total_offices,
            "ValidCoordCount": valid_count,
            "QualityStatus": quality,
            "SpreadKM": f"{spread:.1f}",
            "SelectionRule": rule
        })

    # Backup existing files if present
    if OUTPUT_MASTER_CSV.exists():
        backup_master = DATA_DIR / "india_pincode_master.csv.bak"
        shutil.copy2(OUTPUT_MASTER_CSV, backup_master)

    if OUTPUT_UNRESOLVED_CSV.exists():
        backup_unresolved = DATA_DIR / "unresolved_pins.csv.bak"
        shutil.copy2(OUTPUT_UNRESOLVED_CSV, backup_unresolved)

    # Standard Output Fieldnames (approved schema)
    fieldnames = [
        "PIN", "Latitude", "Longitude", "OfficeName", "OfficeType",
        "District", "State", "RecordCount", "ValidCoordCount",
        "QualityStatus", "SpreadKM", "SelectionRule"
    ]

    # Atomic write for master CSV
    tmp_master = OUTPUT_MASTER_CSV.with_suffix(".tmp")
    with open(tmp_master, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(normalized_records)
    os.replace(tmp_master, OUTPUT_MASTER_CSV)

    # Atomic write for unresolved CSV
    tmp_unresolved = OUTPUT_UNRESOLVED_CSV.with_suffix(".tmp")
    with open(tmp_unresolved, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(unresolved_records)
    os.replace(tmp_unresolved, OUTPUT_UNRESOLVED_CSV)

    build_time = time.time() - t0
    print(f"\nSuccessfully generated normalized master ({len(normalized_records):,} rows) in {build_time:.2f}s:")
    print(f"  Master: {OUTPUT_MASTER_CSV} ({OUTPUT_MASTER_CSV.stat().st_size:,} bytes)")
    print(f"  Unresolved: {OUTPUT_UNRESOLVED_CSV} ({OUTPUT_UNRESOLVED_CSV.stat().st_size:,} bytes, {len(unresolved_records)} PINs)")
    print("\nQuality Breakdown:")
    for k in ["HIGH", "MEDIUM", "LOW_FLAGGED", "UNRESOLVED"]:
        cnt = quality_counts[k]
        print(f"  {k}: {cnt:,} ({cnt/unique_pins*100:.2f}%)")
    print("\nSelection Rule Breakdown:")
    for rule, count in rule_counts.most_common():
        print(f"  {rule}: {count:,}")

    return normalized_records, unresolved_records


if __name__ == "__main__":
    raw_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_RAW_CSV
    build_pin_master(raw_path)
