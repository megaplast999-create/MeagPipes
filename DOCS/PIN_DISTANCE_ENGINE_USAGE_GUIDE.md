# Mega Pipes — India PIN-Code Distance & Marketing Territory Engine
## Developer & User Guide — Phase 1

This guide documents the usage of the normalized India PIN-code dataset and the Python distance calculation and marketing territory discovery engine implemented in `tools/pin_distance_engine.py`.

---

## 1. Quick Start

```python
from tools.pin_distance_engine import (
    validate_pin,
    get_pin_details,
    straight_line_distance_km,
    get_marketing_territory,
    get_distance_cache_stats,
    export_territory_to_csv,
    export_territory_to_json,
    generate_territory_text_report
)

# 1. Validate a PIN code (preserves 6-digit string with leading zeros)
clean_pin = validate_pin("482001")
print(clean_pin)  # Output: "482001"

# 2. Retrieve location details
details = get_pin_details("482001")
print(details["OfficeName"], details["District"], details["State"], details["QualityStatus"])
# Output: Jabalpur H.O JABALPUR MADHYA PRADESH MEDIUM

# 3. Calculate straight-line distance in kilometres
dist = straight_line_distance_km("482001", "410203")  # Jabalpur to Khopoli
print(f"Distance: {dist:.3f} km")
# Output: Distance: 840.428 km

# 4. Discover marketing territory within a radius (e.g. 50 km)
territory = get_marketing_territory("482001", radius_km=50.0)
print(f"Discovered {territory['total_territory_pins']} delivery PINs within 50 km.")

# 5. Export to CSV or JSON
export_territory_to_csv(territory, "EXPORTS/jabalpur_50km_territory.csv")
export_territory_to_json(territory, "EXPORTS/jabalpur_50km_territory.json")

# 6. Generate human-readable report
report_text = generate_territory_text_report(territory)
print(report_text[:500])
```

---

## 2. Public API Reference

### `validate_pin(pin: Any) -> str`
- **Description:** Validates and normalizes an Indian PIN code string.
- **Rules:** Must be a 6-digit numeric string. Rejects empty strings, strings with letters or symbols, and lengths $\ne 6$. Preserves leading zeros.
- **Raises:** `ValueError` on invalid format.

### `get_pin_details(pin: Any) -> Optional[Dict[str, Any]]`
- **Description:** Retrieves metadata for a PIN code from the normalized master.
- **Returns:** Dictionary with keys: `PIN`, `Latitude`, `Longitude`, `OfficeName`, `OfficeType`, `District`, `State`, `RecordCount`, `ValidCoordCount`, `QualityStatus`, `SpreadKM`, `SelectionRule`.
- **Note:** If the PIN is unresolved, `Latitude` and `Longitude` are `None`.

### `straight_line_distance_km(origin_pin: Any, destination_pin: Any) -> float`
- **Description:** Calculates straight-line great-circle distance between two PINs using the Haversine formula ($R = 6,371.0\text{ km}$).
- **Guarantees:**
  - Identical PINs return strictly `0.000 km`.
  - Cached automatically in persistent SQLite cache `DATA/pin_distance_cache.sqlite`.
  - Undirected normalization: `distance(A, B) == distance(B, A)`.
- **Raises:** `ValueError` if either PIN is invalid, missing, or has `UNRESOLVED` coordinates.

### `get_marketing_territory(anchor_pin: Any, radius_km: float = 50.0) -> Dict[str, Any]`
- **Description:** Queries the 3D Spatial KD-Tree index and calculates exact Haversine distances to return all postal zones within `0 <= distance <= radius_km`.
- **Parameters:**
  - `anchor_pin`: 6-digit Indian PIN code.
  - `radius_km`: Float radius in km (default `50.0`). Configurable for any distance (e.g. `25.0`, `50.0`, `100.0`).
- **Returns:**
  - `anchor`: Anchor PIN metadata.
  - `radius_km`: Search radius applied.
  - `metric`: Strictly `"STRAIGHT_LINE_KM"`.
  - `total_territory_pins`: Count of matching delivery postal zones.
  - `pins`: List of surrounding PIN dicts sorted by straight-line distance ascending.
  - `district_summary`: PIN count distribution by administrative district.
  - `state_summary`: PIN count distribution by state.
  - `quality_summary`: Quality breakdown (`HIGH`, `MEDIUM`, `LOW_FLAGGED`).
  - `radial_clusters`: Segmented into `inner_0_15km`, `middle_15_30km`, `outer_30_50km`.
  - `coverage_gaps`: List of unresolved PINs in surrounding districts.
  - `screening_disclaimer`: Mandatory geographical screening notice.
  - `road_distance_note`: Policy statement separating straight-line marketing from commercial road freight.

### `get_distance_cache_stats() -> Dict[str, Any]`
- **Description:** Returns cache statistics including `total_cached_pairs`, `hits`, `misses`, `hit_ratio_pct`, and `db_path`.

---

## 3. Sample Marketing Territory Output (Jabalpur `482001`, 50 km)

```text
================================================================================
MEGA PIPES — MARKETING TERRITORY DISCOVERY REPORT
================================================================================
Anchor PIN:             482001 (Jabalpur H.O, HO)
Anchor Location:        JABALPUR, MADHYA PRADESH (Lat: 23.161490, Lon: 79.954832)
Anchor Quality Status:  MEDIUM (Selection Rule: HO_PO_HUB_WITH_VARIATION)
Exploratory Radius:     50.0 km (Straight-Line Great-Circle)
Total Eligible PINs:    32 unique delivery postal zones
--------------------------------------------------------------------------------
DISTRICT DISTRIBUTION:
  - JABALPUR: 19 PINs (59.4%)
  - KATNI: 3 PINs (9.4%)
  - NARSINGHPUR: 2 PINs (6.2%)
  - MANDLA: 2 PINs (6.2%)
  - REWA: 1 PINs (3.1%)
  - NA: 1 PINs (3.1%)
  - DHAR: 1 PINs (3.1%)
  - DAMOH: 1 PINs (3.1%)
  - TIKAMGARH: 1 PINs (3.1%)
  - NADIA: 1 PINs (3.1%)
--------------------------------------------------------------------------------
COORDINATE QUALITY BREAKDOWN:
  - MEDIUM: 24 PINs (75.0%)
  - HIGH: 8 PINs (25.0%)
--------------------------------------------------------------------------------
RADIAL CLUSTERS:
  - Inner Core (0-15 km):    17 PINs
  - Middle Band (15-30 km):  5 PINs
  - Outer Band (30-50 km):  10 PINs
--------------------------------------------------------------------------------
NEAREST 10 PIN CODES TO ANCHOR 482001:
PIN      Distance     Office Name                  District           State            Quality   
482001   0.000 km     Jabalpur H.O                 JABALPUR           MADHYA PRADESH   MEDIUM    
482002   1.465 km     Lordganj S.O                 JABALPUR           MADHYA PRADESH   HIGH      
482008   3.782 km     Suhagi S.O                   JABALPUR           MADHYA PRADESH   HIGH      
482004   4.041 km     C.O.D. S.O                   JABALPUR           MADHYA PRADESH   HIGH      
482003   4.685 km     Gorakhpur S.O                JABALPUR           MADHYA PRADESH   HIGH      
482005   7.170 km     Vehicle Factory S.O          JABALPUR           MADHYA PRADESH   HIGH      
483001   8.356 km     Barela S.O                   JABALPUR           MADHYA PRADESH   MEDIUM    
483225   9.680 km     Ghamapur S.O                 JABALPUR           MADHYA PRADESH   MEDIUM    
482011   10.231 km    Krishinagar S.O              JABALPUR           MADHYA PRADESH   HIGH      
482020   11.516 km    G.C.F. Jabalpur S.O          JABALPUR           MADHYA PRADESH   HIGH      
--------------------------------------------------------------------------------
```

---

## 4. Rebuilding the Normalized PIN Master

If a new India Post government CSV is supplied, re-run the builder script:

```bash
python tools/build_pin_master.py
```

The script:
1. Audits and validates raw coordinate syntax and bounds.
2. Selects representative coordinates using the 3-tier algorithm.
3. Automatically writes atomically to `.tmp` files and swaps them into `DATA/india_pincode_master.csv` and `DATA/unresolved_pins.csv`.
4. Creates automatic `.bak` backups before swapping.
5. Invalidates the SQLite cache when the master file's SHA-256 hash changes.

---

## 5. Strict Freight Separation Policy

> [!WARNING]
> **Commercial Freight Road Distance Isolation**  
> Straight-line distance (`STRAIGHT_LINE_KM`) is strictly for marketing territory exploration and geographic density screening.  
> It must **NEVER** be substituted for road freight routing or quotation freight calculation. Commercial freight calculations remain exclusively governed by the approved OpenRouteService `driving-hgv` profile cached in `tblDistanceCache` within the Quotation System.
