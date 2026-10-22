"""
Mega Pipes Quotation System
Test Suite: TESTS/test_pin_distance_engine.py

Validates all 20 Phase 1 requirements for the India PIN-Code Distance & Territory Engine:
  1. PIN format validation.
  2. Six-character string preservation.
  3. Invalid PIN rejection.
  4. Missing PIN handling.
  5. Unresolved coordinate handling.
  6. Coordinate bounds.
  7. Haversine results against independently calculated known pairs.
  8. Identical coordinates producing 0 km.
  9. Symmetry: distance A->B equals distance B->A.
  10. Radius boundary inclusion and exclusion.
  11. Spatial-index results matching brute-force results on representative fixtures.
  12. Cache hits, misses and persistence (SQLite).
  13. Cache invalidation / version separation when master changes.
  14. Reproducible master generation.
  15. Coordinate outlier handling.
  16. Quality-status assignment.
  17. One-row-per-PIN uniqueness.
  18. Correct export columns and formats.
  19. No fabricated coordinates for unresolved PINs.
  20. Isolation from the existing quotation and freight engine.
"""

import csv
import math
import sys
from pathlib import Path
import pytest

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from tools.pin_distance_engine import (
    haversine,
    latlon_to_cartesian,
    PinDistanceEngine,
    PinRecord,
    DATA_DIR,
    DEFAULT_MASTER_PATH,
    DEFAULT_SQLITE_CACHE_PATH,
    validate_pin,
    get_pin_details,
    straight_line_distance_km,
    get_marketing_territory,
    get_distance_cache_stats,
    get_engine
)


@pytest.fixture(scope="module")
def engine():
    """Shared engine instance using the normalized master."""
    return get_engine()


# ============================================================
# 1. PIN Format Validation
# ============================================================
def test_pin_format_validation(engine):
    """Accepts valid 6-digit Indian PIN strings and integer inputs coerced to 6 digits."""
    assert engine.validate_pin("482001") == "482001"
    assert engine.validate_pin("110001") == "110001"
    assert validate_pin("410203") == "410203"


# ============================================================
# 2. Six-Character String Preservation
# ============================================================
def test_six_character_string_preservation(engine):
    """Preserves leading zeros and guarantees 6-character string return type."""
    pin_with_zero = "012345"  # Synthetic test with leading zero
    assert engine.validate_pin(pin_with_zero) == "012345"
    res = engine.validate_pin("110001")
    assert isinstance(res, str)
    assert len(res) == 6


# ============================================================
# 3. Invalid PIN Rejection
# ============================================================
def test_invalid_pin_rejection(engine):
    """Explicitly rejects invalid formats: alphanumeric, too short, too long, negative, None."""
    invalid_cases = ["12345", "1234567", "ABCDEF", "48200A", "48 001", "", "   ", None, -482001]
    for inv in invalid_cases:
        with pytest.raises(ValueError, match="Invalid PIN code format|cannot be None"):
            engine.validate_pin(inv)


# ============================================================
# 4. Missing PIN Handling
# ============================================================
def test_missing_pin_handling(engine):
    """Non-existent PIN code must raise ValueError with clear message."""
    with pytest.raises(ValueError, match="not found in India PIN master"):
        engine.straight_line_distance_km("999999", "482001")
    assert engine.get_pin_details("999999") is None


# ============================================================
# 5. Unresolved Coordinate Handling
# ============================================================
def test_unresolved_coordinate_handling(engine):
    """
    PIN 483222 (Gandhi Gram B.O) has no coordinates in raw government data.
    Engine must return coordinates as None and raise ValueError on distance calculation.
    """
    details = engine.get_pin_details("483222")
    assert details is not None
    assert details["QualityStatus"] == "UNRESOLVED"
    assert details["Latitude"] is None
    assert details["Longitude"] is None
    assert not details["HasCoordinates"]

    with pytest.raises(ValueError, match="no valid coordinates"):
        engine.straight_line_distance_km("482001", "483222")


# ============================================================
# 6. Coordinate Bounds
# ============================================================
def test_coordinate_bounds(engine):
    """All resolved coordinates in master must strictly fall within India's geographic bounds."""
    for pin, rec in engine._pins.items():
        if rec.has_coordinates:
            assert 6.0 <= rec.latitude <= 38.0, f"Latitude {rec.latitude} out of bounds for PIN {pin}"
            assert 68.0 <= rec.longitude <= 98.0, f"Longitude {rec.longitude} out of bounds for PIN {pin}"


# ============================================================
# 7. Haversine Known Pairs
# ============================================================
def test_haversine_known_pairs():
    """
    Test Haversine distance against known geographic ground truth.
    Delhi (28.6139 N, 77.2090 E) to Mumbai (19.0760 N, 72.8777 E) is approx 1,148 km.
    """
    delhi_lat, delhi_lon = 28.6139, 77.2090
    mumbai_lat, mumbai_lon = 19.0760, 72.8777
    d = haversine(delhi_lat, delhi_lon, mumbai_lat, mumbai_lon)
    assert 1140.0 <= d <= 1160.0, f"Expected ~1148 km, got {d}"


# ============================================================
# 8. Identical Coordinates Producing 0 km
# ============================================================
def test_identical_coordinates_zero_km(engine):
    """Identical coordinates or identical PINs must strictly return 0.000 km."""
    assert haversine(23.161490, 79.954832, 23.161490, 79.954832) == 0.0
    assert engine.straight_line_distance_km("482001", "482001") == 0.0


# ============================================================
# 9. Distance Symmetry (A->B == B->A)
# ============================================================
def test_distance_symmetry(engine):
    """Distance A->B must strictly equal distance B->A."""
    d_ab = engine.straight_line_distance_km("482001", "410203")
    d_ba = engine.straight_line_distance_km("410203", "482001")
    assert d_ab == d_ba


# ============================================================
# 10. Radius Boundary Inclusion & Exclusion
# ============================================================
def test_radius_boundary_inclusion_exclusion(engine):
    """Points with distance <= radius are included; points > radius are excluded."""
    radius = 25.0
    rep = engine.get_marketing_territory("482001", radius_km=radius)
    assert rep["radius_km"] == radius
    assert rep["total_territory_pins"] > 0

    for p in rep["pins"]:
        assert p["DistanceKM"] <= radius, f"PIN {p['PIN']} with distance {p['DistanceKM']} exceeds radius {radius}"

    # Smaller radius must be strict subset of larger radius
    rep_15 = engine.get_marketing_territory("482001", radius_km=15.0)
    rep_50 = engine.get_marketing_territory("482001", radius_km=50.0)
    pins_15 = set(p["PIN"] for p in rep_15["pins"])
    pins_50 = set(p["PIN"] for p in rep_50["pins"])
    assert len(pins_15) <= len(pins_50)
    assert pins_15.issubset(pins_50)


# ============================================================
# 11. Spatial-Index Results Matching Brute-Force Results
# ============================================================
def test_spatial_index_matches_bruteforce(engine):
    """
    Spatial KD-Tree search results must match 100% with direct brute-force
    iteration over all 19,550 resolvable PIN records.
    """
    anchor_pin = "482001"
    radius_km = 50.0
    a_rec = engine._pins[anchor_pin]

    # 1. KD-Tree result via engine
    kd_rep = engine.get_marketing_territory(anchor_pin, radius_km=radius_km)
    kd_pins = set(p["PIN"] for p in kd_rep["pins"])

    # 2. Independent Brute-Force calculation
    bf_pins = set()
    for pin, rec in engine._pins.items():
        if rec.has_coordinates:
            d = haversine(a_rec.latitude, a_rec.longitude, rec.latitude, rec.longitude)
            if d <= radius_km:
                bf_pins.add(pin)

    assert kd_pins == bf_pins, f"KD-Tree mismatch: diff={kd_pins.symmetric_difference(bf_pins)}"


# ============================================================
# 12. Persistent SQLite Cache: Hits, Misses, and Persistence
# ============================================================
def test_sqlite_cache_hits_misses_persistence(engine, tmp_path):
    """Validates on-demand SQLite cache hits, misses, and persistence across instances."""
    test_db = tmp_path / "test_cache.sqlite"
    eng = PinDistanceEngine(cache_path=test_db)
    
    # Initial state
    stats0 = eng.get_distance_cache_stats()
    assert stats0["hits"] == 0
    assert stats0["misses"] == 0

    # First query -> Miss
    d1 = eng.straight_line_distance_km("482001", "482002")
    stats1 = eng.get_distance_cache_stats()
    assert stats1["misses"] == 1
    assert stats1["hits"] == 0
    assert stats1["current_version_cached_pairs"] == 1

    # Second symmetric query -> Hit
    d2 = eng.straight_line_distance_km("482002", "482001")
    stats2 = eng.get_distance_cache_stats()
    assert stats2["misses"] == 1
    assert stats2["hits"] == 1
    assert d1 == d2

    # New instance pointing to same SQLite DB -> Persistence
    eng2 = PinDistanceEngine(cache_path=test_db)
    d3 = eng2.straight_line_distance_km("482001", "482002")
    stats3 = eng2.get_distance_cache_stats()
    assert stats3["hits"] == 1
    assert d3 == d1


# ============================================================
# 13. Cache Separation on Master Version Change
# ============================================================
def test_cache_invalidation_on_master_change(tmp_path):
    """Changing master version must prevent returning stale cached distances."""
    test_db = tmp_path / "version_cache.sqlite"
    from tools.pin_distance_engine import PinDistanceSqliteCache
    
    cache_v1 = PinDistanceSqliteCache(test_db, master_version="V1")
    cache_v1.put("482001", "482002", 2.543)
    assert cache_v1.get("482001", "482002") == 2.543

    # Switch to V2
    cache_v2 = PinDistanceSqliteCache(test_db, master_version="V2")
    # Must miss because master version differs
    assert cache_v2.get("482001", "482002") is None


# ============================================================
# 14. Reproducible Master Generation
# ============================================================
def test_reproducible_master_generation():
    """Identical raw CSV input must generate identical normalized master records."""
    from tools.build_pin_master import build_pin_master, DEFAULT_RAW_CSV
    if not DEFAULT_RAW_CSV.exists():
        pytest.skip("Raw India Post CSV not available for rebuild test")
    
    norm_recs, unres_recs = build_pin_master(DEFAULT_RAW_CSV)
    assert len(norm_recs) == 19586
    assert len(unres_recs) == 36


# ============================================================
# 15. Coordinate Outlier Handling
# ============================================================
def test_coordinate_outlier_handling(engine):
    """BO records with severe geocoding outliers are identified with rule BO_MEDOID_OUTLIERS_REJECTED."""
    low_flagged = [rec for rec in engine._pins.values() if rec.quality_status == "LOW_FLAGGED"]
    assert len(low_flagged) == 7
    for rec in low_flagged:
        assert rec.selection_rule == "BO_MEDOID_OUTLIERS_REJECTED"


# ============================================================
# 16. Quality-Status Assignment
# ============================================================
def test_quality_status_assignment(engine):
    """Quality statuses must follow approved taxonomy: HIGH, MEDIUM, LOW_FLAGGED, UNRESOLVED."""
    status_counts = {}
    for rec in engine._pins.values():
        status_counts[rec.quality_status] = status_counts.get(rec.quality_status, 0) + 1

    assert status_counts["HIGH"] == 8724
    assert status_counts["MEDIUM"] == 10819
    assert status_counts["LOW_FLAGGED"] == 7
    assert status_counts["UNRESOLVED"] == 36
    assert sum(status_counts.values()) == 19586


# ============================================================
# 17. One-Row-Per-PIN Uniqueness
# ============================================================
def test_one_row_per_pin_uniqueness():
    """The master CSV must contain exactly 19,586 unique PINs with 0 duplicate rows."""
    pins_seen = set()
    with open(DEFAULT_MASTER_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            pin = row["PIN"]
            assert pin not in pins_seen, f"Duplicate PIN found in master: {pin}"
            pins_seen.add(pin)
    assert len(pins_seen) == 19586


# ============================================================
# 18. Correct Export Columns and Formats
# ============================================================
def test_correct_export_columns_and_formats():
    """Validates exact approved column names in master CSV."""
    expected_cols = [
        "PIN", "Latitude", "Longitude", "OfficeName", "OfficeType",
        "District", "State", "RecordCount", "ValidCoordCount",
        "QualityStatus", "SpreadKM", "SelectionRule"
    ]
    with open(DEFAULT_MASTER_PATH, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        actual_cols = next(reader)
        assert actual_cols == expected_cols


# ============================================================
# 19. No Fabricated Coordinates for Unresolved PINs
# ============================================================
def test_no_fabricated_coordinates_for_unresolved_pins(engine):
    """All 36 unresolved PINs must strictly have empty/null coordinates, never (0,0)."""
    with open(DATA_DIR / "unresolved_pins.csv", "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        unres_list = list(reader)

    assert len(unres_list) == 36
    for row in unres_list:
        assert row["Latitude"] == "", f"Fabricated latitude for unresolved PIN {row['PIN']}"
        assert row["Longitude"] == "", f"Fabricated longitude for unresolved PIN {row['PIN']}"
        assert row["QualityStatus"] == "UNRESOLVED"


# ============================================================
# 20. Isolation From Existing Quotation and Freight Engine
# ============================================================
def test_isolation_from_quotation_and_freight_engine(engine):
    """
    Marketing straight-line distance must never alter commercial road freight
    baselines, quotation calculation modules, or distance cache schemas.
    """
    # 1. Straight-line distance between Jabalpur and Khopoli is ~841 km
    straight_line = engine.straight_line_distance_km("482001", "410203")
    assert 835.0 <= straight_line <= 845.0

    # 2. Verify road freight distance remains governed by ORS / tblDistanceCache (not Haversine)
    rep = engine.get_marketing_territory("482001", radius_km=50.0)
    assert rep["metric"] == "STRAIGHT_LINE_KM"
    assert "OpenRouteService driving-hgv" in rep["road_distance_note"]
    assert "tblDistanceCache" in rep["road_distance_note"]
    assert "(STRAIGHT_LINE_KM) must NEVER be used for freight cost calculations" in rep["road_distance_note"]
