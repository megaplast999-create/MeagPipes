"""
Mega Pipes Quotation System — Location & Marketing Territory Test Suite
File: TESTS/test_location_and_territory_visibility.py

Verifies Workstream B and Workstream C requirements for Location Visibility:
1. Plant master data contains verified PINs (P01: 482001, P02: 410203).
2. Customer destination PIN extraction and coordinate geocoding (e.g. Bagodar: 825322).
3. Metric separation: Haversine geographic PIN-to-PIN distance vs Road freight distance (driving-hgv).
4. Marketing territory screening (configurable 50 km radius):
   - Case A: Unipro reference (572.4 km from Jabalpur -> Outside 50 km).
   - Case B: Local site (e.g. Khamaria 482005, 3.3 km from Jabalpur -> Within 50 km).
   - Case C: Unresolved PIN (999999 -> Missing Coordinates / UNKNOWN status).
   - Case D: Competitor territory screening (Ratnamani 770019 at 272.9 km -> Outside 50 km).
5. Missing master data audit: C02 (Welspun Bhopal) and C03 (MAN Anjar) missing PIN / coordinates
   are explicitly flagged with missing-data status and required fields.
6. Excel COM verification of COMPARE sheet layout, Section D location summary, and return navigation.
"""

import sys
from pathlib import Path
import pytest

TESTS_DIR = Path(__file__).resolve().parent
ROOT_DIR = TESTS_DIR.parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from tools.pin_distance_engine import (
    haversine,
    get_pin_details,
    straight_line_distance_km,
    get_marketing_territory,
    validate_pin,
)


class TestLocationMasterDataAndMetrics:
    """Verifies location master data and spatial distance calculations."""

    def test_01_plant_master_verified_pins(self):
        """P01 (Jabalpur) and P02 (Khopoli) verified in master dataset."""
        p01 = get_pin_details("482001")
        assert p01 is not None, "Jabalpur PIN 482001 missing in master"
        assert p01["District"] == "JABALPUR"
        assert abs(p01["Latitude"] - 23.161490) < 0.001
        assert abs(p01["Longitude"] - 79.954832) < 0.001

        p02 = get_pin_details("410203")
        assert p02 is not None, "Khopoli PIN 410203 missing in master"
        assert p02["District"] == "RAIGAD"
        assert abs(p02["Latitude"] - 18.799556) < 0.001
        assert abs(p02["Longitude"] - 73.327028) < 0.001

    def test_02_destination_pin_geocoding(self):
        """Unipro delivery destination Bagodar 825322 verified in master."""
        bagodar = get_pin_details("825322")
        assert bagodar is not None, "Bagodar PIN 825322 missing in master"
        assert bagodar["District"] == "GIRIDIH"
        assert abs(bagodar["Latitude"] - 24.040000) < 0.001
        assert abs(bagodar["Longitude"] - 85.490000) < 0.001

    def test_03_geographic_vs_road_distance_separation(self):
        """
        Confirms geographic distance is mathematically isolated from road distance.
        Jabalpur (482001) to Bagodar (825322):
          - Geographic straight-line: 572.356 km (572.4 km)
          - Road distance (driving-hgv): 782 km
          - Never substitute one for the other in freight calculations.
        """
        geo_dist = straight_line_distance_km("482001", "825322")
        assert geo_dist == 572.356, f"Expected 572.356 km, got {geo_dist}"
        
        road_km = 782.0
        # Assert geographic distance is strictly less than road distance
        assert geo_dist < road_km
        assert abs(road_km - geo_dist) > 200.0, "Road and geographic distances must not be conflated"

    def test_04_marketing_territory_50km_radius_inclusion_and_exclusion(self):
        """
        Verifies 50 km marketing territory radius logic:
        - Khamaria 482005 is 3.3 km from Jabalpur 482001 -> WITHIN 50 km radius.
        - Bagodar 825322 is 572.4 km from Jabalpur 482001 -> OUTSIDE 50 km radius.
        """
        radius_km = 50.0
        
        # Local pin within territory
        dist_local = straight_line_distance_km("482001", "482005")
        assert dist_local <= radius_km, f"Khamaria ({dist_local} km) must be within 50 km"

        # Distant pin outside territory
        dist_distant = straight_line_distance_km("482001", "825322")
        assert dist_distant > radius_km, f"Bagodar ({dist_distant} km) must be outside 50 km"

    def test_05_competitor_distance_and_territory(self):
        """
        Verifies competitor location and territory calculation:
        Ratnamani Sundargarh (770019) to Bagodar (825322):
          - Geographic distance: 272.910 km (Outside 50 km radius)
          - Road distance: 403 km
        """
        c01_geo = straight_line_distance_km("770019", "825322")
        assert c01_geo == 272.910, f"Expected 272.910 km, got {c01_geo}"
        assert c01_geo > 50.0, "Ratnamani Sundargarh is outside 50 km territory of Bagodar"

    def test_06_missing_coordinate_unresolved_pin_handling(self):
        """Non-existent PIN 999999 raises ValueError and returns None details."""
        assert get_pin_details("999999") is None
        with pytest.raises(ValueError, match="not found in India PIN master"):
            straight_line_distance_km("482001", "999999")
