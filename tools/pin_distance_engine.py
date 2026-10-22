#!/usr/bin/env python3
"""
Mega Pipes Quotation System
Module: tools/pin_distance_engine.py

Purpose:
  High-performance All-India PIN-code distance calculation and marketing territory engine.
  Features:
    1. Geographic straight-line distance via the Haversine formula (km) to 3 decimal places.
    2. Persistent on-demand SQLite pair caching with master-version separation and symmetric keys.
    3. Spatial KD-Tree index on 3D spherical coordinates for instantaneous radius discovery.
    4. Configurable radius marketing territory discovery (25 km, 50 km, 100 km).
    5. Export utilities for territory data (CSV, JSON, and formatted plain-text reports).
    6. Strict metric separation: straight-line distance is NEVER used as road freight distance.
    7. Full validation: 6-character PIN preservation, explicit rejection of invalid or unresolved PINs.

Required Public API:
  - validate_pin(pin: str) -> str
  - get_pin_details(pin: str) -> Optional[Dict[str, Any]]
  - straight_line_distance_km(origin_pin: str, destination_pin: str) -> float
  - get_marketing_territory(anchor_pin: str, radius_km: float = 50.0) -> Dict[str, Any]
  - get_distance_cache_stats() -> Dict[str, Any]
"""

import csv
import json
import math
import os
import sys
import sqlite3
import hashlib
from collections import Counter
from datetime import datetime
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Set, Union

# Project paths
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "DATA"
DEFAULT_MASTER_PATH = DATA_DIR / "india_pincode_master.csv"
DEFAULT_SQLITE_CACHE_PATH = DATA_DIR / "pin_distance_cache.sqlite"

# Earth mean radius in kilometres (WGS84 spherical approximation)
EARTH_RADIUS_KM = 6371.0


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """
    Calculate the great-circle distance between two points on the Earth's surface
    using the Haversine formula, in kilometres.
    
    Guarantees:
      - Identical coordinates strictly return 0.0.
      - Returns non-negative distance rounded to 3 decimal places (1 metre precision).
    """
    if lat1 == lat2 and lon1 == lon2:
        return 0.0

    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)

    a = (math.sin(delta_phi / 2.0) ** 2 +
         math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2.0) ** 2)
    a = min(1.0, max(0.0, a))
    c = 2.0 * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))
    
    return round(EARTH_RADIUS_KM * c, 3)


def latlon_to_cartesian(lat_deg: float, lon_deg: float) -> Tuple[float, float, float]:
    """Convert spherical latitude/longitude in degrees to 3D unit Cartesian coordinates (x, y, z)."""
    phi = math.radians(lat_deg)
    lam = math.radians(lon_deg)
    x = math.cos(phi) * math.cos(lam)
    y = math.cos(phi) * math.sin(lam)
    z = math.sin(phi)
    return (x, y, z)


# ============================================================
# PURE-PYTHON SPATIAL KD-TREE FOR SPHERICAL COORDINATES
# ============================================================

class KDNode:
    __slots__ = ("point", "data", "axis", "left", "right")

    def __init__(self, point: Tuple[float, float, float], data: Any, axis: int):
        self.point = point  # (x, y, z) on unit sphere
        self.data = data    # PinRecord
        self.axis = axis
        self.left: Optional['KDNode'] = None
        self.right: Optional['KDNode'] = None


class SpatialKDTree:
    """
    Pure-Python 3D KD-Tree for spatial coordinates on the unit sphere.
    Enables O(log N) radius range searching without requiring external C libraries (scipy/sklearn).
    """

    def __init__(self, items: List[Tuple[Tuple[float, float, float], Any]]):
        self.root = self._build_tree(items, depth=0)

    def _build_tree(self, items: List[Tuple[Tuple[float, float, float], Any]], depth: int) -> Optional[KDNode]:
        if not items:
            return None
        axis = depth % 3
        items.sort(key=lambda item: item[0][axis])
        median_idx = len(items) // 2
        median_pt, median_data = items[median_idx]

        node = KDNode(median_pt, median_data, axis)
        node.left = self._build_tree(items[:median_idx], depth + 1)
        node.right = self._build_tree(items[median_idx + 1:], depth + 1)
        return node

    def query_radius(self, center_pt: Tuple[float, float, float], max_chord_dist: float) -> List[Any]:
        """Find all items within Euclidean chord distance <= max_chord_dist."""
        results: List[Any] = []
        max_chord_sq = max_chord_dist * max_chord_dist

        def _search(node: Optional[KDNode]):
            if node is None:
                return
            pt = node.point
            # Squared 3D Euclidean distance
            dx = pt[0] - center_pt[0]
            dy = pt[1] - center_pt[1]
            dz = pt[2] - center_pt[2]
            d_sq = dx * dx + dy * dy + dz * dz

            if d_sq <= max_chord_sq:
                results.append(node.data)

            axis = node.axis
            diff = center_pt[axis] - pt[axis]

            # Search branch containing the point first
            if diff <= 0:
                _search(node.left)
                if diff * diff <= max_chord_sq:
                    _search(node.right)
            else:
                _search(node.right)
                if diff * diff <= max_chord_sq:
                    _search(node.left)

        _search(self.root)
        return results


# ============================================================
# DATA MODELS
# ============================================================

@dataclass
class PinRecord:
    pin: str
    latitude: Optional[float]
    longitude: Optional[float]
    office_name: str
    office_type: str
    district: str
    state: str
    record_count: int
    valid_coord_count: int
    quality_status: str
    spread_km: float
    selection_rule: str

    @property
    def has_coordinates(self) -> bool:
        return self.latitude is not None and self.longitude is not None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "PIN": self.pin,
            "Latitude": self.latitude,
            "Longitude": self.longitude,
            "OfficeName": self.office_name,
            "OfficeType": self.office_type,
            "District": self.district,
            "State": self.state,
            "RecordCount": self.record_count,
            "ValidCoordCount": self.valid_coord_count,
            "QualityStatus": self.quality_status,
            "SpreadKM": self.spread_km,
            "SelectionRule": self.selection_rule,
            "HasCoordinates": self.has_coordinates
        }


# ============================================================
# PERSISTENT SQLITE ON-DEMAND CACHE
# ============================================================

class PinDistanceSqliteCache:
    """Persistent on-demand SQLite cache for undirected PIN-pair distances."""

    def __init__(self, db_path: Path, master_version: str):
        self.db_path = db_path
        self.master_version = master_version
        self.hits = 0
        self.misses = 0
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        return sqlite3.connect(str(self.db_path), timeout=30.0)

    def _init_db(self) -> None:
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS pin_pair_distance_cache (
                    origin_pin TEXT NOT NULL,
                    destination_pin TEXT NOT NULL,
                    distance_km REAL NOT NULL,
                    master_version TEXT NOT NULL,
                    calculated_at TEXT NOT NULL,
                    PRIMARY KEY (origin_pin, destination_pin)
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cache_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
            """)
            cur.execute("INSERT OR REPLACE INTO cache_metadata (key, value) VALUES ('engine', 'MegaPipes_Haversine_v1');")
            cur.execute("INSERT OR REPLACE INTO cache_metadata (key, value) VALUES ('master_version', ?);", (self.master_version,))
            conn.commit()

    def get(self, pin1: str, pin2: str) -> Optional[float]:
        """Lookup cached distance for undirected pair (origin <= destination)."""
        p_min, p_max = (pin1, pin2) if pin1 <= pin2 else (pin2, pin1)
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT distance_km FROM pin_pair_distance_cache
                WHERE origin_pin = ? AND destination_pin = ? AND master_version = ?;
            """, (p_min, p_max, self.master_version))
            row = cur.fetchone()
            if row is not None:
                self.hits += 1
                return float(row[0])
            self.misses += 1
            return None

    def put(self, pin1: str, pin2: str, distance_km: float) -> None:
        """Store calculated distance for undirected pair."""
        p_min, p_max = (pin1, pin2) if pin1 <= pin2 else (pin2, pin1)
        ts = datetime.now().isoformat()
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT OR REPLACE INTO pin_pair_distance_cache
                (origin_pin, destination_pin, distance_km, master_version, calculated_at)
                VALUES (?, ?, ?, ?, ?);
            """, (p_min, p_max, float(distance_km), self.master_version, ts))
            conn.commit()

    def clear(self) -> None:
        """Purge all cached pairs."""
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM pin_pair_distance_cache;")
            conn.commit()

    def get_stats(self) -> Dict[str, Any]:
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM pin_pair_distance_cache WHERE master_version = ?;", (self.master_version,))
            count = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM pin_pair_distance_cache;")
            total_all_versions = cur.fetchone()[0]
        return {
            "backend": "sqlite",
            "db_path": str(self.db_path),
            "hits": self.hits,
            "misses": self.misses,
            "current_version_cached_pairs": count,
            "total_cached_pairs_all_versions": total_all_versions,
            "master_version": self.master_version
        }


# ============================================================
# PIN DISTANCE & MARKETING TERRITORY ENGINE
# ============================================================

class PinDistanceEngine:
    """Core geospatial distance calculation and marketing territory engine."""

    def __init__(self, master_path: Path = DEFAULT_MASTER_PATH, cache_path: Path = DEFAULT_SQLITE_CACHE_PATH):
        self.master_path = master_path
        self.cache_path = cache_path
        self._pins: Dict[str, PinRecord] = {}
        self._spatial_tree: Optional[SpatialKDTree] = None
        self._master_version = self._compute_master_version()
        self._cache = PinDistanceSqliteCache(self.cache_path, self._master_version)

        self._load_master()
        self._build_spatial_index()

    def _compute_master_version(self) -> str:
        """Compute SHA256 checksum of master CSV file for cache separation."""
        if not self.master_path.exists():
            return "UNINITIALIZED"
        hasher = hashlib.sha256()
        with open(self.master_path, "rb") as f:
            while chunk := f.read(65536):
                hasher.update(chunk)
        return hasher.hexdigest()[:16]

    def _load_master(self) -> None:
        """Load normalized PIN master into memory."""
        if not self.master_path.exists():
            raise FileNotFoundError(
                f"Normalized PIN master not found at {self.master_path}. "
                "Run tools/build_pin_master.py first."
            )

        with open(self.master_path, mode="r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            # Map canonical and alternative header names
            for row in reader:
                # Support both 'PIN' and 'pincode'
                pin_raw = row.get("PIN") or row.get("pincode") or ""
                pin = pin_raw.strip().zfill(6)
                if not pin:
                    continue

                lat_str = (row.get("Latitude") or row.get("latitude") or "").strip()
                lon_str = (row.get("Longitude") or row.get("longitude") or "").strip()
                lat = float(lat_str) if lat_str else None
                lon = float(lon_str) if lon_str else None

                office_name = row.get("OfficeName") or row.get("office_name") or ""
                office_type = row.get("OfficeType") or row.get("office_type") or ""
                district = row.get("District") or row.get("district") or ""
                state = row.get("State") or row.get("state") or ""

                rec_cnt_str = (row.get("RecordCount") or row.get("office_count") or "1").strip()
                valid_cnt_str = (row.get("ValidCoordCount") or row.get("valid_coord_count") or "0").strip()
                rec_cnt = int(rec_cnt_str) if rec_cnt_str.isdigit() else 1
                valid_cnt = int(valid_cnt_str) if valid_cnt_str.isdigit() else 0

                quality = (row.get("QualityStatus") or row.get("quality_status") or "UNKNOWN").strip()
                spread_str = (row.get("SpreadKM") or row.get("internal_spread_km") or "0.0").strip()
                spread = float(spread_str) if spread_str else 0.0
                rule = (row.get("SelectionRule") or row.get("selection_rule") or "").strip()

                self._pins[pin] = PinRecord(
                    pin=pin,
                    latitude=lat,
                    longitude=lon,
                    office_name=office_name,
                    office_type=office_type,
                    district=district,
                    state=state,
                    record_count=rec_cnt,
                    valid_coord_count=valid_cnt,
                    quality_status=quality,
                    spread_km=spread,
                    selection_rule=rule
                )

    def _build_spatial_index(self) -> None:
        """Build 3D KD-Tree index from all resolvable PIN records."""
        items: List[Tuple[Tuple[float, float, float], PinRecord]] = []
        for rec in self._pins.values():
            if rec.has_coordinates:
                cart_pt = latlon_to_cartesian(rec.latitude, rec.longitude)
                items.append((cart_pt, rec))
        self._spatial_tree = SpatialKDTree(items)

    def validate_pin(self, pin: Any) -> str:
        """
        Validate and normalize an Indian PIN code string.
        
        Rules:
          - Must be a 6-digit numeric string.
          - Rejects non-numeric, None, empty, strings of length != 6.
          - Preserves leading zeros.
        """
        if pin is None:
            raise ValueError("PIN code cannot be None")
        pin_str = str(pin).strip()
        if not (len(pin_str) == 6 and pin_str.isdigit()):
            raise ValueError(f"Invalid PIN code format: '{pin}'. Must be a 6-digit numeric string.")
        return pin_str

    # Backward compatibility alias
    normalize_pin = validate_pin

    def get_pin_details(self, pin: Any) -> Optional[Dict[str, Any]]:
        """Retrieve full details dictionary for a given PIN code."""
        clean_pin = self.validate_pin(pin)
        rec = self._pins.get(clean_pin)
        return rec.to_dict() if rec else None

    # Backward compatibility method
    def get_pin(self, pin: Any) -> Optional[PinRecord]:
        clean_pin = self.validate_pin(pin)
        return self._pins.get(clean_pin)

    def straight_line_distance_km(self, origin_pin: Any, destination_pin: Any) -> float:
        """
        Calculate straight-line geographic distance between two PIN codes in kilometres.
        
        Guarantees:
          - Identical PIN codes strictly return 0.000 km.
          - Rejects invalid PIN formats, unknown PINs, and unresolved coordinates.
          - Uses persistent SQLite on-demand cache.
        """
        p1 = self.validate_pin(origin_pin)
        p2 = self.validate_pin(destination_pin)

        if p1 == p2:
            return 0.0

        # Check SQLite cache
        cached_d = self._cache.get(p1, p2)
        if cached_d is not None:
            return cached_d

        rec1 = self._pins.get(p1)
        if not rec1:
            raise ValueError(f"PIN code '{p1}' not found in India PIN master.")
        if not rec1.has_coordinates:
            raise ValueError(f"PIN code '{p1}' has no valid coordinates (status={rec1.quality_status}).")

        rec2 = self._pins.get(p2)
        if not rec2:
            raise ValueError(f"PIN code '{p2}' not found in India PIN master.")
        if not rec2.has_coordinates:
            raise ValueError(f"PIN code '{p2}' has no valid coordinates (status={rec2.quality_status}).")

        dist = haversine(rec1.latitude, rec1.longitude, rec2.latitude, rec2.longitude)
        self._cache.put(p1, p2, dist)
        return dist

    # Backward compatibility alias
    calculate_distance = straight_line_distance_km

    def get_distance_cache_stats(self) -> Dict[str, Any]:
        """Return cache hit, miss, and persistence statistics."""
        return self._cache.get_stats()

    def get_marketing_territory(self, anchor_pin: Any, radius_km: float = 50.0) -> Dict[str, Any]:
        """
        Discover marketing territory around an anchor PIN within a configurable geographic radius.
        
        Parameters:
          anchor_pin: 6-digit Indian PIN string
          radius_km: Float radius in kilometres (inclusive: 0 <= distance <= radius_km)
          
        Returns structured dictionary containing:
          - anchor: Details of anchor PIN
          - radius_km: Search radius applied
          - total_territory_pins: Count of PINs within radius (including anchor)
          - pins: List of surrounding PINs sorted by distance ascending
          - district_summary: Distribution of pins by district
          - state_summary: Distribution of pins by state
          - quality_summary: Distribution of coordinate quality
          - radial_clusters: Tiered radial clusters (0-15km inner, 15-30km middle, 30-radius_km outer)
          - coverage_gaps: Unresolved PINs in the anchor's district or surrounding territory districts
          - screening_disclaimer: Statement that this is geographical screening data
          - road_distance_note: Explicit statement of metric separation from road freight
        """
        if radius_km < 0.0:
            raise ValueError("Radius must be non-negative.")

        clean_anchor = self.validate_pin(anchor_pin)
        anchor_rec = self._pins.get(clean_anchor)

        if not anchor_rec:
            raise ValueError(f"Anchor PIN '{clean_anchor}' not found in India PIN master.")
        if not anchor_rec.has_coordinates:
            raise ValueError(f"Anchor PIN '{clean_anchor}' has no valid coordinates (status={anchor_rec.quality_status}).")

        a_lat = anchor_rec.latitude
        a_lon = anchor_rec.longitude

        # 1. Query Spatial KD-Tree Index
        # Max chord distance in 3D Euclidean space corresponding to angular radius
        theta = radius_km / EARTH_RADIUS_KM
        max_chord_dist = 2.0 * math.sin(theta / 2.0)
        center_cart = latlon_to_cartesian(a_lat, a_lon)

        candidates: List[PinRecord] = self._spatial_tree.query_radius(center_cart, max_chord_dist)

        # 2. Exact Haversine Calculation & Filtering (0 <= d <= radius_km)
        territory_pins = []
        territory_districts: Set[str] = set()
        district_counts = Counter()
        state_counts = Counter()
        quality_counts = Counter()
        cluster_inner = []   # 0 to 15 km
        cluster_middle = []  # 15 to 30 km
        cluster_outer = []   # 30 to radius_km

        for rec in candidates:
            d = haversine(a_lat, a_lon, rec.latitude, rec.longitude)
            # Exact inclusive boundary condition
            if d <= radius_km:
                pin_info = {
                    "PIN": rec.pin,
                    "OfficeName": rec.office_name,
                    "OfficeType": rec.office_type,
                    "District": rec.district,
                    "State": rec.state,
                    "QualityStatus": rec.quality_status,
                    "DistanceKM": d,
                    "Latitude": rec.latitude,
                    "Longitude": rec.longitude,
                    "RecordCount": rec.record_count
                }
                territory_pins.append(pin_info)
                territory_districts.add(rec.district.upper())
                district_counts[rec.district] += 1
                state_counts[rec.state] += 1
                quality_counts[rec.quality_status] += 1

                if d <= 15.0:
                    cluster_inner.append(rec.pin)
                elif d <= 30.0:
                    cluster_middle.append(rec.pin)
                else:
                    cluster_outer.append(rec.pin)

        # Sort pins by straight-line distance ascending
        territory_pins.sort(key=lambda x: x["DistanceKM"])

        # 3. Identify Coverage Gaps (Unresolved PINs in the territory districts)
        coverage_gaps = []
        for pin, rec in self._pins.items():
            if not rec.has_coordinates:
                if rec.district.upper() in territory_districts:
                    coverage_gaps.append({
                        "PIN": rec.pin,
                        "OfficeName": rec.office_name,
                        "District": rec.district,
                        "State": rec.state,
                        "Status": rec.quality_status,
                        "Reason": "Zero physical coordinates in official government directory"
                    })

        return {
            "anchor": anchor_rec.to_dict(),
            "radius_km": float(radius_km),
            "metric": "STRAIGHT_LINE_KM",
            "total_territory_pins": len(territory_pins),
            "pins": territory_pins,
            "district_summary": dict(district_counts.most_common()),
            "state_summary": dict(state_counts.most_common()),
            "quality_summary": dict(quality_counts),
            "radial_clusters": {
                "inner_0_15km": cluster_inner,
                "middle_15_30km": cluster_middle,
                "outer_30_50km": cluster_outer
            },
            "coverage_gaps": coverage_gaps,
            "screening_disclaimer": (
                "GEOGRAPHICAL SCREENING NOTICE: This territory query identifies postal delivery zones "
                f"within a {radius_km:.1f} km straight-line radius of anchor PIN {clean_anchor}. "
                "Inclusion does NOT imply confirmed sales leads, actual customer demand, or immediate "
                "commercial feasibility. Straight-line distance is an exploratory screening metric."
            ),
            "road_distance_note": (
                "METRIC SEPARATION POLICY: Straight-line distance (STRAIGHT_LINE_KM) must NEVER be used "
                "for freight cost calculations, trailer logistics, or quotation pricing. All commercial "
                "freight calculations are governed exclusively by approved OpenRouteService driving-hgv "
                "road distances cached in tblDistanceCache."
            )
        }


# ============================================================
# EXPORT & REPORTING HELPERS
# ============================================================

def export_territory_to_csv(territory_data: Dict[str, Any], filepath: Union[str, Path]) -> Path:
    """Export territory PIN list to CSV."""
    p = Path(filepath)
    p.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ["PIN", "OfficeName", "OfficeType", "District", "State", "QualityStatus", "DistanceKM", "Latitude", "Longitude", "RecordCount"]
    with open(p, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(territory_data["pins"])
    return p


def export_territory_to_json(territory_data: Dict[str, Any], filepath: Union[str, Path]) -> Path:
    """Export complete territory report to JSON."""
    p = Path(filepath)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, mode="w", encoding="utf-8") as f:
        json.dump(territory_data, f, indent=2)
    return p


def generate_territory_text_report(territory_data: Dict[str, Any]) -> str:
    """Generate a clean human-readable text report of marketing territory."""
    a = territory_data["anchor"]
    r = territory_data["radius_km"]
    tot = territory_data["total_territory_pins"]
    
    lines = [
        "=" * 80,
        f"MEGA PIPES — MARKETING TERRITORY DISCOVERY REPORT",
        "=" * 80,
        f"Anchor PIN:             {a['PIN']} ({a['OfficeName']}, {a['OfficeType']})",
        f"Anchor Location:        {a['District']}, {a['State']} (Lat: {a['Latitude']:.6f}, Lon: {a['Longitude']:.6f})",
        f"Anchor Quality Status:  {a['QualityStatus']} (Selection Rule: {a['SelectionRule']})",
        f"Exploratory Radius:     {r:.1f} km (Straight-Line Great-Circle)",
        f"Total Eligible PINs:    {tot:,} unique delivery postal zones",
        "-" * 80,
        "DISTRICT DISTRIBUTION:",
    ]
    for dist, cnt in territory_data["district_summary"].items():
        lines.append(f"  - {dist}: {cnt} PINs ({cnt/tot*100:.1f}%)")
        
    lines.extend([
        "-" * 80,
        "COORDINATE QUALITY BREAKDOWN:",
    ])
    for q, cnt in territory_data["quality_summary"].items():
        lines.append(f"  - {q}: {cnt} PINs ({cnt/tot*100:.1f}%)")
        
    lines.extend([
        "-" * 80,
        "RADIAL CLUSTERS:",
        f"  - Inner Core (0-15 km):    {len(territory_data['radial_clusters']['inner_0_15km'])} PINs",
        f"  - Middle Band (15-30 km):  {len(territory_data['radial_clusters']['middle_15_30km'])} PINs",
        f"  - Outer Band (30-{r:.0f} km):  {len(territory_data['radial_clusters']['outer_30_50km'])} PINs",
        "-" * 80,
        f"NEAREST 10 PIN CODES TO ANCHOR {a['PIN']}:",
        f"{'PIN':<8} {'Distance':<12} {'Office Name':<28} {'District':<18} {'State':<16} {'Quality':<10}"
    ])
    for p in territory_data["pins"][:10]:
        lines.append(f"{p['PIN']:<8} {p['DistanceKM']:<12.3f} {p['OfficeName'][:26]:<28} {p['District'][:16]:<18} {p['State'][:14]:<16} {p['QualityStatus']:<10}")

    if territory_data["coverage_gaps"]:
        lines.extend([
            "-" * 80,
            f"IDENTIFIED COVERAGE GAPS IN TERRITORY ({len(territory_data['coverage_gaps'])} Unresolved PINs):",
        ])
        for g in territory_data["coverage_gaps"][:5]:
            lines.append(f"  - PIN {g['PIN']}: {g['OfficeName']} ({g['District']}, {g['State']}) — No coordinates in raw directory")
            
    lines.extend([
        "=" * 80,
        territory_data["screening_disclaimer"],
        "-" * 80,
        territory_data["road_distance_note"],
        "=" * 80
    ])
    return "\n".join(lines)


# ============================================================
# SINGLETON & MODULE-LEVEL CONVENIENCE API
# ============================================================

_DEFAULT_ENGINE: Optional[PinDistanceEngine] = None

def get_engine() -> PinDistanceEngine:
    """Retrieve shared singleton engine instance."""
    global _DEFAULT_ENGINE
    if _DEFAULT_ENGINE is None:
        _DEFAULT_ENGINE = PinDistanceEngine()
    return _DEFAULT_ENGINE

def validate_pin(pin: Any) -> str:
    """Module-level validation of 6-digit Indian PIN."""
    return get_engine().validate_pin(pin)

def get_pin_details(pin: Any) -> Optional[Dict[str, Any]]:
    """Module-level retrieval of PIN details dictionary."""
    return get_engine().get_pin_details(pin)

def straight_line_distance_km(origin_pin: Any, destination_pin: Any) -> float:
    """Module-level straight-line distance calculation in km."""
    return get_engine().straight_line_distance_km(origin_pin, destination_pin)

def get_marketing_territory(anchor_pin: Any, radius_km: float = 50.0) -> Dict[str, Any]:
    """Module-level marketing territory query within radius_km."""
    return get_engine().get_marketing_territory(anchor_pin, radius_km)

def get_distance_cache_stats() -> Dict[str, Any]:
    """Module-level cache statistics."""
    return get_engine().get_distance_cache_stats()


if __name__ == "__main__":
    eng = get_engine()
    print("Mega Pipes PIN Distance Engine Initialized.")
    print("Master records:", len(eng._pins))
    
    # Sample Query: Jabalpur Plant (482001)
    rep = eng.get_marketing_territory("482001", radius_km=50.0)
    print("\n" + generate_territory_text_report(rep))
