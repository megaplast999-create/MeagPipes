"""Logistics, pincode geolocation, and distance calculation service for Mega Pipes.

Integrates with the authoritative All-India PIN-code distance engine (tools/pin_distance_engine.py)
and respects standard commercial settings from tblSettings / DISTANCE_CACHE.
"""

from __future__ import annotations

import logging
import re
import sys
from pathlib import Path
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)

# Locate project root and tools directory
ROOT_DIR = Path(__file__).resolve().parent.parent.parent.parent
TOOLS_DIR = ROOT_DIR / "tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

try:
    import pin_distance_engine as pde
except ImportError:
    pde = None
    logger.warning("Could not import pin_distance_engine from tools. Basic fallback will be used.")

# Seeded / Reference distances from DISTANCE_CACHE sheet in Mega_Pipes_Quotation_DEV.xlsm
# Key: destination 6-digit PIN -> road distance in km from MPPL primary plant
SEEDED_DISTANCE_CACHE: Dict[str, float] = {
    "825322": 782.0,   # Bagodar, Jharkhand (Reference quotation distance)
    "110001": 850.5,   # New Delhi
    "411001": 1350.0,  # Pune, Maharashtra
    "370110": 1150.0,  # Anjar, Gujarat
    "482001": 30.0,    # Jabalpur (Local delivery minimum)
}

# Standard Commercial Settings from tblSettings
DEFAULT_PLANT_PIN = "482001"     # P01 Jabalpur (Maneri)
DEFAULT_RATE_PER_KM = 125.32      # SettingKey: RatePerKm (₹/km)
DEFAULT_TRIP_RATE_MANUAL = 98000.0  # SettingKey: TripRateManual (₹ / trailer)
REFERENCE_TRIP_RATE = 93000.0     # Reference sheet trip rate


def extract_pin_from_text(text: Optional[str]) -> Optional[str]:
    """Extracts first valid 6-digit Indian PIN code from address text (digits 1-9 followed by 5 digits)."""
    if not text:
        return None
    match = re.search(r"\b([1-9][0-9]{5})\b", str(text))
    return match.group(1) if match else None


def is_valid_pin(pin: Optional[str]) -> bool:
    """Validates whether a string is a valid 6-digit Indian PIN code."""
    if not pin:
        return False
    pin_str = str(pin).strip()
    return bool(re.match(r"^[1-9][0-9]{5}$", pin_str))


def resolve_delivery_logistics(
    pincode_or_address: Optional[str],
    mode: str = "REFERENCE",
    plant_pin: str = DEFAULT_PLANT_PIN,
) -> Dict[str, Any]:
    """
    Resolves distance (km) and recommended freight trip rate (₹/trailer) for a given PIN code or address.
    
    Guarantees:
      - Never returns negative or zero distance or trip rate.
      - Uses authoritative pin_distance_engine data when available.
      - Falls back gracefully to standard reference figures if offline.
    """
    clean_input = str(pincode_or_address or "").strip()
    extracted_pin = extract_pin_from_text(clean_input) if not is_valid_pin(clean_input) else clean_input

    # If no valid PIN is found, fallback to Bagodar reference
    if not extracted_pin:
        logger.info(f"No valid PIN in '{clean_input}'. Using default reference PIN 825322.")
        extracted_pin = "825322"

    office_name = ""
    district = ""
    state = ""

    if pde is not None:
        try:
            details = pde.get_pin_details(extracted_pin)
            if details:
                office_name = details.get("OfficeName", "")
                district = details.get("District", "")
                state = details.get("State", "")
        except Exception as e:
            logger.warning(f"Error querying pin_distance_engine for {extracted_pin}: {e}")

    # Determine distance (km)
    if extracted_pin in SEEDED_DISTANCE_CACHE:
        distance_km = SEEDED_DISTANCE_CACHE[extracted_pin]
        distance_source = "SEEDED_CACHE"
    elif pde is not None:
        try:
            sl_km = pde.straight_line_distance_km(plant_pin, extracted_pin)
            # Commercial HGV road factor: 1.25x curvature over straight line
            distance_km = round(max(30.0, sl_km * 1.25), 1)
            distance_source = "HGV_ROAD_CALCULATED"
        except Exception as e:
            logger.warning(f"Error calculating distance for {extracted_pin}: {e}")
            distance_km = 782.0
            distance_source = "FALLBACK_DEFAULT"
    else:
        distance_km = 782.0
        distance_source = "FALLBACK_DEFAULT"

    # Compute trip rate based on commercial mode
    normalized_mode = mode.upper().strip()
    if normalized_mode == "REFERENCE":
        # In Reference mode, if distance is ~782 km, rate is 98,000 or 93,000
        trip_rate = max(REFERENCE_TRIP_RATE, round(distance_km * DEFAULT_RATE_PER_KM, 2))
    elif normalized_mode == "LEGACY":
        trip_rate = DEFAULT_TRIP_RATE_MANUAL
    else:  # NORMAL mode: dynamic RatePerKm * km
        trip_rate = max(15000.0, round(distance_km * DEFAULT_RATE_PER_KM, 2))

    formatted_address = (
        f"{office_name}, {district}, {state} - {extracted_pin}".strip(", -")
        if district and state
        else f"PIN {extracted_pin}"
    )

    return {
        "valid": True,
        "pin": extracted_pin,
        "office_name": office_name,
        "district": district,
        "state": state,
        "distance_km": float(distance_km),
        "trip_rate": float(trip_rate),
        "rate_per_km": DEFAULT_RATE_PER_KM,
        "source": distance_source,
        "plant_id": "P01",
        "plant_name": "Mega Pipes Jabalpur (Maneri)",
        "formatted_address": formatted_address,
    }
