"""Configuration and constants for Mega Pipes Streamlit Application."""

from __future__ import annotations

import os
from typing import Any, Dict

# Detect backend URL from environment or Streamlit secrets, defaulting to live production Render backend
BACKEND_URL = os.environ.get("MEGAPIPES_BACKEND_URL", "https://meagpipes-1.onrender.com").rstrip("/")

# Cookie names defined in backend security settings
SESSION_COOKIE_NAME = "megapipes_session"
CSRF_COOKIE_NAME = "megapipes_csrf"

# Commercial defaults matching DEV/Mega_Pipes_Quotation_DEV.xlsm SETTINGS
DEFAULT_CALCULATION_MODE = "REFERENCE"
DEFAULT_TRIP_RATE = 98000.0
DEFAULT_STEEL_RATE = 61000.0
DEFAULT_CONVERSION = 4000.0
DEFAULT_EXT_RATE = 500.0
DEFAULT_INT_RATE = 250.0
DEFAULT_HANDLING = 100.0
DEFAULT_FINANCIAL_PCT = 1.0
DEFAULT_GST_PCT = 18.0
DEFAULT_DELIVERY_PIN = "825322"
DEFAULT_RATE_PER_KM = 125.32

# Master Option Lists matching DEV/Mega_Pipes_Quotation_DEV.xlsm sheet QUOTE
PRODUCT_STANDARDS = ["IS-3589-2001", "API 5L", "ASTM A139", "IS 1239 Part-1", "IS 1161"]
PIPE_FAMILIES = ["HSAW", "HFW"]
DIAMETER_BASES = ["ID", "OD"]

STANDARD_SIZES = {
    ("HSAW", "ID"): [1226.0, 1300.0, 1350.0, 1400.0],
    ("HSAW", "OD"): [457.0, 508.0, 610.0, 762.0, 914.0, 1016.0, 1219.0, 1422.0, 1524.0, 1829.0, 2134.0, 2540.0],
    ("HFW", "OD"): [168.3, 219.1, 273.1, 323.9, 355.6, 406.4, 457.2, 508.0],
    ("HFW", "ID"): [168.3, 219.1, 273.1, 323.9, 355.6, 406.4, 457.2, 508.0],
}

STANDARD_THICKNESSES = {
    "HSAW": [6.4, 7.1, 7.9, 8.4, 9.5, 11.1, 12.7, 15.9, 19.1, 25.4],
    "HFW": [3.5, 4.0, 4.5, 5.0, 5.6, 6.4, 7.1, 8.0, 8.8, 10.0, 12.5, 14.2, 16.0],
}

GRADES_BY_FAMILY = {
    "HSAW": [
        "Fe 410",
        "IS 3589 Fe 330",
        "IS 3589 Fe 450",
        "API 5L B",
        "X42 PSL1",
        "X52 PSL1",
        "X60 PSL1",
        "X65 PSL1",
        "X70 PSL1",
        "X52 PSL2",
        "X65 PSL2",
        "X70 PSL2",
        "X80M PSL2",
    ],
    "HFW": [
        "Fe 410",
        "IS 3589 Fe 330",
        "IS 3589 Fe 450",
        "IS 1239 Part-1",
        "IS 1161 YST210",
        "IS 1161 YST240",
        "IS 1161 YST310",
        "API 5L A25",
        "API 5L B",
        "API 5L X42",
        "API 5L X52",
        "API 5L X60",
        "API 5L X65",
        "API 5L X70",
        "API 5L X80M PSL2",
        "API 5CT H40",
        "API 5CT J55",
        "API 5CT K55",
    ],
}

STEEL_GRADES = GRADES_BY_FAMILY["HSAW"]

EXTERNAL_COATINGS = ["Bare", "3LPE", "3LPP", "FBE", "DFBE", "PU", "Zinc Rich Epoxy"]
INTERNAL_COATINGS = ["None", "Liquid Epoxy", "Food Grade Epoxy", "Solvent-Free Epoxy", "PU", "Zinc Rich Epoxy", "Cement Lining"]
QUANTITY_UNITS = ["Metres", "MT", "Nos"]
CUSTOM_PIPE_SIZE_OPTIONS = ["NO", "YES"]
PRICE_UNITS = ["₹/m", "₹/MT"]

DEFAULT_LINE: Dict[str, Any] = {
    "line_number": 1,
    "product_standard": "IS-3589-2001",
    "pipe_family": "HSAW",
    "basis": "ID",
    "size_mm": 508.0,
    "thickness_mm": 6.4,
    "steel_grade": "Fe 410",
    "external_coating": "3LPE",
    "internal_coating": "Liquid Epoxy",
    "quantity": 1000.0,
    "unit": "Metres",
    "pipe_length_m": 12.0,
    "custom_pipe_size": "NO",
    "pipes_per_trailer": 4.0,
    "has_external": True,
    "has_internal": True,
    "steel_rate": DEFAULT_STEEL_RATE,
    "grade_extra": 0.0,
    "yield_pct": 2.0,
    "conversion": DEFAULT_CONVERSION,
    "testing": 0.0,
    "margin": 0.0,
    "ext_rate": DEFAULT_EXT_RATE,
    "int_rate": DEFAULT_INT_RATE,
    "handling": DEFAULT_HANDLING,
    "financial_pct": DEFAULT_FINANCIAL_PCT,
    "gst_pct": DEFAULT_GST_PCT,
    "manual_fc": 0.0,
    "manual_selling_price": 0.0,
    "price_unit": "₹/m",
    "override_reason": "",
}


def format_inr(val: float | None, decimals: int = 2) -> str:
    """Formats a number as Indian Rupee currency (INR)."""
    if val is None:
        return "₹0.00"
    try:
        n = float(val)
        return f"₹{n:,.{decimals}f}"
    except (ValueError, TypeError):
        return "₹0.00"


def format_number(val: float | None, decimals: int = 2) -> str:
    """Formats a float with comma separation."""
    if val is None:
        return "0.00"
    try:
        n = float(val)
        return f"{n:,.{decimals}f}"
    except (ValueError, TypeError):
        return "0.00"
