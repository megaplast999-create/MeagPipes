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

DEFAULT_LINE: Dict[str, Any] = {
    "line_number": 1,
    "size_mm": 508.0,
    "thickness_mm": 6.4,
    "basis": "ID",
    "quantity": 1000.0,
    "unit": "Metres",
    "pipe_length_m": 12.0,
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
