"""Unit tests for Streamlit application configuration and formatting."""

from __future__ import annotations

import pytest
from streamlit_app.config import format_inr, format_number, DEFAULT_LINE, DEFAULT_TRIP_RATE


def test_format_inr_valid() -> None:
    assert format_inr(0) == "₹0.00"
    assert format_inr(176117500.0) == "₹176,117,500.00"
    assert format_inr(98000.5, 1) == "₹98,000.5"


def test_format_inr_none_or_invalid() -> None:
    assert format_inr(None) == "₹0.00"
    assert format_inr("not_a_number") == "₹0.00"


def test_format_number_valid() -> None:
    assert format_number(1234.567, 2) == "1,234.57"
    assert format_number(0, 0) == "0"
    assert format_number(None) == "0.00"


def test_default_line_schema() -> None:
    assert DEFAULT_LINE["size_mm"] == 508.0
    assert DEFAULT_LINE["thickness_mm"] == 6.4
    assert DEFAULT_LINE["basis"] == "ID"
    assert DEFAULT_LINE["unit"] == "Metres"
    assert DEFAULT_LINE["steel_rate"] == 61000.0
    assert DEFAULT_TRIP_RATE == 98000.0
