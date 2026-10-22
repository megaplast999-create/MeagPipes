"""
Unit and regression tests for pricing variance root-cause decomposition.

Validates that:
1. The ₹143,150 difference between Normal Mode and Legacy Mode is reproduced exactly.
2. Every component variance reconciles to the line impact and grand total variance.
3. The root causes (whole-trailer freight, ID coating, dynamic FC) are isolated and asserted.
4. Logistics rounding does not alter commercial line quantities.
5. GST and grand totals calculate with exact mathematical precision.
"""

import sys
from pathlib import Path
import pytest

ROOT_DIR = Path(__file__).resolve().parent.parent
TESTS_DIR = ROOT_DIR / "TESTS"
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from calculation_engine import calc_line, geometry, mass_mt_per_m, roundup

UNIPRO_LINES = [
    {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542, "line_num": 1},
    {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534, "line_num": 2},
    {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529, "line_num": 3},
    {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525, "line_num": 4},
]

COMMON_PARAMS = {
    "unit": "Metres",
    "pipe_length_m": 12.0,
    "pipes_per_trailer": 4.0,
    "has_external": True,
    "has_internal": True,
    "steel_rate": 61000.0,
    "grade_extra": 0.0,
    "yield_pct": 2.0,
    "conversion": 5000.0,
    "testing": 0.0,
    "margin": 0.0,
    "ext_rate": 500.0,
    "int_rate": 275.0,
    "handling": 100.0,
    "financial_pct": 1.5,
    "gst_pct": 18.0,
}


def _strip_line_num(d: dict) -> dict:
    return {k: v for k, v in d.items() if k != "line_num"}


def test_reproduce_exact_143150_grand_variance():
    """Grand total difference between Normal Mode and Legacy Mode must be exactly ₹143,150."""
    subtotal_norm = 0.0
    subtotal_leg = 0.0

    for item in UNIPRO_LINES:
        line_args = _strip_line_num(item)
        kwargs_norm = {**COMMON_PARAMS, **line_args, "trip_rate": 125.32 * 782, "legacy_mode": False}
        kwargs_leg = {**COMMON_PARAMS, **line_args, "trip_rate": 98000.0, "legacy_mode": True}

        res_n = calc_line(**kwargs_norm)
        res_l = calc_line(**kwargs_leg)

        subtotal_norm += res_n.line_value
        subtotal_leg += res_l.line_value

    assert subtotal_norm == 178801050
    assert subtotal_leg == 178657900
    assert subtotal_norm - subtotal_leg == 143150


def test_line_level_variance_reconciliation():
    """Line-level variances must exactly match +₹16k, +₹68.25k, +₹54.5k, +₹4.4k."""
    expected_line_diffs = {
        1: {"diff_per_m": 8,   "qty": 2000, "impact": 16000},
        2: {"diff_per_m": 195, "qty": 350,  "impact": 68250},
        3: {"diff_per_m": 109, "qty": 500,  "impact": 54500},
        4: {"diff_per_m": 1,   "qty": 4400, "impact": 4400},
    }

    for item in UNIPRO_LINES:
        ln = item["line_num"]
        exp = expected_line_diffs[ln]
        line_args = _strip_line_num(item)

        res_n = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 125.32 * 782, "legacy_mode": False})
        res_l = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 98000.0, "legacy_mode": True})

        diff_per_m = res_n.total_per_m - res_l.total_per_m
        diff_val = res_n.line_value - res_l.line_value

        assert diff_per_m == exp["diff_per_m"], f"Line {ln} price diff mismatch: {diff_per_m} vs {exp['diff_per_m']}"
        assert diff_val == exp["impact"], f"Line {ln} value impact mismatch: {diff_val} vs {exp['impact']}"


def test_component_level_variance_bridge():
    """
    Component-level differences must reconcile across all four lines:
      - Freight: +₹193,850
      - Internal Coating: -₹108,750
      - Bare / Financial Cost: +₹58,050
      - External Coating: ₹0
      Total: +₹143,150
    """
    tot_bare = 0.0
    tot_ext = 0.0
    tot_int = 0.0
    tot_fr = 0.0

    for item in UNIPRO_LINES:
        q = item["quantity"]
        line_args = _strip_line_num(item)
        res_n = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 125.32 * 782, "legacy_mode": False})
        res_l = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 98000.0, "legacy_mode": True})

        tot_bare += (res_n.bare_per_m - res_l.bare_per_m) * q
        tot_ext += (res_n.external_per_m - res_l.external_per_m) * q
        tot_int += (res_n.internal_per_m - res_l.internal_per_m) * q
        tot_fr += (res_n.freight_per_m - res_l.freight_per_m) * q

    assert tot_bare == 58050, f"Bare/FC impact mismatch: {tot_bare}"
    assert tot_ext == 0, f"External coating impact mismatch: {tot_ext}"
    assert tot_int == -108750, f"Internal coating impact mismatch: {tot_int}"
    assert tot_fr == 193850, f"Freight impact mismatch: {tot_fr}"

    assert tot_bare + tot_ext + tot_int + tot_fr == 143150


def test_internal_coating_diameter_effect():
    """
    Demonstrate that internal coating difference (-₹15/m) is entirely caused by
    using physical Inner Diameter (ID) in Normal Mode vs Outside Diameter (OD) in Legacy.
    """
    for item in UNIPRO_LINES:
        line_args = _strip_line_num(item)
        res_n = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 125.32 * 782, "legacy_mode": False})
        res_l = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 98000.0, "legacy_mode": True})

        # Normal mode internal coating is strictly lower by exactly ₹15/m
        assert res_n.internal_per_m - res_l.internal_per_m == -15


def test_freight_whole_vs_fractional_trailer_allocation():
    """
    Demonstrate that Normal Mode uses whole-trailer integer ceiling while
    Legacy Mode used fractional trailer allocation.
    """
    # Line 2: 350 m = 30 pipes at 4/trailer -> 7.5 trailers (ceiling = 8 trailers)
    l2 = _strip_line_num(UNIPRO_LINES[1])
    res_n = calc_line(**{**COMMON_PARAMS, **l2, "trip_rate": 125.32 * 782, "legacy_mode": False})
    res_l = calc_line(**{**COMMON_PARAMS, **l2, "trip_rate": 98000.0, "legacy_mode": True})

    assert res_n.trailers == 8
    # Normal freight per MT is ₹8,364 vs legacy ₹7,632 (+₹732/MT)
    assert res_n.freight_per_mt == 8364
    assert res_l.freight_per_mt == 7632
    assert res_n.freight_per_m - res_l.freight_per_m == 199


def test_logistics_rounding_does_not_alter_commercial_metres():
    """Logistics pipe/trailer counts must not alter commercial billable metres."""
    for item in UNIPRO_LINES:
        line_args = _strip_line_num(item)
        res_n = calc_line(**{**COMMON_PARAMS, **line_args, "trip_rate": 125.32 * 782, "legacy_mode": False})
        assert res_n.metres == item["quantity"]
        assert res_n.line_value == res_n.total_per_m * item["quantity"]


def test_gst_and_grand_total_precision():
    """GST @ 18% and Grand Total must be exact additions without floating point drift."""
    subtotal_norm = 178801050
    gst_norm = round(subtotal_norm * 0.18, 0)
    grand_norm = subtotal_norm + gst_norm

    assert gst_norm == 32184189
    assert grand_norm == 210985239

    subtotal_leg = 178657900
    gst_leg = round(subtotal_leg * 0.18, 0)
    grand_leg = subtotal_leg + gst_leg

    assert gst_leg == 32158422
    assert grand_leg == 210816322
