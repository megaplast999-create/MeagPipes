"""Unit tests for pure-Python calculation engine."""

import pytest
from calculation_engine import (
    CalculationError,
    PI_CONSTANT,
    coating_cost_per_mt,
    external_coating_per_mt,
    financial_cost,
    freight_per_mt,
    geometry,
    internal_coating_per_mt,
    mass_mt_per_m,
    quantity_conversion,
    roundup,
    selling_price,
    selling_price_per_m,
    calc_line,
)


def test_roundup_basic():
    assert roundup(1226 + 2 * 8.4, 0) == 1243
    assert roundup(17593.4478, 0) == 17594
    assert roundup(1951.51, 0) == 1952


def test_mass_mt_per_m():
    # OD=1243, WT=8.4
    # (1243 - 8.4) * 8.4 * 0.0246615 / 1000 = 0.25575553836
    m = mass_mt_per_m(1243, 8.4)
    assert round(m, 11) == round(0.25575553836, 11)


def test_geometry_od():
    od, inner, mass = geometry(1243, 8.4, "OD")
    assert od == 1243
    assert inner == 1226.2
    assert round(mass, 8) == round(255.75553836, 8)


def test_geometry_id_rounds_od_up():
    od, inner, mass = geometry(1226, 8.4, "ID")
    assert od == 1243
    assert inner == 1226
    assert round(mass, 8) == round(255.75553836, 8)


def test_quantity_conversion_preserves_exact_metres():
    metres, pipes, weight = quantity_conversion(2000, "Metres", 0.25575553836, 12)
    assert metres == 2000
    assert pipes == 167
    assert round(weight, 6) == round(511.51107672, 6)


def test_coating_cost_is_rate_area_over_mass():
    # Corrected formula: PI = 3.14, mass in MT/m (0.25575553836)
    result = coating_cost_per_mt(500, 1243, 0.25575553836)
    expected = (500 * 3.14 * 1.243) / 0.25575553836
    assert round(result, 6) == round(expected, 6)
    assert round(result, 6) == round(7630.372395897313, 6)


def test_legacy_coating_cost_uses_mass_mt_per_m():
    result = coating_cost_per_mt(500, 1243, 0.25575553836, legacy_mode=True)
    assert round(result, 6) == round(7630.372395897313, 6)


def test_internal_coating_normal_uses_inner_d():
    # Normal mode uses InnerD (1226)
    result = internal_coating_per_mt(275, 1243, 1226, 0.25575553836, legacy_mode=False)
    expected = (275 * 3.14 * 1.226) / 0.25575553836
    assert round(result, 6) == round(expected, 6)
    assert round(result, 6) == round(4139.308211225709, 6)


def test_internal_coating_legacy_uses_od():
    # Legacy mode uses OD (1243)
    result = internal_coating_per_mt(275, 1243, 1226, 0.25575553836, legacy_mode=True)
    expected = (275 * 3.14 * 1.243) / 0.25575553836
    assert round(result, 6) == round(expected, 6)
    assert round(result, 6) == round(4196.704817743521, 6)


def test_legacy_freight_uses_fractional_trailers():
    trailers, result = freight_per_mt(
        98000,
        2000,
        12,
        0.25575553836,
        4,
        100,
        legacy_mode=True,
    )
    assert trailers == 42
    assert result == 8083


def test_normal_freight_uses_whole_trailers():
    trailers, result = freight_per_mt(
        125.32 * 782,
        2000,
        12,
        0.25575553836,
        4,
        100,
        legacy_mode=False,
    )
    assert trailers == 42
    assert result == 8147


def test_financial_cost_normal_algebraic():
    # pre=67220, ext=7630.37..., int=4139.30..., fr=8147
    ext = (500 * 3.14 * 1.243) / 0.25575553836
    intc = (275 * 3.14 * 1.226) / 0.25575553836
    fc = financial_cost(67220, ext, intc, 8147, 1.5, 18.0)
    assert round(fc, 4) == round(1570.1102, 4)


def test_financial_cost_legacy_manual():
    fc = financial_cost(67220, 1000, 500, 8000, 1.5, 18.0, legacy_mode=True, manual_fc=1542)
    assert fc == 1542


def test_selling_price_uses_positive_rounding():
    result = selling_price(61000.2, 100.1, 500.2, 275.1, 8000.1)
    assert result.total_per_mt == 69879


def test_selling_price_per_m_components():
    # Line 1 normal: pre=67220, fc=1570.1102, ext=7630.37..., int=4139.30..., fr=8147
    ext = (500 * 3.14 * 1.243) / 0.25575553836
    intc = (275 * 3.14 * 1.226) / 0.25575553836
    fc = 1570.110197237175
    mass = 0.25575553836
    pm = selling_price_per_m(67220, fc, ext, intc, 8147, mass, 2000)
    assert pm.bare_per_m == 17594
    assert pm.external_per_m == 1952
    assert pm.internal_per_m == 1059
    assert pm.freight_per_m == 2084
    assert pm.total_per_m == 22689
    assert pm.line_value == 45378000


def test_invalid_geometry_is_rejected():
    with pytest.raises(CalculationError):
        geometry(100, 60, "OD")


def test_negative_wall_thickness_rejected():
    with pytest.raises(CalculationError):
        geometry(1226, -8.4, "ID")


def test_invalid_unit_rejected():
    with pytest.raises(CalculationError):
        quantity_conversion(100, "INVALID_UNIT", 0.25, 12)
