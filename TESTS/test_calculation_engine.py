from calculation_engine import (
    CalculationError,
    coating_cost_per_mt,
    financial_cost,
    freight_per_mt,
    geometry,
    quantity_conversion,
    selling_price,
)


def test_geometry_od():
    od, inner, mass = geometry(1243, 8.4, "OD")
    assert od == 1243
    assert inner == 1226.2
    assert round(mass, 8) == round(255.75553836, 8)


def test_geometry_id_rounds_od_up():
    od, inner, _ = geometry(1226, 8.4, "ID")
    assert od == 1243
    assert inner == 1226


def test_quantity_conversion_preserves_exact_metres():
    metres, pipes, weight = quantity_conversion(2000, "Metres", 0.25575553836, 12)
    assert metres == 2000
    assert pipes == 167
    assert round(weight, 6) == round(511.51107672, 6)


def test_coating_cost_is_rate_area_over_mass():
    result = coating_cost_per_mt(500, 1243, 255.75553836)
    assert round(result, 6) == round(500 * 3.141592653589793 * 1.243 / 255.75553836, 6)


def test_legacy_coating_cost_uses_mass_mt_per_m():
    result = coating_cost_per_mt(500, 1243, 255.75553836, legacy_mode=True)
    assert round(result, 6) == round(7630.372395897313, 6)


def test_legacy_freight_uses_fractional_trailers():
    _, result = freight_per_mt(
        98000,
        2000,
        12,
        0.25575553836,
        4,
        100,
        legacy_mode=True,
    )
    assert result == 8083


def test_financial_cost():
    result = financial_cost(61000, 1000, 500, 8000, 1.5, 18)
    assert result > 0


def test_selling_price_uses_positive_rounding():
    result = selling_price(61000.2, 100.1, 500.2, 275.1, 8000.1)
    assert result.total_per_mt == 69879


def test_invalid_geometry_is_rejected():
    try:
        geometry(100, 60, "OD")
    except CalculationError:
        pass
    else:
        raise AssertionError("invalid geometry was accepted")
