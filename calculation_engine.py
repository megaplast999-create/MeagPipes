"""Pure-Python reference calculations for independent validation.

This module is intentionally independent of Excel and VBA. It implements
the authoritative geometry, quantity, coating, freight, financial, and
selling price rules defined in DOCS/ANTIGRAVITY_MASTER_PROMPT.md.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_CEILING
from math import ceil

PI_CONSTANT: float = 3.14


class CalculationError(ValueError):
    """Raised when documented input constraints are violated."""


def roundup(value: float | Decimal, digits: int = 0) -> float:
    """Excel-style ROUNDUP: always rounds away from zero."""
    factor = Decimal(10) ** digits
    d_val = Decimal(str(value))
    return float((d_val * factor).to_integral_value(rounding=ROUND_CEILING) / factor)


def mass_mt_per_m(od_mm: float, thickness_mm: float) -> float:
    """Calculates pipe mass in metric tonnes per metre (full precision)."""
    if od_mm <= 0 or thickness_mm <= 0:
        raise CalculationError("OD and wall thickness must be positive")
    d_od = Decimal(str(od_mm))
    d_wt = Decimal(str(thickness_mm))
    mass_mt = ((d_od - d_wt) * d_wt * Decimal("0.0246615")) / Decimal(1000)
    return float(mass_mt)


def geometry(size_mm: float, thickness_mm: float, basis: str) -> tuple[float, float, float]:
    """Calculates OD, ID, and mass in kg/metre from specified basis."""
    if size_mm <= 0 or thickness_mm <= 0:
        raise CalculationError("size and thickness must be positive")
    normalized_basis = basis.upper().strip()
    if normalized_basis == "OD":
        od_mm = float(size_mm)
        id_mm = od_mm - 2 * thickness_mm
    elif normalized_basis == "ID":
        id_mm = float(size_mm)
        od_mm = roundup(id_mm + 2 * thickness_mm, 0)
    else:
        raise CalculationError("basis must be OD or ID")
    if id_mm <= 0:
        raise CalculationError("calculated ID must be positive")
    if id_mm <= 2 * thickness_mm:
        raise CalculationError("inner diameter must be greater than twice wall thickness")
    kg_per_m = float(((Decimal(str(od_mm)) - Decimal(str(thickness_mm))) * Decimal(str(thickness_mm)) * Decimal("0.0246615")))
    return od_mm, id_mm, kg_per_m


def quantity_conversion(
    quantity: float,
    unit: str,
    mass_mt_per_m_val: float,
    pipe_length_m: float,
) -> tuple[float, int, float]:
    """Converts enquiry quantity to metres, pipe count, and total weight in MT."""
    if quantity <= 0 or mass_mt_per_m_val <= 0:
        raise CalculationError("quantity and mass must be positive")
    if not 6 <= pipe_length_m <= 13:
        raise CalculationError("pipe length must be between 6 and 13 metres")
    normalized_unit = unit.upper().strip()
    if normalized_unit == "METRES":
        metres = float(quantity)
    elif normalized_unit == "MT":
        metres = float(quantity) / mass_mt_per_m_val
    elif normalized_unit == "NOS":
        metres = float(quantity) * pipe_length_m
    else:
        raise CalculationError("unit must be Metres, MT, or Nos")
    pipe_count = int(roundup(metres / pipe_length_m, 0))
    weight_mt = metres * mass_mt_per_m_val
    return metres, pipe_count, weight_mt


def coating_cost_per_mt(
    rate_rs_per_m2: float,
    diameter_mm: float,
    mass_mt_per_m_val: float,
    *,
    legacy_mode: bool = False,
) -> float:
    """Calculates coating cost per MT using PI=3.14 and mass in MT/metre."""
    if rate_rs_per_m2 < 0 or diameter_mm <= 0 or mass_mt_per_m_val <= 0:
        raise CalculationError("coating rate, diameter, and mass are invalid")
    # If caller provided mass in kg/m (>10), convert to MT/m
    mt = mass_mt_per_m_val / 1000.0 if mass_mt_per_m_val > 10.0 else mass_mt_per_m_val
    d_rate = Decimal(str(rate_rs_per_m2))
    d_pi = Decimal(str(PI_CONSTANT))
    d_dia = Decimal(str(diameter_mm))
    d_mass = Decimal(str(mt))
    cost = (d_rate * d_pi * (d_dia / Decimal(1000))) / d_mass
    return float(cost)


def external_coating_per_mt(
    rate_rs_per_m2: float,
    od_mm: float,
    mass_mt_per_m_val: float,
) -> float:
    """Calculates external coating cost per MT using outer diameter."""
    return coating_cost_per_mt(rate_rs_per_m2, od_mm, mass_mt_per_m_val)


def internal_coating_per_mt(
    rate_rs_per_m2: float,
    od_mm: float,
    inner_d_mm: float,
    mass_mt_per_m_val: float,
    *,
    legacy_mode: bool = False,
    approved_mode: bool = False,
) -> float:
    """Calculates internal coating cost per MT (uses InnerD for Option 1 normal, OD for legacy or approved)."""
    diameter = od_mm if (legacy_mode or approved_mode) else inner_d_mm
    return coating_cost_per_mt(rate_rs_per_m2, diameter, mass_mt_per_m_val, legacy_mode=legacy_mode)


def freight_per_mt(
    trip_rate: float,
    metres: float,
    pipe_length_m: float,
    mass_mt_per_m_val: float,
    pipes_per_trailer: float,
    handling_rs_per_mt: float,
    *,
    legacy_mode: bool = False,
    approved_mode: bool = False,
) -> tuple[int, float]:
    """Calculates required trailer count and freight cost per MT."""
    if min(trip_rate, metres, pipe_length_m, mass_mt_per_m_val, pipes_per_trailer) <= 0:
        raise CalculationError("freight inputs must be positive")
    d_metres = Decimal(str(metres))
    d_len = Decimal(str(pipe_length_m))
    d_per_trailer = Decimal(str(pipes_per_trailer))
    d_trip = Decimal(str(trip_rate))
    d_mass = Decimal(str(mass_mt_per_m_val))
    weight_mt = d_metres * d_mass

    pipes = roundup(d_metres / d_len, 0)
    trailers = int(roundup(Decimal(str(pipes)) / d_per_trailer, 0))

    if legacy_mode or approved_mode:
        fractional_trailers = (d_metres / d_len) / d_per_trailer
        freight = roundup(d_trip * fractional_trailers / weight_mt, 0) + handling_rs_per_mt
    else:
        freight = roundup(d_trip * Decimal(trailers) / weight_mt, 0) + handling_rs_per_mt
    return trailers, float(freight)


def financial_cost(
    pre_cost: float,
    external: float,
    internal: float,
    freight: float,
    financial_pct: float,
    gst_pct: float,
    *,
    legacy_mode: bool = False,
    approved_mode: bool = False,
    manual_fc: float = 0.0,
) -> float:
    """Calculates financial cost algebraically in normal option 1 mode, or manual override / approved formula in legacy or approved mode."""
    if legacy_mode or approved_mode:
        if manual_fc > 100.0:
            return float(manual_fc)
        d_fin = Decimal(str(financial_pct)) / Decimal(100)
        d_gst = Decimal(str(gst_pct)) / Decimal(100)
        base = Decimal(str(pre_cost)) + Decimal(str(external)) + Decimal(str(internal)) + Decimal(str(freight))
        # Approved Excel CostSheet formula: ROUND((PreCost + Ext + Int + Freight) * (1 + GST) * FinPct, 0)
        fc = round(float(base * (Decimal(1) + d_gst) * d_fin))
        return float(fc)
    d_fin = Decimal(str(financial_pct))
    d_gst = Decimal(str(gst_pct))
    k = (d_fin / Decimal(100)) * (Decimal(1) + d_gst / Decimal(100))
    if not 0 <= k < 1:
        raise CalculationError("financial and GST percentages produce an invalid factor")
    base = Decimal(str(pre_cost)) + Decimal(str(external)) + Decimal(str(internal)) + Decimal(str(freight))
    fc = k * base / (Decimal(1) - k)
    return float(fc)


@dataclass(frozen=True)
class SellingPrice:
    """Legacy container for per-MT selling price components."""
    bare_per_mt: float
    external_per_mt: float
    internal_per_mt: float
    freight_per_mt: float
    total_per_mt: float


def selling_price(pre_cost: float, financial: float, external: float, internal: float, freight: float) -> SellingPrice:
    """Calculates per-MT selling price with upward rounding."""
    bare = roundup(pre_cost + financial)
    ext = roundup(external)
    intc = roundup(internal)
    fr = roundup(freight)
    return SellingPrice(bare, ext, intc, fr, bare + ext + intc + fr)


@dataclass(frozen=True)
class SellingPricePerM:
    """Per-metre selling price components and commercial line value."""
    bare_per_m: float
    external_per_m: float
    internal_per_m: float
    freight_per_m: float
    total_per_m: float
    line_value: float


def selling_price_per_m(
    pre_cost: float,
    financial_cost_val: float,
    external_per_mt: float,
    internal_per_mt: float,
    freight_per_mt_val: float,
    mass_mt_per_m_val: float,
    metres: float,
) -> SellingPricePerM:
    """Calculates per-metre selling prices rounded UP to nearest whole rupee, then summed."""
    d_pre = Decimal(str(pre_cost))
    d_fc = Decimal(str(financial_cost_val))
    d_ext = Decimal(str(external_per_mt))
    d_int = Decimal(str(internal_per_mt))
    d_fr = Decimal(str(freight_per_mt_val))
    d_mass = Decimal(str(mass_mt_per_m_val))
    d_metres = Decimal(str(metres))

    bare = float(roundup((d_pre + d_fc) * d_mass, 0))
    ext = float(roundup(d_ext * d_mass, 0))
    intc = float(roundup(d_int * d_mass, 0))
    fr = float(roundup(d_fr * d_mass, 0))
    total = bare + ext + intc + fr
    line_val = float(Decimal(str(total)) * d_metres)
    return SellingPricePerM(bare, ext, intc, fr, total, line_val)


@dataclass(frozen=True)
class QuoteLineResult:
    """Comprehensive calculated results for a single quotation line."""
    od_mm: float
    inner_d_mm: float
    mass_mt_per_m: float
    mass_kg_per_m: float
    metres: float
    weight_mt: float
    pipes: int
    trailers: int
    pre_cost: float
    external_cost_per_mt: float
    internal_cost_per_mt: float
    freight_per_mt: float
    financial_cost_per_mt: float
    bare_per_m: float
    external_per_m: float
    internal_per_m: float
    freight_per_m: float
    total_per_m: float
    total_per_mt: float
    line_value: float
    # Dynamic Staged Costing & Override Fields
    calculated_unit_price: float = 0.0
    applied_unit_price: float = 0.0
    price_source: str = "CALCULATED"
    price_variance: float = 0.0
    override_reason: str = ""
    raw_material_cost_per_mt: float = 0.0
    bare_pipe_cost_per_mt: float = 0.0


def calc_line(
    size_mm: float,
    thickness_mm: float,
    basis: str,
    quantity: float,
    unit: str = "Metres",
    pipe_length_m: float = 12.0,
    pipes_per_trailer: float = 4.0,
    has_external: bool = True,
    has_internal: bool = True,
    steel_rate: float = 61000.0,
    grade_extra: float = 0.0,
    yield_pct: float = 2.0,
    conversion: float = 5000.0,
    testing: float = 0.0,
    margin: float = 0.0,
    ext_rate: float = 500.0,
    int_rate: float = 275.0,
    trip_rate: float = 98000.0,
    handling: float = 100.0,
    financial_pct: float = 1.5,
    gst_pct: float = 18.0,
    manual_fc: float = 0.0,
    legacy_mode: bool = False,
    price_by_mt: bool = False,
    approved_mode: bool | None = None,
    manual_selling_price: float = 0.0,
    override_reason: str = "",
) -> QuoteLineResult:
    """Calculates a complete quotation line mirroring VBA modPricing.CalcLine."""
    if approved_mode is None:
        approved_mode = (steel_rate == 62000.0 or conversion == 4000.0)

    od_mm, inner_d_mm, mass_kg = geometry(size_mm, thickness_mm, basis)
    mass_mt = mass_mt_per_m(od_mm, thickness_mm)

    metres, pipes, weight_mt = quantity_conversion(quantity, unit, mass_mt, pipe_length_m)

    d_steel = Decimal(str(steel_rate))
    d_grade = Decimal(str(grade_extra))
    d_yield = Decimal(str(yield_pct))
    d_conv = Decimal(str(conversion))
    d_test = Decimal(str(testing))
    d_margin = Decimal(str(margin))

    delivered_steel_d = d_steel + d_grade
    yield_loss_d = delivered_steel_d * (d_yield / Decimal(100))
    raw_material_d = delivered_steel_d + yield_loss_d
    pre_d = raw_material_d + d_conv + d_test + d_margin
    pre_cost = float(pre_d)

    ext_cost = external_coating_per_mt(ext_rate, od_mm, mass_mt) if has_external else 0.0
    int_cost = (
        internal_coating_per_mt(int_rate, od_mm, inner_d_mm, mass_mt, legacy_mode=legacy_mode, approved_mode=approved_mode)
        if has_internal
        else 0.0
    )

    trailers, freight = freight_per_mt(
        trip_rate, metres, pipe_length_m, mass_mt, pipes_per_trailer, handling,
        legacy_mode=legacy_mode, approved_mode=approved_mode
    )

    fc = financial_cost(
        pre_cost, ext_cost, int_cost, freight, financial_pct, gst_pct,
        legacy_mode=legacy_mode, approved_mode=approved_mode, manual_fc=manual_fc
    )

    per_m = selling_price_per_m(pre_cost, fc, ext_cost, int_cost, freight, mass_mt, metres)

    d_total_mt = (
        roundup(pre_cost + fc, 0)
        + roundup(ext_cost, 0)
        + roundup(int_cost, 0)
        + freight
    )
    total_per_mt = float(d_total_mt)

    calc_price = total_per_mt if price_by_mt else per_m.total_per_m

    if manual_selling_price > 0:
        applied_price = float(manual_selling_price)
        price_src = "MANUAL_OVERRIDE"
        price_var = applied_price - calc_price
        line_value = float(Decimal(str(applied_price)) * Decimal(str(weight_mt))) if price_by_mt else float(Decimal(str(applied_price)) * Decimal(str(metres)))
    else:
        applied_price = calc_price
        price_src = "CALCULATED"
        price_var = 0.0
        line_value = float(Decimal(str(total_per_mt)) * Decimal(str(weight_mt))) if price_by_mt else per_m.line_value

    return QuoteLineResult(
        od_mm=od_mm,
        inner_d_mm=inner_d_mm,
        mass_mt_per_m=mass_mt,
        mass_kg_per_m=mass_kg,
        metres=metres,
        weight_mt=weight_mt,
        pipes=pipes,
        trailers=trailers,
        pre_cost=pre_cost,
        external_cost_per_mt=ext_cost,
        internal_cost_per_mt=int_cost,
        freight_per_mt=freight,
        financial_cost_per_mt=fc,
        bare_per_m=per_m.bare_per_m,
        external_per_m=per_m.external_per_m,
        internal_per_m=per_m.internal_per_m,
        freight_per_m=per_m.freight_per_m,
        total_per_m=per_m.total_per_m,
        total_per_mt=total_per_mt,
        line_value=line_value,
        calculated_unit_price=calc_price,
        applied_unit_price=applied_price,
        price_source=price_src,
        price_variance=price_var,
        override_reason=override_reason,
        raw_material_cost_per_mt=float(raw_material_d),
        bare_pipe_cost_per_mt=pre_cost,
    )


def calc_quote(lines: list[dict], *, legacy_mode: bool = False) -> tuple[list[QuoteLineResult], float]:
    """Calculates a multi-line quotation and returns line results and grand total."""
    results = [calc_line(**line_kwargs, legacy_mode=legacy_mode) for line_kwargs in lines]
    grand_total = sum(r.line_value for r in results)
    return results, grand_total
