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

# The policy of the reference cost sheet ('CS Stacking') - the one the sales
# team actually quotes with.  Pass as calc_line(..., **REFERENCE_POLICY).
REFERENCE_POLICY: dict = {
    "freight_trailers": "FRACTIONAL",
    "int_coating_basis": "OD",
    "fc_method": "ALGEBRAIC",
}


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
    basis: str | None = None,
) -> float:
    """Internal coating cost per MT.

    basis='OD' reproduces the reference cost sheet (row 35 reuses the OD-based
    coating factor of row 20); basis='ID' coats the true inside surface.
    With basis=None the legacy flags decide (backward compatible).
    """
    if basis is None:
        basis = "OD" if (legacy_mode or approved_mode) else "ID"
    basis = basis.upper().strip()
    if basis not in ("OD", "ID"):
        raise CalculationError("internal coating basis must be OD or ID")
    diameter = od_mm if basis == "OD" else inner_d_mm
    return coating_cost_per_mt(rate_rs_per_m2, diameter, mass_mt_per_m_val, legacy_mode=legacy_mode)


def freight_per_mt(*args, **kwargs) -> tuple[int, float]:
    """Backward-compatible wrapper: returns (physical whole trailers, freight Rs/MT)."""
    physical, _trips, freight = freight_detail(*args, **kwargs)
    return physical, freight


def freight_detail(
    trip_rate: float,
    metres: float,
    pipe_length_m: float,
    mass_mt_per_m_val: float,
    pipes_per_trailer: float,
    handling_rs_per_mt: float,
    *,
    legacy_mode: bool = False,
    approved_mode: bool = False,
    whole_trailers: bool | None = None,
    trips_override: float | None = None,
) -> tuple[int, float, float]:
    """Returns (physical_trailers, costed_trips, freight_rs_per_mt).

    physical_trailers = whole trailers that physically move the pipes;
    costed_trips      = the trip count the freight Rs/MT is actually built on
                        (fractional in the reference sheet).

    Reference cost-sheet formula (CS Stacking, rows 57-60, col P):
        trips   = (metres / pipe_length) / pipes_per_trailer      (fractional)
        freight = ROUNDUP(trip_rate * trips / line_MT, 0) + handling

    whole_trailers=True switches `trips` to CEILING(CEILING(metres/len)/ppt).
    trips_override lets the user key trips taken from a manual loading plan
    (e.g. the nesting plan); it wins over both rules.
    When whole_trailers is None the legacy flags decide (backward compatible).
    """
    if min(trip_rate, metres, pipe_length_m, mass_mt_per_m_val, pipes_per_trailer) <= 0:
        raise CalculationError("freight inputs must be positive")
    if whole_trailers is None:
        whole_trailers = not (legacy_mode or approved_mode)
    d_metres = Decimal(str(metres))
    d_len = Decimal(str(pipe_length_m))
    d_per_trailer = Decimal(str(pipes_per_trailer))
    d_trip = Decimal(str(trip_rate))
    d_mass = Decimal(str(mass_mt_per_m_val))
    weight_mt = d_metres * d_mass

    if trips_override is not None and trips_override > 0:
        trips = Decimal(str(trips_override))
    elif whole_trailers:
        pipes = roundup(d_metres / d_len, 0)
        trips = Decimal(str(int(roundup(Decimal(str(pipes)) / d_per_trailer, 0))))
    else:
        trips = (d_metres / d_len) / d_per_trailer
    freight = roundup(d_trip * trips / weight_mt, 0) + handling_rs_per_mt
    physical = int(roundup(Decimal(str(roundup(d_metres / d_len, 0))) / d_per_trailer, 0))
    return physical, float(trips), float(freight)


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
    method: str | None = None,
) -> float:
    """Financial cost (Rs/MT).

    A manual FC (> 0) always wins - it is the typed value in cost-sheet row 32.
    Otherwise:
      ALGEBRAIC (default): FC = k*base/(1-k), k = fin% * (1+GST).
          This is the exact fixed point of the reference sheet's own check
          cell (row 41 = Grand Total incl. GST x fin%), because row 32 feeds
          the Bare price which feeds that Grand Total.
      SIMPLE (old 'approved' mode): FC = ROUND(base * (1+GST) * fin%) - it leaves
          FC itself out of the base, so it under-states FC by ~1.2% of itself
          (Rs 1,019 vs the sheet's Rs 1,031 for Unipro line 1).
    """
    explicit = method is not None
    if method is None:
        method = "SIMPLE" if (legacy_mode or approved_mode) else "ALGEBRAIC"
    method = method.upper().strip()
    # Old flag-only 'normal' mode deliberately ignored typed FC; the explicit
    # policy (and legacy/approved) honour it.
    if explicit:
        honour_manual = manual_fc > 0
    else:
        honour_manual = (legacy_mode or approved_mode) and manual_fc > 100.0
    if honour_manual:
        return float(manual_fc)
    d_fin = Decimal(str(financial_pct)) / Decimal(100)
    d_gst = Decimal(str(gst_pct)) / Decimal(100)
    base = Decimal(str(pre_cost)) + Decimal(str(external)) + Decimal(str(internal)) + Decimal(str(freight))
    if method == "SIMPLE":
        return float(round(float(base * (Decimal(1) + d_gst) * d_fin)))
    if method != "ALGEBRAIC":
        raise CalculationError("fc method must be ALGEBRAIC or SIMPLE")
    k = d_fin * (Decimal(1) + d_gst)
    if not 0 <= k < 1:
        raise CalculationError("financial and GST percentages produce an invalid factor")
    return float(k * base / (Decimal(1) - k))


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
    freight_trips: float = 0.0


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
    freight_trailers: str | None = None,
    int_coating_basis: str | None = None,
    fc_method: str | None = None,
    trips_override: float | None = None,
) -> QuoteLineResult:
    """Calculates a complete quotation line mirroring VBA modPricing.CalcLine.

    Cost policy - pass these three explicitly (the VBA reads them from SETTINGS):
      freight_trailers  'FRACTIONAL' (reference sheet) | 'WHOLE'
      int_coating_basis 'OD' (reference sheet)         | 'ID'
      fc_method         'ALGEBRAIC' (reference fixed point) | 'SIMPLE'
    Any left as None falls back to the old legacy/approved/normal behaviour,
    including the old heuristic (steel_rate==62000 or conversion==4000 =>
    approved).  That heuristic is DEPRECATED: it silently changes the price
    method when somebody happens to type 62000.  Use REFERENCE_POLICY instead.
    """
    explicit_policy = (freight_trailers is not None or int_coating_basis is not None
                       or fc_method is not None)
    if approved_mode is None:
        approved_mode = False if explicit_policy else (steel_rate == 62000.0 or conversion == 4000.0)
    whole_trailers = None if freight_trailers is None else (freight_trailers.upper().strip() == "WHOLE")

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
        internal_coating_per_mt(int_rate, od_mm, inner_d_mm, mass_mt, legacy_mode=legacy_mode,
                                approved_mode=approved_mode, basis=int_coating_basis)
        if has_internal
        else 0.0
    )

    trailers, freight_trips, freight = freight_detail(
        trip_rate, metres, pipe_length_m, mass_mt, pipes_per_trailer, handling,
        legacy_mode=legacy_mode, approved_mode=approved_mode,
        whole_trailers=whole_trailers, trips_override=trips_override,
    )

    fc = financial_cost(
        pre_cost, ext_cost, int_cost, freight, financial_pct, gst_pct,
        legacy_mode=legacy_mode, approved_mode=approved_mode, manual_fc=manual_fc,
        method=fc_method,
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
        freight_trips=freight_trips,
    )


def calc_quote(lines: list[dict], *, legacy_mode: bool = False) -> tuple[list[QuoteLineResult], float]:
    """Calculates a multi-line quotation and returns line results and grand total."""
    results = [calc_line(**line_kwargs, legacy_mode=legacy_mode) for line_kwargs in lines]
    grand_total = sum(r.line_value for r in results)
    return results, grand_total


@dataclass(frozen=True)
class CostSheet:
    """Whole-quotation result laid out like the reference cost sheet."""
    lines: list
    total_metres: float
    total_mt: float
    total_trips: float
    subtotal_excl_gst: float
    gst: float
    grand_total: float
    # 'total' column (G) of the sheet = SUMPRODUCT(row, line MT) / total MT
    weighted: dict


def _mt_weighted(values: list[float], weights: list[float]) -> float:
    total_w = sum(weights)
    return sum(v * w for v, w in zip(values, weights)) / total_w if total_w else 0.0


def calc_costsheet(
    lines: list[dict],
    *,
    pooled_trips: float | None = None,
    gst_pct: float = 18.0,
    policy: dict | None = None,
) -> CostSheet:
    """Calculate a whole quotation the way the reference cost sheet does.

    lines        list of calc_line keyword dicts (size_mm, thickness_mm, basis,
                 quantity, steel_rate, manual_fc, ... ).  Every rate may differ
                 per line, exactly like a column of the sheet.
    pooled_trips manual loading-plan trips for the WHOLE quote (the 'nesting'
                 sheet).  Freight is then one pooled figure
                 ROUNDUP(trip_rate*pooled_trips/total_MT)+handling for all
                 lines.  Leave None for the stacking method (per-line trips).
    policy       defaults to REFERENCE_POLICY.
    """
    pol = dict(REFERENCE_POLICY if policy is None else policy)
    prepared = []
    for raw in lines:
        kw = {**raw, **pol}
        kw.setdefault("gst_pct", gst_pct)
        prepared.append(kw)

    if pooled_trips:
        mts = []
        for kw in prepared:
            od, _id, _kg = geometry(kw["size_mm"], kw["thickness_mm"], kw["basis"])
            mass = mass_mt_per_m(od, kw["thickness_mm"])
            metres, _p, wt = quantity_conversion(kw["quantity"], kw.get("unit", "Metres"), mass,
                                                 kw.get("pipe_length_m", 12.0))
            mts.append(wt)
        total_mt_pre = sum(mts)
        for kw, wt in zip(prepared, mts):
            kw["trips_override"] = pooled_trips * wt / total_mt_pre

    results = [calc_line(**kw) for kw in prepared]
    mts = [r.weight_mt for r in results]
    subtotal = sum(r.line_value for r in results)
    gst = subtotal * gst_pct / 100.0
    weighted = {
        "financial_cost_per_mt": _mt_weighted([r.financial_cost_per_mt for r in results], mts),
        "freight_per_mt": _mt_weighted([r.freight_per_mt for r in results], mts),
        "external_per_mt": _mt_weighted([r.external_cost_per_mt for r in results], mts),
        "internal_per_mt": _mt_weighted([r.internal_cost_per_mt for r in results], mts),
        "bare_per_mt": _mt_weighted([r.pre_cost + r.financial_cost_per_mt for r in results], mts),
        "subtotal_per_mt": subtotal / sum(mts) if sum(mts) else 0.0,
    }
    return CostSheet(
        lines=results,
        total_metres=sum(r.metres for r in results),
        total_mt=sum(mts),
        total_trips=sum(r.freight_trips for r in results),
        subtotal_excl_gst=subtotal,
        gst=gst,
        grand_total=subtotal + gst,
        weighted=weighted,
    )
