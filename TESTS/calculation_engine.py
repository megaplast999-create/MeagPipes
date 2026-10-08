"""Pure-Python reference calculations for independent validation.

This module is intentionally independent of Excel and VBA. It implements only
the documented geometry, quantity, coating, freight, financial, and selling
price rules.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_CEILING
from math import ceil, pi


class CalculationError(ValueError):
    """Raised when documented input constraints are violated."""


def roundup(value: float, digits: int = 0) -> float:
    factor = Decimal(10) ** digits
    return float((Decimal(str(value)) * factor).to_integral_value(rounding=ROUND_CEILING) / factor)


def geometry(size_mm: float, thickness_mm: float, basis: str) -> tuple[float, float, float]:
    if size_mm <= 0 or thickness_mm <= 0:
        raise CalculationError("size and thickness must be positive")
    normalized_basis = basis.upper()
    if normalized_basis == "OD":
        od_mm = size_mm
        id_mm = od_mm - 2 * thickness_mm
    elif normalized_basis == "ID":
        id_mm = size_mm
        od_mm = roundup(id_mm + 2 * thickness_mm)
    else:
        raise CalculationError("basis must be OD or ID")
    if id_mm <= 0:
        raise CalculationError("calculated ID must be positive")
    kg_per_m = (od_mm - thickness_mm) * thickness_mm * 0.0246615
    return od_mm, id_mm, kg_per_m


def quantity_conversion(
    quantity: float,
    unit: str,
    mass_mt_per_m: float,
    pipe_length_m: float,
) -> tuple[float, int, float]:
    if quantity <= 0 or mass_mt_per_m <= 0:
        raise CalculationError("quantity and mass must be positive")
    if not 6 <= pipe_length_m <= 13:
        raise CalculationError("pipe length must be between 6 and 13 metres")
    normalized_unit = unit.upper()
    if normalized_unit == "METRES":
        metres = quantity
    elif normalized_unit == "MT":
        metres = quantity / mass_mt_per_m
    elif normalized_unit == "NOS":
        metres = quantity * pipe_length_m
    else:
        raise CalculationError("unit must be Metres, MT, or Nos")
    pipe_count = ceil(metres / pipe_length_m)
    return metres, pipe_count, metres * mass_mt_per_m


def coating_cost_per_mt(rate_rs_per_m2: float, diameter_mm: float, mass_kg_per_m: float) -> float:
    if rate_rs_per_m2 < 0 or diameter_mm <= 0 or mass_kg_per_m <= 0:
        raise CalculationError("coating rate, diameter, and mass are invalid")
    return rate_rs_per_m2 * pi * diameter_mm / 1000 / mass_kg_per_m


def freight_per_mt(
    trip_rate: float,
    metres: float,
    pipe_length_m: float,
    mass_mt_per_m: float,
    pipes_per_trailer: float,
    handling_rs_per_mt: float,
) -> tuple[int, float]:
    if min(trip_rate, metres, pipe_length_m, mass_mt_per_m, pipes_per_trailer) <= 0:
        raise CalculationError("freight inputs must be positive")
    trailers = ceil((metres / pipe_length_m) / pipes_per_trailer)
    freight = roundup(trip_rate * trailers / (metres * mass_mt_per_m)) + handling_rs_per_mt
    return trailers, freight


def financial_cost(pre_cost: float, external: float, internal: float, freight: float, financial_pct: float, gst_pct: float) -> float:
    k = (financial_pct / 100) * (1 + gst_pct / 100)
    if not 0 <= k < 1:
        raise CalculationError("financial and GST percentages produce an invalid factor")
    base = pre_cost + external + internal + freight
    return k * base / (1 - k)


@dataclass(frozen=True)
class SellingPrice:
    bare_per_mt: float
    external_per_mt: float
    internal_per_mt: float
    freight_per_mt: float
    total_per_mt: float


def selling_price(pre_cost: float, financial: float, external: float, internal: float, freight: float) -> SellingPrice:
    bare = roundup(pre_cost + financial)
    ext = roundup(external)
    intc = roundup(internal)
    fr = roundup(freight)
    return SellingPrice(bare, ext, intc, fr, bare + ext + intc + fr)
