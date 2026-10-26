"""Thin adapter importing the authoritative calculation engine without formula duplication."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Locate and add repository root to sys.path so calculation_engine is directly importable
ROOT_DIR = Path(__file__).resolve().parent.parent.parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

# Also support backend parent if deployed in backend/ subdirectory
BACKEND_DIR = Path(__file__).resolve().parent.parent.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

try:
    import calculation_engine as ce
    from calculation_engine import (
        REFERENCE_POLICY,
        CalculationError,
        CostSheet,
        QuoteLineResult,
        calc_costsheet,
        calc_line,
        calc_quote,
    )
except ImportError as err:
    raise RuntimeError(
        f"Failed to import authoritative calculation_engine from {ROOT_DIR}. "
        "Ensure Quotation System/calculation_engine.py is accessible."
    ) from err


class CalculationService:
    """Production adapter for the pure-Python calculation engine."""

    @staticmethod
    def calculate_line(
        line_data: Dict[str, Any],
        trip_rate: float = 98000.0,
        mode: str = "REFERENCE",
    ) -> QuoteLineResult:
        """Calculates a single quotation line according to selected commercial mode."""
        kw = dict(line_data)
        kw.setdefault("trip_rate", trip_rate)

        # Apply calculation mode policy
        normalized_mode = mode.upper().strip()
        if normalized_mode == "REFERENCE":
            kw.update(REFERENCE_POLICY)
        elif normalized_mode == "LEGACY":
            kw["legacy_mode"] = True
        elif normalized_mode == "NORMAL":
            kw["legacy_mode"] = False
        else:
            # Default to reference policy
            kw.update(REFERENCE_POLICY)

        return calc_line(**kw)

    @classmethod
    def calculate_quotation(
        cls,
        lines: List[Dict[str, Any]],
        trip_rate: float = 98000.0,
        mode: str = "REFERENCE",
        gst_pct: float = 18.0,
        pooled_trips: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Calculates an entire quotation and produces line outputs, weighted stats, and totals."""
        if not lines:
            return {
                "lines": [],
                "total_metres": 0.0,
                "total_mt": 0.0,
                "total_trips": 0.0,
                "subtotal_excl_gst": 0.0,
                "gst_amount": 0.0,
                "grand_total": 0.0,
                "calculation_mode": mode,
            }

        normalized_mode = mode.upper().strip()
        
        # Prepare line keyword arguments
        prepared_lines = []
        for line in lines:
            item_kw = dict(line)
            item_kw.setdefault("trip_rate", trip_rate)
            item_kw.setdefault("gst_pct", gst_pct)
            prepared_lines.append(item_kw)

        if normalized_mode == "REFERENCE":
            cost_sheet = calc_costsheet(
                prepared_lines,
                pooled_trips=pooled_trips,
                gst_pct=gst_pct,
                policy=REFERENCE_POLICY,
            )
            calculated_lines = [cls._line_result_to_dict(r) for r in cost_sheet.lines]
            return {
                "lines": calculated_lines,
                "total_metres": cost_sheet.total_metres,
                "total_mt": cost_sheet.total_mt,
                "total_trips": cost_sheet.total_trips,
                "subtotal_excl_gst": cost_sheet.subtotal_excl_gst,
                "gst_amount": cost_sheet.gst,
                "grand_total": cost_sheet.grand_total,
                "weighted": cost_sheet.weighted,
                "calculation_mode": "REFERENCE",
            }
        else:
            # Per-line calculation for NORMAL or LEGACY modes
            results: List[QuoteLineResult] = []
            for item in prepared_lines:
                res = cls.calculate_line(item, trip_rate=trip_rate, mode=normalized_mode)
                results.append(res)

            subtotal = sum(r.line_value for r in results)
            gst = round(subtotal * (gst_pct / 100.0), 0)
            grand = subtotal + gst

            return {
                "lines": [cls._line_result_to_dict(r) for r in results],
                "total_metres": sum(r.metres for r in results),
                "total_mt": sum(r.weight_mt for r in results),
                "total_trips": sum(r.freight_trips for r in results),
                "subtotal_excl_gst": subtotal,
                "gst_amount": gst,
                "grand_total": grand,
                "calculation_mode": normalized_mode,
            }

    @staticmethod
    def _line_result_to_dict(res: QuoteLineResult) -> Dict[str, Any]:
        """Serializes QuoteLineResult dataclass into a JSON-compatible dictionary."""
        return {
            "od_mm": res.od_mm,
            "inner_d_mm": res.inner_d_mm,
            "mass_mt_per_m": res.mass_mt_per_m,
            "mass_kg_per_m": res.mass_kg_per_m,
            "metres": res.metres,
            "weight_mt": res.weight_mt,
            "pipes": res.pipes,
            "trailers": res.trailers,
            "pre_cost": res.pre_cost,
            "external_cost_per_mt": res.external_cost_per_mt,
            "internal_cost_per_mt": res.internal_cost_per_mt,
            "freight_per_mt": res.freight_per_mt,
            "financial_cost_per_mt": res.financial_cost_per_mt,
            "bare_per_m": res.bare_per_m,
            "external_per_m": res.external_per_m,
            "internal_per_m": res.internal_per_m,
            "freight_per_m": res.freight_per_m,
            "total_per_m": res.total_per_m,
            "total_per_mt": res.total_per_mt,
            "line_value": res.line_value,
            "calculated_unit_price": res.calculated_unit_price,
            "applied_unit_price": res.applied_unit_price,
            "price_source": res.price_source,
            "price_variance": res.price_variance,
            "override_reason": res.override_reason,
            "raw_material_cost_per_mt": res.raw_material_cost_per_mt,
            "bare_pipe_cost_per_mt": res.bare_pipe_cost_per_mt,
            "freight_trips": res.freight_trips,
        }
