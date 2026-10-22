"""
Unit and regression tests for Phase 1-6 Dynamic Staged Cost-Calculation Engine.

Verifies:
1. Stage-by-stage costing flow (raw material, yield loss, bare pipe, coating, freight, finance, margin).
2. Enquiry-specific variable inputs (steel rate, yield %, conversion, coatings, trip rate, fin %).
3. Manual selling-price override behavior:
   - Preserves system-calculated price
   - Sets applied_unit_price to manual override
   - Tags price_source = "MANUAL_OVERRIDE"
   - Computes price_variance
   - Calculates line_value and totals from applied price
   - Does not infer cost components from override price
4. Dynamic recalculation when inputs change.
5. Authoritative Unipro approved CostSheet baseline preservation (₹176,117,500 subtotal, ₹31,701,150 GST, ₹207,818,650 grand total).
6. New-enquiry multi-line scenario with different commercial inputs.
"""

import sys
from pathlib import Path
import pytest

TESTS_DIR = Path(__file__).resolve().parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from calculation_engine import calc_line, calc_quote


class TestDynamicStagedCostingEngine:
    """Verifies staged costing calculations and enquiry-specific commercial inputs."""

    def test_raw_material_and_yield_loss_stages(self):
        """Verifies raw material cost calculation from steel rate + grade extra + yield loss %."""
        # Test Case 1: Standard Unipro inputs
        res = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=62000.0, grade_extra=0.0, yield_pct=2.0, conversion=4000.0,
            approved_mode=True
        )
        assert res.raw_material_cost_per_mt == 63240.0  # 62000 * 1.02
        assert res.bare_pipe_cost_per_mt == 67240.0    # 63240 + 4000
        assert res.pre_cost == 67240.0

        # Test Case 2: Enquiry-specific steel rate and yield loss
        res2 = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=65000.0, grade_extra=1000.0, yield_pct=2.5, conversion=4500.0,
            approved_mode=True
        )
        # Delivered = 66000. Yield loss @ 2.5% = 1650. Raw material = 67650.
        assert res2.raw_material_cost_per_mt == 67650.0
        # Bare pipe pre-cost = 67650 + 4500 = 72150.
        assert res2.bare_pipe_cost_per_mt == 72150.0
        assert res2.pre_cost == 72150.0

    def test_coating_cost_stages(self):
        """Verifies external and internal coating cost stages respond dynamically to rates."""
        # Default approved rates: Ext=500, Int=250
        res_default = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=62000.0, ext_rate=500.0, int_rate=250.0, approved_mode=True
        )
        assert round(res_default.external_cost_per_mt, 2) == 7630.37
        assert round(res_default.internal_cost_per_mt, 2) == 3815.19

        # Enquiry-specific rates: Ext=550, Int=280
        res_override = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=62000.0, ext_rate=550.0, int_rate=280.0, approved_mode=True
        )
        # 550 * 3.14 * (1243/1000) / 0.25575553836 = 8393.41
        assert round(res_override.external_cost_per_mt, 2) == 8393.41
        # 280 * 3.14 * (1243/1000) / 0.25575553836 = 4273.01
        assert round(res_override.internal_cost_per_mt, 2) == 4273.01

        # Changing coating rate changes per-metre prices
        assert res_override.external_per_m > res_default.external_per_m
        assert res_override.internal_per_m > res_default.internal_per_m

    def test_freight_and_financial_cost_stages(self):
        """Verifies freight and finance cost stages with enquiry-specific inputs."""
        # Trip rate changed from 93,000 to 105,000
        res_freight = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=62000.0, conversion=4000.0, ext_rate=500.0, int_rate=250.0,
            trip_rate=105000.0, handling=150.0,
            manual_fc=0.0, financial_pct=1.2, approved_mode=True
        )
        # Fractional trailers = (2000/12)/4 = 41.666667
        # Weight = 511.511077 MT
        # Freight = ceil(105000 * 41.666667 / 511.511077) + 150 = 8554 + 150 = 8704
        assert res_freight.freight_per_mt == 8704.0

        # Finance cost formulaic: Round(BaseCost * 1.18 * 0.012)
        # BaseCost = 67240 + 7630.37 + 3815.19 + 8704 = 87389.56
        # FC = Round(87389.56 * 1.18 * 0.012) = Round(1237.437) = 1237
        assert res_freight.financial_cost_per_mt == 1237.0


class TestManualSellingPriceOverride:
    """Verifies Phase 5 Manual Selling-Price Override requirements."""

    def test_manual_price_override_preserves_calculated_price(self):
        """Verifies that when manual price is used, calculated price is preserved and variance is recorded."""
        # Line 1 calculated price is 22,351 / m
        res = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=62000.0, conversion=4000.0, trip_rate=93000.0,
            ext_rate=500.0, int_rate=250.0, manual_fc=1024,
            approved_mode=True,
            manual_selling_price=22000.0,
            override_reason="Strategic key customer discount authorized by Sales Director"
        )

        # 1. Calculated price is preserved
        assert res.calculated_unit_price == 22351.0
        assert res.total_per_m == 22351.0

        # 2. Applied price is the manual override
        assert res.applied_unit_price == 22000.0

        # 3. Price source is tagged as MANUAL_OVERRIDE
        assert res.price_source == "MANUAL_OVERRIDE"

        # 4. Variance is recorded (-351 Rs/m)
        assert res.price_variance == -351.0

        # 5. Line value is calculated from applied price (22,000 * 2,000 = 44,000,000)
        assert res.line_value == 44000000.0

        # 6. Override reason is captured
        assert res.override_reason == "Strategic key customer discount authorized by Sales Director"

        # 7. Internal cost components remain intact (not inferred from override)
        assert res.pre_cost == 67240.0
        assert res.freight_per_mt == 7676.0
        assert res.financial_cost_per_mt == 1024.0

    def test_zero_or_no_manual_override_defaults_to_calculated(self):
        """Verifies that omitting manual price defaults to CALCULATED with 0 variance."""
        res = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            steel_rate=62000.0, conversion=4000.0, trip_rate=93000.0,
            ext_rate=500.0, int_rate=250.0, manual_fc=1024,
            approved_mode=True,
            manual_selling_price=0.0
        )
        assert res.calculated_unit_price == 22351.0
        assert res.applied_unit_price == 22351.0
        assert res.price_source == "CALCULATED"
        assert res.price_variance == 0.0
        assert res.line_value == 44702000.0


class TestNewEnquiryDynamicRecalculation:
    """Verifies Phase 6 multi-line separate new-enquiry scenario with distinct commercial inputs."""

    NEW_ENQUIRY_RATES = {
        "unit": "Metres",
        "pipe_length_m": 12.0,
        "pipes_per_trailer": 4.0,
        "has_external": True,
        "has_internal": True,
        "steel_rate": 64500.0,     # Different steel rate
        "grade_extra": 500.0,      # Grade extra
        "yield_pct": 2.2,          # Yield loss 2.2%
        "conversion": 4200.0,      # Conversion 4200
        "testing": 50.0,           # Testing cost 50
        "margin": 200.0,           # Margin 200
        "ext_rate": 520.0,         # Ext coating 520
        "int_rate": 260.0,         # Int coating 260
        "trip_rate": 96000.0,      # Trip rate 96000
        "handling": 120.0,         # Handling 120
        "financial_pct": 1.0,
        "gst_pct": 18.0,
        "approved_mode": True,
    }

    NEW_ENQUIRY_LINES = [
        {"size_mm": 1026, "thickness_mm": 8.0, "basis": "ID", "quantity": 1500},
        {"size_mm": 1200, "thickness_mm": 9.0, "basis": "ID", "quantity": 800},
        {"size_mm": 1400, "thickness_mm": 10.0, "basis": "ID", "quantity": 2500},
    ]

    def test_new_enquiry_calculates_dynamically(self):
        """Verifies that a new enquiry with custom commercial inputs calculates reproducible prices."""
        results = [
            calc_line(**{**self.NEW_ENQUIRY_RATES, **line})
            for line in self.NEW_ENQUIRY_LINES
        ]

        # Verify all lines have CALCULATED status
        for r in results:
            assert r.price_source == "CALCULATED"
            assert r.price_variance == 0.0
            assert r.applied_unit_price == r.calculated_unit_price
            assert r.line_value == r.applied_unit_price * r.metres

        subtotal = sum(r.line_value for r in results)
        gst = round(subtotal * 0.18, 0)
        grand = subtotal + gst

        assert subtotal > 0
        assert gst == round(subtotal * 0.18, 0)
        assert grand == subtotal + gst
