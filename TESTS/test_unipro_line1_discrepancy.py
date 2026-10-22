"""
Mega Pipes Quotation System — Regression & Audit Test Suite
File: TESTS/test_unipro_line1_discrepancy.py

Investigates and validates the forensic root cause of the ₹1/m discrepancy
observed between:
- Unipro Line 1 displayed at ₹22,350/m (Dynamic Landed Formulaic Finance Cost)
- Acceptance report recorded at ₹22,351/m (Historical Row 32 Pre-Cost Manual FC)

Strict Governance & Traceability:
- Traced to Decision Record 1 in DOCS/COMMERCIAL_POLICY_SIGNOFF.md and
  DOCS/COMMERCIAL_POLICY_DECISION_REGISTER.md (Pending Executive Sign-Off).
- Authoritative reference: MKTGENQ26-27-594 - Unipro 2.xlsx (sheet CS Stacking 10-09-26).
"""

import sys
from decimal import Decimal
from pathlib import Path
import pytest

TESTS_DIR = Path(__file__).resolve().parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from calculation_engine import calc_line

UNIPRO_APPROVED_RATES = {
    "unit": "Metres",
    "pipe_length_m": 12.0,
    "pipes_per_trailer": 4.0,
    "has_external": True,
    "has_internal": True,
    "steel_rate": 62000.0,
    "grade_extra": 0.0,
    "yield_pct": 2.0,
    "conversion": 4000.0,
    "testing": 0.0,
    "margin": 0.0,
    "ext_rate": 500.0,
    "int_rate": 250.0,
    "trip_rate": 93000.0,
    "handling": 100.0,
    "financial_pct": 1.0,
    "gst_pct": 18.0,
    "approved_mode": True,
}

UNIPRO_LINE_1_SPEC = {
    "size_mm": 1226,
    "thickness_mm": 8.4,
    "basis": "ID",
    "quantity": 2000,
}


class TestUniproLine1DiscrepancyForensicAnalysis:
    """Forensic verification of the ₹22,350 vs ₹22,351 unit price discrepancy."""

    def test_01_historical_unipro_costsheet_replication_22351(self):
        """
        Replicates historical approved CostSheet (CS Stacking 10-09-26) with Row 32 Manual FC = 1,024 Rs/MT.
        
        Evaluates to:
          - Bare Price: ₹17,459 / m
          - Ext Coating: ₹1,952 / m
          - Int Coating: ₹976 / m
          - Freight: ₹1,964 / m
          - Total Unit Price: ₹22,351 / m
          - Line Total (2,000 m): ₹44,702,000
        """
        res = calc_line(**{**UNIPRO_APPROVED_RATES, **UNIPRO_LINE_1_SPEC, "manual_fc": 1024})
        
        assert res.od_mm == 1243.0
        assert res.financial_cost_per_mt == 1024
        assert res.bare_per_m == 17459
        assert res.external_per_m == 1952
        assert res.internal_per_m == 976
        assert res.freight_per_m == 1964
        assert res.total_per_m == 22351
        assert res.line_value == 44702000
        assert res.applied_unit_price == 22351
        assert res.price_source == "CALCULATED"

    def test_02_dynamic_formulaic_landed_finance_cost_22350(self):
        """
        Evaluates dynamic formulaic landed finance cost (manual_fc = 0 / omitted in UI):
          FC/MT = Round((PreCost + Ext + Int + Freight) * 1.18 * Fin% / 100, 0)
                = Round((67240 + 7630.34 + 3815.19 + 7676.45) * 1.18 * 0.01, 0)
                = 1,019 Rs/MT
        
        Evaluates to:
          - Bare Price: ₹17,458 / m (-₹1/m due to -₹5/MT on 0.255756 MT/m)
          - Ext Coating: ₹1,952 / m
          - Int Coating: ₹976 / m
          - Freight: ₹1,964 / m
          - Total Unit Price: ₹22,350 / m
          - Line Total (2,000 m): ₹44,700,000
        """
        res = calc_line(**{**UNIPRO_APPROVED_RATES, **UNIPRO_LINE_1_SPEC, "manual_fc": 0})
        
        assert res.od_mm == 1243.0
        assert res.financial_cost_per_mt == 1019
        assert res.bare_per_m == 17458
        assert res.external_per_m == 1952
        assert res.internal_per_m == 976
        assert res.freight_per_m == 1964
        assert res.total_per_m == 22350
        assert res.line_value == 44700000
        assert res.applied_unit_price == 22350
        assert res.price_source == "CALCULATED"

    def test_03_exact_mathematical_rounding_bridge(self):
        """
        Traces the exact mathematical bridge down to the decimal paisa level:
          Delta FC = 1,024 - 1,019 = 5.00 Rs/MT
          Mass MT/m = 0.25575635 MT/m
          Raw bare pre-cost before FC = 67,240.00 * 0.25575635 = 17,196.997 Rs/m
          
          With FC=1024: (67,240 + 1024) * 0.25575635 = 17,458.851 Rs/m -> Round() = 17,459 Rs/m
          With FC=1019: (67,240 + 1019) * 0.25575635 = 17,457.572 Rs/m -> Round() = 17,458 Rs/m
          Differential: Exactly -1.00 Rs/m (-0.0045%)
        """
        mass = (1243.0 - 8.4) * 8.4 * 0.0246615 / 1000.0
        raw_bare = (62000.0 * 1.02 + 4000.0) * mass
        
        val_with_1024 = round((67240.0 + 1024.0) * mass)
        val_with_1019 = round((67240.0 + 1019.0) * mass)
        
        assert val_with_1024 == 17459
        assert val_with_1019 == 17458
        assert val_with_1024 - val_with_1019 == 1

    def test_04_manual_selling_price_override_preservation(self):
        """
        Verifies that if a user intentionally enters a manual selling price of ₹22,350/m
        or ₹22,351/m, the engine preserves that exact price and tags it as MANUAL_OVERRIDE.
        """
        # User explicitly overrides to 22,350
        res_override_22350 = calc_line(
            **{**UNIPRO_APPROVED_RATES, **UNIPRO_LINE_1_SPEC, "manual_fc": 1024, "manual_selling_price": 22350.0}
        )
        assert res_override_22350.price_source == "MANUAL_OVERRIDE"
        assert res_override_22350.applied_unit_price == 22350.0
        assert res_override_22350.calculated_unit_price == 22351.0
        assert res_override_22350.price_variance == -1.0
        assert res_override_22350.line_value == 44700000.0

        # User explicitly overrides to 22,351
        res_override_22351 = calc_line(
            **{**UNIPRO_APPROVED_RATES, **UNIPRO_LINE_1_SPEC, "manual_fc": 0, "manual_selling_price": 22351.0}
        )
        assert res_override_22351.price_source == "MANUAL_OVERRIDE"
        assert res_override_22351.applied_unit_price == 22351.0
        assert res_override_22351.calculated_unit_price == 22350.0
        assert res_override_22351.price_variance == +1.0
        assert res_override_22351.line_value == 44702000.0
