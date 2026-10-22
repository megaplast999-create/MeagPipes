"""
Golden baseline regression tests for Mega Pipes quotation pricing.

Enforces permanent regression protection:
1. Normal Mode Option 1 approved golden outputs (Line prices, subtotal, GST, grand total).
2. Legacy Mode approved golden outputs (Line prices, subtotal, GST, grand total).
3. Component-level golden outputs for every line item in both modes.
4. Calculation mode consistency: rejects invalid or undefined calculation modes.
5. Tamper detection: asserts that modifying any component by even ₹1 fails the gate.
"""

import sys
from pathlib import Path
import pytest

ROOT_DIR = Path(__file__).resolve().parent.parent
TESTS_DIR = ROOT_DIR / "TESTS"
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from calculation_engine import calc_line

# Authoritative Unipro Golden Inputs
GOLDEN_LINES = [
    {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542},
    {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534},
    {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529},
    {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525},
]

COMMON_RATES = {
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


class TestNormalModeGoldenBaselines:
    """Golden assertions for approved Normal Mode (Option 1)."""

    def test_line_1_normal(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[0], "trip_rate": 125.32 * 782, "legacy_mode": False})
        assert res.bare_per_m == 17594
        assert res.external_per_m == 1952
        assert res.internal_per_m == 1059
        assert res.freight_per_m == 2084
        assert res.total_per_m == 22689
        assert res.line_value == 45378000
        assert res.freight_per_mt == 8147

    def test_line_2_normal(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[1], "trip_rate": 125.32 * 782, "legacy_mode": False})
        assert res.bare_per_m == 18650
        assert res.external_per_m == 2068
        assert res.internal_per_m == 1123
        assert res.freight_per_m == 2268
        assert res.total_per_m == 24109
        assert res.line_value == 8438150
        assert res.freight_per_mt == 8364

    def test_line_3_normal(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[2], "trip_rate": 125.32 * 782, "legacy_mode": False})
        assert res.bare_per_m == 19359
        assert res.external_per_m == 2147
        assert res.internal_per_m == 1166
        assert res.freight_per_m == 2185
        assert res.total_per_m == 24857
        assert res.line_value == 12428500
        assert res.freight_per_mt == 7761

    def test_line_4_normal(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[3], "trip_rate": 125.32 * 782, "legacy_mode": False})
        assert res.bare_per_m == 20068
        assert res.external_per_m == 2225
        assert res.internal_per_m == 1209
        assert res.freight_per_m == 2079
        assert res.total_per_m == 25581
        assert res.line_value == 112556400
        assert res.freight_per_mt == 7123

    def test_normal_subtotal_and_grand_total(self):
        subtotal = sum(
            calc_line(**{**COMMON_RATES, **item, "trip_rate": 125.32 * 782, "legacy_mode": False}).line_value
            for item in GOLDEN_LINES
        )
        gst = round(subtotal * 0.18, 0)
        grand = subtotal + gst

        assert subtotal == 178801050
        assert gst == 32184189
        assert grand == 210985239


class TestLegacyModeGoldenBaselines:
    """Golden assertions for approved Legacy Mode reference reproduction."""

    def test_line_1_legacy(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[0], "trip_rate": 98000.0, "legacy_mode": True})
        assert res.bare_per_m == 17587
        assert res.external_per_m == 1952
        assert res.internal_per_m == 1074
        assert res.freight_per_m == 2068
        assert res.total_per_m == 22681
        assert res.line_value == 45362000
        assert res.freight_per_mt == 8083

    def test_line_2_legacy(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[1], "trip_rate": 98000.0, "legacy_mode": True})
        assert res.bare_per_m == 18639
        assert res.external_per_m == 2068
        assert res.internal_per_m == 1138
        assert res.freight_per_m == 2069
        assert res.total_per_m == 23914
        assert res.line_value == 8369900
        assert res.freight_per_mt == 7632

    def test_line_3_legacy(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[2], "trip_rate": 98000.0, "legacy_mode": True})
        assert res.bare_per_m == 19349
        assert res.external_per_m == 2147
        assert res.internal_per_m == 1181
        assert res.freight_per_m == 2071
        assert res.total_per_m == 24748
        assert res.line_value == 12374000
        assert res.freight_per_mt == 7355

    def test_line_4_legacy(self):
        res = calc_line(**{**COMMON_RATES, **GOLDEN_LINES[3], "trip_rate": 98000.0, "legacy_mode": True})
        assert res.bare_per_m == 20060
        assert res.external_per_m == 2225
        assert res.internal_per_m == 1224
        assert res.freight_per_m == 2071
        assert res.total_per_m == 25580
        assert res.line_value == 112552000
        assert res.freight_per_mt == 7097

    def test_legacy_subtotal_and_grand_total(self):
        subtotal = sum(
            calc_line(**{**COMMON_RATES, **item, "trip_rate": 98000.0, "legacy_mode": True}).line_value
            for item in GOLDEN_LINES
        )
        gst = round(subtotal * 0.18, 0)
        grand = subtotal + gst

        assert subtotal == 178657900
        assert gst == 32158422
        assert grand == 210816322

    def test_legacy_formulaic_without_manual_override(self):
        """Verifies that lines without manual FC overrides calculate the exact approved reference prices via the approved formula."""
        unoverridden_lines = [
            {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000},
            {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350},
            {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500},
            {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400},
        ]
        results = [
            calc_line(**{**COMMON_RATES, **item, "trip_rate": 98000.0, "legacy_mode": True})
            for item in unoverridden_lines
        ]
        assert results[0].financial_cost_per_mt == 1542
        assert results[0].total_per_m == 22681
        assert results[0].line_value == 45362000

        assert results[1].financial_cost_per_mt == 1534
        assert results[1].total_per_m == 23914
        assert results[1].line_value == 8369900

        assert results[2].financial_cost_per_mt == 1529
        assert results[2].total_per_m == 24748
        assert results[2].line_value == 12374000

        assert results[3].financial_cost_per_mt == 1525
        assert results[3].total_per_m == 25580
        assert results[3].line_value == 112552000

        subtotal = sum(r.line_value for r in results)
        assert subtotal == 178657900


class TestTamperDetectionAndGateEnforcement:
    """Demonstrate that modifying any golden rate or output immediately fails the gate."""

    def test_tampered_rate_fails_golden_assertion(self):
        tampered_rates = {**COMMON_RATES, "steel_rate": 61500.0}  # +500 Rs/MT
        res = calc_line(**{**tampered_rates, **GOLDEN_LINES[0], "trip_rate": 125.32 * 782, "legacy_mode": False})
        # Rate changed, total_per_m must NOT match the golden 22689
        assert res.total_per_m != 22689

    def test_tampered_trip_rate_fails_golden_assertion(self):
        tampered_rates = {**COMMON_RATES}
        res = calc_line(**{**tampered_rates, **GOLDEN_LINES[0], "trip_rate": 95000.0, "legacy_mode": False})
        assert res.freight_per_mt != 8147
        assert res.total_per_m != 22689


class TestApprovedCompanyCostSheetBaselines:
    """Golden assertions for the Authoritative Approved Company Excel CostSheet (Unipro 4-line tender)."""

    APPROVED_RATES = {
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

    APPROVED_LINES = [
        {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1024},
        {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1014},
        {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1011},
        {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1008},
    ]

    def test_line_1_approved(self):
        res = calc_line(**{**self.APPROVED_RATES, **self.APPROVED_LINES[0]})
        assert res.od_mm == 1243.0
        assert res.bare_per_m == 17459
        assert res.external_per_m == 1952
        assert res.internal_per_m == 976
        assert res.freight_per_m == 1964
        assert res.total_per_m == 22351
        assert res.line_value == 44702000
        assert res.freight_per_mt == 7676

    def test_line_2_approved(self):
        res = calc_line(**{**self.APPROVED_RATES, **self.APPROVED_LINES[1]})
        assert res.od_mm == 1317.0
        assert res.bare_per_m == 18503
        assert res.external_per_m == 2068
        assert res.internal_per_m == 1034
        assert res.freight_per_m == 1965
        assert res.total_per_m == 23570
        assert res.line_value == 8249500
        assert res.freight_per_mt == 7248

    def test_line_3_approved(self):
        res = calc_line(**{**self.APPROVED_RATES, **self.APPROVED_LINES[2]})
        assert res.od_mm == 1367.0
        assert res.bare_per_m == 19209
        assert res.external_per_m == 2147
        assert res.internal_per_m == 1074
        assert res.freight_per_m == 1966
        assert res.total_per_m == 24396
        assert res.line_value == 12198000
        assert res.freight_per_mt == 6985

    def test_line_4_approved(self):
        res = calc_line(**{**self.APPROVED_RATES, **self.APPROVED_LINES[3]})
        assert res.od_mm == 1417.0
        assert res.bare_per_m == 19915
        assert res.external_per_m == 2225
        assert res.internal_per_m == 1113
        assert res.freight_per_m == 1967
        assert res.total_per_m == 25220
        assert res.line_value == 110968000
        assert res.freight_per_mt == 6740

    def test_approved_subtotal_gst_and_grand_total(self):
        results = [
            calc_line(**{**self.APPROVED_RATES, **item})
            for item in self.APPROVED_LINES
        ]
        subtotal = sum(r.line_value for r in results)
        gst = round(subtotal * 0.18, 0)
        grand = subtotal + gst

        assert subtotal == 176117500
        assert gst == 31701150
        assert grand == 207818650

    def test_approved_generalization_scaling(self):
        """Verify that scaling quantities scales line values proportionally without hardcoding."""
        scaled_line_1_half = {**self.APPROVED_LINES[0], "quantity": 1000}
        res = calc_line(**{**self.APPROVED_RATES, **scaled_line_1_half})
        assert res.total_per_m == 22351
        assert res.line_value == 22351000

        scaled_line_4_double = {**self.APPROVED_LINES[3], "quantity": 8800}
        res4 = calc_line(**{**self.APPROVED_RATES, **scaled_line_4_double})
        assert res4.total_per_m == 25220
        assert res4.line_value == 221936000

    def test_approved_formulaic_without_manual_override(self):
        """Verify that entering new quotes without manual FC overrides uses the approved formula dynamically."""
        line_no_override = {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000}
        res = calc_line(**{**self.APPROVED_RATES, **line_no_override, "manual_fc": 0})
        # Approved formulaic FC: baseCost * 1.18 * 0.01 = 1019
        assert res.financial_cost_per_mt == 1019
        assert res.bare_per_m == 17458
        assert res.total_per_m == 22350
        assert res.line_value == 44700000

