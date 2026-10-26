"""Golden-baseline integration tests verifying zero commercial calculation drift through CalculationService."""

import pytest
from app.services.calculation import CalculationService


class TestCalculationEngineGoldenBaselines:
    """Verifies strict mathematical parity across all 3 approved golden figures."""

    def test_approved_company_costsheet_baseline_parity(self):
        """Approved Company CostSheet baseline: exactly ₹17,61,17,500 excluding GST."""
        approved_lines = [
            {
                "size_mm": 1226.0,
                "thickness_mm": 8.4,
                "basis": "ID",
                "quantity": 2000.0,
                "unit": "Metres",
                "steel_rate": 62000.0,
                "conversion": 4000.0,
                "yield_pct": 2.0,
                "margin": 0.0,
                "testing": 0.0,
                "ext_rate": 500.0,
                "int_rate": 250.0,
                "trip_rate": 93000.0,
                "handling": 100.0,
                "financial_pct": 1.0,
                "manual_fc": 1024.0,
                "pipes_per_trailer": 4.0,
            },
            {
                "size_mm": 1300.0,
                "thickness_mm": 8.4,
                "basis": "ID",
                "quantity": 350.0,
                "unit": "Metres",
                "steel_rate": 62000.0,
                "conversion": 4000.0,
                "yield_pct": 2.0,
                "margin": 0.0,
                "testing": 0.0,
                "ext_rate": 500.0,
                "int_rate": 250.0,
                "trip_rate": 93000.0,
                "handling": 100.0,
                "financial_pct": 1.0,
                "manual_fc": 1014.0,
                "pipes_per_trailer": 4.0,
            },
            {
                "size_mm": 1350.0,
                "thickness_mm": 8.4,
                "basis": "ID",
                "quantity": 500.0,
                "unit": "Metres",
                "steel_rate": 62000.0,
                "conversion": 4000.0,
                "yield_pct": 2.0,
                "margin": 0.0,
                "testing": 0.0,
                "ext_rate": 500.0,
                "int_rate": 250.0,
                "trip_rate": 93000.0,
                "handling": 100.0,
                "financial_pct": 1.0,
                "manual_fc": 1011.0,
                "pipes_per_trailer": 4.0,
            },
            {
                "size_mm": 1400.0,
                "thickness_mm": 8.4,
                "basis": "ID",
                "quantity": 4400.0,
                "unit": "Metres",
                "steel_rate": 62000.0,
                "conversion": 4000.0,
                "yield_pct": 2.0,
                "margin": 0.0,
                "testing": 0.0,
                "ext_rate": 500.0,
                "int_rate": 250.0,
                "trip_rate": 93000.0,
                "handling": 100.0,
                "financial_pct": 1.0,
                "manual_fc": 1008.0,
                "pipes_per_trailer": 4.0,
            },
        ]

        result = CalculationService.calculate_quotation(
            lines=approved_lines,
            trip_rate=93000.0,
            mode="REFERENCE",
        )

        lines = result["lines"]
        assert len(lines) == 4

        # Assert individual line value parity
        assert lines[0]["total_per_m"] == 22351.0
        assert lines[0]["line_value"] == 44702000.0

        assert lines[1]["total_per_m"] == 23570.0
        assert lines[1]["line_value"] == 8249500.0

        assert lines[2]["total_per_m"] == 24396.0
        assert lines[2]["line_value"] == 12198000.0

        assert lines[3]["total_per_m"] == 25220.0
        assert lines[3]["line_value"] == 110968000.0

        # Assert total project value excluding GST is exactly ₹17,61,17,500
        assert result["subtotal_excl_gst"] == 176117500.0

    def test_legacy_baseline_parity(self):
        """Legacy baseline: exactly ₹17,86,57,900."""
        common = {
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
            "trip_rate": 98000.0,
        }
        legacy_lines = [
            {**common, "size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542},
            {**common, "size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534},
            {**common, "size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529},
            {**common, "size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525},
        ]

        result = CalculationService.calculate_quotation(
            lines=legacy_lines,
            trip_rate=98000.0,
            mode="LEGACY",
        )

        lines = result["lines"]
        assert lines[0]["line_value"] == 45362000
        assert lines[1]["line_value"] == 8369900
        assert lines[2]["line_value"] == 12374000
        assert lines[3]["line_value"] == 112552000

        assert result["subtotal_excl_gst"] == 178657900
        assert result["gst_amount"] == 32158422
        assert result["grand_total"] == 210816322

    def test_historical_standard_tender_baseline_parity(self):
        """Historical standard-tender baseline: exactly ₹17,88,01,050."""
        trip_rate = 125.32 * 782
        common = {
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
            "trip_rate": trip_rate,
        }
        tender_lines = [
            {**common, "size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542},
            {**common, "size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534},
            {**common, "size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529},
            {**common, "size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525},
        ]

        result = CalculationService.calculate_quotation(
            lines=tender_lines,
            trip_rate=trip_rate,
            mode="NORMAL",
        )

        lines = result["lines"]
        assert lines[0]["line_value"] == 45378000
        assert lines[1]["line_value"] == 8438150
        assert lines[2]["line_value"] == 12428500
        assert lines[3]["line_value"] == 112556400

        assert result["subtotal_excl_gst"] == 178801050
        assert result["gst_amount"] == 32184189
        assert result["grand_total"] == 210985239

    def test_costsheet_automation_reference_policy_golden_parity(self):
        """CostSheet automation parity: typed FC totals 176,117,500; auto FC totals 176,142,150."""
        unipro = [(1226, 2000), (1300, 350), (1350, 500), (1400, 4400)]
        base = {
            "basis": "ID",
            "thickness_mm": 8.4,
            "steel_rate": 62000.0,
            "yield_pct": 2.0,
            "conversion": 4000.0,
            "ext_rate": 500.0,
            "int_rate": 250.0,
            "trip_rate": 93000.0,
            "handling": 100.0,
            "financial_pct": 1.0,
        }

        # 1. Typed FC: 1024, 1014, 1011, 1008 -> exact reference CostSheet total 176,117,500
        typed_fcs = [1024, 1014, 1011, 1008]
        typed_lines = [{**base, "size_mm": s, "quantity": q, "manual_fc": fc} for (s, q), fc in zip(unipro, typed_fcs)]
        res_typed = CalculationService.calculate_quotation(typed_lines, mode="REFERENCE")
        assert [r["total_per_m"] for r in res_typed["lines"]] == [22351.0, 23570.0, 24396.0, 25220.0]
        assert res_typed["subtotal_excl_gst"] == 176117500.0

        # 2. Auto FC: algebraic fixed point -> [22353, 23573, 24400, 25224], subtotal 176,142,150
        auto_lines = [{**base, "size_mm": s, "quantity": q} for s, q in unipro]
        res_auto = CalculationService.calculate_quotation(auto_lines, mode="REFERENCE")
        assert [r["total_per_m"] for r in res_auto["lines"]] == [22353.0, 23573.0, 24400.0, 25224.0]
        assert res_auto["subtotal_excl_gst"] == 176142150.0

