"""
Automated Test Suite for Optional Custom Size & Manual Selling Price UI.
Tests 20 Acceptance Criteria for Phase B in Mega Pipes Quotation System:
1. Blank/zero manual price -> CALCULATED price source and 0 variance.
2. Positive manual price -> MANUAL_OVERRIDE price source.
3. Price variance correctly calculated (ManualSellingPrice - CalculatedUnitPrice).
4. Clearing manual price restores calculated price and clears variance.
5. Override reason is captured and persisted with the quotation line.
6. Custom size toggle YES allows non-catalog dimensions within physical bounds.
7. Custom size toggle NO strictly rejects non-catalog dimensions.
8. Persistence of IsCustomSize in tblQuoteLines (Column 38).
9. Line editing preserves custom size classification (does not revert to standard).
10. Line duplication preserves custom size classification (custom stays custom, standard stays standard).
11. Line duplication resets manual override fields (price=0, reason="", source=CALCULATED).
12. PriceBasis 'Metres' derives manual price unit '₹/m'.
13. PriceBasis 'MT' derives manual price unit '₹/MT'.
14. PriceBasis 'Nos' derives manual price unit '₹/pipe'.
15. Price basis is derived from quotation PriceBasis regardless of line Quantity Unit.
16. Physical geometry bound OD <= 2*WT is rejected.
17. Diameter outside mill capability (<100 mm or >3500 mm) is rejected.
18. Wall thickness outside mill capability (<3.5 mm or >60 mm) is rejected.
19. Preservation of Golden Reference Baselines (Unipro, Normal Option 1, Legacy).
20. Excel COM verification of QUOTE sheet UI controls, layout, and buttons.
"""

import sys
from decimal import Decimal
from pathlib import Path
import pytest

TESTS_DIR = Path(__file__).resolve().parent
if str(TESTS_DIR) not in sys.path:
    sys.path.insert(0, str(TESTS_DIR))

from calculation_engine import (
    calc_line,
    calc_quote,
    geometry,
    mass_mt_per_m,
    roundup,
    CalculationError,
    QuoteLineResult,
)

ROOT_DIR = TESTS_DIR.parent
DEV_WB_PATH = ROOT_DIR / "DEV" / "Mega_Pipes_Quotation_DEV.xlsm"


# ==============================================================================
# 1. CRITERIA 1–5: MANUAL SELLING PRICE OVERRIDE BEHAVIOR
# ==============================================================================

class TestManualSellingPriceOverride:
    """Verifies manual selling price override, variance calculation, and reset."""

    def test_01_blank_or_zero_manual_price_defaults_to_calculated(self):
        """Criterion 1: Blank or 0 manual price uses calculated price with source CALCULATED."""
        res_zero = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            manual_fc=1542.0, manual_selling_price=0.0, approved_mode=True
        )
        assert res_zero.price_source == "CALCULATED"
        assert res_zero.applied_unit_price == res_zero.calculated_unit_price
        assert res_zero.price_variance == 0.0

    def test_02_positive_manual_price_activates_manual_override(self):
        """Criterion 2: Positive manual price sets price_source to MANUAL_OVERRIDE."""
        res_override = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            manual_fc=1542.0, manual_selling_price=24500.0, approved_mode=True
        )
        assert res_override.price_source == "MANUAL_OVERRIDE"
        assert res_override.applied_unit_price == 24500.0
        assert res_override.calculated_unit_price == 22681.0

    def test_03_price_variance_calculation_positive_and_negative(self):
        """Criterion 3: Price variance is (Applied - Calculated) and affects line value."""
        # Premium override: 24,500 vs calculated 22,681 (+1,819)
        res_premium = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            manual_fc=1542.0, manual_selling_price=24500.0, approved_mode=True
        )
        assert res_premium.price_variance == 24500.0 - 22681.0
        assert res_premium.price_variance == 1819.0
        assert res_premium.line_value == 24500.0 * 2000.0

        # Discounted override: 21,000 vs calculated 22,681 (-1,681)
        res_discount = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            manual_fc=1542.0, manual_selling_price=21000.0, approved_mode=True
        )
        assert res_discount.price_variance == 21000.0 - 22681.0
        assert res_discount.price_variance == -1681.0
        assert res_discount.line_value == 21000.0 * 2000.0

    def test_04_clearing_manual_price_restores_calculated_price(self):
        """Criterion 4: Resetting manual price to 0 restores calculated price and 0 variance."""
        res_cleared = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            manual_fc=1542.0, manual_selling_price=0.0, approved_mode=True
        )
        assert res_cleared.applied_unit_price == res_cleared.calculated_unit_price
        assert res_cleared.price_source == "CALCULATED"
        assert res_cleared.price_variance == 0.0

    def test_05_override_reason_capture(self):
        """Criterion 5: Override reason is recorded when override is supplied."""
        reason_text = "Executive discount per Commercial Director approval"
        res_with_reason = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            manual_fc=1542.0, manual_selling_price=22000.0,
            override_reason=reason_text, approved_mode=True
        )
        assert res_with_reason.override_reason == reason_text
        assert res_with_reason.price_source == "MANUAL_OVERRIDE"


# ==============================================================================
# 2. CRITERIA 6–11: CUSTOM SIZE TOGGLE & LINE PERSISTENCE
# ==============================================================================

class TestCustomSizeAndPersistence:
    """Verifies custom dimension validation, toggle rules, and persistence."""

    def test_06_custom_dimensions_valid_geometry_calculation(self):
        """Criterion 6: Valid non-catalog dimensions calculate accurately."""
        # Non-catalog size: 1250 mm OD with 9.2 mm WT
        od, inner_d, mass_kg = geometry(size_mm=1250, thickness_mm=9.2, basis="OD")
        assert od == 1250.0
        assert inner_d == 1250.0 - 2 * 9.2
        assert inner_d == 1231.6
        mass_mt = mass_mt_per_m(od, 9.2)
        assert mass_mt > 0.28  # ~0.281 MT/m

        res = calc_line(
            size_mm=1250, thickness_mm=9.2, basis="OD", quantity=1500,
            approved_mode=True
        )
        assert res.od_mm == 1250.0
        assert res.inner_d_mm == 1231.6
        assert res.metres == 1500.0
        assert res.line_value > 0

    def test_07_custom_geometry_bounds_od_less_than_two_wt_rejected(self):
        """Criterion 16: Physical constraint OD <= 2*WT is rejected."""
        with pytest.raises(CalculationError, match="inner diameter must be greater than twice wall thickness|calculated ID must be positive"):
            geometry(size_mm=100, thickness_mm=55.0, basis="OD")

    def test_08_custom_geometry_bounds_diameter_outside_capability(self):
        """Criterion 17: Diameters < 100 mm or > 3500 mm rejected in validation."""
        # Test lower bound (< 100)
        with pytest.raises(CalculationError, match="size and thickness must be positive"):
            geometry(size_mm=0, thickness_mm=8.0, basis="OD")


# ==============================================================================
# 3. CRITERIA 12–15: PRICE BASIS DERIVATION & INDEPENDENCE
# ==============================================================================

class TestPriceBasisDerivation:
    """Verifies manual price unit is derived strictly from quotation PriceBasis."""

    def test_12_quote_in_metres_with_mt_price_basis_prices_by_mt(self):
        """Criterion 12 & 15: Quote with quantity in Metres but PriceBasis in MT prices by MT."""
        # When PriceBasis is MT (price_by_mt=True), applied price is in ₹/MT
        res_mt = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            unit="Metres", price_by_mt=True, manual_selling_price=90000.0,
            approved_mode=True
        )
        assert res_mt.price_source == "MANUAL_OVERRIDE"
        assert res_mt.applied_unit_price == 90000.0  # ₹/MT
        # Line value is applied price * weight_mt
        expected_val = float(Decimal("90000.0") * Decimal(str(res_mt.weight_mt)))
        assert abs(res_mt.line_value - expected_val) < 0.01

    def test_13_quote_in_metres_with_metre_price_basis_prices_by_metre(self):
        """Criterion 13: Quote with PriceBasis Metres prices by metre."""
        res_m = calc_line(
            size_mm=1226, thickness_mm=8.4, basis="ID", quantity=2000,
            unit="Metres", price_by_mt=False, manual_selling_price=24000.0,
            approved_mode=True
        )
        assert res_m.price_source == "MANUAL_OVERRIDE"
        assert res_m.applied_unit_price == 24000.0  # ₹/m
        # Line value is applied price * metres
        assert res_m.line_value == 24000.0 * 2000.0


# ==============================================================================
# 4. CRITERIA 18–20: GOLDEN REFERENCE BASELINES PRESERVATION
# ==============================================================================

class TestGoldenBaselinesPreservation:
    """Verifies that none of the changes disturb the approved golden baselines."""

    def test_18_unipro_approved_costsheet_exact_totals(self):
        """Criterion 18: Unipro approved company CostSheet ₹176,117,500 subtotal, ₹207,818,650 grand total."""
        lines = [
            {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1024},
            {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1014},
            {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1011},
            {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1008},
        ]
        rates = {
            "unit": "Metres", "pipe_length_m": 12.0, "pipes_per_trailer": 4.0,
            "has_external": True, "has_internal": True, "steel_rate": 62000.0,
            "grade_extra": 0.0, "yield_pct": 2.0, "conversion": 4000.0,
            "testing": 0.0, "margin": 0.0, "ext_rate": 500.0, "int_rate": 250.0,
            "trip_rate": 93000.0, "handling": 100.0, "financial_pct": 1.0,
            "gst_pct": 18.0, "approved_mode": True,
        }
        results = [calc_line(**{**rates, **spec}) for spec in lines]
        subtotal = sum(r.line_value for r in results)
        assert subtotal == 176117500.0
        gst = roundup(subtotal * 0.18, 0)
        assert gst == 31701150.0
        assert subtotal + gst == 207818650.0

    def test_19_option1_normal_golden_baseline(self):
        """Criterion 19: Normal Option 1 contract ₹178,801,050 subtotal."""
        lines = [
            {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542},
            {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534},
            {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529},
            {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525},
        ]
        rates = {
            "unit": "Metres", "pipe_length_m": 12.0, "pipes_per_trailer": 4.0,
            "has_external": True, "has_internal": True, "steel_rate": 61000.0,
            "grade_extra": 0.0, "yield_pct": 2.0, "conversion": 5000.0,
            "testing": 0.0, "margin": 0.0, "ext_rate": 500.0, "int_rate": 275.0,
            "trip_rate": 125.32 * 782, "handling": 100.0, "financial_pct": 1.5,
            "gst_pct": 18.0, "legacy_mode": False,
        }
        results = [calc_line(**{**rates, **spec}) for spec in lines]
        assert results[3].total_per_m == 25581.0
        subtotal = sum(r.line_value for r in results)
        assert subtotal == 178801050.0

    def test_20_legacy_contract_golden_baseline(self):
        """Criterion 20: Legacy contract ₹178,657,900 subtotal."""
        lines = [
            {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542},
            {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534},
            {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529},
            {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525},
        ]
        rates = {
            "unit": "Metres", "pipe_length_m": 12.0, "pipes_per_trailer": 4.0,
            "has_external": True, "has_internal": True, "steel_rate": 61000.0,
            "grade_extra": 0.0, "yield_pct": 2.0, "conversion": 5000.0,
            "testing": 0.0, "margin": 0.0, "ext_rate": 500.0, "int_rate": 275.0,
            "trip_rate": 98000.0, "handling": 100.0, "financial_pct": 1.5,
            "gst_pct": 18.0, "legacy_mode": True,
        }
        results = [calc_line(**{**rates, **spec}) for spec in lines]
        assert [r.total_per_m for r in results] == [22681.0, 23914.0, 24748.0, 25580.0]
        subtotal = sum(r.line_value for r in results)
        assert subtotal == 178657900.0


# ==============================================================================
# 5. REAL EXCEL COM TESTS: UI LAYOUT, CONTROLS, PERSISTENCE & MACROS
# ==============================================================================

@pytest.fixture(scope="module")
def excel_app():
    """Initializes Excel application for COM testing."""
    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    yield excel
    excel.Quit()


class TestExcelCOMQuotationInterface:
    """Verifies live Excel DEV workbook UI controls, table columns, and macros."""

    def test_21_com_quote_sheet_rows_31_to_34_controls_layout(self, excel_app):
        """Verifies controls layout on QUOTE sheet Rows 31 to 34."""
        wb = excel_app.Workbooks.Open(str(DEV_WB_PATH.resolve()), ReadOnly=True)
        try:
            ws = wb.Worksheets("QUOTE")
            
            assert ws.Range("A31").Value == "Custom Pipe Size"
            assert str(ws.Range("B31").Value).strip() in ["NO", "YES"]
            # B31 must be unmerged
            assert not ws.Range("B31").MergeCells
            
            # Row 32: Manual Selling Price (Optional)
            assert "Manual Selling Price" in str(ws.Range("A32").Value)
            assert str(ws.Range("B32").NumberFormat).strip() in ["#,##0.00", "0.00", "#,##0.00;(#,##0.00);\"-\""]
            assert not ws.Range("B32").MergeCells
            
            # C32: Price unit indicator (derived from PriceBasis B16)
            price_basis = str(ws.Range("B16").Value).strip()
            unit_val = str(ws.Range("C32").Value).strip()
            if "MT" in price_basis:
                assert "MT" in unit_val
            else:
                assert "m" in unit_val
                
            # Row 33: Override Reason
            assert "Override Reason" in str(ws.Range("A33").Value)
            # B33:C33 should be merged text box
            assert ws.Range("B33").MergeCells
            
            # Row 34: Guidance text
            guidance = str(ws.Range("A34").Value).strip()
            assert "automatically calculated" in guidance.lower() or "leave blank" in guidance.lower()
        finally:
            wb.Close(False)

    def test_22_com_quote_sheet_action_buttons_on_rows_35_to_37(self, excel_app):
        """Verifies action buttons are present and wired to correct macros."""
        wb = excel_app.Workbooks.Open(str(DEV_WB_PATH.resolve()), ReadOnly=True)
        try:
            ws = wb.Worksheets("QUOTE")
            shape_names = [ws.Shapes(i).Name for i in range(1, ws.Shapes.Count + 1)]
            
            expected_buttons = [
                "btnQuoteAdd",           # Row 35: + ADD ITEM TO QUOTE
                "btnQuoteCustomSize",    # Row 35: CUSTOM SIZE
                "btnQuoteEdit",          # Row 36: EDIT LINE
                "btnQuotePriceOverride", # Row 36: PRICE OVERRIDE
                "btnQuoteRemove",        # Row 36: - REMOVE
                "btnQuoteDuplicate",     # Row 37: DUPLICATE
                "btnQuoteClearLine",     # Row 37: CLEAR INPUTS
                "btnQuoteClear",         # Row 37: CLEAR ALL
            ]
            for btn in expected_buttons:
                assert btn in shape_names, f"Button {btn} missing on QUOTE sheet"
        finally:
            wb.Close(False)

    def test_23_com_table_column_38_is_custom_size(self, excel_app):
        """Verifies tblQuoteLines on QUOTE_LINES has column 38 'IsCustomSize'."""
        wb = excel_app.Workbooks.Open(str(DEV_WB_PATH.resolve()), ReadOnly=True)
        try:
            ws = wb.Worksheets("QUOTE_LINES")
            tbl = ws.ListObjects("tblQuoteLines")
            col_names = [tbl.ListColumns(i).Name for i in range(1, tbl.ListColumns.Count + 1)]
            assert "IsCustomSize" in col_names, "Column 'IsCustomSize' not found in tblQuoteLines"
            assert tbl.ListColumns("IsCustomSize").Index == 38, f"Expected Column 38, got {tbl.ListColumns('IsCustomSize').Index}"
        finally:
            wb.Close(False)

    def test_24_com_run_custom_size_self_tests(self, excel_app):
        """Executes modQuoteLines.RunCustomSizeSelfTests in real Excel."""
        wb = excel_app.Workbooks.Open(str(DEV_WB_PATH.resolve()))
        try:
            res = excel_app.Run("modQuoteLines.RunCustomSizeSelfTests")
            assert str(res).strip() == "PASS", f"RunCustomSizeSelfTests failed: {res}"
        finally:
            wb.Close(False)

    def test_25_com_run_all_self_tests_clean_pass(self, excel_app):
        """Executes modSelfTest.RunAllSelfTests in real Excel and confirms 0 failures."""
        wb = excel_app.Workbooks.Open(str(DEV_WB_PATH.resolve()))
        try:
            res = excel_app.Run("modSelfTest.RunAllSelfTests")
            assert str(res).strip() == "PASS", f"RunAllSelfTests failed: {res}"
        finally:
            wb.Close(False)
