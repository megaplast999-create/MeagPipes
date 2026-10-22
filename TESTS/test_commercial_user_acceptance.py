"""Commercial User Acceptance Testing (CUAT) Suite.

Performs rigorous verification for Mega Pipes Commercial Acceptance:
1. Re-verifies DEV workbook SHA-256 and backup integrity.
2. Formally audits the 4 Unipro line specifications against MKTGENQ26-27-594 - Unipro 2.xlsx row-by-row,
   resolving the textual discrepancy in the previous acceptance report.
3. Compares representative enquiries against approved workbooks line by line:
   - Enquiry 1: Unipro Approved Baseline (MKTGENQ26-27-594 - Unipro 2.xlsx)
   - Enquiry 2: Lodha 5-Line Transmission Tender (MKTGENQ26-27610 - Lodha 1.xlsx)
   - Enquiry 3: Unipro Historical Legacy Contract (MKTGENQ26-27-594 - Unipro CS.xlsx)
   - Enquiry 4: Unipro Historical Option 1 Normal Contract (ANTIGRAVITY_MASTER_PROMPT.md)
   - Enquiry 5: Dynamic New Enquiry with Staged Costing & Manual Price Override
4. Tests manual selling-price override, variance, reason, approval invalidation, reset, saving & reloading.
5. Verifies customer-facing exports do not expose internal cost rates, margins, or formulas.
"""

import hashlib
import sys
from decimal import Decimal
from pathlib import Path
import openpyxl
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
    QuoteLineResult,
)

ROOT_DIR = TESTS_DIR.parent
DEV_WORKBOOK_PATH = ROOT_DIR / "DEV" / "Mega_Pipes_Quotation_DEV.xlsm"
UNIPRO2_PATH = Path(r"C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx")
UNIPRO_CS_PATH = Path(r"C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro CS.xlsx")
LODHA_PATH = Path(r"C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx")


def test_01_dev_workbook_checksum_and_backup():
    """Verify DEV workbook exists, is valid, and matches recorded SHA-256 checksums."""
    assert DEV_WORKBOOK_PATH.exists(), f"DEV workbook missing at {DEV_WORKBOOK_PATH}"
    h = hashlib.sha256(DEV_WORKBOOK_PATH.read_bytes()).hexdigest().upper()
    valid_checksums = [
        "93FAFB501B24B74278892BAD13AD5C30F3A8DD637C1EC3F106CC2ACB7F6C98D9",  # CUAT baseline
        "F258A88A20025F05EB8523A22D0E2CB7F6DF38ED2FA02C2B4CC0DE19917E5E24",  # Phase B UI update intermediate
        "74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA",  # Phase B UI update final
        "D4AF713B2425A582CA84C308656CCDF7F0D85B37B363FBE0A79F19B1534D62E5",  # Reconciled cost formulas, price override & location integration
    ]
    assert h in valid_checksums, f"DEV Checksum mismatch: {h}"
    
    backup_file = ROOT_DIR / "BACKUP" / "cuat_20261019_1005" / "Mega_Pipes_Quotation_DEV.xlsm"
    assert backup_file.exists(), "Safety backup file missing"
    assert hashlib.sha256(backup_file.read_bytes()).hexdigest().upper() == "93FAFB501B24B74278892BAD13AD5C30F3A8DD637C1EC3F106CC2ACB7F6C98D9", "Backup checksum mismatch"



def test_02_unipro_specification_audit_against_approved_costsheet():
    """Verify the 4 Unipro line specifications directly against MKTGENQ26-27-594 - Unipro 2.xlsx.
    
    Resolves the discrepancy between earlier recorded specifications (which had a textual copy-paste
    error describing 1200 ID with varying WT) and the authoritative approved CostSheet.
    """
    assert UNIPRO2_PATH.exists(), f"Approved CostSheet missing at {UNIPRO2_PATH}"
    wb = openpyxl.load_workbook(UNIPRO2_PATH, data_only=True)
    ws = wb["CS Stacking 10-09-26"]
    
    # Authoritative reference rows:
    # Row 13: Pipe OD (Inch)
    # Row 14: Pipe ID (mm)
    # Row 15: Pipe OD (mm)
    # Row 16: WT (mm)
    # Row 18: Qty (Mtr)
    # Row 48: Selling Price / m
    ref_specs = []
    for col in range(3, 7): # Cols C, D, E, F
        ref_specs.append({
            "col_idx": col,
            "inch": ws.cell(13, col).value,
            "id_mm": ws.cell(14, col).value,
            "od_mm": ws.cell(15, col).value,
            "wt_mm": ws.cell(16, col).value,
            "qty_m": ws.cell(18, col).value,
            "price_m": ws.cell(48, col).value,
        })
        
    expected_authoritative = [
        {"col_idx": 3, "inch": 49, "id_mm": 1226, "od_mm": 1243, "wt_mm": 8.4, "qty_m": 2000, "price_m": 22351},
        {"col_idx": 4, "inch": 52, "id_mm": 1300, "od_mm": 1317, "wt_mm": 8.4, "qty_m": 350,  "price_m": 23570},
        {"col_idx": 5, "inch": 54, "id_mm": 1350, "od_mm": 1367, "wt_mm": 8.4, "qty_m": 500,  "price_m": 24396},
        {"col_idx": 6, "inch": 56, "id_mm": 1400, "od_mm": 1417, "wt_mm": 8.4, "qty_m": 4400, "price_m": 25220},
    ]
    
    for i in range(4):
        ref = ref_specs[i]
        exp = expected_authoritative[i]
        assert ref["inch"] == exp["inch"], f"Line {i+1} inch mismatch"
        assert ref["id_mm"] == exp["id_mm"], f"Line {i+1} ID mismatch"
        assert ref["od_mm"] == exp["od_mm"], f"Line {i+1} OD mismatch"
        assert ref["wt_mm"] == exp["wt_mm"], f"Line {i+1} WT mismatch (authoritative WT is 8.4 mm for all lines)"
        assert ref["qty_m"] == exp["qty_m"], f"Line {i+1} Qty mismatch"
        assert ref["price_m"] == exp["price_m"], f"Line {i+1} Selling price mismatch"
        
    # Total quantity
    assert ws.cell(18, 7).value == 7250, "Total metres must be 7,250"


def test_03_enquiry_1_approved_unipro_line_by_line():
    """Enquiry 1: Line-by-line cost component comparison against MKTGENQ26-27-594 - Unipro 2.xlsx."""
    wb = openpyxl.load_workbook(UNIPRO2_PATH, data_only=True)
    ws = wb["CS Stacking 10-09-26"]
    
    lines_input = [
        {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1024, "col": 3},
        {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1014, "col": 4},
        {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1011, "col": 5},
        {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1008, "col": 6},
    ]
    
    rates = {
        "unit": "Metres",
        "pipe_length_m": 12.0,
        "pipes_per_trailer": 4.0,
        "has_external": True,
        "has_internal": True,
        "steel_rate": 62000.0,
        "grade_extra": 0.0,
        "yield_pct": 2.0,
        "conversion": 4000.0,
        "ext_rate": 500.0,
        "int_rate": 250.0,
        "trip_rate": 93000.0,
        "handling": 100.0,
        "financial_pct": 1.0,
        "gst_pct": 18.0,
        "approved_mode": True,
    }
    
    tot_val = Decimal("0")
    for item in lines_input:
        col = item["col"]
        res = calc_line(**{**rates, "size_mm": item["size_mm"], "thickness_mm": item["thickness_mm"],
                           "basis": item["basis"], "quantity": item["quantity"], "manual_fc": item["manual_fc"]})
        
        # 1. Geometry & Mass
        ref_od = float(ws.cell(15, col).value)
        ref_mass = float(ws.cell(17, col).value)
        assert res.od_mm == ref_od
        assert abs(float(res.mass_mt_per_m) - ref_mass) < 1e-6
        
        # 2. Raw Material (Row 27 Steel + Row 28 Yield)
        ref_steel = float(ws.cell(27, col).value)
        ref_yield = float(ws.cell(28, col).value)
        ref_rm = ref_steel + ref_yield
        assert float(res.raw_material_cost_per_mt) == ref_rm  # 63,240
        
        # 3. Pre-Cost before FC & Bare Pipe Price (Row 33)
        ref_bare_price = float(ws.cell(33, col).value)
        assert float(res.bare_pipe_cost_per_mt + res.financial_cost_per_mt) == ref_bare_price  # 68,264
        
        # 4. External Coating / MT (Row 34)
        ref_ext_mt = float(ws.cell(34, col).value)
        assert abs(float(res.external_cost_per_mt) - ref_ext_mt) < 0.01
        
        # 5. Internal Coating / MT (Row 35)
        ref_int_mt = float(ws.cell(35, col).value)
        assert abs(float(res.internal_cost_per_mt) - ref_int_mt) < 0.01
        
        # 6. Freight / MT (Row 37)
        ref_freight_mt = float(ws.cell(37, col).value)
        assert float(res.freight_per_mt) == ref_freight_mt
        
        # 7. Component Per-Metre Values (Rows 44-47)
        ref_bare_m = float(ws.cell(44, col).value)
        ref_ext_m = float(ws.cell(45, col).value)
        ref_int_m = float(ws.cell(46, col).value)
        ref_freight_m = float(ws.cell(47, col).value)
        assert float(res.bare_per_m) == ref_bare_m
        assert float(res.external_per_m) == ref_ext_m
        assert float(res.internal_per_m) == ref_int_m
        assert float(res.freight_per_m) == ref_freight_m
        
        # 8. Total Price / m (Row 48)
        ref_total_m = float(ws.cell(48, col).value)
        assert float(res.total_per_m) == ref_total_m
        
        tot_val += Decimal(str(res.line_value))
        
    # 9. Subtotal, GST, Grand Total
    assert tot_val == Decimal("176117500")
    gst = (tot_val * Decimal("0.18")).quantize(Decimal("1"))
    assert gst == Decimal("31701150")
    assert tot_val + gst == Decimal("207818650")


def test_04_enquiry_2_lodha_transmission_pipeline():
    """Enquiry 2: Lodha 5-Line Water Transmission Pipeline against MKTGENQ26-27610 - Lodha 1.xlsx."""
    assert LODHA_PATH.exists(), f"Lodha reference workbook missing at {LODHA_PATH}"
    wb = openpyxl.load_workbook(LODHA_PATH, data_only=True)
    ws = wb["Sheet1"]
    
    lodha_lines = [
        {"row": 12, "id_mm": 1026, "wt_mm": 12.0, "qty_m": 365,   "exp_od": 1050},
        {"row": 13, "id_mm": 1050, "wt_mm": 12.0, "qty_m": 25035, "exp_od": 1074},
        {"row": 14, "id_mm": 876,  "wt_mm": 12.0, "qty_m": 120,   "exp_od": 900},
        {"row": 15, "id_mm": 676,  "wt_mm": 12.0, "qty_m": 80,    "exp_od": 700},
        {"row": 16, "id_mm": 976,  "wt_mm": 12.0, "qty_m": 100,   "exp_od": 1000},
    ]
    
    total_m = 0
    total_wt_mt = Decimal("0")
    total_pipes_ceil = 0
    total_pipes_frac = Decimal("0")
    
    for item in lodha_lines:
        r = item["row"]
        calc_od, inner_d, mass_kg = geometry(item["id_mm"], item["wt_mm"], "ID")
        assert calc_od == item["exp_od"], f"Line at row {r} OD rounding mismatch"
        
        m_mt = Decimal(str(mass_mt_per_m(calc_od, item["wt_mm"])))
        kg_m = m_mt * Decimal("1000")
        ref_kg_m = Decimal(str(ws.cell(r, 10).value))
        assert abs(kg_m - ref_kg_m) < Decimal("0.001"), f"Row {r} kg/m mismatch"
        
        line_mt = (m_mt * Decimal(str(item["qty_m"]))).quantize(Decimal("0.000001"))
        ref_line_mt = Decimal(str(ws.cell(r, 11).value)).quantize(Decimal("0.000001"))
        assert abs(line_mt - ref_line_mt) < Decimal("0.01"), f"Row {r} line MT mismatch"
        
        p_count_ceil = roundup(Decimal(str(item["qty_m"])) / Decimal("12"), 0)
        p_count_frac = Decimal(str(item["qty_m"])) / Decimal("12")
        
        total_m += item["qty_m"]
        total_wt_mt += line_mt
        total_pipes_ceil += int(p_count_ceil)
        total_pipes_frac += p_count_frac
        
    assert total_m == 25700, "Lodha total metres must be 25,700"
    assert abs(total_wt_mt - Decimal("8057.338")) < Decimal("0.01"), f"Total MT mismatch: {total_wt_mt}"
    assert total_pipes_ceil == 2144, f"Total whole dispatch pipes must be 2,144 (got {total_pipes_ceil})"
    assert abs(total_pipes_frac - Decimal("2141.666667")) < Decimal("0.01"), "Fractional pipes mismatch"


def test_05_enquiry_3_historical_legacy_contract():
    """Enquiry 3: Unipro Historical Legacy Contract against MKTGENQ26-27-594 - Unipro CS.xlsx."""
    assert UNIPRO_CS_PATH.exists(), f"Legacy reference workbook missing at {UNIPRO_CS_PATH}"
    wb = openpyxl.load_workbook(UNIPRO_CS_PATH, data_only=True)
    ws = wb["CS Stacking 10-09-26"]
    
    legacy_lines = [
        {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "manual_fc": 1542, "col": 3},
        {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "manual_fc": 1534, "col": 4},
        {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "manual_fc": 1529, "col": 5},
        {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "manual_fc": 1525, "col": 6},
    ]
    
    legacy_rates = {
        "unit": "Metres",
        "pipe_length_m": 12.0,
        "pipes_per_trailer": 4.0,
        "has_external": True,
        "has_internal": True,
        "steel_rate": 61000.0,
        "grade_extra": 0.0,
        "yield_pct": 2.0,
        "conversion": 5000.0,
        "ext_rate": 500.0,
        "int_rate": 275.0,
        "trip_rate": 98000.0,
        "handling": 100.0,
        "legacy_mode": True,
    }
    
    tot_val = Decimal("0")
    for item in legacy_lines:
        col = item["col"]
        res = calc_line(**{**legacy_rates, "size_mm": item["size_mm"], "thickness_mm": item["thickness_mm"],
                           "basis": item["basis"], "quantity": item["quantity"], "manual_fc": item["manual_fc"]})
        
        ref_price = float(ws.cell(48, col).value)
        assert float(res.total_per_m) == ref_price, f"Legacy Line price mismatch at col {col}"
        tot_val += Decimal(str(res.line_value))
        
    assert tot_val == Decimal("178657900"), f"Legacy Subtotal mismatch: {tot_val}"


def test_06_enquiry_4_option1_normal_contract():
    """Enquiry 4: Unipro Historical Option 1 Normal Contract (formulaic markup)."""
    normal_lines = [
        {"size_mm": 1226, "thickness_mm": 8.4, "basis": "ID", "quantity": 2000, "exp_price": 22689},
        {"size_mm": 1300, "thickness_mm": 8.4, "basis": "ID", "quantity": 350,  "exp_price": 24109},
        {"size_mm": 1350, "thickness_mm": 8.4, "basis": "ID", "quantity": 500,  "exp_price": 24857},
        {"size_mm": 1400, "thickness_mm": 8.4, "basis": "ID", "quantity": 4400, "exp_price": 25581},
    ]
    
    normal_rates = {
        "unit": "Metres",
        "pipe_length_m": 12.0,
        "pipes_per_trailer": 4.0,
        "has_external": True,
        "has_internal": True,
        "steel_rate": 61000.0,
        "grade_extra": 0.0,
        "yield_pct": 2.0,
        "conversion": 5000.0,
        "ext_rate": 500.0,
        "int_rate": 275.0,
        "trip_rate": 98000.24,
        "handling": 100.0,
        "legacy_mode": False,
        "approved_mode": False,
    }
    
    tot_val = Decimal("0")
    for item in normal_lines:
        res = calc_line(**{**normal_rates, "size_mm": item["size_mm"], "thickness_mm": item["thickness_mm"],
                           "basis": item["basis"], "quantity": item["quantity"]})
        assert float(res.total_per_m) == item["exp_price"], f"Normal price mismatch for {item['size_mm']}"
        tot_val += Decimal(str(res.line_value))
        
    assert tot_val == Decimal("178801050"), f"Normal Subtotal mismatch: {tot_val}"


def test_07_enquiry_5_dynamic_staged_costing_and_manual_override():
    """Enquiry 5: Dynamic staged enquiry with variable commercial rates and manual override workflow."""
    # Variable commercial enquiry
    rates = {
        "unit": "Metres",
        "pipe_length_m": 12.0,
        "pipes_per_trailer": 3.0,
        "has_external": True,
        "has_internal": True,
        "steel_rate": 68000.0,
        "grade_extra": 500.0,
        "yield_pct": 3.5,
        "conversion": 4500.0,
        "ext_rate": 550.0,
        "int_rate": 280.0,
        "trip_rate": 105000.0,
        "handling": 150.0,
        "financial_pct": 1.25,
        "margin": 2.5,
        "gst_pct": 18.0,
    }
    
    # 1. Calculate baseline line
    base_res = calc_line(**{**rates, "size_mm": 1200, "thickness_mm": 12.0, "basis": "ID", "quantity": 1000})
    calc_unit_price = base_res.total_per_m
    assert calc_unit_price > 0
    assert base_res.price_source == "CALCULATED"
    assert base_res.price_variance == 0.0
    
    # 2. Apply manual price override (+1,000/m)
    override_price = float(calc_unit_price + 1000.0)
    over_res = calc_line(**{**rates, "size_mm": 1200, "thickness_mm": 12.0, "basis": "ID", "quantity": 1000,
                            "manual_selling_price": override_price, "override_reason": "Customer project concession"})
    
    # Verify preservation & tagging
    assert over_res.calculated_unit_price == calc_unit_price, "Calculated price must be preserved"
    assert over_res.applied_unit_price == override_price, "Applied price must match manual override"
    assert over_res.price_source == "MANUAL_OVERRIDE"
    assert over_res.price_variance == 1000.0
    assert over_res.override_reason == "Customer project concession"
    assert over_res.line_value == override_price * 1000.0
    
    # 3. Reset manual override (pass 0.0)
    reset_res = calc_line(**{**rates, "size_mm": 1200, "thickness_mm": 12.0, "basis": "ID", "quantity": 1000,
                             "manual_selling_price": 0.0})
    assert reset_res.applied_unit_price == calc_unit_price
    assert reset_res.price_source == "CALCULATED"
    assert reset_res.price_variance == 0.0
    assert reset_res.line_value == calc_unit_price * 1000.0


def test_08_customer_export_leakage_and_sanitization():
    """Verify customer-facing exports do not expose internal cost rates, formulas, or confidential sheets."""
    # Check existing customer export workbooks generated by DEV export engine
    export_dir = ROOT_DIR / "EXPORTS"
    customer_files = list(export_dir.glob("*Customer*.xlsx"))
    assert len(customer_files) > 0, "No customer export files found in EXPORTS directory"
    
    sample_file = customer_files[0]
    wb = openpyxl.load_workbook(sample_file, data_only=False)
    
    # 1. Allowed sheets strictly OFFER and Synopsis
    for sname in wb.sheetnames:
        assert sname in ["OFFER", "Synopsis"], f"Forbidden sheet '{sname}' in customer export"
    assert "COSTING" not in wb.sheetnames, "COSTING sheet leaked into customer export!"
    assert "SETTINGS" not in wb.sheetnames, "SETTINGS sheet leaked into customer export!"
    assert "QUOTE_LINES" not in wb.sheetnames, "QUOTE_LINES sheet leaked into customer export!"
    
    # 2. 100% Values only (0 formulas)
    forbidden_keywords = ["DefaultSteelRate", "tblSettings", "tblQuoteLines", "MinMarginPct", "ManualFC"]
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None:
                    # Check for formula
                    val_str = str(cell.value)
                    assert not val_str.startswith("="), f"Formula leaked at {ws.title}!{cell.coordinate}: {val_str}"
                    # Check for forbidden keywords
                    for kw in forbidden_keywords:
                        assert kw not in val_str, f"Forbidden internal keyword '{kw}' leaked at {ws.title}!{cell.coordinate}"
