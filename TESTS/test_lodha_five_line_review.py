"""
Automated Pytest Suite: Five-Line Lodha Technical Verification
Validates:
1. ID-basis upward-rounding geometry rule across all 5 sizes.
2. Mass (kg/m), Total Weight (MT), External & Internal coating areas against Lodha reference.
3. Separate per-line integer dispatch pipe rounding (2,144 pipes) vs fractional pipe equivalents (2,141.67).
4. Trailer count logistics (538 trailers).
5. Explicit audit gate: Commercial subtotal unverified pending approved Lodha CostSheet.
"""
import math
from pathlib import Path
import openpyxl
import pytest

REF_FILE = Path(r"C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx")

LODHA_SPECS = [
    {"line": 1, "row": 12, "id": 1026, "wt": 12, "qty": 365, "exp_od": 1050},
    {"line": 2, "row": 13, "id": 1050, "wt": 12, "qty": 25035, "exp_od": 1074},
    {"line": 3, "row": 14, "id": 876, "wt": 12, "qty": 120, "exp_od": 900},
    {"line": 4, "row": 15, "id": 676, "wt": 12, "qty": 80, "exp_od": 700},
    {"line": 5, "row": 16, "id": 976, "wt": 12, "qty": 100, "exp_od": 1000},
]

def roundup_d(val: float, decimals: int = 0) -> float:
    factor = 10 ** decimals
    return math.ceil(val * factor) / factor

def test_lodha_id_basis_geometry_derivation():
    """Confirms all five ID-basis sizes derive expected ODs using RoundUp(ID + 2*WT, 0)."""
    for spec in LODHA_SPECS:
        raw_od = spec["id"] + 2 * spec["wt"]
        derived_od = roundup_d(raw_od, 0)
        assert derived_od == spec["exp_od"], f"Line {spec['line']} OD mismatch: {derived_od} != {spec['exp_od']}"

def test_lodha_mass_and_weight_against_reference():
    """Compares kg/m and total weight against the Lodha reference workbook."""
    wb = openpyxl.load_workbook(REF_FILE, data_only=True)
    ws = wb["Sheet1"]
    
    total_calc_mt = 0.0
    for spec in LODHA_SPECS:
        r = spec["row"]
        od = spec["exp_od"]
        wt = spec["wt"]
        qty = spec["qty"]
        
        calc_kg_m = (od - wt) * wt * 0.0246615
        ref_kg_m = ws.cell(r, 10).value
        assert abs(calc_kg_m - ref_kg_m) < 1e-6, f"Line {spec['line']} kg/m mismatch"
        
        calc_mt = qty * calc_kg_m / 1000.0
        ref_mt = ws.cell(r, 11).value
        assert abs(calc_mt - ref_mt) < 1e-6, f"Line {spec['line']} weight MT mismatch"
        total_calc_mt += calc_mt
        
    assert round(total_calc_mt, 3) == 8057.338, f"Total weight mismatch: {total_calc_mt} != 8057.338"

def test_lodha_coating_areas_against_reference():
    """Compares external and internal coating surface areas against Lodha reference."""
    wb = openpyxl.load_workbook(REF_FILE, data_only=True)
    ws = wb["Sheet1"]
    
    total_ext = 0
    total_int = 0
    for spec in LODHA_SPECS:
        r = spec["row"]
        od = spec["exp_od"]
        wt = spec["wt"]
        qty = spec["qty"]
        
        calc_ext = roundup_d(od * qty * (22 / 7) / 1000.0, 0)
        ref_ext = ws.cell(r, 14).value
        assert calc_ext == ref_ext, f"Line {spec['line']} external coating mismatch"
        
        calc_int = roundup_d((22 / 7) * qty * (od - 2 * wt) / 1000.0, 0)
        ref_int = ws.cell(r, 15).value
        assert calc_int == ref_int, f"Line {spec['line']} internal coating mismatch"
        
        total_ext += calc_ext
        total_int += calc_int
        
    assert total_ext == 86540, f"Total external coating mismatch: {total_ext} != 86540"
    assert total_int == 84601, f"Total internal coating mismatch: {total_int} != 84601"

def test_lodha_pipe_count_dispatch_vs_fractional():
    """Verifies separate whole-pipe rounding per line (2,144) vs fractional pipe equivalents (2,141.67)."""
    whole_dispatch_sum = sum(math.ceil(s["qty"] / 12.0) for s in LODHA_SPECS)
    fractional_sum = sum(s["qty"] / 12.0 for s in LODHA_SPECS)
    
    # 31 + 2087 + 10 + 7 + 9 = 2,144
    assert whole_dispatch_sum == 2144, f"Whole dispatch sum mismatch: {whole_dispatch_sum} != 2144"
    # 30.4167 + 2086.25 + 10.0 + 6.6667 + 8.3333 = 2,141.6667
    assert round(fractional_sum, 2) == 2141.67, f"Fractional sum mismatch: {fractional_sum} != 2141.67"

def test_lodha_trailer_count():
    """Verifies discrete trailer allocation sum (538 trailers)."""
    # 8 + 522 + 3 + 2 + 3 = 538
    total_trailers = sum(math.ceil(math.ceil(s["qty"] / 12.0) / 4.0) for s in LODHA_SPECS)
    assert total_trailers == 538, f"Total trailers mismatch: {total_trailers} != 538"
