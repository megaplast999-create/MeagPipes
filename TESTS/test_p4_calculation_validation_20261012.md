# Mega Pipes Quotation System — P4 Calculation Engine Validation Report

**Phase:** P4 — Calculation engine (VBA CalcLine agrees with Python within 1 rupee)  
**Date:** 2026-10-12  
**Baseline Git Commit:** `46b64f61189451baf8c5466075da47ee3f406489`  
**Baseline Git Tag:** `p3-golden-fix`  
**Target:** `master`  

---

## 1. Executive Summary

Phase 4 (P4) calculation engine alignment and automated validation has been successfully implemented and verified.
- The Python reference calculation engine (`TESTS/calculation_engine.py`) has been fully aligned with `DOCS/ANTIGRAVITY_MASTER_PROMPT.md` and the approved Option 1 normal-mode contract.
- Coating calculations now divide strictly by `mass_mt_per_m` (metric tonnes per metre), eliminating the legacy denominator defect that divided by kg/m in normal mode.
- Circumference calculations use `PI = 3.14`, matching `tblSettings` and the reference workbook.
- Internal coating calculates perimeter from `inner_d` in normal mode and `OD` in legacy mode.
- Per-metre selling price component rounding (`BarePerM`, `ExtPerM`, `IntPerM`, `FreightPerM`) has been implemented according to the master specification.
- Dedicated automated tests in `TESTS/test_p4_line_agreement.py` verify 100% agreement ($\Delta = ₹0$) across all 4 quotation lines in both Normal and Legacy modes against both the approved Option 1 golden targets and the verified desktop Excel VBA `CalcLine` fixture (`TESTS/vba_p3_golden_20261008.txt`).
- The entire Python test suite (39 tests) passes in 0.11 seconds with zero errors or warnings.

---

## 2. Changes Made

1. **`TESTS/calculation_engine.py`**:
   - Added `mass_mt_per_m(od_mm, thickness_mm) -> float`.
   - Updated `coating_cost_per_mt` to use `PI_CONSTANT = 3.14` and divide by mass in MT/metre.
   - Added `external_coating_per_mt` and `internal_coating_per_mt` (supporting normal `inner_d` and legacy `OD`).
   - Implemented `selling_price_per_m` calculating rounded whole-rupee components and extended line values.
   - Implemented `calc_line` and `calc_quote` providing end-to-end multi-line quotation calculations matching VBA `modPricing.CalcLine`.

2. **`TESTS/test_calculation_engine.py`**:
   - Replaced obsolete assertion of buggy normal coating formula with corrected `mass_mt_per_m` and `PI = 3.14` assertions.
   - Added tests for `roundup`, `mass_mt_per_m`, `internal_coating_per_mt` (normal vs legacy), whole trailer rounding in normal mode, algebraic financial cost calculation, per-metre selling price components, and geometry guard rails.

3. **`TESTS/test_p4_line_agreement.py`** (New test suite):
   - Parameterized tests across all 4 golden lines in Normal mode against Option 1 targets.
   - Parameterized tests across all 4 golden lines in Legacy mode against reference targets.
   - Automated comparison asserting $\le ₹1$ tolerance and exact integer equality between Python results and verified desktop Excel VBA output from `TESTS/vba_p3_golden_20261008.txt`.
   - Grand total assertions for Normal mode (₹178,801,050) and Legacy mode (₹178,657,900).
   - Geometry validation guard tests.

4. **`BACKUP/`**:
   - Timestamped pre-modification backups preserved:
     - `BACKUP/calculation_engine_pre_p4_20261012_170500.py`
     - `BACKUP/test_calculation_engine_pre_p4_20261012_170500.py`

---

## 3. Numerical Verification Evidence

### Normal Mode (Option 1 Contract)
| Line | Size | WT | Metres | Python Total/m | VBA Total/m | Target Total/m | Diff (Py - VBA) | Line Value |
|---|---|---|---|---|---|---|---|---|
| Line 1 | ID 1226 (OD 1243) | 8.4 | 2,000 | ₹22,689 | ₹22,689 | ₹22,689 | ₹0 | ₹45,378,000 |
| Line 2 | ID 1300 (OD 1317) | 8.4 | 350 | ₹24,109 | ₹24,109 | ₹24,109 | ₹0 | ₹8,438,150 |
| Line 3 | ID 1350 (OD 1367) | 8.4 | 500 | ₹24,857 | ₹24,857 | ₹24,857 | ₹0 | ₹12,428,500 |
| Line 4 | ID 1400 (OD 1417) | 8.4 | 4,400 | ₹25,581 | ₹25,581 | ₹25,581 | ₹0 | ₹112,556,400 |
| **Total** | | | **7,250** | | | | **₹0** | **₹178,801,050** |

### Legacy Mode (Historical Contract)
| Line | Size | WT | Metres | Python Total/m | VBA Total/m | Target Total/m | Diff (Py - VBA) | Line Value |
|---|---|---|---|---|---|---|---|---|
| Line 1 | ID 1226 (OD 1243) | 8.4 | 2,000 | ₹22,681 | ₹22,681 | ₹22,681 | ₹0 | ₹45,362,000 |
| Line 2 | ID 1300 (OD 1317) | 8.4 | 350 | ₹23,914 | ₹23,914 | ₹23,914 | ₹0 | ₹8,369,900 |
| Line 3 | ID 1350 (OD 1367) | 8.4 | 500 | ₹24,748 | ₹24,748 | ₹24,748 | ₹0 | ₹12,374,000 |
| Line 4 | ID 1400 (OD 1417) | 8.4 | 4,400 | ₹25,580 | ₹25,580 | ₹25,580 | ₹0 | ₹112,552,000 |
| **Total** | | | **7,250** | | | | **₹0** | **₹178,657,900** |

---

## 4. Test Execution Results

Command executed:
```powershell
python -B -m pytest TESTS -v
```

Terminal output:
```text
============================= test session starts =============================
platform win32 -- Python 3.14.8, pytest-9.1.1, pluggy-1.6.0 -- C:\Python314\python.exe
cachedir: .pytest_cache
rootdir: C:\Users\Administrator\Desktop\Mega Pipes\Quotation System
collecting ... collected 39 items

TESTS/test_calculation_engine.py::test_roundup_basic PASSED              [  2%]
TESTS/test_calculation_engine.py::test_mass_mt_per_m PASSED              [  5%]
TESTS/test_calculation_engine.py::test_geometry_od PASSED                [  7%]
TESTS/test_calculation_engine.py::test_geometry_id_rounds_od_up PASSED   [ 10%]
TESTS/test_calculation_engine.py::test_quantity_conversion_preserves_exact_metres PASSED [ 12%]
TESTS/test_calculation_engine.py::test_coating_cost_is_rate_area_over_mass PASSED [ 15%]
TESTS/test_calculation_engine.py::test_legacy_coating_cost_uses_mass_mt_per_m PASSED [ 17%]
TESTS/test_calculation_engine.py::test_internal_coating_normal_uses_inner_d PASSED [ 20%]
TESTS/test_calculation_engine.py::test_internal_coating_legacy_uses_od PASSED [ 23%]
TESTS/test_calculation_engine.py::test_legacy_freight_uses_fractional_trailers PASSED [ 25%]
TESTS/test_calculation_engine.py::test_normal_freight_uses_whole_trailers PASSED [ 28%]
TESTS/test_financial_cost_normal_algebraic PASSED                        [ 30%]
TESTS/test_financial_cost_legacy_manual PASSED                           [ 33%]
TESTS/test_selling_price_uses_positive_rounding PASSED                   [ 35%]
TESTS/test_selling_price_per_m_components PASSED                         [ 38%]
TESTS/test_invalid_geometry_is_rejected PASSED                           [ 41%]
TESTS/test_negative_wall_thickness_rejected PASSED                       [ 43%]
TESTS/test_invalid_unit_rejected PASSED                                  [ 46%]
TESTS/test_p4_line_agreement.py::test_normal_mode_per_line_against_golden_option1[0] PASSED [ 48%]
TESTS/test_p4_line_agreement.py::test_normal_mode_per_line_against_golden_option1[1] PASSED [ 51%]
TESTS/test_p4_line_agreement.py::test_normal_mode_per_line_against_golden_option1[2] PASSED [ 53%]
TESTS/test_p4_line_agreement.py::test_normal_mode_per_line_against_golden_option1[3] PASSED [ 56%]
TESTS/test_p4_line_agreement.py::test_normal_mode_grand_total_option1 PASSED [ 58%]
TESTS/test_p4_line_agreement.py::test_legacy_mode_per_line_against_golden[0] PASSED [ 61%]
TESTS/test_p4_line_agreement.py::test_legacy_mode_per_line_against_golden[1] PASSED [ 64%]
TESTS/test_p4_line_agreement.py::test_legacy_mode_per_line_against_golden[2] PASSED [ 66%]
TESTS/test_p4_line_agreement.py::test_legacy_mode_per_line_against_golden[3] PASSED [ 69%]
TESTS/test_p4_line_agreement.py::test_legacy_mode_grand_total PASSED     [ 71%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_normal_mode_agreement[0] PASSED [ 74%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_normal_mode_agreement[1] PASSED [ 76%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_normal_mode_agreement[2] PASSED [ 79%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_normal_mode_agreement[3] PASSED [ 82%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_legacy_mode_agreement[0] PASSED [ 84%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_legacy_mode_agreement[1] PASSED [ 87%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_legacy_mode_agreement[2] PASSED [ 89%]
TESTS/test_p4_line_agreement.py::test_python_vs_vba_legacy_mode_agreement[3] PASSED [ 92%]
TESTS/test_p4_line_agreement.py::test_invalid_geometry_guards PASSED     [ 94%]
TESTS/test_workbook_runtime_audit.py::test_macro_enabled_package_and_structure PASSED [ 97%]
TESTS/test_workbook_runtime_audit.py::test_customer_master_has_one_nonblank_customer PASSED [100%]

============================= 39 passed in 0.11s ==============================
```

---

## 5. Acceptance Decision

**Verdict: P4 PASS**
- Acceptance criteria met: VBA `CalcLine` and Python calculation reference agree within 1 rupee ($\Delta = 0$).
- Option 1 normal total (₹178,801,050) and legacy total (₹178,657,900) verified.
- Safe execution: Workbook state unmodified; DEV and LIVE untouched; untracked safety backups preserved.
