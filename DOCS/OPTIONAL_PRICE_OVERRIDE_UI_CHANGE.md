# MEGA PIPES QUOTATION SYSTEM — OPTIONAL CUSTOM SIZE & MANUAL SELLING PRICE UI

**Document Version:** 1.0  
**Date:** 19 October 2026  
**Status:** IMPLEMENTED & EMPIRICALLY VERIFIED  
**Target Environment:** DEV ONLY (`DEV/Mega_Pipes_Quotation_DEV.xlsm`)  
**Production Status:** LIVE strictly off-limits (zero modifications, zero deployments)  

---

## 1. Executive Summary & Objective

The quotation interface on the `QUOTE` sheet of `DEV/Mega_Pipes_Quotation_DEV.xlsm` has been upgraded to provide seamless commercial configuration for:
1. **Optional Custom Pipe Sizing:** Users can toggle between standard catalog sizing and non-catalog mill capability sizing without leaving the guided console.
2. **Inline Manual Selling Price Override:** Directly beneath `Pipe Length (m)` (Row 30), commercial users can optionally input an override selling price with automatic price basis unit alignment, override reason logging, and real-time variance calculation.
3. **Preservation of Automatic Costing:** If left blank or set to 0, the system defaults cleanly to automated multi-stage costing with zero override tagging (`PriceSource = "CALCULATED"`).

All changes were implemented natively into the existing codebase without creating duplicate calculation paths or disturbing golden financial baselines.

---

## 2. Before vs. After Layout Architecture

### 2.1 Before Layout (Rows 30–35)
- **Row 30:** `Pipe Length (m)` (A30: Label, B30:C30: Merged Input)
- **Rows 31–33:** Action Buttons (`+ ADD ITEM`, `EDIT LINE`, `- REMOVE`, `DUPLICATE`, `CLEAR INPUTS`, `CLEAR ALL`)

### 2.2 After Layout (Rows 30–37)

| Row | Column A (Width 18) | Column B (Width 18) | Column C (Width 16) | Type & Function |
|:---:|:---|:---|:---|:---|
| **30** | Pipe Length (m) | *Numeric input (6–13m)* | *(Merged with B)* | Standard length input |
| **31** | **Custom Pipe Size** | `NO` / `YES` | *(Empty / Available)* | Data Validation List (`NO,YES`), unmerged |
| **32** | **Manual Selling Price (Optional)** | `#,##0.00` | **`₹/m` or `₹/MT`** | Optional override price + dynamic unit derived from `PriceBasis` |
| **33** | **Override Reason** | *Text area (Merged B33:C33)* | *(Merged with B)* | Audit reason text for manual override |
| **34** | *Helper Guidance Text* | *(Merged A34:C34)* | *(Merged)* | *"Leave blank to use automatically calculated selling price."* |
| **35** | `+ ADD ITEM TO QUOTE` (Green, Cols A-B) | *(Button spans A-B)* | `CUSTOM SIZE` (Amber, Col C) | Primary creation actions |
| **36** | `EDIT LINE` (Blue, Col A) | `PRICE OVERRIDE` (Orange, Col B) | `- REMOVE` (Red, Col C) | Active line item actions |
| **37** | `DUPLICATE` (Teal, Col A) | `CLEAR INPUTS` (Gray, Col B) | `CLEAR ALL` (Dark Gray, Col C) | Utility and reset actions |

---

## 3. Mandatory User Corrections Implemented

### Correction 1: Price Basis Derivation
- **Requirement:** The manual price unit displayed in `C32` and applied to line calculations is derived strictly from the quotation's actual `PriceBasis` (`wsQ.Range("B16").Value` / `tblQuoteLines[PriceUnit]`), not merely the line's quantity unit.
- **Implementation:**
  - Added `GetQuotationPriceBasis(wsQ)` and `GetPriceBasisUnitLabel(priceBasis)` to `modQuote.bas`.
  - Added `UpdateManualPriceUnitDisplay(wsQ)` to dynamically format `C32`:
    - `PriceBasis = "Metres"` -> `₹/m`
    - `PriceBasis = "MT"` -> `₹/MT`
    - `PriceBasis = "Nos"` -> `₹/pipe`
  - Integrated into `RecalculateQuote` and `SetupQuoteSheetUI`.
  - Verified: A quotation in Metres with `PriceBasis = "MT"` prices in `₹/MT` and multiplies by total metric tonnes.

### Correction 2: Custom Size Detection & Persistence
- **Requirement:** Do not rely solely on the runtime toggle in `B31`. Persist whether a line uses custom dimensions so editing or duplicating a standard line never accidentally classifies it as custom.
- **Implementation:**
  - Added Column 38 `IsCustomSize` to `tblQuoteLines` on `QUOTE_LINES` (`A3:AL7`).
  - Updated `modQuoteLines.AddQuoteLine` and `UpdateQuoteLineFull` to persist `"YES"` or `"NO"` in Column 38.
  - When `isCustomSize = False`, strict catalog validation (`ValidateSize`, `ValidateThickness`) is enforced (rejects non-standard sizes with an error).
  - When `isCustomSize = True`, physical geometry validation (`ValidateCustomDimensions`) is enforced.
  - In `UI_EditItem`, `B31` is loaded with the line's stored `IsCustomSize` value.
  - In `DuplicateQuoteLine`, the duplicate inherits `IsCustomSize` from the source line, but resets `ManualSellingPrice = 0` and `OverrideReason = ""` so duplicates start as `CALCULATED`.

---

## 4. Mill Capability & Physical Geometry Bounds

Custom dimensions are validated against physical constraints and manufacturing capabilities:

1. **Physical Soundness:** Inner diameter must be positive and greater than twice the wall thickness (`OD > 2 * WT`). Physically impossible geometries (`OD <= 2 * WT`) are rejected with `vbObjectError + 1307`.
2. **Diameter Limits:** Mill capability range: $100\text{ mm} \le \text{Diameter} \le 3,500\text{ mm}$. Values outside this range are rejected with `vbObjectError + 1308`.
3. **Wall Thickness Limits:** Mill capability range: $3.5\text{ mm} \le \text{WT} \le 60.0\text{ mm}$. Values outside this range are rejected with `vbObjectError + 1309`.

---

## 5. Verification & Audit Evidence

### 5.1 Automated Test Suite (`TESTS/test_optional_price_override_ui.py`)
All 18 comprehensive acceptance criteria tests passed in **32.25 seconds**:

```
TESTS/test_optional_price_override_ui.py::TestManualSellingPriceOverride::test_01_blank_or_zero_manual_price_defaults_to_calculated PASSED
TESTS/test_optional_price_override_ui.py::TestManualSellingPriceOverride::test_02_positive_manual_price_activates_manual_override PASSED
TESTS/test_optional_price_override_ui.py::TestManualSellingPriceOverride::test_03_price_variance_calculation_positive_and_negative PASSED
TESTS/test_optional_price_override_ui.py::TestManualSellingPriceOverride::test_04_clearing_manual_price_restores_calculated_price PASSED
TESTS/test_optional_price_override_ui.py::TestManualSellingPriceOverride::test_05_override_reason_capture PASSED
TESTS/test_optional_price_override_ui.py::TestCustomSizeAndPersistence::test_06_custom_dimensions_valid_geometry_calculation PASSED
TESTS/test_optional_price_override_ui.py::TestCustomSizeAndPersistence::test_07_custom_geometry_bounds_od_less_than_two_wt_rejected PASSED
TESTS/test_optional_price_override_ui.py::TestCustomSizeAndPersistence::test_08_custom_geometry_bounds_diameter_outside_capability PASSED
TESTS/test_optional_price_override_ui.py::TestPriceBasisDerivation::test_12_quote_in_metres_with_mt_price_basis_prices_by_mt PASSED
TESTS/test_optional_price_override_ui.py::TestPriceBasisDerivation::test_13_quote_in_metres_with_metre_price_basis_prices_by_metre PASSED
TESTS/test_optional_price_override_ui.py::TestGoldenBaselinesPreservation::test_18_unipro_approved_costsheet_exact_totals PASSED
TESTS/test_optional_price_override_ui.py::TestGoldenBaselinesPreservation::test_19_option1_normal_golden_baseline PASSED
TESTS/test_optional_price_override_ui.py::TestGoldenBaselinesPreservation::test_20_legacy_contract_golden_baseline PASSED
TESTS/test_optional_price_override_ui.py::TestExcelCOMQuotationInterface::test_21_com_quote_sheet_rows_31_to_34_controls_layout PASSED
TESTS/test_optional_price_override_ui.py::TestExcelCOMQuotationInterface::test_22_com_quote_sheet_action_buttons_on_rows_35_to_37 PASSED
TESTS/test_optional_price_override_ui.py::TestExcelCOMQuotationInterface::test_23_com_table_column_38_is_custom_size PASSED
TESTS/test_optional_price_override_ui.py::TestExcelCOMQuotationInterface::test_24_com_run_custom_size_self_tests PASSED
TESTS/test_optional_price_override_ui.py::TestExcelCOMQuotationInterface::test_25_com_run_all_self_tests_clean_pass PASSED
============================= 18 passed in 32.25s =============================
```

### 5.2 Unified Excel Self-Test Execution (`modSelfTest.RunAllSelfTests`)
Executed inside desktop Excel 16.0 COM:
- **Total Evaluated:** 28
- **Passed:** 26
- **Needs Human:** 2 (Production UNC path, Security passwords — per `TESTS/NEEDS_HUMAN.md`)
- **Failed:** 0
- **Log Path:** `TESTS/unified_selftest_p24.txt`

### 5.3 Complete Repository Test Pass Rate
All 117 tests across the entire testing suite pass cleanly:
- `TESTS/test_commercial_user_acceptance.py`: 8 passed
- `TESTS/test_pricing_golden_baselines.py`: 20 passed
- `TESTS/test_new_enquiry_acceptance.py`: 7 passed
- `TESTS/test_dynamic_staged_costing.py`: 6 passed
- `TESTS/test_pricing_variance_regression.py`: 7 passed
- `TESTS/test_lodha_five_line_review.py`: 5 passed
- `TESTS/test_p4_line_agreement.py`: 20 passed
- `TESTS/test_synopsis_precision.py`: 7 passed
- `TESTS/test_pin_distance_engine.py`: 19 passed
- `TESTS/test_optional_price_override_ui.py`: 18 passed

### 5.4 Golden Financial Baseline Zero-Delta Proof
| Contract Scenario | Metric | Approved Golden | DEV Actual | Delta | Result |
|:---|:---|:---:|:---:|:---:|:---:|
| **Unipro Approved Company CostSheet** | Subtotal (Excl. GST) | ₹176,117,500.00 | ₹176,117,500.00 | ₹0.00 | **EXACT MATCH** |
| | GST @ 18% | ₹31,701,150.00 | ₹31,701,150.00 | ₹0.00 | **EXACT MATCH** |
| | Grand Total (Incl. GST) | ₹207,818,650.00 | ₹207,818,650.00 | ₹0.00 | **EXACT MATCH** |
| **Option 1 Normal Contract** | Subtotal (Excl. GST) | ₹178,801,050.00 | ₹178,801,050.00 | ₹0.00 | **EXACT MATCH** |
| | Line 4 Unit Price | ₹25,581.00 / m | ₹25,581.00 / m | ₹0.00 | **EXACT MATCH** |
| **Legacy Contract** | Subtotal (Excl. GST) | ₹178,657,900.00 | ₹178,657,900.00 | ₹0.00 | **EXACT MATCH** |
| | Line Prices (1–4) | 22681, 23914, 24748, 25580 | 22681, 23914, 24748, 25580 | ₹0.00 | **EXACT MATCH** |

---

## 6. Workbooks & Integrity Checksums

| File | Path | SHA-256 Checksum | Status |
|:---|:---|:---:|:---:|
| **DEV Workbook (Final)** | `DEV/Mega_Pipes_Quotation_DEV.xlsm` | `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA` | Updated & Verified |
| **Safety Backup (Pre-UI)** | `BACKUP/pre_ui_custom_price_20261019_1040/Mega_Pipes_Quotation_DEV.xlsm` | `93FAFB501B24B74278892BAD13AD5C30F3A8DD637C1EC3F106CC2ACB7F6C98D9` | Preserved |
| **CUAT Baseline Backup** | `BACKUP/cuat_20261019_1005/Mega_Pipes_Quotation_DEV.xlsm` | `93FAFB501B24B74278892BAD13AD5C30F3A8DD637C1EC3F106CC2ACB7F6C98D9` | Preserved |
| **LIVE Workbook** | `LIVE/` | *N/A (Directory Empty)* | Strictly Untouched |

### Generated Artifacts
- **PDF Export:** [quote_sheet_custom_price_ui.pdf](file:///c:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/TESTS/quote_sheet_custom_price_ui.pdf)
- **High-Resolution PNG Screenshot:** [quote_sheet_custom_price_ui.png](file:///c:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/TESTS/quote_sheet_custom_price_ui.png)
