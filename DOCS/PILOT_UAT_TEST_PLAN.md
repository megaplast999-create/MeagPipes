# MEGA PIPES — PILOT USER ACCEPTANCE TEST (UAT) PLAN

**Document Ref:** `DOCS/PILOT_UAT_TEST_PLAN.md`  
**Date:** 2026-10-21  
**Scope:** Pilot User Acceptance Testing on Isolated Copies of `Mega_Pipes_Quotation_PILOT_COPY.xlsm`  
**Target Workbook:** `DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm`  
**Protected Workbook:** `DEV/Mega_Pipes_Quotation_DEV.xlsm` (Strictly Off-Limits)

---

## 1. Test Objectives & Scope

This UAT plan defines 10 end-to-end commercial scenarios to validate that the reconciled costing methods, formula automation, manual override pathways, and persistence mechanisms operate with 100% reliability in the Excel pilot environment.

### Core Testing Principles:
1. **Isolated Execution:** All tests must execute against a working copy (e.g. `TESTS/pilot_uat_work.xlsm`). Never run destructive tests on `PILOT_COPY.xlsm` directly, and never touch `DEV`.
2. **Empirical Evidence:** Every test result must be validated by inspecting cell contents, workbook state, and COM return codes.
3. **No Silent Changes:** Overrides, formulas, and persistence must be audited explicitly.

---

## 2. Test Environment & Setup Instructions

1. **Preconditions:**
   - Operating System: Windows 10/11 / Server.
   - Microsoft Excel: Installed with VBA and COM enabled.
   - Python: 3.10+ with `pywin32` and `openpyxl`.
   - Master Workbook Check: Verify `DEV/Mega_Pipes_Quotation_DEV.xlsm` SHA-256 equals `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA`.
2. **Execution Script:**
   - Automated UAT suite: `TESTS/run_pilot_uat_scenarios.py`.
   - Command: `python TESTS/run_pilot_uat_scenarios.py`.

---

## 3. Detailed UAT Test Scenarios

### Scenario 1 (UAT-01): 4-Line Quotation Structure
- **Objective:** Verify that a standard 4-line quotation is correctly loaded and structured in `tblQuoteLines`.
- **Preconditions:** Isolated copy of `PILOT_COPY.xlsm` initialized.
- **Input:** 4 sizes from Unipro (1226 ID, 1300 ID, 1350 ID, 1400 ID; WT 8.4 mm; HSAW-002; Metres: 2000, 350, 500, 4400).
- **Expected Outcome:** `tblQuoteLines.ListRows.Count = 4`. Line IDs L0001 to L0004 generated. Total metres equal 7,250 m.
- **Pass Criteria:** Exact line count and total metres match.

### Scenario 2 (UAT-02): Dual-Mode ID / OD Input Sizing
- **Objective:** Verify that the system allows entering either ID (auto-calculating OD) or OD (auto-calculating ID) per line.
- **Preconditions:** UAT-01 complete.
- **Input:**
  - Line 1: `Basis = "ID"`, `Size = 1226`, `WT = 8.4`.
  - Line 2: `Basis = "OD"`, `Size = 1050`, `WT = 12.0` (Lodha pattern).
- **Expected Outcome:**
  - Line 1: `CalculatedODMM = 1243`, `CalculatedIDMM = 1226`.
  - Line 2: `CalculatedODMM = 1050`, `CalculatedIDMM = 1026`.
- **Pass Criteria:** Source rounding rule $\text{OD} = \text{ROUNDUP}(\text{ID} + 2 \times \text{WT}, 0)$ preserved.

### Scenario 3 (UAT-03): Dimensional Calculations & Mass Conversion
- **Objective:** Verify that linear mass (kg/m), order weight (MT), and pipe piece counts (Nos) calculate automatically.
- **Preconditions:** Line 1 active (1243 mm OD, 8.4 mm WT, 2,000 m).
- **Input:** Mass formula: $(\text{OD} - \text{WT}) \times \text{WT} \times 0.0246615$.
- **Expected Outcome:**
  - `MassKgPerM = 255.756 kg/m`.
  - `WeightMT = 511.511 MT`.
  - `PipeCount = 167 Nos`.
- **Pass Criteria:** Absolute delta $< 0.001$.

### Scenario 4 (UAT-04): Stacking vs Nesting Costing Model
- **Objective:** Verify that stacking freight calculates line-specific freight per MT based on trailer payload and trip rate.
- **Preconditions:** Supplying Plant = MPPL Jabalpur, Delivery = Bagodar, Trip Rate = ₹93,000.
- **Input:** 4 pipes per trailer ($2 + 2$). Line 1 trips = $41.67$.
- **Expected Outcome:** `Freight_per_MT = ROUNDUP(93000 * 41.6667 / 511.511, 0) + 100 = ₹7,676 / MT`.
- **Pass Criteria:** Freight in Costing Row 37 evaluates to ₹7,676 / MT.

### Scenario 5 (UAT-05): Suggested Finance Cost Formula Calculation
- **Objective:** Verify that the system computes suggested FC using the delivered price and payment term rate.
- **Preconditions:** Payment terms: 25% Advance, Balance against LC ($r = 1.0\%$).
- **Input:** Pre-cost delivered subtotal with GST @ 18%.
- **Expected Outcome:** Suggested FC evaluates between ₹1,019/MT and ₹1,031/MT in Costing Row 41.
- **Pass Criteria:** Suggested FC is positive, non-zero, and within documented tolerance.

### Scenario 6 (UAT-06): Manual FC Override Retention
- **Objective:** Verify that entering a manual FC in Costing Row 32 overrides the suggested value and persists.
- **Preconditions:** Line 1 active.
- **Input:** Enter `1024.0` in cell `C32`.
- **Expected Outcome:** `Applied_FC = 1024.0`. `FC_Mode` flagged as manual override.
- **Pass Criteria:** Value 1024.0 retained across save and recalculation.

### Scenario 7 (UAT-07): Manual Selling Price Override Pathway
- **Objective:** Verify that the commercial user can customize the selling price per metre on the QUOTE sheet.
- **Preconditions:** Calculated price is ₹22,351/m.
- **Input:** Enter Manual Selling Price = ₹22,350/m. Override reason = "Agreed client concession".
- **Expected Outcome:**
  - `Applied_Unit_Price = ₹22,350.00 / m`.
  - `Price_Source = "MANUAL_OVERRIDE"`.
  - `Price_Variance = -₹1.00`.
  - `Line_Total = 22350 * 2000 = ₹4,47,00,000.00`.
- **Pass Criteria:** Effective unit price updates immediately without formula damage.

### Scenario 8 (UAT-08): Recalculate & Reopen Persistence
- **Objective:** Verify quotation price stability across explicit recalculation and workbook reopen.
- **Preconditions:** Overrides and lines configured in UAT-06 and UAT-07.
- **Execution:** Execute `modQuote.RecalculateQuote`, save workbook, close Excel, reopen workbook.
- **Expected Outcome:** All line prices, subtotal (E8), grand total (E10), and manual FC overrides retain their exact values.
- **Pass Criteria:** Zero price drift upon recalculation or reopening.

### Scenario 9 (UAT-09): Generated Commercial Offer Total Alignment
- **Objective:** Verify that the customer-facing OFFER sheet displays the active commercial valuation.
- **Preconditions:** 4 lines active.
- **Input:** Inspect `OFFER!H16` and `QUOTE!E8`.
- **Expected Outcome:** Offer sheet aligns with active quote totals.
- **Pass Criteria:** Offer document reflects valid commercial amounts.

### Scenario 10 (UAT-10): Boundary & Input Validation Handling
- **Objective:** Verify that invalid dimensions, zero lengths, and boundary inputs are rejected safely without crashing VBA.
- **Preconditions:** Active quotation sheet.
- **Input:** Test invalid basis (not OD or ID), negative thickness, zero quantity.
- **Expected Outcome:** System traps error gracefully, prevents corrupted table states, and maintains valid workflow status.
- **Pass Criteria:** System state remains `Calculated` or `Ready` without unhandled runtime exceptions.

---

## 4. Pass / Fail Evaluation Matrix

| Test ID | Scenario Name | Primary Success Criteria | Critical Failure Threshold |
|---|---|---|---|
| **UAT-01** | 4-Line Quotation Structure | 4 rows in `tblQuoteLines`, 7,250m total | $< 4$ lines or truncated rows |
| **UAT-02** | Dual-Mode ID/OD Sizing | Exact OD calculated from ID; ID from OD | Incorrect rounding or formula error |
| **UAT-03** | Dimensional & Physical Mass | kg/m and MT match IS 3589 formula | Variance $> 0.01\text{ MT}$ |
| **UAT-04** | Stacking Freight Costing | Line-specific freight per MT computed | Missing or zero freight |
| **UAT-05** | Suggested FC Calculation | Suggested FC evaluates correctly | Zero or `#VALUE!` error |
| **UAT-06** | Manual FC Override | Value 1024.0 retained and applied | Overwritten by recalculation |
| **UAT-07** | Manual Selling Price Override | Applied price = override, variance tracked | Override ignored or formula overwritten |
| **UAT-08** | Recalculate & Reopen Stability | Zero price drift on recalculate/reopen | Price changes after save/load |
| **UAT-09** | Commercial Offer Consistency | Valid commercial alignment | `#REF!` or mismatched figures |
| **UAT-10** | Validation & Error Handling | Graceful rejection of invalid inputs | Unhandled VBA crash or corrupted state |
