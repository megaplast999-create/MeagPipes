# MEGA PIPES — COSTSHEET PARITY VERIFICATION RESULTS

**Document Ref:** `TESTS/COSTSHEET_PARITY_RESULTS.md`  
**Date:** 2026-10-21  
**Test Suite Executed:** `TESTS/test_costsheet_source_parity.py`  
**Execution Environment:** Python 3.13 / Windows / Isolated PILOT Test Environment  
**Status:** 100% Passed (14/14 Tests Passed, 0 Failures, 0 Errors)  
**Checksum Integrity:** `DEV/Mega_Pipes_Quotation_DEV.xlsm` SHA-256 Verified: `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA`

---

## 1. Executive Summary

This report documents the empirical regression and parity results achieved by running automated tests against the three original Excel workbooks:
1. `MKTGENQ26-27594 - Unipro.xlsx`
2. `MKTGENQ26-27-594 - Unipro 2.xlsx`
3. `MKTGENQ26-27610 - Lodha 1.xlsx`

Every calculated value from the software calculation engine was compared directly to the evaluated cell values in openpyxl. **Zero variance was detected across all primary physical, geometrical, costing, per-metre unit prices, and project total values.**

---

## 2. Unipro Stacking Parity Results (Approved Golden Baseline)

**Source Sheet:** `MKTGENQ26-27-594 - Unipro 2.xlsx!CS Stacking 10-09-26`  
**Configuration:** MPPL Jabalpur Plant, Steel ₹62,000/MT, Conversion ₹4,000/MT, PU Ext ₹500/m², FGE Int ₹250/m², Freight ₹93,000/trip, Handling ₹100/MT, 4 pipes/trailer.

### Side-by-Side Verification Table:

| Metric | Source Excel Value | Software Engine Output | Absolute Variance | Percentage Variance | Status |
|---|---|---|---|---|---|
| **Line 1 OD (mm)** | `1243.0` | `1243.0` | `0.00 mm` | `0.00%` | **MATCH** |
| **Line 1 Mass (MT/m)** | `0.25575553836` | `0.25575553836` | `0.00000000 MT/m` | `0.00%` | **MATCH** |
| **Line 1 Weight (MT)** | `511.51107672` | `511.51107672` | `0.0000 MT` | `0.00%` | **MATCH** |
| **Line 1 Pipes (Nos)** | `167` | `167` | `0 Nos` | `0.00%` | **MATCH** |
| **Line 1 Ext PU (₹/MT)** | `₹7,630.37` | `₹7,630.37` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Int FGE (₹/MT)**| `₹3,815.19` | `₹3,815.19` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Freight (₹/MT)**| `₹7,676.00` | `₹7,676.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Bare Price / m** | `₹17,459.00` | `₹17,459.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Ext Coating / m**| `₹1,952.00` | `₹1,952.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Int Coating / m**| `₹976.00` | `₹976.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Freight / m** | `₹1,964.00` | `₹1,964.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 1 Total Price / m**| **₹22,351.00** | **₹22,351.00** | **₹0.00** | **0.00%** | **MATCH** |
| **Line 1 Total Value** | **₹4,47,02,000.00**| **₹4,47,02,000.00**| **₹0.00** | **0.00%** | **MATCH** |
|---|---|---|---|---|---|
| **Line 2 OD (mm)** | `1317.0` | `1317.0` | `0.00 mm` | `0.00%` | **MATCH** |
| **Line 2 Mass (MT/m)** | `0.27108512676` | `0.27108512676` | `0.00000000 MT/m` | `0.00%` | **MATCH** |
| **Line 2 Weight (MT)** | `94.87979437` | `94.87979437` | `0.0000 MT` | `0.00%` | **MATCH** |
| **Line 2 Pipes (Nos)** | `30` | `30` | `0 Nos` | `0.00%` | **MATCH** |
| **Line 2 Ext PU (₹/MT)** | `₹7,627.46` | `₹7,627.46` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Int FGE (₹/MT)**| `₹3,813.73` | `₹3,813.73` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Freight (₹/MT)**| `₹7,248.00` | `₹7,248.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Bare Price / m** | `₹18,503.00` | `₹18,503.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Ext Coating / m**| `₹2,068.00` | `₹2,068.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Int Coating / m**| `₹1,034.00` | `₹1,034.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Freight / m** | `₹1,965.00` | `₹1,965.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 2 Total Price / m**| **₹23,570.00** | **₹23,570.00** | **₹0.00** | **0.00%** | **MATCH** |
| **Line 2 Total Value** | **₹82,49,500.00** | **₹82,49,500.00** | **₹0.00** | **0.00%** | **MATCH** |
|---|---|---|---|---|---|
| **Line 3 OD (mm)** | `1367.0` | `1367.0` | `0.00 mm` | `0.00%` | **MATCH** |
| **Line 3 Mass (MT/m)** | `0.28144295676` | `0.28144295676` | `0.00000000 MT/m` | `0.00%` | **MATCH** |
| **Line 3 Weight (MT)** | `140.72147838` | `140.72147838` | `0.0000 MT` | `0.00%` | **MATCH** |
| **Line 3 Pipes (Nos)** | `42` | `42` | `0 Nos` | `0.00%` | **MATCH** |
| **Line 3 Freight (₹/MT)**| `₹6,985.00` | `₹6,985.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 3 Total Price / m**| **₹24,396.00** | **₹24,396.00** | **₹0.00** | **0.00%** | **MATCH** |
| **Line 3 Total Value** | **₹1,21,98,000.00**| **₹1,21,98,000.00**| **₹0.00** | **0.00%** | **MATCH** |
|---|---|---|---|---|---|
| **Line 4 OD (mm)** | `1417.0` | `1417.0` | `0.00 mm` | `0.00%` | **MATCH** |
| **Line 4 Mass (MT/m)** | `0.29180078676` | `0.29180078676` | `0.00000000 MT/m` | `0.00%` | **MATCH** |
| **Line 4 Weight (MT)** | `1283.92346174` | `1283.92346174` | `0.0000 MT` | `0.00%` | **MATCH** |
| **Line 4 Pipes (Nos)** | `367` | `367` | `0 Nos` | `0.00%` | **MATCH** |
| **Line 4 Freight (₹/MT)**| `₹6,740.00` | `₹6,740.00` | `₹0.00` | `0.00%` | **MATCH** |
| **Line 4 Total Price / m**| **₹25,220.00** | **₹25,220.00** | **₹0.00** | **0.00%** | **MATCH** |
| **Line 4 Total Value** | **₹11,09,68,000.00**| **₹11,09,68,000.00**| **₹0.00** | **0.00%** | **MATCH** |
|---|---|---|---|---|---|
| **PROJECT TOTAL VALUE** | **₹17,61,17,500.00**| **₹17,61,17,500.00**| **₹0.00** | **0.00%** | **MATCH (ZERO DELTA)** |

---

## 3. Unipro Nesting Parity Results

**Source Sheet:** `MKTGENQ26-27-594 - Unipro 2.xlsx!CS NESTING10-09-2026`  
**Configuration:** MPPL Khopoli Plant, Steel ₹61,000/MT, Conversion ₹5,000/MT (L1 & L4) / ₹10,000/MT (L2 & L3), Freight ₹98,000/trip across 103 trips (₹5,070/MT flat).

| Parameter | Source Excel Value | Software Engine Output | Absolute Variance | Status |
|---|---|---|---|---|
| **Total Nesting Trips** | `103` | `103` | `0` | **MATCH** |
| **Total Freight Value** | `₹1,00,94,000` | `₹1,00,94,000` | `₹0` | **MATCH** |
| **Pooled Freight Rate** | `₹5,069.88 / MT` | `₹5,069.88 / MT` | `₹0.00` | **MATCH** |
| **Line 1 Total Price / m**| `₹21,796 / m` | `₹21,796 / m` | `₹0.00` | **MATCH** |
| **Line 2 Total Price / m**| `₹24,481 / m` | `₹24,481 / m` | `₹0.00` | **MATCH** |
| **Line 3 Total Price / m**| `₹25,416 / m` | `₹25,416 / m` | `₹0.00` | **MATCH** |
| **Line 4 Total Price / m**| `₹24,865 / m` | `₹24,865 / m` | `₹0.00` | **MATCH** |
| **Project Total Value** | `₹17,42,74,350` | `₹17,42,74,350` | `₹0` | **MATCH** |

---

## 4. Lodha Enquiry Sizing & Geometry Parity Results

**Source Sheet:** `MKTGENQ26-27610 - Lodha 1.xlsx!Sheet1`  
**Configuration:** 5 sizes, mixed OD and ID specifications, WT 12.0 mm.

| Line | Sizing Basis | Specified | Calculated Opposite | Mass (kg/m) Source | Mass (kg/m) Code | Order MT Source | Order MT Code | Status |
|---|---|---|---|---|---|---|---|---|
| **L1** | **OD** | `1050 mm` | `ID: 1026.0 mm` | `307.183644` | `307.183644` | `112.122030` | `112.122030` | **MATCH** |
| **L2** | **ID** | `1050 mm` | `OD: 1074.0 mm` | `314.286156` | `314.286156` | `7868.153915` | `7868.153915` | **MATCH** |
| **L3** | **OD** | `900 mm` | `ID: 876.0 mm` | `262.792944` | `262.792944` | `31.535153` | `31.535153` | **MATCH** |
| **L4** | **OD** | `700 mm` | `ID: 676.0 mm` | `203.605344` | `203.605344` | `16.288428` | `16.288428` | **MATCH** |
| **L5** | **OD** | `1000 mm` | `ID: 976.0 mm` | `292.386744` | `292.386744` | `29.238674` | `29.238674` | **MATCH** |
| **TOTAL**| — | — | — | — | — | **8,057.338 MT** | **8,057.338 MT** | **MATCH** |

---

## 5. Manual Selling Price Override Pathway Audit

| Scenario | Calculated Price | Manual Selling Price | Price Source Flag | Price Variance | Line Value Computed |
|---|---|---|---|---|---|
| **Unipro L1 Default** | `₹22,351.00` | `None (0.0)` | `CALCULATED` | `₹0.00` | `₹4,47,02,000.00` |
| **Unipro L1 Overridden**| `₹22,351.00` | `₹22,350.00` | `MANUAL_OVERRIDE` | `-₹1.00` | `₹4,47,00,000.00` |
| **Unipro L2 Overridden**| `₹23,570.00` | `₹23,500.00` | `MANUAL_OVERRIDE` | `-₹70.00` | `₹82,25,000.00` |

---

## 6. Required Human Approvals

The following approvals must be obtained before any code changes are deployed to the core workbook:

```
[ ] APPROVAL 1: Ratification of Manual FC Data Model
    Allows sales engineers to enter manual FC per MT or accept the auto-suggested value.
    Status: USER CONFIRMED VIA DIRECTIVE (2026-10-21). Pending formal sign-off.

[ ] APPROVAL 2: Dual-Mode Sizing UI Confirmation
    Authorizes per-line toggle between ID-specified and OD-specified pipe entries.
    Status: USER CONFIRMED VIA DIRECTIVE (2026-10-21). Pending formal sign-off.

[ ] APPROVAL 3: Location UI Migration to DEV
    Authorizes merging the location picker, PIN distance engine, and competitor benchmark UI from PILOT_COPY into DEV.
    Status: DEFERRED BY USER DIRECTIVE (PILOT testing continues).
```

---

## 7. Next Safe Implementation Step

1. Keep `DEV/Mega_Pipes_Quotation_DEV.xlsm` locked and unchanged.
2. Complete testing in `DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm`.
3. Await final written confirmation to implement the reconciled formula interface.
