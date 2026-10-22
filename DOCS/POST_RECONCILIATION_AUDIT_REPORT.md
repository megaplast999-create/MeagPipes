# MEGA PIPES — POST-RECONCILIATION AUDIT REPORT

**Document Ref:** `DOCS/POST_RECONCILIATION_AUDIT_REPORT.md`  
**Date:** 2026-10-21  
**Audit Author:** Senior Costing Systems Auditor & QA Lead  
**Audit Status:** Complete — Verified with Empirical Test Evidence  
**Scope:** Post-Reconciliation Verification, Source Workbook Forensics, Pilot UAT, and DEV Master Protection

---

## 1. Executive Summary

This independent technical audit evaluates the outcomes of the formula reconciliation gate conducted on the Mega Pipes Quotation System. The audit verified the mathematical, commercial, and structural integrity of the costing engine against three live commercial enquiry files:
1. `MKTGENQ26-27594 - Unipro.xlsx`
2. `MKTGENQ26-27-594 - Unipro 2.xlsx`
3. `MKTGENQ26-27610 - Lodha 1.xlsx`

### Summary of Audit Verdicts:
- **Baseline Checksum Verification:** **PASSED (100% Intact).**  
  `DEV/Mega_Pipes_Quotation_DEV.xlsm` was independently hashed. Its SHA-256 is `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA`, matching the protected master baseline byte-for-byte.
- **Formula Parity Verification:** **PASSED (14/14 Automated Tests OK).**  
  Zero variance detected across all 4 Unipro Stacking items (₹22,351, ₹23,570, ₹24,396, ₹25,220; project total ₹17,61,17,500), all 4 Unipro Nesting items (103 trips, ₹5,069.88/MT pooled freight, project total ₹17,42,74,350), all 5 Lodha lines (8,057.338 MT), and the manual selling price override pathway.
- **Pilot UAT Workflows:** **PASSED (10/10 Scenarios OK).**  
  Successfully verified on an isolated copy of `PILOT_COPY.xlsm`.
- **Location UI Migration Boundary:** **DEFERRED & SAFE.**  
  Location and territory summary UI remains strictly confined to `PILOT_COPY.xlsm`. No code was migrated to `DEV`.

---

## 2. Protected Files & Independent Hash Verification

| File Description | Absolute Path | Reported / Expected SHA-256 | Actual Measured SHA-256 | Status |
|---|---|---|---|---|
| **DEV Master Workbook** | `C:\Users\Administrator\Desktop\Mega Pipes\Quotation System\DEV\Mega_Pipes_Quotation_DEV.xlsm` | `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA` | `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA` | **VERIFIED (MATCH)** |
| **Pilot Copy Workbook** | `C:\Users\Administrator\Desktop\Mega Pipes\Quotation System\DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm` | N/A (Working Copy) | `47AD1649D12BE2C3A8BDC20FDF5EA42AD902D168BA43FC12CF751865BD47F604` | **STABLE PILOT** |
| **Python Calculation Engine** | `C:\Users\Administrator\Desktop\Mega Pipes\Quotation System\calculation_engine.py` | 408 lines, 14,586 bytes | `688E237936E7B6C07FD9CDA2E05FF7D6749FEE73FEFBDA46E26FBCBE28B10BC7` | **VERIFIED** |
| **Source Parity Test Suite** | `C:\Users\Administrator\Desktop\Mega Pipes\Quotation System\TESTS\test_costsheet_source_parity.py` | 277 lines, 10,750 bytes | `9A72BC77D8C5B32194FB094DE741D8B416D41FD6B1A006323FD8B2D5D44A14EC` | **VERIFIED** |

---

## 3. Independent Reproduction of Formula Parity Tests

The automated test suite `TESTS/test_costsheet_source_parity.py` was executed in a clean Python sub-process against the live openpyxl objects of the downloaded Excel workbooks.

### Test Run Log Output:
```
Ran 14 tests in 4.386s — OK (0 Failures, 0 Errors)
```

### Detailed Results by Test Case:
1. `test_line_1_stacking_golden_baseline` — **PASS:**  
   - Specification: 1243 mm OD $\times$ 8.4 mm WT, 2,000 metres.  
   - Physical: Mass = $0.25575553836$ MT/m, Total = $511.51108$ MT, Pipes = 167 Nos.  
   - Rates: Stacking Freight = ₹7,676.00/MT, Finance Cost = ₹1,024.00/MT.  
   - Unit Components: Bare = ₹17,459/m, Ext PU = ₹1,952/m, Int FGE = ₹976/m, Freight = ₹1,964/m.  
   - Evaluated Unit Price: **₹22,351.00 / m** (Source: `CS Stacking!C48` = `22351`). Variance = **₹0.00**.  
   - Line Total: **₹4,47,02,000.00** (Source: `CS Stacking!C49` = `44702000`). Variance = **₹0.00**.

2. `test_line_2_stacking_golden_baseline` — **PASS:**  
   - Specification: 1317 mm OD $\times$ 8.4 mm WT, 350 metres.  
   - Physical: Mass = $0.27108512676$ MT/m, Total = $94.87979$ MT, Pipes = 30 Nos.  
   - Rates: Stacking Freight = ₹7,248.00/MT, Finance Cost = ₹1,014.00/MT.  
   - Unit Components: Bare = ₹18,503/m, Ext PU = ₹2,068/m, Int FGE = ₹1,034/m, Freight = ₹1,965/m.  
   - Evaluated Unit Price: **₹23,570.00 / m** (Source: `CS Stacking!D48` = `23570`). Variance = **₹0.00**.  
   - Line Total: **₹82,49,500.00** (Source: `CS Stacking!D49` = `8249500`). Variance = **₹0.00**.

3. `test_line_3_stacking_golden_baseline` — **PASS:**  
   - Specification: 1367 mm OD $\times$ 8.4 mm WT, 500 metres.  
   - Physical: Mass = $0.28144295676$ MT/m, Total = $140.72148$ MT, Pipes = 42 Nos.  
   - Rates: Stacking Freight = ₹6,985.00/MT, Finance Cost = ₹1,011.00/MT.  
   - Unit Components: Bare = ₹19,209/m, Ext PU = ₹2,147/m, Int FGE = ₹1,074/m, Freight = ₹1,966/m.  
   - Evaluated Unit Price: **₹24,396.00 / m** (Source: `CS Stacking!E48` = `24396`). Variance = **₹0.00**.  
   - Line Total: **₹1,21,98,000.00** (Source: `CS Stacking!E49` = `12198000`). Variance = **₹0.00**.

4. `test_line_4_stacking_golden_baseline` — **PASS:**  
   - Specification: 1417 mm OD $\times$ 8.4 mm WT, 4,400 metres.  
   - Physical: Mass = $0.29180078676$ MT/m, Total = $1,283.92346$ MT, Pipes = 367 Nos.  
   - Rates: Stacking Freight = ₹6,740.00/MT, Finance Cost = ₹1,008.00/MT.  
   - Unit Components: Bare = ₹19,915/m, Ext PU = ₹2,225/m, Int FGE = ₹1,113/m, Freight = ₹1,967/m.  
   - Evaluated Unit Price: **₹25,220.00 / m** (Source: `CS Stacking!F48` = `25220`). Variance = **₹0.00**.  
   - Line Total: **₹11,09,68,000.00** (Source: `CS Stacking!F49` = `110968000`). Variance = **₹0.00**.

5. `test_unipro_total_project_value_parity` — **PASS:**  
   - Sum of 4 Line Totals: $44702000 + 8249500 + 12198000 + 110968000 = \mathbf{₹17,61,17,500.00}$.  
   - Commercial Offer Source: `MKTGENQ26-27-594 - Unipro 2.xlsx!Sheet1!H16` = `176117500`.  
   - Delta: **₹0.00 (Zero variance).**

6. `test_nesting_trips_and_freight_pool` — **PASS:**  
   - Total Nesting Trips: 103 trips across 7 nesting combinations (Source: `CS NESTING!N103` = `103`).  
   - Total Freight: $103 \times 98,000 = \mathbf{₹1,00,94,000}$ (Source: `CS NESTING!Q66` = `10094000`).  
   - Pooled Freight per MT: $\frac{10094000}{2031.036} + 100 = \mathbf{₹5,069.88 / MT}$ (Source: `CS NESTING!Q67` = `5069.8779`).

7. `test_nesting_per_metre_prices` — **PASS:**  
   - Per-Metre Prices: Line 1 = ₹21,796; Line 2 = ₹24,481; Line 3 = ₹25,416; Line 4 = ₹24,865.  
   - Project Value: **₹17,42,74,350** (Source: `CS NESTING!G49` = `174274350`).

8. `test_lodha_lines_1_to_5` (5 tests) — **PASS:**  
   - Line 1 (OD 1050, WT 12): ID = 1026.0 mm, Mass = 307.184 kg/m, MT = 112.122 MT.  
   - Line 2 (ID 1050, WT 12): OD = 1074.0 mm, Mass = 314.286 kg/m, MT = 7868.154 MT.  
   - Line 3 (OD 900, WT 12): ID = 876.0 mm, Mass = 262.793 kg/m, MT = 31.535 MT.  
   - Line 4 (OD 700, WT 12): ID = 676.0 mm, Mass = 203.605 kg/m, MT = 16.288 MT.  
   - Line 5 (OD 1000, WT 12): ID = 976.0 mm, Mass = 292.387 kg/m, MT = 29.239 MT.  
   - Total Lodha Tonnage: **8,057.338 MT** (Source: `Lodha 1!Sheet1!K17` = `8057.3382`).

9. `test_manual_override_audit_trail` — **PASS:**  
   - Line 1 calculated price (₹22,351) overridden to ₹22,350.  
   - Verified that `applied_unit_price = 22350.0`, `price_source = "MANUAL_OVERRIDE"`, `price_variance = -1.0`, and line extension = ₹4,47,00,000.

---

## 4. Finance Cost (FC) Control Audit

### Findings:
1. **Row 32 vs Row 41 Relationship:**  
   In both `CS NESTING` and `CS Stacking`, Row 32 contains manually entered numbers, while Row 41 evaluates `=Grand_Total_Incl_GST * FC%`.
2. **Behavioral Verification:**
   - **Suggested FC Formula:** Evaluated algebraically as $\text{Suggested\_FC} = \text{ROUND}(\text{Delivered\_Incl\_GST} \times \text{FC\%}, 0)$, producing ₹1,019/MT (pre-cost basis) or ₹1,031/MT (closed-loop circular solution).
   - **Manual FC Override:** When a historical value (e.g. ₹1,024) is entered into `manual_fc`, the calculation engine overrides the suggested value and uses ₹1,024 as the active cost component.
   - **Audit Trail:** The engine records `price_source` / `FC_Mode` as `"MANUAL_OVERRIDE"`.
   - **Persistence:** Recalculation does not overwrite manual inputs. Clearing the field reverts to `Suggested_FC`.

---

## 5. Audit of Known Workbook Issues

### A. Broken Reference `[=#REF!]` in Cell C11
- **Affected Sheets:** `MKTGENQ26-27-594 - Unipro 2.xlsx` $\rightarrow$ `CS NESTING10-09-2026!C11` and `CS Stacking 10-09-26!C11`.
- **Finding:** Cell B11 contains label `HSAW`, and C11 contains `[=#REF!]`. This was originally a link to the enquiry type cell on the Synopsis sheet (`Synopsis!L4`), which was broken when the sheet template was reorganized.
- **Downstream Dependency:** **None.** No formula in either sheet references `C11`.
- **Classification:** **Confirmed Benign Source Issue.** No repair to source files is required or permitted.

### B. Incomplete Synopsis SUM Ranges
- **Affected Sheets:** `MKTGENQ26-27594 - Unipro.xlsx!Sheet1` and `MKTGENQ26-27-594 - Unipro 2.xlsx!Synopsis`.
- **Observed Formulas:**
  - `I16: =SUM(I11:I14)` = 239 Nos (Omits Line 4 / Row 15: 367 Nos).
  - `K16: =SUM(K11:K14)` = 747.11 MT (Omits Line 4 / Row 15: 1,283.92 MT).
  - `N16: =SUM(N11:N14)` = 11,412 m² (Omits Line 4 / Row 15: 19,596 m²).
  - `O16: =SUM(O11:O14)` = 11,261 m² (Omits Line 4 / Row 15: 19,363 m²).
- **Evidence of True Total:**
  - In `CS Stacking 10-09-26`, all 4 sizes are correctly summed in Row 18 (`G18: =SUM(C18:F18) = 7250 m`), Row 19 (`G19: =SUM(C19:F19) = 2031.036 MT`), and Row 22 (`G22: =SUM(C22:F22) = 606 Nos`).
  - In RMR coating rows 68 and 79, total SQM is `30,622.271` and `31,005`, exactly matching the 4-line total.
- **Classification:** **Confirmed Source Template Defect.** The software engine correctly aggregates all active line items without truncating the ranges.

### C. Stacking vs Nesting Freight Separation
- **Stacking:** Confirmed un-nested model using 151 trips @ ₹93,000/trip, with size-specific freight per MT.
- **Nesting:** Confirmed telescoped model using 103 trips @ ₹98,000/trip, with pooled flat freight of ₹5,070/MT.
- **Classification:** **Confirmed Intentional Dual-Model Behavior.**

---

## 6. Commercial Risk Assessment

1. **Risk of Commercial Price Drift:** **ZERO.** The engine produces the exact rupee values of the approved Unipro baseline (₹22,351, ₹23,570, etc.).
2. **Risk of Accidental Master Corruption:** **ZERO.** The DEV master workbook was verified before and after all runs, remaining byte-for-byte identical to the baseline SHA-256.
3. **Risk of Uncontrolled Location Migration:** **ZERO.** Location and competitor distance UI remains strictly isolated in `PILOT_COPY.xlsm`.
