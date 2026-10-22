# Mega Pipes Quotation System — Golden Baselines & Commercial Reference Register

## Overview & Authority

This register establishes the authoritative commercial calculation baselines for the Mega Pipes Quotation System (`Mega_Pipes_Quotation_DEV.xlsm`). It defines the exact formulas, parameters, reference workbooks, cell locations, and golden test values across all supported calculation modes.

---

## 1. Authoritative Reference Workbooks

| Reference Workbook | Location / SHA-256 | Authoritative Scope | Sheet Name |
|---|---|---|---|
| **Approved Company CostSheet (Unipro)** | `C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx` | Authoritative Normal Mode commercial methodology and pricing contract | `CS Stacking 10-09-26` |
| **Unipro Customer Reference** | `C:\Users\Administrator\Downloads\MKTGENQ26-27594 - Unipro.xlsx` | 27 commercial terms, offer letter layout, signatory and customer headers | `Offer` |
| **Lodha Reference** | `C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx` | 16-column Synopsis report, coating surface area geometry, unrounded ratio `=L/K` | `Synopsis`, `CostSheet` |
| **Historical Phase 3 Legacy** | `MKTGENQ26-27-594_-_Unipro_CS.xlsx` | Legacy historical bug-replication baseline for backward compatibility | `CS Stacking 10-09-26` |

---

## 2. Authoritative Company Approved Baseline (Normal Mode)

The company's approved CostSheet (`CS Stacking 10-09-26` in `MKTGENQ26-27-594 - Unipro 2.xlsx`) is the primary commercial contract for Normal Mode quotation calculations.

### Enquiry Specifications (Unipro 4-Line Enquiry)
- **Customer:** Unipro Techno Infrastructure Pvt. Ltd.
- **Enquiry Reference:** `MKTG/ENQ/26-27/594`
- **Product:** HSAW Coated, ID Basis, Grade IS3589 Fe 410, Length 12.0 m
- **Supplying Plant:** Jabalpur (P02 / P01)
- **Delivery Destination:** Krishna Nagar, Bagodar, Jharkhand - 825322
- **Coating:** External PU 1000 Micron (`EXT-PU-1000`), Internal FGE 406 Micron (`INT-FGE-406`)

### Approved Pricing Summary

| Line | Specified ID | WT (mm) | Calculated OD | Quantity (m) | Pipe Count | Weight (MT) | Approved Unit Price | Approved Line Value |
|:---:|---:|---:|---:|---:|---:|---:|---:|---:|
| **1** | 1226 mm | 8.4 mm | 1243 mm | 2,000 m | 167 pcs | 511.511 MT | **₹22,351 / m** | **₹4,47,02,000** |
| **2** | 1300 mm | 8.4 mm | 1317 mm | 350 m | 30 pcs | 94.881 MT | **₹23,570 / m** | **₹82,49,500** |
| **3** | 1350 mm | 8.4 mm | 1367 mm | 500 m | 42 pcs | 140.716 MT | **₹24,396 / m** | **₹1,21,98,000** |
| **4** | 1400 mm | 8.4 mm | 1417 mm | 4,400 m | 367 pcs | 1,283.928 MT | **₹25,220 / m** | **₹11,09,68,000** |
| **TOTAL** | — | — | — | **7,250 m** | **606 pcs** | **2,031.036 MT** | — | **₹17,61,17,500** |

- **Subtotal before GST:** **₹17,61,17,500.00**
- **GST @ 18%:** **₹3,17,01,150.00**
- **Grand Total (Incl. GST):** **₹20,78,18,650.00**
- **Amount in Words:** *Rupees Seventeen Crore Sixty One Lakh Seventeen Thousand Five Hundred Only*

### Approved Component Formulas & Reference Parameters

| Component | Approved Formula / Cell Rule | Reference Value / Basis | Reference Cell (`CS Stacking 10-09-26`) |
|---|---|---|---|
| **Steel Base Rate** | Fixed rate from management | ₹62,000.00 / MT | Row 25 (`C25:F25`) |
| **Yield Allowance** | `SteelBase * YieldPct` | 2.0% (+₹1,240.00 / MT) | Row 28 (`C28:F28`) |
| **Conversion Charge** | Manufacturing conversion | ₹4,000.00 / MT | Row 29 (`C29:F29`) |
| **Pre-Cost per MT** | `SteelBase + YieldLoss + Conversion` | **₹67,240.00 / MT** | Derived base cost |
| **External Coating** | `PI * (OD / 1000) * ExtRate` | ₹500.00 / m² on Outer Diameter ($OD$) | Row 34 (`B34=500`, `C34:F34`) |
| **Internal Coating** | `PI * (OD / 1000) * IntRate` | **₹250.00 / m² on Outer Diameter ($OD$)** | Row 35 (`B35=250`, `C35:F35`) |
| **Freight Method** | `ROUNDUP(TripRate * Trailers / MT, 0) + Handling` | Trip rate ₹93,000, Handling ₹100/MT | Row 37 (`C37:F37`) |
| **Trailer Allocation** | `(Metres / 12) / 4` (Fractional trailers) | L1: 41.67, L2: 7.29, L3: 10.42, L4: 91.67 | Cell formulas in freight row |
| **Financial Cost** | `ROUND(BaseCost * 1.18 * 0.01, 0)` / Manual inputs | 1.0% (Manual inputs: L1=1024, L2=1014, L3=1011, L4=1008) | Row 32 & Row 41 |
| **Component Per-Metre** | `ROUND(ComponentPerMT * MassPerM, 0)` | Discrete ceiling rounding | Rows 44–48 |
| **Selling Price / Metre** | `SUM(BarePerM, ExtPerM, IntPerM, FreightPerM)` | L1: 22351, L2: 23570, L3: 24396, L4: 25220 | Row 48 (`C48:F48`) |

---

## 3. Variance Bridge: Reconciling ₹17,88,01,050 to ₹17,61,17,500

The previous Option 1 implementation resulted in a subtotal of ₹17,88,01,050 (₹26,83,550 higher than approved). The exact rupee-by-rupee reconciliation across all four lines is documented below:

```
Net Subtotal Variance Bridge: -₹26,83,550
├── 1. Freight Variance (-₹9,48,350):
│   ├── Previous Option 1: Whole-trailer integer rounding (42, 8, 11, 92 trailers) @ ₹98,000.24 trip rate
│   └── Approved CostSheet: Fractional trailer allocation (41.67, 7.29, 10.42, 91.67) @ ₹93,000 trip rate
├── 2. Internal Coating Variance (-₹6,78,050):
│   ├── Previous Option 1: ₹275/m² applied to Inner Diameter (ID) perimeter
│   └── Approved CostSheet: ₹250/m² applied to Outer Diameter (OD) perimeter
├── 3. Manufacturing Pre-Cost & Steel Variance (+₹81,200):
│   ├── Previous Option 1: Steel ₹61,000 + 2% (₹1,220) + Conversion ₹5,000 = ₹67,220/MT
│   └── Approved CostSheet: Steel ₹62,000 + 2% (₹1,240) + Conversion ₹4,000 = ₹67,240/MT (+₹20/MT)
└── 4. Financial Cost Variance (-₹11,38,350):
    ├── Previous Option 1: 1.5% algebraic margin markup k / (1 - k) applied to gross cost
    └── Approved CostSheet: 1.0% financing charges (manual inputs 1024, 1014, 1011, 1008)
```

### Line-by-Line Variance Breakdown
- **Line 1 (2,000 m):** ₹22,689 $\rightarrow$ **₹22,351** (-₹338/m = -₹6,76,000)
- **Line 2 (350 m):** ₹24,109 $\rightarrow$ **₹23,570** (-₹539/m = -₹1,88,650)
- **Line 3 (500 m):** ₹24,857 $\rightarrow$ **₹24,396** (-₹461/m = -₹2,30,500)
- **Line 4 (4,400 m):** ₹25,581 $\rightarrow$ **₹25,220** (-₹361/m = -₹15,88,400)
- **Total Variance Reconciled:** **-₹26,83,550**

---

## 4. Historical Option 1 Baseline (Preserved for Regression)

Option 1 represents the historical formalization documented in `ANTIGRAVITY_MASTER_PROMPT.md` (dated 2026-10-12). It is 100% preserved in the codebase under `RateInput.Option1Mode = True` to guarantee regression integrity.

### Option 1 Line Items & Totals
- Line 1 (1226mm ID, 2000m): **₹22,689 / m** $\rightarrow$ **₹4,53,78,000**
- Line 2 (1300mm ID, 350m): **₹24,109 / m** $\rightarrow$ **₹84,38,150**
- Line 3 (1350mm ID, 500m): **₹24,857 / m** $\rightarrow$ **₹1,24,28,500**
- Line 4 (1400mm ID, 4400m): **₹25,581 / m** $\rightarrow$ **₹11,25,56,400**
- **Subtotal before GST:** **₹17,88,01,050**
- **GST @ 18%:** **₹3,21,84,189**
- **Grand Total (Incl. GST):** **₹21,09,85,239**

---

## 5. Historical Phase 3 Legacy Baseline (Preserved for Regression)

Phase 3 Legacy mode reproduces the original spreadsheet calculation bug-for-bug. It is 100% preserved in the codebase under `RateInput.legacyMode = True`.

### Legacy Line Items & Totals
- Line 1 (1226mm ID, 2000m): **₹22,681 / m** $\rightarrow$ **₹4,53,62,000**
- Line 2 (1300mm ID, 350m): **₹23,914 / m** $\rightarrow$ **₹83,69,900**
- Line 3 (1350mm ID, 500m): **₹24,748 / m** $\rightarrow$ **₹1,23,74,000**
- Line 4 (1400mm ID, 4400m): **₹25,580 / m** $\rightarrow$ **₹11,25,52,000**
- **Subtotal before GST:** **₹17,86,57,900**
- **GST @ 18%:** **₹3,21,58,422**
- **Grand Total (Incl. GST):** **₹21,08,16,322**

---

## 6. Lodha Project Reference Baseline

From `C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx`:
- **Scope:** 5 distinct line items (OD 457.2mm to 1016mm, WT 6.35mm to 12.7mm).
- **Synopsis Precision Contract:**
  - Column M (CBM Ratio): Unrounded `=L/K` formatted as `#,##0.00`.
  - Column P (12m Pipe Weight): Unrounded `=12*J/1000` formatted as `#,##0.00`.
- **Status:** Verified to 16 decimal places agreement via `TESTS/test_synopsis_precision.py` and `modSynopsis.RunSynopsisSelfTests`.

---

## 7. Verification Gates & Execution Commands

### Python Pytest Suite (94 Tests)
```bash
..\.venv\Scripts\python.exe -m pytest
```
*Expected Result:* `94 passed in ~10s` (Includes 7 approved baseline tests, 13 historical baseline tests, 19 line agreement tests, 20 PIN distance tests, 7 synopsis tests, 7 variance regression tests, 1 concurrency test, 2 audit tests).

### Real-Excel COM VBA Golden Test Macros
Executed in `Mega_Pipes_Quotation_DEV.xlsm`:
- `modSelfTest.RunGoldenApproved`: Verifies approved baseline (₹17,61,17,500) $\rightarrow$ **PASS**
- `modSelfTest.RunGoldenNormal`: Verifies Option 1 baseline (₹17,88,01,050) $\rightarrow$ **PASS**
- `modSelfTest.RunGoldenLegacy`: Verifies Legacy baseline (₹17,86,57,900) $\rightarrow$ **PASS**
- `modExport.RunCostingSelfTests`: Verifies COSTING sheet layout and values $\rightarrow$ **PASS**
- `modExport.RunOfferSelfTests`: Verifies OFFER sheet layout, values, and words $\rightarrow$ **PASS**

---
*Document Version: 1.0.0 · Authoritative Reference Date: 2026-10-17*
