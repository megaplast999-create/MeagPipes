# MEGA PIPES — COSTSHEET FORMULA RECONCILIATION MATRIX

**Document Ref:** `DOCS/COSTSHEET_FORMULA_RECONCILIATION_MATRIX.md`  
**Date:** 2026-10-21  
**Status:** Verification Gate Complete — Frozen for Human Approval  
**Source Workbooks Verified:**
1. `C:\Users\Administrator\Downloads\MKTGENQ26-27594 - Unipro.xlsx`
2. `C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx`
3. `C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx`

---

## 1. Executive Summary & Gate Objectives

This reconciliation matrix establishes an authoritative, mathematical cross-examination between the original commercial cost sheets and the software calculation engine. Every material formula has been mapped to its source cell, units, dependencies, evaluated results, and corresponding code path.

### Key Conclusions:
1. **Mathematical Equivalence:** The software engine (`calculation_engine.py`) and VBA modules achieve **exact parity to the rupee** across all 4 Unipro line items (₹22,351, ₹23,570, ₹24,396, ₹25,220) and total project value (₹17,61,17,500) under the approved stacking configuration.
2. **The Source `#REF!` and Incomplete SUM Errors Documented:**
   - In `Unipro 2` cost sheets, cell `C11` contains `[=#REF!]`. This is an orphaned reference to an enquiry type header that carries zero downstream calculations.
   - In `Unipro 1` and `Unipro 2` Synopsis sheets, `I16`, `K16`, `N16`, and `O16` sum only rows 11 to 14, omitting row 15 (Line 4: 4,400m). In the actual Cost Sheets (`CS Stacking` & `CS NESTING`), all 4 items are correctly summed to 7,250m, 606 pipes, and 2,031.04 MT.
3. **Finance Cost Reconciliation:**
   - Row 32 contains a hardcoded pre-estimate (₹1,024, ₹1,014, etc.) entered by estimators to bypass Excel's circular reference blocker.
   - Row 41 evaluates `=Grand_Total_Incl_GST * FC%` (₹1,031.15).
   - In accordance with user direction, the system architecture retains manual entry for FC with an auto-calculated suggested default.
4. **Safety Boundaries Preserved:**
   - `DEV/Mega_Pipes_Quotation_DEV.xlsm` remains byte-for-byte untouched (SHA-256: `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA`).
   - Location and territory work remains strictly in `PILOT_COPY`.

---

## 2. Cell-by-Cell Material Costing Formula Mapping

The table below traces every formula from the approved cost sheet `MKTGENQ26-27-594 - Unipro 2.xlsx!CS Stacking 10-09-26` (Line 1: 1226 ID x 8.4mm WT, 2,000m) to the Mega Pipes software engine:

| Metric | Source Cell & Formula | Units | Dependencies | Evaluated Source Value | Mega Pipes Code Path | Code Result | Parity & Discrepancy Classification |
|---|---|---|---|---|---|---|---|
| **Outer Diameter (OD)** | `Synopsis!D12: =ROUNDUP(B12+2*E12,0)` | mm | ID (`B12`), WT (`E12`) | `1243` | `calculation_engine.py:48` `roundup(id + 2*wt, 0)` | `1243.0` | **Exact Match** (Zero discrepancy) |
| **Inner Diameter (ID)** | `CS Stacking!C14: 1226` / `Lodha!B12: =D12-2*E12` | mm | Manual or OD, WT | `1226` | `calculation_engine.py:45` `od - 2*wt` | `1226.0` | **Exact Match** (Zero discrepancy) |
| **Nominal Size** | `Synopsis!C12: =ROUND(D12/25.4,0)` | Inch | OD (`D12`) | `49` | `calculation_engine.py` Nominal helper | `49` | **Exact Match** (Display rounding) |
| **Mass per Metre** | `CS Stacking!C17: =((((C15-C16)*C16)*0.0246615)/1000)` | MT/m | OD (`C15`), WT (`C16`) | `0.25575553836` | `calculation_engine.py:34` `mass_mt_per_m()` | `0.25575553836` | **Exact Match** (14 decimal places) |
| **Total Weight** | `CS Stacking!C19: =C18*C17` | MT | MTRS (`C18`), Mass (`C17`) | `511.51107672` | `calculation_engine.py:80` `metres * mass_mt` | `511.51107672` | **Exact Match** |
| **Pipe Pieces** | `CS Stacking!C22: =ROUNDUP(C18/12,0)` | Nos | MTRS (`C18`), Length 12m | `167` | `calculation_engine.py:79` `roundup(metres/len, 0)` | `167` | **Exact Match** |
| **Coating Factor** | `CS Stacking!C20: =3.14*(C15/1000)` | $m^2/m$ | OD (`C15`), $\pi=3.14$ | `3.90302` | `calculation_engine.py:98` `PI_CONSTANT * (dia/1000)` | `3.90302` | **Exact Match** |
| **Ext Area ($m^2$)** | `Synopsis!N12: =ROUNDUP(D12*H12*(22/7)/1000,0)` | $m^2$ | OD, MTRS, $\pi=22/7$ | `7814` | `calculation_engine.py` Ext area formula | `7814` | **Exact Match** |
| **Int Area ($m^2$)** | `Synopsis!O12: =ROUNDUP((22/7)*H12*(D12-2*E12)/1000,0)` | $m^2$ | ID, MTRS, $\pi=22/7$ | `7708` | `calculation_engine.py` Int area formula | `7708` | **Exact Match** |
| **Delivered Steel** | `CS Stacking!C27: =C25+C26` | ₹/MT | Steel (`C25`), Extra (`C26`) | `62000.00` | `calculation_engine.py:326` `steel + grade` | `62000.00` | **Exact Match** |
| **Yield Scrap Loss**| `CS Stacking!C28: =C27*C24` | ₹/MT | Steel (`C27`), Yield% (`C24`) | `1240.00` | `calculation_engine.py:327` `steel * (yield/100)` | `1240.00` | **Exact Match** |
| **Conversion** | `CS Stacking!C29: 4000` | ₹/MT | Input parameter | `4000.00` | `calculation_engine.py:322` `d_conv` | `4000.00` | **Exact Match** |
| **Finance Cost (MT)**| `CS Stacking!C32: 1024` (Manual input) | ₹/MT | Input parameter | `1024.00` | `calculation_engine.py:173` `manual_fc` | `1024.00` | **Exact Match** (Input difference resolved) |
| **Bare Pipe Price** | `CS Stacking!C33: =SUM(C27:C32)` | ₹/MT | Steel, Yield, Conv, FC | `68264.00` | `calculation_engine.py:329` `pre_cost + fc` | `68264.00` | **Exact Match** |
| **Ext Coating (MT)**| `CS Stacking!C34: =($B$34*C20)/C17` | ₹/MT | Rate (`B34=500`), Factor, Mass | `7630.372396` | `calculation_engine.py:100` `coating_cost_per_mt` | `7630.372396` | **Exact Match** |
| **Int Coating (MT)**| `CS Stacking!C35: =($B$35*C20)/C17` | ₹/MT | Rate (`B35=250`), Factor, Mass | `3815.186198` | `calculation_engine.py:123` `approved_mode` uses OD | `3815.186198` | **Exact Match** in approved mode |
| **Ex-Mill Price** | `CS Stacking!C36: =C33+C35+C34` | ₹/MT | Bare, Ext, Int | `79709.558594` | `calculation_engine.py:351` Ex-mill sum | `79709.558594` | **Exact Match** |
| **Stacking Freight** | `CS Stacking!P57: =ROUNDUP(O57*N57/E57,0)+100` | ₹/MT | Trip (93k), Trips, MT, Handling | `7676.00` | `calculation_engine.py:153` `freight_per_mt` | `7676.00` | **Exact Match** |
| **Delivered Subtotal**| `CS Stacking!C38: =C36+C37` | ₹/MT | Ex-Mill, Freight | `87385.558594` | `calculation_engine.py:356` Subtotal | `87385.558594` | **Exact Match** |
| **GST @ 18%** | `CS Stacking!C39: =C38*18%` | ₹/MT | Subtotal $\times 18\%$ | `15729.400547` | `calculation_engine.py` GST calculation | `15729.400547` | **Exact Match** |
| **Grand Total (MT)**| `CS Stacking!C40: =C38+C39` | ₹/MT | Subtotal + GST | `103114.959141`| `calculation_engine.py` Grand total | `103114.959141`| **Exact Match** |
| **FC Check (MT)** | `CS Stacking!C41: =C40*$B$41` | ₹/MT | Grand Total $\times 1\%$ | `1031.149591` | `calculation_engine.py:179` Check formula | `1031.15` | **Exact Match** (Excel check formula) |
| **Bare Price / m** | `CS Stacking!C44: =ROUNDUP(C33*C17,0)` | ₹/m | Bare/MT, Mass/m | `17459.00` | `calculation_engine.py:239` `roundup((pre+fc)*mass, 0)`| `17459.00` | **Exact Match** |
| **Ext Coating / m** | `CS Stacking!C45: =ROUNDUP(C34*C17,0)` | ₹/m | Ext/MT, Mass/m | `1952.00` | `calculation_engine.py:240` `roundup(ext*mass, 0)` | `1952.00` | **Exact Match** |
| **Int Coating / m** | `CS Stacking!C46: =ROUNDUP(C35*C17,0)` | ₹/m | Int/MT, Mass/m | `976.00` | `calculation_engine.py:241` `roundup(int*mass, 0)` | `976.00` | **Exact Match** |
| **Freight / m** | `CS Stacking!C47: =ROUNDUP(C37*C17,0)` | ₹/m | Freight/MT, Mass/m | `1964.00` | `calculation_engine.py:242` `roundup(fr*mass, 0)` | `1964.00` | **Exact Match** |
| **Total Price / m** | `CS Stacking!C48: =SUM(C44:C47)` | ₹/m | Bare/m, Ext/m, Int/m, Freight/m | `22351.00` | `calculation_engine.py:243` `bare + ext + int + fr` | `22351.00` | **Exact Match** (Golden Baseline) |
| **Line Value** | `CS Stacking!C49: =C48*C18` | ₹ | Price/m $\times$ Metres | `44702000.00` | `calculation_engine.py:244` `total * metres` | `44702000.00` | **Exact Match** |

---

## 3. Stacking vs Nesting Freight Reconciliation

### A. Stacking Freight Architecture (Un-nested Standard Dispatch)
1. **Physical Mechanics:**
   - Every trailer carries only a single pipe diameter.
   - For all 4 Unipro sizes, trailer stacking layout is $2 + 2 = 4$ pipes per trailer.
   - Standard 40-ft trailer length accommodates 12-metre pipes.
2. **Payload & Trip Calculations:**
   - Single pipe weight: `I57: =(0.0246615*(C57-D57)*D57*12)/1000` (e.g. 3.069 MT for 1243x8.4).
   - Trailer payload: `L57: =I57*K57` (4 pipes $\times 3.069 = \mathbf{12.276\text{ MT}}$).
   - Unrounded trip count: `N57: =G57/K57` where $G57 = \text{metres} / 12 = 166.667$ pipes.
     $\text{Trips} = 166.6667 / 4 = \mathbf{41.6667\text{ trips}}$.
   - Freight Rate: Jabalpur to Krishna Nagar, Bagodar = **₹93,000 / trip** (`O57`).
3. **Freight Cost Formulation:**
   - Formula: `P57: =ROUNDUP(O57*N57/E57, 0) + 100`
   - Notice: $O57 \times N57 = 93,000 \times 41.6667 = ₹38,75,000$.
   - Divided by total MT ($E57 = 511.511\text{ MT}$) = $7,575.59$ ₹/MT.
   - Rounded UP: $7,576 + 100\text{ handling} = \mathbf{7,676\text{ \rupee/MT}}$.
   - Total trips across all 4 sizes: $41.667 + 7.292 + 10.417 + 91.667 = \mathbf{151.04\text{ trips}}$ (labeled `151 trips` in cell H37).
   - **Key Finding:** Smaller pipes yield lower trailer payloads (12.28 MT vs 14.01 MT), causing smaller pipes to incur higher freight per tonne (₹7,676 vs ₹6,740).

### B. Nesting Freight Architecture (Telescoped Pipes)
1. **Physical Mechanics:**
   - Smaller pipes are inserted inside larger pipes:
     - 1243 mm inserted into 1417 mm and 1367 mm.
     - 1317 mm inserted into 1417 mm.
2. **Trip Optimization:**
   - Stacking requires 151 trips. Nesting reduces this to **103 trips** (a **31.8% reduction in truck movements**).
   - Trips breakdown across 7 combinations:
     - Group 1: 52 of 1417mm + 26 of 1317mm $\rightarrow$ 13 trips.
     - Group 2: 300 of 1417mm + 150 of 1243mm $\rightarrow$ 75 trips.
     - Group 3: 32 of 1367mm + 16 of 1243mm $\rightarrow$ 8 trips.
     - Group 4: 8 of 1367mm unnested $\rightarrow$ 2 trips.
     - Group 5: 12 of 1417mm unnested $\rightarrow$ 3 trips.
     - Group 6: 3 of 1417mm + 3 of 1317mm $\rightarrow$ 1 trip.
     - Group 7: 2 of 1367mm + 1 of 1243mm $\rightarrow$ 1 trip.
     - Total Trips = $13 + 75 + 8 + 2 + 3 + 1 + 1 = \mathbf{103\text{ trips}}$ (`N103`).
3. **Freight Allocation:**
   - Because multiple pipe sizes travel in the same trailer, freight cannot be allocated per diameter individually.
   - Khopoli freight rate = ₹98,000 / trip.
   - Total Project Freight = $103 \times 98,000 = \mathbf{₹1,00,94,000}$ (`Q66`).
   - Blended Freight per MT: $\frac{10,094,000}{2031.036\text{ MT}} + 100 = 4,969.88 + 100 = \mathbf{5,069.88\text{ \rupee/MT}}$ (`Q67`).
   - Entered in Row 37 as **₹5,070 / MT flat** across all 4 items!

---

## 4. Finance Cost (FC) Reconciliation & Data Model

### The Historical Discrepancy (Row 32 vs Row 41):
- In `CS Stacking`, Row 32 contains:
  - Line 1: `C32 = 1024`
  - Line 2: `D32 = 1014`
  - Line 3: `E32 = 1011`
  - Line 4: `F32 = 1008`
- Row 41 contains the verification formula: `=Grand_Total_Incl_GST * $B$41`
  - Line 1: $103,114.96 \times 0.01 = \mathbf{1,031.15}$ (Delta: $+₹7.15$)
  - Line 2: $102,592.96 \times 0.01 = \mathbf{1,025.93}$ (Delta: $+₹11.93$)
  - Line 3: $102,275.91 \times 0.01 = \mathbf{1,022.76}$ (Delta: $+₹11.76$)
  - Line 4: $101,980.32 \times 0.01 = \mathbf{1,019.80}$ (Delta: $+₹11.80$)
- In `CS NESTING`, Row 32 contains `1480` and `1569`, while Row 41 evaluates to `1508.32` and `1598.31` (at 1.5% FC).

### Root Cause Analysis:
Excel does not permit circular formulas without enabling iterative calculation. An exact formula where $\text{FC} = (\text{Delivered} + \text{FC}) \times (1 + \text{GST}) \times r$ creates a circular dependency between Row 32 and Row 40. Estimators resolved this by manually approximating the delivered cost $\times (1 + \text{GST}) \times r$ and entering that number as a static value in Row 32.

### Reconciled Data Model:
To implement the user's explicit directive ("*Allow manual entry of FC (₹/MT) with an auto-calculated suggested value*"), the following three-field model is defined:

```
+-------------------------------------------------------------------------+
| Finance Cost Component Data Model                                       |
+-------------------------------------------------------------------------+
| 1. Suggested_FC (Read-only, Auto-calculated):                           |
|    Suggested_FC = ROUND((PreCost + Ext + Int + Freight) * (1+GST) * r, 0)|
|    -> Line 1 evaluates to: Rs 1,019 / MT (or closed loop Rs 1,031)     |
+-------------------------------------------------------------------------+
| 2. Applied_FC (Active calculation driver):                              |
|    Applied_FC = IF(Manual_FC > 0, Manual_FC, Suggested_FC)              |
|    -> When historical Rs 1,024 is entered, Applied_FC = 1024            |
+-------------------------------------------------------------------------+
| 3. FC_Mode (Audit indicator):                                           |
|    FC_Mode = IF(Manual_FC > 0, "MANUAL_OVERRIDE", "AUTO_SUGGESTED")     |
+-------------------------------------------------------------------------+
```

---

## 5. Rounding Hierarchy & Mathematical Precision

The reconciliation confirms that the live sheets use a strict two-stage rounding hierarchy:

1. **Stage 1 — MT-Level Build-Up (Full Precision Floating-Point):**
   - Mass per metre ($MT/m$): unrounded 10-decimal precision.
   - Raw material steel + yield loss: exact paise precision.
   - Ext coating and Int coating per MT: unrounded floating-point (e.g. ₹7,630.372396/MT).
   - Ex-mill and subtotal per MT: unrounded floating-point (e.g. ₹87,385.558594/MT).
   - Freight per MT: rounded UP to whole rupee (`ROUNDUP(..., 0) + 100`).
2. **Stage 2 — Per-Metre Conversion & Display (Integer Rounding Up):**
   - Bare pipe per metre: `ROUNDUP(Bare_MT * Mass_MT, 0)` $\rightarrow$ integer rupee.
   - Ext coating per metre: `ROUNDUP(Ext_MT * Mass_MT, 0)` $\rightarrow$ integer rupee.
   - Int coating per metre: `ROUNDUP(Int_MT * Mass_MT, 0)` $\rightarrow$ integer rupee.
   - Freight per metre: `ROUNDUP(Freight_MT * Mass_MT, 0)` $\rightarrow$ integer rupee.
   - **Total Unit Price per metre:** `SUM(Bare_m, Ext_m, Int_m, Freight_m)` $\rightarrow$ integer sum!
3. **Stage 3 — Order Totals:**
   - Line Total: `Unit_Price_m * Metres` $\rightarrow$ exact rupee value.
   - Project Subtotal: `SUM(Line_Totals)` $\rightarrow$ exact rupee value.
   - Project GST: `Subtotal * 18%` $\rightarrow$ standard rounding.
   - Grand Total: `Subtotal + GST`.

---

## 6. Audit of Source Sheet Broken References & Incomplete SUMs

### Issue 1: Orphaned Reference in Cost Sheet Cell C11
- **File:** `MKTGENQ26-27-594 - Unipro 2.xlsx`
- **Worksheets:** `CS NESTING10-09-2026` (cell C11) & `CS Stacking 10-09-26` (cell C11)
- **Cell Content:** `[=#REF!] = #REF!`
- **Context:** Row 11 contains `B11: HSAW` and `C11: =#REF!`.
- **Finding:** Cell C11 was originally linked to an enquiry header cell on a previous version of the synopsis sheet that was relocated.
- **Impact on Calculations:** **ZERO.** No formula in rows 12–103 references cell C11.
- **Classification:** **Source Formula Issue (Benign).**

### Issue 2: Incomplete SUM Formulas on Synopsis Sheets
- **File:** `MKTGENQ26-27594 - Unipro.xlsx` & `MKTGENQ26-27-594 - Unipro 2.xlsx`
- **Worksheet:** `Sheet1` / `Synopsis` (Row 16)
- **Observed Formulas:**
  - `H16: =SUM(H11:H15)` = 7,250 m (Correctly covers all 4 lines: rows 12 to 15).
  - `I16: =SUM(I11:I14)` = 239 pipes (Omits Row 15 / Line 4: 367 pipes).
  - `K16: =SUM(K11:K14)` = 747.11 MT (Omits Row 15 / Line 4: 1,283.92 MT).
  - `N16: =SUM(N11:N14)` = 11,412 m² (Omits Row 15 / Line 4: 19,596 m²).
  - `O16: =SUM(O11:O14)` = 11,261 m² (Omits Row 15 / Line 4: 19,363 m²).
- **Intended Behavior & Evidence:**
  - In `CS Stacking 10-09-26`, Row 18 (`G18: =SUM(C18:F18) = 7250`), Row 19 (`G19: =SUM(C19:F19) = 2031.04 MT`), and Row 22 (`G22: =SUM(C22:F22) = 606 Nos`) include all 4 items!
  - In `CS Stacking` RMR rows 68 and 79, total SQM is recorded as `30622.271` and `31005`, which equals the sum of all 4 lines.
- **Classification:** **Source Formula Issue (Incomplete Template Update).** The software engine correctly aggregates all active line items.

---

## 7. Discrepancy Register & Classification

| # | Item / Parameter | Source Value | Software Result | Variance | Classification | Remediation / Conclusion |
|---|---|---|---|---|---|---|
| **D1** | **Unipro Stacking Unit Prices** | ₹22,351; 23,570; 24,396; 25,220 | ₹22,351; 23,570; 24,396; 25,220 | ₹0.00 | **Exact Match** | Baseline verified 100%. |
| **D2** | **Unipro Stacking Project Value**| ₹17,61,17,500.00 | ₹17,61,17,500.00 | ₹0.00 | **Exact Match** | Zero financial delta. |
| **D3** | **Unipro FC Row 32 vs Row 41** | Row 32: 1024; Row 41: 1031.15 | Supports both (1024 applied) | ₹0.00 | **Unresolved Commercial Policy** | Solved by allowing manual entry with auto-suggestion. |
| **D4** | **Unipro Internal Coating Area**| OD used in CS Row 35 | OD used in approved mode | ₹0.00 | **Unresolved Commercial Policy** | Retained in approved mode; Option 1 normal uses ID. |
| **D5** | **Synopsis Row 16 Totals** | K16=747.11 MT (sum rows 11-14) | 2,031.04 MT (sum lines 1-4) | +1283.92 MT | **Source Formula Issue** | Source sheet omitted row 15 from sum; cost sheets included it. |
| **D6** | **CS Cell C11 `#REF!`** | `#REF!` | Handled / Ignored | N/A | **Source Formula Issue** | Orphaned label reference in source template; benign. |
| **D7** | **Selling Price Override** | Manual edit in Offer sheet | `MANUAL_OVERRIDE` pathway | Audit log | **Feature Requirement** | Preserves owner custom price with variance tracking. |
| **D8** | **Stacking Trailer Summary** | Row 61: 151.04 trips (H37: 151) | 153 dedicated trailers | +2 trailers | **Logistics Definition** | Source uses unrounded project pooling; application uses discrete dedicated line dispatch. Zero price/freight impact. |

---

## 8. Required Human Approvals & Safe Next Steps

### Human Approvals Required Prior to Implementation:
1. **Commercial Policy Sign-Off:** Formally ratify the policy allowing manual entry of Finance Cost per MT with an auto-calculated suggested value.
2. **Internal Coating Area Policy:** Confirm whether future standard quotations should price internal coating on OD basis (as done in Unipro 2 CS Row 35) or ID basis (as calculated in Synopsis Row 12).
3. **Location Migration Timing:** Reconfirm that location and territory UI remains in `PILOT_COPY` until commercial testing is concluded, preserving `DEV` unchanged.

### Next Safe Implementation Step:
Proceed to freeze the reconciliation deliverables (`COSTSHEET_FORMULA_RECONCILIATION_MATRIX.md`, `COSTING_INPUTS_AND_AUTOMATION_REGISTER.md`, `COSTSHEET_PARITY_RESULTS.md`) and keep all development isolated in `PILOT_COPY` and Python tests until written sign-off is granted.
