# Mega Pipes — Unipro Pricing Variance Forensic Root-Cause Analysis
**Project:** Mega Pipes Quotation System  
**Document:** `DOCS/UNIPRO_PRICE_VARIANCE_ROOT_CAUSE.md`  
**Reference Tender:** `MKTGENQ26-27594 - Unipro.xlsx` (NIT MKTG/ENQ/26-27/594)  
**Reference CostSheet:** `MKTGENQ26-27-594 - Unipro CS.xlsx`  
**Authoritative Development Workbook:** `DEV/Mega_Pipes_Quotation_DEV.xlsm`  
**Date of Audit:** October 16, 2026  

---

## 1. Executive Summary

This forensic audit investigates the **₹143,150 (+0.08%) difference** between the historical reference CostSheet (`MKTGENQ26-27-594 - Unipro CS.xlsx`) and the automated Normal Mode calculation in `DEV\Mega_Pipes_Quotation_DEV.xlsm`.

| Metric | Reference / Legacy Mode | Normal Mode (Option 1) | Variance | Status |
| :--- | :---: | :---: | :---: | :---: |
| **Line 1 (1226mm ID x 8.4mm WT, 2,000m)** | ₹22,681 / m | ₹22,689 / m | **+₹8 / m** | **+₹16,000** |
| **Line 2 (1300mm ID x 8.4mm WT, 350m)** | ₹23,914 / m | ₹24,109 / m | **+₹195 / m** | **+₹68,250** |
| **Line 3 (1350mm ID x 8.4mm WT, 500m)** | ₹24,748 / m | ₹24,857 / m | **+₹109 / m** | **+₹54,500** |
| **Line 4 (1400mm ID x 8.4mm WT, 4,400m)** | ₹25,580 / m | ₹25,581 / m | **+₹1 / m** | **+₹4,400** |
| **Subtotal Excl. GST (7,250 metres)** | **₹178,657,900** | **₹178,801,050** | — | **+₹143,150** |
| **GST @ 18%** | ₹32,158,422 | ₹32,184,189 | — | **+₹25,767** |
| **Grand Total Incl. GST** | **₹210,816,322** | **₹210,985,239** | — | **+₹168,917** |

### Key Conclusion
The ₹143,150 variance is **NOT a software defect or calculation bug**. It is an **expected, mathematically verified methodology variance** resulting from three intentional commercial business rules implemented in the approved Normal Mode contract:
1. **Whole-Trailer Ceiling Logistics Allocation** (+₹193,850 impact): Real-world commercial transport requires paying for whole dispatched flatbed trailers, replacing the legacy sheet's unviable fractional trailer calculation.
2. **Physical Inner Diameter (ID) for Internal Lining** (-₹108,750 impact): Normal Mode applies internal epoxy coating to the actual inside surface area ($\pi \times ID \times L$), correcting a legacy spreadsheet defect that mistakenly reused the outside diameter ($OD$).
3. **Dynamic Working Capital Financing** (+₹58,050 impact): Normal Mode computes financing via the approved algebraic formula $k \times baseCost / (1 - k)$, replacing unverified static manual entries.
4. **External Coating Parity** (₹0 impact): Identical in both modes.

$$\text{Net Variance} = +\text{₹}193,850 \text{ (Freight)} - \text{₹}108,750 \text{ (Internal Coating)} + \text{₹}58,050 \text{ (Financial)} + \text{₹}0 = +\mathbf{\text{₹}143,150}$$

Every single rupee reconciles across all four lines and all individual cost components.

---

## 2. Component-by-Component Variance Bridge

The following table provides the exhaustive component-level reconciliation for all 4 enquiry lines:

| Pipe ID | Qty (m) | Component | Legacy Input / Value | Normal Input / Value | Unit Variance | Line Impact (₹) | Root Cause | Classification | Action |
| :--- | :---: | :--- | :---: | :---: | :---: | :---: | :--- | :--- | :--- |
| **1226mm ID** | 2,000 | Bare Price (FC) | ₹17,587 / m (FC=₹1542) | ₹17,594 / m (FC=₹1570.11) | +₹7 / m | +₹14,000 | Dynamic $k/(1-k)$ on delivered base cost | Expected Rule | No change |
| | | Ext Coating | ₹1,952 / m | ₹1,952 / m | ₹0 / m | ₹0 | Identical OD coating area formula | Expected Agreement | No change |
| | | Int Coating | ₹1,074 / m ($D=OD=1243$) | ₹1,059 / m ($D=ID=1226$) | -₹15 / m | -₹30,000 | Physical inside diameter (ID) used | Corrected Defect | No change |
| | | Freight | ₹2,068 / m (41.67 trailers) | ₹2,084 / m (42 trailers) | +₹16 / m | +₹32,000 | Whole-trailer ceiling allocation | Expected Rule | No change |
| **Line 1 Total** | **2,000** | **Delivered Price** | **₹22,681 / m** | **₹22,689 / m** | **+₹8 / m** | **+₹16,000** | Net line reconciliation | Expected Rule | **Approved** |
| :--- | :---: | :--- | :---: | :---: | :---: | :---: | :--- | :--- | :--- |
| **1300mm ID** | 350 | Bare Price (FC) | ₹18,639 / m (FC=₹1534) | ₹18,650 / m (FC=₹1574.00) | +₹11 / m | +₹3,850 | Dynamic $k/(1-k)$ on delivered base cost | Expected Rule | No change |
| | | Ext Coating | ₹2,068 / m | ₹2,068 / m | ₹0 / m | ₹0 | Identical OD coating area formula | Expected Agreement | No change |
| | | Int Coating | ₹1,138 / m ($D=OD=1317$) | ₹1,123 / m ($D=ID=1300$) | -₹15 / m | -₹5,250 | Physical inside diameter (ID) used | Corrected Defect | No change |
| | | Freight | ₹2,069 / m (7.29 trailers) | ₹2,268 / m (8 trailers) | +₹199 / m | +₹69,650 | Whole-trailer ceiling (8 vs 7.29) | Expected Rule | No change |
| **Line 2 Total** | **350** | **Delivered Price** | **₹23,914 / m** | **₹24,109 / m** | **+₹195 / m** | **+₹68,250** | Net line reconciliation | Expected Rule | **Approved** |
| :--- | :---: | :--- | :---: | :---: | :---: | :---: | :--- | :--- | :--- |
| **1350mm ID** | 500 | Bare Price (FC) | ₹19,349 / m (FC=₹1529) | ₹19,359 / m (FC=₹1563.12) | +₹10 / m | +₹5,000 | Dynamic $k/(1-k)$ on delivered base cost | Expected Rule | No change |
| | | Ext Coating | ₹2,147 / m | ₹2,147 / m | ₹0 / m | ₹0 | Identical OD coating area formula | Expected Agreement | No change |
| | | Int Coating | ₹1,181 / m ($D=OD=1367$) | ₹1,166 / m ($D=ID=1350$) | -₹15 / m | -₹7,500 | Physical inside diameter (ID) used | Corrected Defect | No change |
| | | Freight | ₹2,071 / m (10.42 trailers) | ₹2,185 / m (11 trailers) | +₹114 / m | +₹57,000 | Whole-trailer ceiling (11 vs 10.42) | Expected Rule | No change |
| **Line 3 Total** | **500** | **Delivered Price** | **₹24,748 / m** | **₹24,857 / m** | **+₹109 / m** | **+₹54,500** | Net line reconciliation | Expected Rule | **Approved** |
| :--- | :---: | :--- | :---: | :---: | :---: | :---: | :--- | :--- | :--- |
| **1400mm ID** | 4,400 | Bare Price (FC) | ₹20,060 / m (FC=₹1525) | ₹20,068 / m (FC=₹1551.61) | +₹8 / m | +₹35,200 | Dynamic $k/(1-k)$ on delivered base cost | Expected Rule | No change |
| | | Ext Coating | ₹2,225 / m | ₹2,225 / m | ₹0 / m | ₹0 | Identical OD coating area formula | Expected Agreement | No change |
| | | Int Coating | ₹1,224 / m ($D=OD=1417$) | ₹1,209 / m ($D=ID=1400$) | -₹15 / m | -₹66,000 | Physical inside diameter (ID) used | Corrected Defect | No change |
| | | Freight | ₹2,071 / m (91.67 trailers) | ₹2,079 / m (92 trailers) | +₹8 / m | +₹35,200 | Whole-trailer ceiling (92 vs 91.67) | Expected Rule | No change |
| **Line 4 Total** | **4,400** | **Delivered Price** | **₹25,580 / m** | **₹25,581 / m** | **+₹1 / m** | **+₹4,400** | Net line reconciliation | Expected Rule | **Approved** |
| **GRAND TOTAL**| **7,250** | **Subtotal Excl. GST**| **₹178,657,900** | **₹178,801,050** | — | **+₹143,150** | **Exact Cumulative Reconciliation** | **Expected Rule** | **Approved** |

---

## 3. Deep Technical Root-Cause Investigation

### 3.1 Root Cause 1: Freight Whole-Trailer Ceiling vs Fractional Allocation (+₹193,850)
- **Legacy Implementation:**  
  The historical CostSheet calculated required trailers as:
  $$\text{Trailers}_{\text{legacy}} = \frac{\text{Metres} / \text{PipeLen}}{\text{PipesPerTrailer}} = \frac{\text{Metres} / 12}{4}$$
  For Line 2 ($350\text{ m} = 29.17 \rightarrow 30\text{ pipes}$), this evaluated to $7.2917$ fractional trailers.
  Legacy freight was computed as:
  $$\text{FreightPerMT}_{\text{legacy}} = \left\lceil \frac{98000 \times 7.2917}{94.880\text{ MT}} \right\rceil + 100 = \text{₹}7,632/\text{MT} \implies \text{₹}2,069/\text{m}$$
- **Normal Mode Implementation:**  
  Real-world heavy transport contracts require booking integer trailers. Dispatching 30 pipes at 4 pipes/trailer requires commissioning **8 full flatbed trailers** (7 full trailers carrying 28 pipes + 1 partial trailer carrying 2 pipes).
  $$\text{Trailers}_{\text{normal}} = \left\lceil \frac{30}{4} \right\rceil = 8\text{ trailers}$$
  $$\text{FreightPerMT}_{\text{normal}} = \left\lceil \frac{(125.32 \times 782) \times 8}{94.880\text{ MT}} \right\rceil + 100 = \text{₹}8,364/\text{MT} \implies \text{₹}2,268/\text{m}$$
- **Commercial Consequence:**  
  For low-volume orders (e.g. Line 2 with 350 m), paying for the 8th trailer creates a +₹199/m (+9.6%) increase in freight cost. For high-volume orders (Line 4 with 4,400 m), the extra partial trailer is amortized over 1,283.9 MT, resulting in only +₹8/m (+0.38%).
- **Verdict:** Authoritative commercial business rule. Mega Pipes cannot absorb unbilled partial-trailer freight charges.

### 3.2 Root Cause 2: Internal Coating Diameter Correction (-₹108,750)
- **Legacy Implementation:**  
  In the reference Excel spreadsheet (`CS Stacking 10-09-26`), the internal coating row copied the formula from the external coating row:
  $$\text{Area}_{\text{internal, legacy}} = \pi \times \frac{OD}{1000} \times L$$
  For Line 1 ($ID=1226\text{ mm}$, $OD=1243\text{ mm}$), it used $OD=1243\text{ mm}$ instead of $ID=1226\text{ mm}$.
  $$\text{IntCostPerM}_{\text{legacy}} = \left\lceil 275 \times 3.14 \times 1.243 \right\rceil = \text{₹}1,074/\text{m}$$
- **Normal Mode Implementation:**  
  Internal epoxy lining is applied strictly to the inside surface of the pipe ($\pi \times ID \times L$):
  $$\text{IntCostPerM}_{\text{normal}} = \left\lceil 275 \times 3.14 \times 1.226 \right\rceil = \text{₹}1,059/\text{m}$$
- **Commercial Consequence:**  
  Normal Mode charges customers for the actual physical surface area coated, saving ₹15/m across all lines:
  $$\Delta \text{IntPerM} = \left\lceil 275 \times 3.14 \times \frac{OD - ID}{1000} \right\rceil = \lceil 14.68 \rceil = \text{₹}15/\text{m}$$
  Total savings passed to customer: $7,250\text{ m} \times \text{₹}15/\text{m} = \mathbf{\text{₹}108,750}$.
- **Verdict:** Corrected engineering and commercial defect. The legacy sheet erroneously billed customers for external surface area on internal lining.

### 3.3 Root Cause 3: Dynamic Algebraic Working Capital Financing (+₹58,050)
- **Legacy Implementation:**  
  The reference sheet had hardcoded, static numbers entered into the Financial Cost column:
  - Line 1: ₹1,542/MT
  - Line 2: ₹1,534/MT
  - Line 3: ₹1,529/MT
  - Line 4: ₹1,525/MT
  These numbers were manually typed overrides from an earlier offline calculation.
- **Normal Mode Implementation:**  
  Normal Mode implements the approved dynamic algebraic formula reflecting working capital financing ($1.5\%$) plus statutory GST ($18\%$):
  $$k = \left(\frac{1.5}{100}\right) \times \left(1 + \frac{18}{100}\right) = 0.0177$$
  $$\text{FinancialPerMT} = \frac{k \times \text{baseCost}}{1 - k} = \frac{0.0177 \times (PreCost + Ext + Int + Freight)}{0.9823}$$
  Because Normal Mode baseCost reflects the higher (whole-trailer) freight, financial cost per MT is dynamically higher:
  - Line 1: ₹1,570.11/MT (+₹28.11/MT $\implies$ +₹7/m)
  - Line 2: ₹1,574.00/MT (+₹40.00/MT $\implies$ +₹11/m)
  - Line 3: ₹1,563.12/MT (+₹34.12/MT $\implies$ +₹10/m)
  - Line 4: ₹1,551.61/MT (+₹26.61/MT $\implies$ +₹8/m)
  Total monetary impact: $\mathbf{+\text{₹}58,050}$.
- **Verdict:** Authoritative financial rule. Dynamic financing ensures finance charges scale accurately with logistics and delivery costs.

---

## 4. Policy Decision & Architectural Action Plan

According to the decision framework established in Section 7 of the specification:

1. **Retain Normal Mode Golden Baseline (₹178,801,050):**  
   Normal Mode is the approved commercial contract for current and future tenders. The ₹143,150 difference must **never** be "fixed" by reverting to fractional trailers, billing customers for phantom internal coating area, or hardcoding manual financial costs.
2. **Preserve Legacy Mode Reproduction (₹178,657,900):**  
   When `LegacyMode=YES` is explicitly toggled, the engine executes the legacy path and reproduces the exact historical reference figure with 0 difference.
3. **Automated Reconciliation Gate:**  
   Future quotation imports run through `tools/pricing_reconciliation_runner.py` which automatically decomposes variances and ensures that both modes remain bit-for-bit compliant with their respective golden baselines.
