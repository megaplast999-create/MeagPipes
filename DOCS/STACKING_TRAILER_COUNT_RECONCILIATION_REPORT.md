# MEGA PIPES — STACKING TRAILER COUNT RECONCILIATION REPORT

**Document Reference:** `DOCS/STACKING_TRAILER_COUNT_RECONCILIATION_REPORT.md`  
**Date:** 2026-10-22  
**Status:** Complete & Forensically Verified  
**Primary Author:** Senior Costing Systems Auditor & QA Automation Engineer  
**Source Workbooks Verified:**  
1. `C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx` (Sheet: `CS Stacking 10-09-26`)  
2. `c:\Users\Administrator\Desktop\Mega Pipes\Quotation System\DEV\Mega_Pipes_Quotation_DEV.xlsm`  
**Backup Image Created:**  
`c:\Users\Administrator\Desktop\Mega Pipes\Quotation System\BACKUP\Mega_Pipes_Quotation_DEV_pre_trailer_reconciliation_20261022_112700.xlsm`  
(SHA-256: `D4AF713B2425A582CA84C308656CCDF7F0D85B37B363FBE0A79F19B1534D62E5`)

---

## 1. Executive Summary

This investigation was commissioned to reconcile an apparent discrepancy between:
- The **151 trips** reported in the original Unipro Stacking CostSheet (`MKTGENQ26-27-594 - Unipro 2.xlsx!CS Stacking 10-09-26`), and
- The **153 total trailers** displayed on the `QUOTE` console (`E7`) in `DEV\Mega_Pipes_Quotation_DEV.xlsm`.

### Key Findings:
1. **Mathematical Invariance of Commercial Prices:**
   The trailer count discrepancy is **purely an aggregation and logistics definition difference** in summary reporting. It has **ZERO impact on freight cost per MT**, **ZERO impact on line item unit prices**, and **ZERO impact on total quotation value**.
2. **Unit Price Parity Confirmed to the Rupee:**
   The approved per-metre selling prices remain **100% identical** to the source CostSheet and Offer sheet:
   - **Line 1 (1226 ID / 1243 OD):** ₹22,351 / metre
   - **Line 2 (1300 ID / 1317 OD):** ₹23,570 / metre
   - **Line 3 (1350 ID / 1367 OD):** ₹24,396 / metre
   - **Line 4 (1400 ID / 1417 OD):** ₹25,220 / metre
   - **Total Quotation Subtotal (Excl. GST):** ₹17,61,17,500.00
3. **Forensic Root Cause:**
   - **Source "151 trips" (Continuous Project Pooling):** The source Excel sheet models logistics as a continuous flow where pipes per line are unrounded floating-point numbers ($2000/12 = 166.6667$, etc.) and trips per line are unrounded fractions ($41.6667 + 7.2917 + 10.4167 + 91.6667 = \mathbf{151.0417\text{ trips}}$). The project total is then rounded to **151 trips** in cell `H37`.
   - **Application "153 trailers" (Dedicated Discrete Loading):** The application engine models physical dispatch where each pipe size is manufactured and dispatched independently in full 12m pipe joints and dedicated 40ft HGV trailers without mixed-size consolidation. Each line's pipe count is rounded UP to the nearest integer ($\lceil \text{metres}/12 \rceil$), and each line's trailer count is rounded UP to the nearest whole trailer ($\lceil \text{pipes}/4 \rceil$):
     $$42 + 8 + 11 + 92 = \mathbf{153\text{ dedicated trailers}}$$
4. **Intermediate Consolidation ("152 trailers"):**
   If whole manufactured pipes are pooled across sizes at the project level before assigning trailers, total project pipes equal $167 + 30 + 42 + 367 = 606$ pipes. Dividing by 4 gives $606 / 4 = 151.5 \rightarrow \lceil 151.5 \rceil = \mathbf{152\text{ trailers}}$.

---

## 2. Line-by-Line Reconciliation Table

The table below reconciles all four line items between the original source workbook and the Mega Pipes DEV application:

| Parameter | Line 1 | Line 2 | Line 3 | Line 4 | Project Total / Summary |
|---|---|---|---|---|---|
| **Pipe Size (Nominal / Basis)** | 49" (1226 mm ID) | 52" (1300 mm ID) | 54" (1350 mm ID) | 56" (1400 mm ID) | 4 Sizes Added |
| **Calculated OD (mm)** | 1243.0 mm | 1317.0 mm | 1367.0 mm | 1417.0 mm | — |
| **Wall Thickness (mm)** | 8.4 mm | 8.4 mm | 8.4 mm | 8.4 mm | — |
| **Ordered Quantity (Metres)** | 2,000 m | 350 m | 500 m | 4,400 m | **7,250 m** |
| **Linear Mass (MT / m)** | 0.255755538 | 0.271085127 | 0.281442957 | 0.291800787 | Weighted: 0.28014287 |
| **Total Weight (MT)** | 511.511 MT | 94.880 MT | 140.721 MT | 1,283.923 MT | **2,031.036 MT** |
| **Single Pipe Weight (MT / 12m)** | 3.069 MT | 3.253 MT | 3.377 MT | 3.502 MT | — |
| **Trailer Stacking Arrangement** | 2 + 2 | 2 + 2 | 2 + 2 | 2 + 2 | 4 pipes per trailer |
| **Pipes per Trailer (K)** | 4 | 4 | 4 | 4 | 4 |
| **Trailer Payload (MT / Trailer)** | 12.276 MT | 13.012 MT | 13.509 MT | 14.006 MT | Avg: 13.447 MT |
| **Source Pipe Count (`G57:G60`)** | 166.6667 (raw) | 29.1667 (raw) | 41.6667 (raw) | 366.6667 (raw) | **604.1667 pipes** |
| **DEV Pipe Count (`result.Pipes`)**| **167 pcs** | **30 pcs** | **42 pcs** | **367 pcs** | **606 pcs** |
| **Source Trip Count (`N57:N60`)** | 41.6667 trips | 7.2917 trips | 10.4167 trips | 91.6667 trips | **151.0417 trips** |
| **Source Displayed Trips (`H37`)** | — | — | — | — | **151 trips** (`ROUND(151.04, 0)`) |
| **DEV Line Trailers (`result.Trailers`)**| **42 trailers** | **8 trailers** | **11 trailers** | **92 trailers** | **153 trailers** (`QUOTE!E7`) |
| **Full Trailers + Remaining Pipes**| 41 full + 3 pcs | 7 full + 2 pcs | 10 full + 2 pcs | 91 full + 3 pcs | 149 full + 4 partial |
| **Spare Trailer Capacity (pcs)** | 1 pipe spare | 2 pipes spare | 2 pipes spare | 1 pipe spare | 6 pipes spare (1.5 trailers) |
| **Source Freight / MT (`P57:P60`)**| ₹7,676 / MT | ₹7,248 / MT | ₹6,985 / MT | ₹6,740 / MT | Weighted: ₹6,937.05 / MT |
| **DEV Freight / MT (`freight`)** | ₹7,676 / MT | ₹7,248 / MT | ₹6,985 / MT | ₹6,740 / MT | Weighted: ₹6,937.05 / MT |
| **Freight Variance (₹ / MT)** | **₹0.00** | **₹0.00** | **₹0.00** | **₹0.00** | **₹0.00 (Exact Parity)** |
| **Approved Unit Price (₹ / m)** | **₹22,351 / m** | **₹23,570 / m** | **₹24,396 / m** | **₹25,220 / m** | Weighted: ₹24,292.07 / m |
| **DEV Unit Price (₹ / m)** | **₹22,351 / m** | **₹23,570 / m** | **₹24,396 / m** | **₹25,220 / m** | Weighted: ₹24,292.07 / m |
| **Unit Price Variance (₹ / m)** | **₹0.00** | **₹0.00** | **₹0.00** | **₹0.00** | **₹0.00 (Exact Parity)** |
| **Line Total Value (INR)** | ₹4,47,02,000 | ₹82,49,500 | ₹1,21,98,000 | ₹11,09,68,000 | **₹17,61,17,500.00** |

---

## 3. Source Formula Tracing (`CS Stacking 10-09-26`)

In `MKTGENQ26-27-594 - Unipro 2.xlsx`, worksheet `CS Stacking 10-09-26`:

1. **Stacking Plan Section (Rows 54 to 61):**
   - **Line 1 (Row 57):**
     - Cell `C57`: `1243` (Pipe Dia MM)
     - Cell `D57`: `8.4` (Wall Thickness MM)
     - Cell `E57`: `=(0.0246615*(C57-D57)*D57*F57)/1000` $\rightarrow 511.511\text{ MT}$
     - Cell `F57`: `2000` (Metres)
     - Cell `G57`: `=F57/12` $\rightarrow 166.6666667\text{ pipes}$ (unrounded)
     - Cell `I57`: `=(0.0246615*(C57-D57)*D57*12)/1000` $\rightarrow 3.069066\text{ MT / pipe}$
     - Cell `J57`: `2+2` (text description of 4 pipes stacked on trailer)
     - Cell `K57`: `4` (Pipes on Trailer)
     - Cell `L57`: `=I57*K57` $\rightarrow 12.276266\text{ MT / trailer payload}$
     - Cell `N57`: `=G57/K57` $\rightarrow 41.6666667\text{ trips}$ (unrounded)
     - Cell `O57`: `93000` (Freight rate per trip in INR)
     - Cell `P57`: `=ROUNDUP(O57*N57/E57, 0) + 100` $\rightarrow \mathbf{7,676\text{ \rupee/MT}}$
   - **Line 2 (Row 58):**
     - Cell `F58`: `350`, Cell `G58`: `=F58/12 = 29.1666667`, Cell `K58`: `4`
     - Cell `N58`: `=G58/K58 = 7.2916667 trips`, Cell `P58`: $\mathbf{7,248\text{ \rupee/MT}}$
   - **Line 3 (Row 59):**
     - Cell `F59`: `500`, Cell `G59`: `=F59/12 = 41.6666667`, Cell `K59`: `4`
     - Cell `N59`: `=G59/K59 = 10.4166667 trips`, Cell `P59`: $\mathbf{6,985\text{ \rupee/MT}}$
   - **Line 4 (Row 60):**
     - Cell `F60`: `4400`, Cell `G60`: `=F60/12 = 366.6666667`, Cell `K60`: `4`
     - Cell `N60`: `=G60/K60 = 91.6666667 trips`, Cell `P60`: $\mathbf{6,740\text{ \rupee/MT}}$
   - **Total Row (Row 61):**
     - Cell `F61`: `=SUM(F57:F60) = 7250`
     - Cell `G61`: `=SUM(G57:G60) = 604.1666667`
     - Cell `N61`: `=SUM(N57:N60) = \mathbf{151.0416667\text{ trips}}`
2. **Cost Breakdown Section (Row 37):**
   - Cell `A37`: `Freight : Bagodar (Stacking)`
   - Cells `C37:F37`: `=P57`, `=P58`, `=P59`, `=P60` (7676, 7248, 6985, 6740)
   - Cell `G37`: `6937.05` (Weighted average freight per MT)
   - Cell `H37`: Text label `"151 trips"` (derived from `ROUND(151.0417, 0)`).

---

## 4. Application Formula Tracing (`Mega_Pipes_Quotation_DEV.xlsm`)

In `DEV\Mega_Pipes_Quotation_DEV.xlsm`:

1. **Pipe and Trailer Calculation per Line (`VBA\Modules\modPricing.bas:308-318`):**
   ```vba
   result.Pipes = RoundUpD(result.Metres / CDec(lineData.PipeLen), 0)
   result.Trailers = RoundUpD(result.Pipes / CDec(lineData.PerTrailer), 0)
   ```
   - **Line 1 (2,000 m):** $\lceil 2000 / 12 \rceil = 167\text{ pipes}$. $\lceil 167 / 4 \rceil = \mathbf{42\text{ trailers}}$.
   - **Line 2 (350 m):** $\lceil 350 / 12 \rceil = 30\text{ pipes}$. $\lceil 30 / 4 \rceil = \mathbf{8\text{ trailers}}$.
   - **Line 3 (500 m):** $\lceil 500 / 12 \rceil = 42\text{ pipes}$. $\lceil 42 / 4 \rceil = \mathbf{11\text{ trailers}}$.
   - **Line 4 (4,400 m):** $\lceil 4400 / 12 \rceil = 367\text{ pipes}$. $\lceil 367 / 4 \rceil = \mathbf{92\text{ trailers}}$.
2. **Summary Aggregation (`VBA\Modules\modQuote.bas:279-280` & `340-341`):**
   ```vba
   summary.TotalPipeCount = summary.TotalPipeCount + CLng(result.Pipes)
   summary.TotalTrailers = summary.TotalTrailers + CLng(result.Trailers)
   ...
   ws.Range("D7").Value = "Total Trailers"
   ws.Range("E7").Value = summary.TotalTrailers
   ```
   - `summary.TotalPipeCount` = $167 + 30 + 42 + 367 = \mathbf{606\text{ pipes}}$ (`QUOTE!E6`)
   - `summary.TotalTrailers` = $42 + 8 + 11 + 92 = \mathbf{153\text{ trailers}}$ (`QUOTE!E7`)
3. **Freight Rate Calculation per Line (`VBA\Modules\modPricing.bas:420-429`):**
   ```vba
   freight = RoundUpD(CDec(rates.TripRate) * _
                      (CDec(result.Metres) / CDec(lineData.PipeLen)) / _
                      CDec(lineData.PerTrailer) / _
                      result.MT, 0) + CDec(rates.Handling)
   ```
   Notice that DEV's approved costing formula evaluates using the **unrounded fractional trips**:
   $$\frac{\text{TripRate} \times \frac{\text{Metres} / 12}{4}}{\text{MT}} + 100$$
   This is 100% mathematically identical to source cell `P57: =ROUNDUP(O57*N57/E57, 0) + 100`!

---

## 5. Mathematical Proof: Invariance of Freight / MT and Unit Prices

Why does the 151 vs 153 difference have zero effect on freight rates and selling prices?

Let:
- $L = 12\text{ m}$ (standard pipe joint length)
- $K = 4$ (pipes per trailer)
- $R = ₹93,000$ (freight rate per trip)
- $m = \text{ordered metres}$
- $w = \text{linear mass in MT per metre}$
- $M = m \times w$ (total MT)

In both the source spreadsheet and DEV:
$$\text{Trips}_{\text{calc}} = \frac{m / L}{K} = \frac{m}{L \times K} = \frac{m}{48}$$
$$\text{Freight}_{\text{pre-handling}} = \frac{R \times \text{Trips}_{\text{calc}}}{M} = \frac{R \times \frac{m}{48}}{m \times w} = \frac{R}{48 \times w} = \frac{R}{\text{Trailer Payload MT}}$$
Because $m$ (metres) cancels out in the numerator and denominator, **the freight rate per tonne depends exclusively on the physical trailer payload weight**, not on whether 151 or 153 is displayed in the project summary header!

Evaluating for all 4 lines:
- **Line 1:** Payload = $12.276266\text{ MT} \implies 93,000 / 12.276266 = 7575.60 \implies \lceil 7575.60 \rceil + 100 = \mathbf{7,676\text{ \rupee/MT}}$
- **Line 2:** Payload = $13.012086\text{ MT} \implies 93,000 / 13.012086 = 7147.20 \implies \lceil 7147.20 \rceil + 100 = \mathbf{7,248\text{ \rupee/MT}}$
- **Line 3:** Payload = $13.509262\text{ MT} \implies 93,000 / 13.509262 = 6884.17 \implies \lceil 6884.17 \rceil + 100 = \mathbf{6,985\text{ \rupee/MT}}$
- **Line 4:** Payload = $14.006438\text{ MT} \implies 93,000 / 14.006438 = 6639.80 \implies \lceil 6639.80 \rceil + 100 = \mathbf{6,740\text{ \rupee/MT}}$

Because the freight per MT is identical, all downstream per-metre conversions:
$$\text{Bare/m} + \text{Ext/m} + \text{Int/m} + \text{Freight/m} = \text{Total/m}$$
evaluate to:
- Line 1: $17459 + 1952 + 976 + 1964 = \mathbf{₹22,351 / m}$
- Line 2: $18503 + 2068 + 1034 + 1965 = \mathbf{₹23,570 / m}$
- Line 3: $19209 + 2147 + 1074 + 1966 = \mathbf{₹24,396 / m}$
- Line 4: $19915 + 2225 + 1113 + 1967 = \mathbf{₹25,220 / m}$
- Total Project Value = $\mathbf{₹17,61,17,500.00}$

---

## 6. Definitions and Logistics Assumptions

The investigation confirms that the two numbers answer two different operational questions:

```
+-----------------------------------------------------------------------------------------+
| LOGISTICS AGGREGATION METHODOLOGY COMPARISON                                            |
+-----------------------------------------------------------------------------------------+
| Dimension                   | Source Workbook (151 Trips) | Application (153 Trailers)  |
+-----------------------------+-----------------------------+-----------------------------+
| Operational Concept         | Continuous Project Flow     | Dedicated Line Dispatches   |
| Pipe Joint Quantization     | Fractional (604.17 pipes)   | Integer Ceil (606 pipes)    |
| Multi-Size Load Pooling     | Implicitly Pooled           | Independent Dispatches      |
| Trailer Ceiling Scope       | Project Level               | Per-Line Item Level         |
| Mathematical Expression     | round( sum(m) / 48 )        | sum( ceil( ceil(m/12) / 4)) |
| Practical Physical Meaning  | Total transport work units  | Physical trucks booked      |
+-----------------------------------------------------------------------------------------+
```

### Breakdown of the +2 Delta:
1. **Pipe Ceiling Delta (+0.46 trailers):**
   Rounding each line's length to integer 12m joints creates 606 physical pipes ($167 + 30 + 42 + 367$) vs $604.17$ fractional pipes ($+1.83$ pipes $\approx +0.46$ trailers).
2. **Line Ceiling Delta (+1.50 trailers):**
   In dedicated loading, partial trailers cannot take pipes of another size. 
   - Line 1 has 3 pipes in its 42nd trailer (1 spare slot).
   - Line 2 has 2 pipes in its 8th trailer (2 spare slots).
   - Line 3 has 2 pipes in its 11th trailer (2 spare slots).
   - Line 4 has 3 pipes in its 92nd trailer (1 spare slot).
   Together, these 4 partially-filled trailers carry $3 + 2 + 2 + 3 = 10$ pipes.
   If cross-size loading was permitted, 10 pipes would occupy $\lceil 10 / 4 \rceil = 3$ trailers, saving 1 trailer ($153 \rightarrow 152$).
   If fractional loading is used, $10 / 4 = 2.5$ trailers ($153 \rightarrow 151$).

---

## 7. Automated Regression Test Suite

A comprehensive regression test has been added to `TESTS/test_costsheet_source_parity.py`:
- **Test Name:** `test_unipro_stacking_trailer_reconciliation_source_vs_engine`
- **Assertions:**
  1. Verifies source `CS Stacking 10-09-26` row 61 has 151.0417 continuous trips and cell `H37` displays `"151 trips"`.
  2. Verifies application produces discrete per-line loading: 42, 8, 11, 92 $\rightarrow$ sum = 153 dedicated trailers.
  3. Verifies project-level integer consolidation yields $\lceil 606 / 4 \rceil = 152$ trailers.
  4. Verifies project continuous pooling yields $\text{round}(7250 / 48) = 151$ trips.
  5. Verifies all 4 freight rates (₹7,676, ₹7,248, ₹6,985, ₹6,740) and unit prices (₹22,351, ₹23,570, ₹24,396, ₹25,220) match with zero variance.

### Test Execution Result:
```
============================ 168 passed in 58.98s =============================
```
Zero failures, zero errors across the entire test suite.

---

## 8. Root Cause, Affected Values, and Smallest Safe Correction

### Exact Root Cause:
The difference stems entirely from **aggregation scope**:
- Source sums unrounded fractional trips across lines and rounds the total ($151.04 \rightarrow 151$).
- Application computes physical dispatch requirements per line by taking the ceiling of each independent size item ($42 + 8 + 11 + 92 = 153$).

### Affected Values:
- **Affected:** Cell `QUOTE!E7` (Summary KPI display of physical trailer dispatches) and internal accumulator `summary.TotalTrailers`.
- **Unaffected:** Freight per MT, Bare pipe price, Ext coating price, Int coating price, Finance cost, Line item unit prices (₹22,351, ₹23,570, ₹24,396, ₹25,220), Line totals, and Grand Total (₹17,61,17,500).

### Smallest Safe Correction:
1. **Maintain Master Workbook Integrity:**
   Do **NOT** alter the calculation engine or VBA logic merely to hardcode or force 151 in `QUOTE!E7`. Displaying 153 dedicated trailers is logistically accurate for un-nested, discrete size dispatch (avoiding under-booking trucks in logistics operations).
2. **UI Clarification (Recommended for Phase 6):**
   In the next scheduled UI release, add an explanatory tooltip or secondary KPI metric:
   - `Total Trailers: 153 (Dedicated Dispatch)`
   - `Continuous Pooled Trips: 151 (Theoretical Project Floor)`
3. **Formal Register Update:**
   Record this reconciliation in `COSTSHEET_FORMULA_RECONCILIATION_MATRIX.md` as Item **D8 (Logistics Aggregation Definition)**.
