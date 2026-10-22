# Mega Pipes — Approved Quotation Calculation Audit & Implementation Inventory
**Document:** `DOCS/APPROVED_QUOTATION_CALCULATION_AUDIT.md`  
**Phase:** Phase 1 — Comprehensive Implementation Audit  
**Authoritative Reference File:** `C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro CS.xlsx` (`CS Stacking 10-09-26` & `Sheet1`)  
**Development Workbook:** `DEV\Mega_Pipes_Quotation_DEV.xlsm`  
**Audit Date:** October 16, 2026  
**Environment Boundary:** Strict DEV only. LIVE strictly untouched.

---

## 1. Executive Summary & Audit Mandate

The primary objective of the Mega Pipes Quotation System is to automate the previously approved commercial quotation process with exact parity to the approved Excel methodology, formulas, fixed values, commercial rules, freight allocation, and rounding behavior.

This audit establishes the comprehensive formula mapping between:
1. The approved reference Excel sheets (`MKTGENQ26-27-594 - Unipro CS.xlsx`).
2. The Excel VBA calculation engine (`modPricing.bas`, `modFreight.bas`, `modQuote.bas`, `modQuoteLines.bas`).
3. The pure-Python reference calculation engine (`TESTS/calculation_engine.py`).

### Approved Commercial Benchmark Target
For the approved Unipro reference tender (7,250 metres across 4 line items):

| Line Item | Size Specification | Commercial Quantity | Approved Reference Rate | Approved Line Total |
| :---: | :--- | :---: | :---: | :---: |
| **Line 1** | 1226 mm ID (1243 mm OD) × 8.4 mm WT | 2,000 m | **₹22,681 / m** | **₹45,362,000** |
| **Line 2** | 1300 mm ID (1317 mm OD) × 8.4 mm WT | 350 m | **₹23,914 / m** | **₹8,369,900** |
| **Line 3** | 1350 mm ID (1367 mm OD) × 8.4 mm WT | 500 m | **₹24,748 / m** | **₹12,374,000** |
| **Line 4** | 1400 mm ID (1417 mm OD) × 8.4 mm WT | 4,400 m | **₹25,580 / m** | **₹112,552,000** |
| **Subtotal** | **Total Order Metres (Excl. GST)** | **7,250 m** | — | **₹178,657,900** |
| **GST** | **Statutory GST @ 18%** | — | — | **₹32,158,422** |
| **Grand Total** | **Delivered Order Total (Incl. GST)** | — | — | **₹210,816,322** |

---

## 2. End-to-End Calculation Flow Mapping

The calculation path across the approved sheets, VBA modules, and Python engine is:
$$\text{User Input} \rightarrow \text{Validation} \rightarrow \text{Geometry} \rightarrow \text{Mass} \rightarrow \text{Manufacturing Cost} \rightarrow \text{Coating} \rightarrow \text{Freight} \rightarrow \text{Financial Cost} \rightarrow \text{Component Rounding} \rightarrow \text{Selling Price} \rightarrow \text{Line Total} \rightarrow \text{Subtotal} \rightarrow \text{GST} \rightarrow \text{Grand Total}$$

### Detailed Component Mapping Table

| Component | Approved Sheet Cell / Formula | Equivalent VBA Procedure (`modPricing.bas`) | Equivalent Python Function (`calculation_engine.py`) | Precision & Rounding Rules | Relevant Test & Expected Result | Confirmed Mismatch & Status |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Pipe OD Derivation (from ID)** | `Synopsis` / `CS Stacking` Row 15: Nominal OD integer (e.g. $1226 + 16.8 \rightarrow 1243$) | `modGeometry.CalculateOD(ID, WT)`: `RoundUpD(ID + 2 * WT, 0)` | `geometry(size, wt, 'ID')`: `roundup(id + 2*wt, 0)` | Integer millimetres (`ROUNDUP(..., 0)`) | `test_geometry_id_rounds_od_up`: OD = 1243, 1317, 1367, 1417 | **PASS** — Complete Parity across Sheet, VBA, and Python. |
| **Pipe Mass (MT/m)** | `CS Stacking` Row 17: `=((((OD-WT)*WT)*0.0246615)/1000)` | `modGeometry.CalculateMassKgPerM`: `(OD - WT) * WT * 0.0246615`, `massMtPerM = kg/1000` | `mass_mt_per_m(od, wt)`: `((OD - WT) * WT * 0.0246615) / 1000` | Full 11-decimal floating/decimal precision (no rounding at mass stage) | `test_mass_mt_per_m`: Line 1 = 0.25575553836 MT/m | **PASS** — Complete Parity. |
| **Commercial Quantity** | `CS Stacking` Row 18: Exact enquiry metres ($2000, 350, 500, 4400$) | `modPricing.CalcLine`: preserves `lineData.Qty` | `quantity_conversion`: preserves `metres` | Exact decimal commercial metres; logistics pipe counts do not alter billable length | `test_quantity_conversion_preserves_exact_metres` | **PASS** — Complete Parity. |
| **Pipe Count (Joints)** | `CS Stacking` Row 22: `=ROUNDUP(Qty/12, 0)` | `modGeometry.PipesForMetres`: `RoundUpD(Qty / PipeLen, 0)` | `quantity_conversion`: `roundup(metres / pipe_len, 0)` | Nearest integer whole pipe joints (12m nominal) | Line 1 = 167, Line 2 = 30, Line 3 = 42, Line 4 = 367 | **PASS** — Complete Parity. |
| **Base Steel & Yield** | `CS Stacking` Row 25 & 28: `=SteelCost * Yield%` ($61000 \times 2\% = 1220$) | `modPricing.CalcLine`: `pre = Steel + Yield% * Steel` | `calc_line`: `pre_d = steel + yield * steel` | Exact currency decimal (₹62,220/MT) | PreCost = ₹67,220/MT | **PASS** — Complete Parity. |
| **Conversion Charge** | `CS Stacking` Row 29: Fixed value ₹5,000/MT | `rates.Conversion`: ₹5,000/MT | `conversion`: ₹5,000.0/MT | Fixed approved rate | Total PreCost = ₹67,220/MT | **PASS** — Complete Parity. |
| **External Coating Cost** | `CS Stacking` Row 34: `=($B$34*C20)/C17` where C20=`3.14*(OD/1000)` | `ext = CDec(rates.ExtRate) * CDec(rates.PI) * od / 1000 / massMtPerM` | `external_coating_per_mt`: `(Rate * 3.14 * (OD/1000)) / MassMT` | Rate: ₹500/$\text{m}^2$, $\pi=3.14$, cost in ₹/MT | Line 1: ₹7,630.37/MT | **PASS** — Complete Parity. |
| **Internal Coating Cost** | `CS Stacking` Row 35: `=($B$35*C20)/C17` where C20=`3.14*(OD/1000)` | `If legacyMode Then intc = IntRate * PI * od / 1000 / massMtPerM Else intc = IntRate * PI * idd / 1000 / massMtPerM` | `internal_coating_per_mt`: uses `od` in legacy, `inner_d` in normal | Rate: ₹275/$\text{m}^2$, $\pi=3.14$, cost in ₹/MT | Approved Reference: Line 1 = ₹4,196.70/MT (uses OD) | **CRITICAL MISMATCH:** In Normal Mode (`LegacyMode=No`), automated system uses $ID$, reducing price by ₹15/m. Approved sheet specifies $OD$. |
| **Trailer Allocation** | `CS Stacking` Row 57: `N57 = G57/K57 = (F57/12)/4` (Fractional: 41.67, 7.29, 10.42, 91.67) | `If legacyMode Then freight = RoundUp(Trip * (Metres/12)/4 / MT, 0) Else freight = RoundUp(Trip * Trailers / MT, 0)` | `freight_per_mt`: Fractional trailers in legacy, whole trailers in normal | Fractional trailers: $\frac{\text{Metres} / 12}{4}$ | Approved Reference: Line 1 = ₹8,083, Line 2 = ₹7,632, Line 3 = ₹7,355, Line 4 = ₹7,097 | **CRITICAL MISMATCH:** In Normal Mode (`LegacyMode=No`), automated system enforces integer ceiling trailers, raising freight by up to ₹199/m on Line 2. |
| **Financial Cost** | `CS Stacking` Row 32 & 41: `=C40 * 1.5%` with manual values in Row 32: 1542, 1534, 1529, 1525 | `If legacyMode Then financial = ManualFC Else financial = k * baseCost / (1 - k)` | `financial_cost`: `manual_fc` in legacy, algebraic in normal | Manual override in Row 32; Base formula: $\text{ROUND}(\text{BaseCost} \times 1.18 \times 1.5\%, 0)$ | Approved Reference: ₹1,542, ₹1,534, ₹1,529, ₹1,525/MT | **CRITICAL MISMATCH:** In Normal Mode (`LegacyMode=No`), automated system uses dynamic factor on whole-trailer base cost, raising financial cost by ₹7 to ₹11/m. |
| **Selling Price / Metre** | `CS Stacking` Row 44–48: `ROUNDUP(Bare*Mass,0) + ROUNDUP(Ext*Mass,0) + ROUNDUP(Int*Mass,0) + ROUNDUP(Freight*Mass,0)` | `result.TotalPerM = BarePerM + ExtPerM + IntPerM + FreightPerM` | `selling_price_per_m`: sum of 4 ceiling components | Ceiling to nearest integer rupee (`ROUNDUP(..., 0)`) per component | Line 1: ₹22,681; Line 2: ₹23,914; Line 3: ₹24,748; Line 4: ₹25,580 | **PARITY CONFIRMED WHEN APPROVED RULES APPLIED.** |
| **Line Subtotal & GST** | `CS Stacking` Row 49 & `Sheet1`: `=TotalPerM * Qty`, `GST = Subtotal * 18%` | `result.LineValue = TotalPerM * Metres`, `GSTAmount = Subtotal * 0.18` | `calc_quote`: `sum(line_value)`, `gst = subtotal * 0.18` | Exact integer currency sum | Subtotal: ₹178,657,900; GST: ₹32,158,422; Grand Total: ₹210,816,322 | **PASS** — Exact parity achieved under approved methodology. |

---

## 3. Analysis of Confirmed Mismatches

The audit confirms three specific formula mismatches when the automated quotation engine executes with `LegacyMode = No` (the current default):

1. **Freight Methodology Mismatch:**
   - *Approved Sheet:* Fractional trailer allocation $\text{Trailers} = (\text{Metres} / 12) / 4$.
   - *Current Default in Software:* Discrete whole-trailer ceiling $\text{Trailers} = \lceil \text{Pipes} / 4 \rceil$.
   - *Impact:* Overcharges Line 2 by +₹199/m (+₹69,650), Line 3 by +₹114/m (+₹57,000), Line 1 by +₹16/m (+₹32,000), Line 4 by +₹8/m (+₹35,200). Net freight inflation: **+₹193,850**.
2. **Internal Coating Methodology Mismatch:**
   - *Approved Sheet:* Internal coating surface area evaluated using Outside Diameter ($OD$).
   - *Current Default in Software:* Evaluated using Inside Diameter ($ID$).
   - *Impact:* Deflates internal coating price by -₹15/m across all lines (-₹108,750).
3. **Financial Cost Methodology Mismatch:**
   - *Approved Sheet:* Static manual inputs (1542, 1534, 1529, 1525) based on $\text{ROUND}(\text{BaseCostWithoutFC} \times 1.18 \times 1.5\%, 0)$.
   - *Current Default in Software:* Dynamic algebraic factor $k / (1 - k)$ on whole-trailer base cost.
   - *Impact:* Inflates bare pipe selling price by +₹7 to +₹11/m across all lines (+₹58,050).

### Root Cause of the Discrepancy
In `tblSettings`, the configuration parameter `LegacyMode` is currently set to `'No'`. When `LegacyMode = 'No'`, the system branches to the alternative "Normal Mode (Option 1)" which applies the three experimental changes above, generating a +₹143,150 difference. When the approved methodology is activated, the engine produces the **exact approved figures** with 0 rupee difference.

---

## 4. Minimal Implementation Plan

To satisfy the primary project objective ("Make the automated quotation system calculate new quotations exactly according to the formulas, fixed values, commercial rules, freight methodology, and rounding behaviour of the previously approved Excel sheets"):

1. **Set Approved Methodology as the Primary Default:**
   Configure `tblSettings.LegacyMode = 'Yes'` (or promote the approved methodology to the default execution path in `modPricing.bas`, `modFreight.bas`, `modQuote.bas`, `modQuoteLines.bas`, and `TESTS/calculation_engine.py`).
2. **Preserve Fractional Trailer Freight:**
   Ensure the default freight calculation uses the approved fractional trailer allocation formula:
   $$\text{FreightPerMT} = \left\lceil \frac{\text{TripRate} \times (\text{Metres} / 12) / 4}{\text{TotalMT}} \right\rceil + \text{Handling}$$
   Keep whole-trailer count available as an informational logistics metric on the quotation console without altering the commercial freight price.
3. **Preserve Approved Coating Area & Financial Cost Formulas:**
   Ensure internal coating and financial cost use the approved reference methodology by default, while supporting manual commercial overrides.
4. **Quotation Workflow & In-App Approval (Phases 4–6):**
   Verify and strengthen `modApproval.bas` and `modUI.bas` to allow creating, editing, recalculating, reviewing, and approving new quotations inside the software, locking the snapshot upon approval.
5. **Permanent Regression Protection (Phase 7):**
   Strengthen the test suite to ensure that creating new quotations with approved inputs strictly asserts ₹22,681, ₹23,914, ₹24,748, ₹25,580 and ₹178,657,900.
