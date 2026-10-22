# Mega Pipes Quotation System — Dynamic Staged Costing Engine Specification

## Status & Authority
- **Document Version:** 1.0.0
- **Authority:** Approved Company CostSheet (`MKTGENQ26-27-594 - Unipro 2.xlsx`, Sheet `CS Stacking 10-09-26`)
- **Status:** Approved Baseline & Engineering Mapping for Phases 1–6
- **Date:** 17-Oct-2026

---

## 1. Executive Summary & Objective

The Mega Pipes Quotation System calculates pipe selling prices stage-by-stage from enquiry-specific commercial inputs. Pipe selling prices depend on actual steel purchase rates, yield loss, bare-pipe manufacturing conversion costs, coating specifications, road freight and stacking rules, finance charges, and profit margin.

This specification details all 12 costing stages:
1. Input fields and units
2. Source of default values
3. Mathematical formulas and business rules
4. Downstream dependencies
5. Output fields
6. Existing test coverage
7. Precedence rules for enquiry-specific inputs and manual selling-price overrides

---

## 2. Complete Costing Stage Mapping (Stages 1 to 12)

### Stage 1: Pipe Specifications, Geometry & Dimension Validation
- **Input Fields:**
  - `Basis` (String): Diameter basis (`"OD"` or `"ID"`).
  - `SizeMM` (Double): Nominal outer or inner diameter in millimetres.
  - `WT` (Double): Wall thickness in millimetres.
  - `PipeLen` (Double): Nominal pipe joint length (default 12.0m; range 6.0m to 13.0m).
- **Units:** mm, mm, mm, metres.
- **Source of Defaults:** Form inputs on `QUOTE` sheet (`B22:B24`, `B30`) or `tblSizes` / `tblThickness`.
- **Formulas & Business Rules:**
  - If `Basis = "OD"`:
    $$\text{OD} = \text{SizeMM}, \quad \text{ID} = \text{OD} - 2 \times \text{WT}$$
  - If `Basis = "ID"`:
    $$\text{ID} = \text{SizeMM}, \quad \text{OD} = \lceil \text{ID} + 2 \times \text{WT} \rceil = \operatorname{RoundUp}(\text{ID} + 2 \times \text{WT}, 0)$$
  - Physical constraints:
    $$\text{OD} > 0, \quad \text{WT} > 0, \quad \text{ID} > 2 \times \text{WT}, \quad 6.0 \le \text{PipeLen} \le 13.0$$
- **Downstream Dependencies:** Linear mass, surface areas, pre-cost, freight stacking, per-metre pricing.
- **Output Fields:** `CalculatedODMM`, `CalculatedIDMM`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_geometry_*`, `TESTS/test_lodha_five_line_review.py::test_lodha_od_derivation_rounding_up`.

---

### Stage 2: Mass, Weight, Quantities & Coating Surface Areas
- **Input Fields:** $\text{OD}$, $\text{ID}$, $\text{WT}$, $\text{Quantity}$, $\text{QuantityUnit}$, $\text{PipeLen}$.
- **Units:** mm, mm, mm, Metres/MT/Nos, Text, metres.
- **Source of Defaults:** Stage 1 outputs, form inputs on `QUOTE` sheet (`B28:B29`).
- **Formulas & Business Rules:**
  - Linear Mass (kg/m):
    $$\text{MassKgPerM} = (\text{OD} - \text{WT}) \times \text{WT} \times 0.0246615$$
  - Linear Mass (MT/m):
    $$\text{MassMtPerM} = \frac{\text{MassKgPerM}}{1000}$$
  - Quantity Conversions:
    - If `QuantityUnit = "Metres"`: $\text{Metres} = \text{Quantity}$
    - If `QuantityUnit = "MT"`: $\text{Metres} = \frac{\text{Quantity}}{\text{MassMtPerM}}$
    - If `QuantityUnit = "Nos"`: $\text{Metres} = \text{Quantity} \times \text{PipeLen}$
  - Total Weight (MT):
    $$\text{WeightMT} = \text{Metres} \times \text{MassMtPerM}$$
  - Whole Dispatch Pipes:
    $$\text{DispatchPipes} = \lceil \text{Metres} / \text{PipeLen} \rceil$$
  - Fractional Pipe Equivalents (for Synopsis):
    $$\text{FractionalPipes} = \text{Metres} / \text{PipeLen}$$
  - External Surface Area Factor ($m^2/m$):
    $$\text{CoatingFactor} = 3.14 \times \left(\frac{\text{OD}}{1000}\right)$$
  - Total Coating Surface Areas ($m^2$):
    $$\text{ExtArea } (m^2) = \lceil \text{OD} \times \text{Metres} \times (22/7) / 1000 \rceil$$
    $$\text{IntArea } (m^2) = \lceil (22/7) \times \text{Metres} \times \text{ID} / 1000 \rceil$$
- **Downstream Dependencies:** Coating cost per MT, freight payload allocation, line total calculation.
- **Output Fields:** `MassKgPerM`, `Metres`, `WeightMT`, `PipeCount`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_mass_mt_per_m_formula`, `TESTS/test_synopsis_precision.py`.

---

### Stage 3: Raw Material Requirements & Yield Loss
- **Input Fields:**
  - `SteelRate` (Double): Base steel purchase rate (Rs/MT).
  - `GradeExtra` (Double): Material/Grade premium (Rs/MT).
  - `YieldPct` (Double): Yield loss percentage (default 2.0%).
- **Units:** Rs/MT, Rs/MT, %.
- **Source of Defaults:** `tblSettings.DefaultSteelRate` (₹62,000), `tblSettings.DefaultGradeExtra` (₹0), `tblSettings.DefaultYieldPct` (2.0%). Overridable per enquiry/line.
- **Formulas & Business Rules:**
  - Delivered Steel Rate (Rs/MT):
    $$\text{DeliveredSteel} = \text{SteelRate} + \text{GradeExtra}$$
  - Yield Loss Amount (Rs/MT):
    $$\text{YieldLoss} = \text{DeliveredSteel} \times \left(\frac{\text{YieldPct}}{100}\right)$$
  - Raw Material Cost (Rs/MT):
    $$\text{RawMaterialCost} = \text{DeliveredSteel} + \text{YieldLoss} = \text{DeliveredSteel} \times \left(1 + \frac{\text{YieldPct}}{100}\right)$$
- **Downstream Dependencies:** Pre-cost / Bare pipe cost, Ex-mill price.
- **Output Fields:** `DeliveredSteel`, `YieldLoss`, `RawMaterialCost`.
- **Existing Test Coverage:** `TESTS/test_pricing_golden_baselines.py::test_approved_pre_cost_breakdown`.

---

### Stage 4: Bare-Pipe Manufacturing & Conversion Cost
- **Input Fields:**
  - `RawMaterialCost` (Double): From Stage 3 (Rs/MT).
  - `ConversionRate` (Double): Pipe manufacturing conversion cost (Rs/MT).
  - `Testing` (Double): Approved testing & hospitality cost (Rs/MT).
  - `PreCostMargin` (Double): Manufacturing margin (Rs/MT).
- **Units:** Rs/MT.
- **Source of Defaults:** `tblSettings.DefaultConversion` (₹4,000), `tblSettings.DefaultTesting` (₹0), `tblSettings.DefaultMargin` (₹0). Overridable per enquiry.
- **Formulas & Business Rules:**
  - Pre-Cost / Bare Pipe Manufacturing Cost:
    $$\text{PreCostPerMT} = \text{RawMaterialCost} + \text{ConversionRate} + \text{Testing} + \text{PreCostMargin}$$
  - Approved Reference Standard:
    $$\text{PreCostPerMT} = (62,000 + 0) \times 1.02 + 4,000 + 0 + 0 = \text{₹}67,240\text{/MT}$$
- **Downstream Dependencies:** Bare pipe price per metre, financial cost base, ex-mill price.
- **Output Fields:** `PreCostPerMT`.
- **Existing Test Coverage:** `TESTS/test_pricing_golden_baselines.py`, `TESTS/test_p4_line_agreement.py`.

---

### Stage 5: External Coating Cost
- **Input Fields:**
  - `HasExt` (Boolean): External coating required flag.
  - `ExtCoatingRate` (Double): External coating application rate ($Rs/m^2$).
  - $\text{OD}$ (Double): Outer diameter in mm.
  - $\text{MassMtPerM}$ (Double): Pipe mass in MT/metre.
  - $\text{PI}$ (Double): Circumference constant = 3.14.
- **Units:** Boolean, $Rs/m^2$, mm, MT/m.
- **Source of Defaults:** `tblCoatings` (`EXT-PU-1000` = ₹500/m²), `tblSettings.PIPct` (3.14). Overridable per enquiry.
- **Formulas & Business Rules:**
  - If `HasExt = True`:
    $$\text{ExtCostPerMT} = \frac{\text{ExtCoatingRate} \times \text{PI} \times (\text{OD} / 1000)}{\text{MassMtPerM}}$$
  - If `HasExt = False`:
    $$\text{ExtCostPerMT} = 0$$
- **Downstream Dependencies:** Ex-mill price, financial cost base, external coating price per metre (`ExtPerM`).
- **Output Fields:** `ExtCostPerMT`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_external_coating_cost_per_mt`.

---

### Stage 6: Internal Coating Cost
- **Input Fields:**
  - `HasInt` (Boolean): Internal coating required flag.
  - `IntCoatingRate` (Double): Internal coating application rate ($Rs/m^2$).
  - $\text{OD}$, $\text{ID}$ (Double): Diameters in mm.
  - $\text{MassMtPerM}$ (Double): Pipe mass in MT/metre.
  - $\text{PI}$ (Double): Circumference constant = 3.14.
  - `Mode` (String): `"NORMAL"` (Approved CostSheet standard) vs `"OPTION1"` (Historical).
- **Units:** Boolean, $Rs/m^2$, mm, MT/m.
- **Source of Defaults:** `tblCoatings` (`INT-FGE-406` = ₹250/m²), `tblSettings.PIPct` (3.14). Overridable per enquiry.
- **Formulas & Business Rules:**
  - Approved Company Normal Standard:
    $$\text{IntCostPerMT} = \frac{\text{IntCoatingRate} \times \text{PI} \times (\text{OD} / 1000)}{\text{MassMtPerM}}$$
    *(Note: Per approved CostSheet `CS Stacking 10-09-26`, internal coating is commercially calculated on the OD surface factor)*.
  - Option 1 Historical Contract:
    $$\text{IntCostPerMT} = \frac{\text{IntCoatingRate} \times \text{PI} \times (\text{ID} / 1000)}{\text{MassMtPerM}}$$
- **Downstream Dependencies:** Ex-mill price, financial cost base, internal coating price per metre (`IntPerM`).
- **Output Fields:** `IntCostPerMT`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_internal_coating_cost_per_mt`.

---

### Stage 7: Other Approved Finishing / Manufacturing Costs
- **Input Fields:** Special bevelling, third-party inspection, hydrostatic proof testing, hospitality.
- **Units:** Rs/MT.
- **Source of Defaults:** Default = 0 unless specified in enquiry commercial terms.
- **Formulas & Business Rules:** Additively incorporated into Pre-Cost or Ex-mill without double counting.
- **Downstream Dependencies:** Ex-mill price.
- **Output Fields:** Finishing cost subcomponents.

---

### Stage 8: Freight and Handling Costs
- **Input Fields:**
  - `FreightMethod`: Method A (Manual Trip Rate), Method B (Rate/km), Method C (Slabs).
  - `TripRate` (Double): Trip rate per trailer (Rs/trip).
  - `RatePerKm` (Double): Rate per km (Rs/km).
  - `DistanceKm` (Double): Road distance (km).
  - `PipesPerTrailer` (Double): Pipes per trailer (from `tblStacking` lookup).
  - `Handling` (Double): Handling charge (Rs/MT).
  - $\text{Metres}$, $\text{PipeLen}$, $\text{WeightMT}$.
- **Units:** Rs/trip, km, pipes, Rs/MT.
- **Source of Defaults:** `tblSettings.TripRateManual` (₹93,000), `tblSettings.HandlingRsPerMT` (₹100), `tblStacking`. Overridable per enquiry.
- **Formulas & Business Rules:**
  - Approved CostSheet Standard (Fractional Trailer Stacking Allocation):
    $$\text{FractionalTrailers} = \frac{\text{Metres} / \text{PipeLen}}{\text{PipesPerTrailer}}$$
    $$\text{FreightPerMT} = \left\lceil \frac{\text{TripRate} \times \text{FractionalTrailers}}{\text{WeightMT}} \right\rceil + \text{Handling}$$
  - Option 1 Historical Standard:
    $$\text{Trailers} = \left\lceil \frac{\lceil \text{Metres} / \text{PipeLen} \rceil}{\text{PipesPerTrailer}} \right\rceil$$
    $$\text{FreightPerMT} = \left\lceil \frac{\text{TripRate} \times \text{Trailers}}{\text{WeightMT}} \right\rceil + \text{Handling}$$
- **Downstream Dependencies:** Landed cost, financial cost base, freight price per metre (`FreightPerM`).
- **Output Fields:** `Trailers`, `FreightPerMT`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_freight_per_mt_approved_and_legacy`.

---

### Stage 9: Financial Cost
- **Input Fields:**
  - `PreCostPerMT`, `ExtCostPerMT`, `IntCostPerMT`, `FreightPerMT`.
  - `FinPct` (Double): Finance cost percentage (default 1.0%).
  - `GSTPct` (Double): GST percentage (default 18.0%).
  - `ManualFC` (Double): Line-level manual financial cost override (Rs/MT).
- **Units:** %, %, Rs/MT.
- **Source of Defaults:** `tblSettings.DefaultFinancialPct` (1.0%), `tblSettings.GSTPct` (18.0%). Overridable per enquiry/line (`tblQuoteLines.ManualFC`).
- **Formulas & Business Rules:**
  - Base Cost:
    $$\text{BaseCost} = \text{PreCostPerMT} + \text{ExtCostPerMT} + \text{IntCostPerMT} + \text{FreightPerMT}$$
  - If `ManualFC > 100`:
    $$\text{FCperMT} = \text{ManualFC} \quad (\text{e.g. 1024, 1014, 1011, 1008 in Unipro reference})$$
  - Else (Approved Formulaic):
    $$\text{FCperMT} = \operatorname{Round}\left(\text{BaseCost} \times \left(1 + \frac{\text{GSTPct}}{100}\right) \times \frac{\text{FinPct}}{100}, 0\right)$$
  - Option 1 Historical Contract:
    $$k = \left(\frac{\text{FinPct}}{100}\right) \times \left(1 + \frac{\text{GSTPct}}{100}\right), \quad \text{FCperMT} = \frac{k \times \text{BaseCost}}{1 - k}$$
- **Downstream Dependencies:** Bare pipe price per metre, ex-mill price per MT, total selling price.
- **Output Fields:** `FCperMT`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_financial_cost_*`, `TESTS/test_pricing_golden_baselines.py`.

---

### Stage 10: Margin & Commercial Markup
- **Input Fields:**
  - `MarginPct` / `MarginPerMT`: Commercial profit margin.
- **Units:** % or Rs/MT.
- **Source of Defaults:** `tblSettings.DefaultMargin` (0), Customer/Project pricing profile.
- **Formulas & Business Rules:** Additively included in pre-cost or ex-mill selling price. Floor protection enforces $\text{SellingPrice} \ge \text{ExMillFloor}$.
- **Downstream Dependencies:** Calculated selling price.
- **Output Fields:** Margin amount.

---

### Stage 11: Selling Price Derivation & Component Rounding Rules
- **Input Fields:**
  - `PreCostPerMT`, `FCperMT`, `ExtCostPerMT`, `IntCostPerMT`, `FreightPerMT`, `MassMtPerM`, `PriceBasis`.
  - `ManualSellingPrice` (Double): Optional manual price override (Rs/m or Rs/MT).
- **Units:** Rs/m, Rs/MT.
- **Source of Defaults:** `PriceBasis` from quotation header (default `"Metres"`).
- **Formulas & Business Rules:**
  - Component-wise per-metre prices (each rounded UP to integer rupee):
    $$\text{BarePerM} = \lceil (\text{PreCostPerMT} + \text{FCperMT}) \times \text{MassMtPerM} \rceil$$
    $$\text{ExtPerM} = \lceil \text{ExtCostPerMT} \times \text{MassMtPerM} \rceil$$
    $$\text{IntPerM} = \lceil \text{IntCostPerMT} \times \text{MassMtPerM} \rceil$$
    $$\text{FreightPerM} = \lceil \text{FreightPerMT} \times \text{MassMtPerM} \rceil$$
    $$\text{CalculatedTotalPerM} = \text{BarePerM} + \text{ExtPerM} + \text{IntPerM} + \text{FreightPerM}$$
  - Calculated Total per MT:
    $$\text{CalculatedTotalPerMT} = \lceil \text{PreCostPerMT} + \text{FCperMT} \rceil + \lceil \text{ExtCostPerMT} \rceil + \lceil \text{IntCostPerMT} \rceil + \text{FreightPerMT}$$
  - Applied Unit Price & Price Source Tag:
    - If `ManualSellingPrice > 0`:
      $$\text{AppliedUnitPrice} = \text{ManualSellingPrice}, \quad \text{PriceSource} = \text{"MANUAL\_OVERRIDE"}$$
      $$\text{PriceVariance} = \text{AppliedUnitPrice} - \text{CalculatedUnitPrice}$$
    - Else:
      $$\text{AppliedUnitPrice} = \text{CalculatedUnitPrice}, \quad \text{PriceSource} = \text{"CALCULATED"}$$
      $$\text{PriceVariance} = 0$$
- **Downstream Dependencies:** Line total value, quotation subtotal, GST, grand total.
- **Output Fields:** `BarePerM`, `ExtPerM`, `IntPerM`, `FreightPerM`, `TotalPerM`, `TotalPerMT`, `UnitPrice`, `CalculatedPrice`, `PriceSource`, `PriceVariance`.
- **Existing Test Coverage:** `TESTS/test_calculation_engine.py::test_selling_price_per_m`, `TESTS/test_pricing_golden_baselines.py`.

---

### Stage 12: Commercial Line Value & Quotation Totals
- **Input Fields:** `AppliedUnitPrice`, `Metres`, `WeightMT`, `PriceBasis`, `GSTPct` (18.0%).
- **Units:** Rupees.
- **Source of Defaults:** `GSTPct = 18.0%` from `tblSettings.GSTPct`.
- **Formulas & Business Rules:**
  - Line Total (Rupees):
    - If `PriceBasis = "MT"`:
      $$\text{LineTotal} = \text{AppliedUnitPrice} \times \text{WeightMT}$$
    - Else (`PriceBasis = "Metres"`):
      $$\text{LineTotal} = \text{AppliedUnitPrice} \times \text{Metres}$$
  - Quotation Subtotal (Excl. GST):
    $$\text{Subtotal} = \sum_{i=1}^{N} \text{LineTotal}_i$$
  - GST Amount:
    $$\text{GSTAmount} = \text{Subtotal} \times \frac{\text{GSTPct}}{100} = \text{Subtotal} \times 0.18$$
  - Grand Total (Incl. GST):
    $$\text{GrandTotal} = \text{Subtotal} + \text{GSTAmount}$$
- **Downstream Dependencies:** Live Summary Card, `OFFER` sheet, `COSTING` sheet, exports, immutable snapshots.
- **Output Fields:** `LineTotal`, `Subtotal`, `GSTAmount`, `GrandTotal`.
- **Existing Test Coverage:** `TESTS/test_pricing_golden_baselines.py::test_approved_quotation_totals` (exact ₹176,117,500; ₹31,701,150; ₹207,818,650).

---

## 3. Commercial Rate Precedence & Isolation Rules

To support enquiry-specific commercial inputs without mutating shared profiles:
1. **Inheritance Order:**
   $$\text{Line Override} \longrightarrow \text{Enquiry Header} \longrightarrow \text{Customer Profile} \longrightarrow \text{System Default (\texttt{tblSettings})}$$
2. **Immutability Guarantee:**
   - Overriding a rate for Enquiry `MKTG/ENQ/26-27/594` creates a local override in `tblQuoteLines` and the enquiry snapshot.
   - It **must never modify** `tblSettings` or `MASTERS` tables.
3. **Price Source Tagging:**
   - `CALCULATED`: Price derived 100% from staged calculation engine.
   - `MANUAL_OVERRIDE`: Price manually entered; calculated price and variance preserved.

---

## 4. Unresolved Business-Rule Questions Register

1. **Internal Coating Surface Area Basis:**
   - The approved Unipro CostSheet (`CS Stacking 10-09-26`) computes internal coating cost per MT using the OD surface factor ($OD / 1000 \times 3.14$), whereas Option 1 used the ID ($ID / 1000 \times 3.14$).
   - *Current Implementation Rule:* Normal/Approved mode follows the approved Unipro CostSheet (OD basis); Option 1 mode preserves ID basis for regression.
2. **Financial Cost Base Iteration:**
   - In row 32 of `CS Stacking 10-09-26`, the values 1024, 1014, 1011, 1008 reflect 1.0% financing with 25% advance / balance LC payment terms. In row 41, the formula is strictly `=C40 * 1%` where C40 includes GST and FC itself.
   - *Current Implementation Rule:* Support both formulaic calculation $\operatorname{Round}(\text{BaseCost} \times 1.18 \times 0.01, 0)$ and explicit line overrides in `ManualFC`.
3. **Manual Price Override Authorization:**
   - Does a manual price override below calculated cost require Admin password (`MEGA_ADMIN_PWD`)?
   - *Current Implementation Rule:* Follow existing approval workflow: if price is below floor or overridden, flag `PriceSource = "MANUAL_OVERRIDE"`, record reason, and require Review/Approval state before issuing offer.
