# Mega Pipes Quotation System — Normal-Mode Golden Calculation Audit

**Date:** 2026-10-12  
**Author:** Senior VBA Engineer & Financial Calculation Auditor  
**Target Workbook:** `DEV\Mega_Pipes_Quotation_DEV.xlsm`  
**Deliverable Location:** `TESTS\ANTIGRAVITY_NORMAL_MODE_CALCULATION_AUDIT.md`  

---

## 1. Executive Summary & Objective

The objective of this audit is to diagnose and independently audit the normal-mode calculation discrepancy in the Mega Pipes Quotation System:

- **Authoritative Required Normal-Mode Total:** ₹178,805,450  
- **Current Normal-Mode Total:** ₹154,499,250  
- **Total Mismatch:** Deficit of ₹24,306,200  
- **Authoritative Legacy Total:** ₹178,657,900 (Currently passing with 0 difference)  
- **Authoritative Line Selling Prices (per metre):**
  - Line 1 (2,000 m): ₹22,689
  - Line 2 (350 m): ₹24,109
  - Line 3 (500 m): ₹24,857
  - Line 4 (4,400 m): ₹25,582

### Key Findings

1. **Primary Defect (Unit Mismatch in Normal-Mode Coating):**  
   In [modPricing.bas](file:///C:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modPricing.bas#L344-L350) and lines 373–379, normal-mode external and internal coating formulas divide by `massKgPerM` (kg/m) instead of `massMtPerM` (MT/m). This produces coating rates in **₹/kg** instead of **₹/MT**, deflating coating costs by a factor of 1,000.
2. **Double Compounding in Per-Metre Selling Price:**  
   In [modPricing.bas](file:///C:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modPricing.bas#L489-L499), `result.ExtPerM = RoundUpD(ext * massMtPerM, 0)`. Because `ext` is already in ₹/kg (1/1,000 of ₹/MT), multiplying by `massMtPerM` ($kg/m \div 1000$) divides by 1,000 a second time. This collapses `ExtPerM` and `IntPerM` from ~₹1,952 and ~₹1,059 down to **₹2/m** and **₹2/m**.
3. **Financial Cost Collateral Damage:**  
   In normal mode, `baseCost = pre + ext + intc + freight`. Because `ext` and `intc` are deflated by 1,000, `baseCost` receives only ₹11.77/MT instead of ~₹11,770/MT of coating cost. This deflates the GST-adjusted financial cost per MT from ~₹1,570/MT down to ~₹1,358/MT, reducing `BarePerM` by ₹54 to ₹62 per metre across all lines.
4. **Freight Rate Verification:**  
   Normal-mode freight is verified to receive the correct trip rate ($125.32 \times 782 = \text{₹}98,000.24$) and produces exact freight values (Line 1: ₹2,084, Line 2: ₹2,268, Line 3: ₹2,185, Line 4: ₹2,079). Freight is **not** defaulting to zero in the calculation engine.
5. **Exact Decomposition of the ₹24,306,200 Difference:**  
   - Line 1 (2,000 m): Deficit of ₹3,061/m = ₹6,122,000
   - Line 2 (350 m): Deficit of ₹3,244/m = ₹1,135,400
   - Line 3 (500 m): Deficit of ₹3,368/m = ₹1,684,000
   - Line 4 (4,400 m): Deficit of ₹3,492/m = ₹15,364,800
   - **Sum across 4 lines:** Exactly ₹24,306,200.
6. **Line 4 Target Analysis (₹25,581 vs ₹25,582):**  
   Correcting the coating unit defect using configured `rates.PI = 3.14` produces:
   - Line 1: ₹22,689 (Exact match)
   - Line 2: ₹24,109 (Exact match)
   - Line 3: ₹24,857 (Exact match)
   - Line 4: ₹25,581 (1 rupee lower than target ₹25,582; Line value ₹112,556,400 vs ₹112,560,800)
   - Reconstructed Grand Total: ₹178,801,050 (Difference: ₹4,400 vs ₹178,805,450).
   The audit reveals a subtle sensitivity where Line 4 internal coating crosses the rounding threshold from ₹1,209 to ₹1,210 if $\pi \ge 3.14026$, which perfectly resolves the ₹4,400 difference without changing any other line.

---

## 2. Codebase & Configuration Verification

### 2.1 File & Module Inventory

- `MASTER_PROMPT_MANUS.md`: Confirmed **not present** in the project repository.
- `VBA\Modules\modPricing.bas`: Primary pricing and line calculation module (`CalcLine`, `LoadDefaultRateInput`, `P3GoldenTotal`, `P3GoldenLine`).
- `VBA\Modules\modSelfTest.bas`: Diagnostic test harness (`RunP3GoldenTests`, `WriteGoldenLine`).
- `VBA\Modules\modGeometry.bas`: Pipe geometry functions (`CalculateOD`, `CalculateID`, `CalculateMassKgPerM`).
- `VBA\Modules\modMath.bas`: Rounding and mass conversion functions (`RoundUpD`, `massMtPerM`).
- `VBA\Modules\modCalculation.bas`: Helper math formulas (`CalculateCoatingCostPerMT`, `CalculateFinancialCost`).
- `TESTS\calculation_engine.py`: Pure-Python reference engine.

### 2.2 Table Settings Verification (`tblSettings`)

Inspected from `DEV\Mega_Pipes_Quotation_DEV.xlsm` (Sheet `SETTINGS`, Table `tblSettings`):

| SettingKey | SettingValue | DataType | Category | Purpose |
| :--- | :--- | :--- | :--- | :--- |
| `DefaultSteelRate` | `61000` | Number | Pricing | Base steel cost (₹/MT) |
| `DefaultGradeExtra` | `0` | Number | Pricing | Grade extra cost (₹/MT) |
| `DefaultYieldPct` | `2` | Number | Pricing | Yield loss allowance (%) |
| `DefaultConversion` | `5000` | Number | Pricing | Conversion cost (₹/MT) |
| `DefaultTesting` | `0` | Number | Pricing | Testing & inspection (₹/MT) |
| `DefaultMargin` | `0` | Number | Pricing | Commercial margin (₹/MT) |
| `DefaultFinancialPct` | `1.5` | Number | Pricing | Working capital financing (%) |
| `GSTPct` | `18` | Number | Pricing | GST rate applied to financial factor (%) |
| `PIPct` | `3.14` | Number | Pricing | Circumference multiplier ($\pi$) |
| `HandlingRsPerMT` | `100` | Number | Freight | Logistics handling surcharge (₹/MT) |
| `TripRateManual` | `98000` | Number | Freight | Legacy lump-sum trip rate (₹/trip) |
| `RatePerKm` | `125.32` | Number | Freight | Dynamic freight rate per km (₹/km) |
| `ReferenceDistanceKm`| `782` | Number | Freight | Reference trip distance from Jabalpur (km) |

---

## 3. Physical Units & Normal-Mode Coating Analysis

### 3.1 Theoretical Dimensional Analysis

Let pipe geometry and coating parameters be:
- Outer Diameter: $OD$ (mm), converted to metres as $OD / 1000$.
- Inner Diameter: $ID$ (mm), converted to metres as $ID / 1000$.
- Circumference Multiplier: $\pi$ (dimensionless).
- Coating Application Rate: $Rate$ ($\text{₹}/\text{m}^2$).
- Pipe Mass: $massKgPerM$ ($\text{kg}/\text{m}$), and $massMtPerM = massKgPerM / 1000$ ($\text{MT}/\text{m}$).

Surface area per linear metre of pipe is:
$$\text{Area per metre} = \pi \times \frac{OD}{1000} \quad \left(\frac{\text{m}^2}{\text{m}}\right)$$

Coating cost per linear metre of pipe is:
$$\text{Cost per linear metre} = Rate \times \pi \times \frac{OD}{1000} \quad \left(\frac{\text{₹}}{\text{m}}\right)$$

To express this cost **per Metric Ton (MT)** of steel:
$$\text{Cost per MT} = \frac{\text{Cost per linear metre}}{\text{Mass in MT per linear metre}} = \frac{Rate \times \pi \times (OD / 1000)}{massMtPerM} \quad \left(\frac{\text{₹}}{\text{MT}}\right)$$

### 3.2 The Defect in `modPricing.bas`

In [modPricing.bas](file:///C:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modPricing.bas#L334-L380):

```vba
' EXTERNAL COATING
If lineData.HasExt Then
    If rates.legacyMode Then
        ext = CDec(rates.ExtRate) * CDec(rates.PI) * od / CDec(1000) / massMtPerM
    Else
        ext = CDec(rates.ExtRate) * CDec(rates.PI) * od / CDec(1000) / massKgPerM
    End If
End If

' INTERNAL COATING
If lineData.HasInt Then
    If rates.legacyMode Then
        intc = CDec(rates.IntRate) * CDec(rates.PI) * od / CDec(1000) / massMtPerM
    Else
        intc = CDec(rates.IntRate) * CDec(rates.PI) * idd / CDec(1000) / massKgPerM
    End If
End If
```

#### What went wrong:
1. In `legacyMode`, the code divides by `massMtPerM`. The result is in **₹/MT** (~₹7,630/MT and ~₹4,195/MT).
2. In Normal mode, the code divides by `massKgPerM`. The result is in **₹/kg**, which is:
   $$\frac{\text{₹/m}}{\text{kg/m}} = \text{₹/kg} = \frac{\text{₹/MT}}{1000}$$
   Instead of ₹7,630/MT, `ext` evaluated to **₹7.63/kg**.
   Instead of ₹4,139/MT, `intc` evaluated to **₹4.14/kg**.
3. In lines 489–499:
   ```vba
   result.ExtPerM = RoundUpD(ext * massMtPerM, 0)
   result.IntPerM = RoundUpD(intc * massMtPerM, 0)
   ```
   Because `ext` was already deflated by 1,000, multiplying by `massMtPerM` ($massKgPerM / 1000$) multiplied by $1/1000$ a second time:
   $$\text{ext} \times massMtPerM = \frac{\text{₹/m}}{1000} \approx \frac{1952}{1000} = 1.952$$
   Applying `RoundUpD(1.952, 0)` rounded this to **₹2/m**!
   Similarly, `intc * massMtPerM` $\approx 1.059$, rounded to **₹2/m**!
4. In lines 432–454, the financial cost formula uses:
   $$baseCost = pre + ext + intc + freight$$
   Because `ext` and `intc` were ₹7.63 and ₹4.14, they contributed only ₹11.77 to `baseCost` instead of ₹11,770. This reduced financial cost per MT by ~₹212/MT.

---

## 4. Freight Rate Verification (Task 7)

A specific concern was whether production normal-mode freight receives the intended trip rate or defaults to zero.

### 4.1 Implementation Inspection

In `modPricing.LoadDefaultRateInput`:
```vba
If legacyMode Then
    .TripRate = CDec(SettingValue("TripRateManual"))
Else
    ' Production freight distance is supplied by the later freight,
    ' plant, and ORS workflow. Golden fixtures override this value.
    .TripRate = CDec(0)
End If
```

In `modPricing.P3GoldenLine`:
```vba
If legacy Then
    rates.TripRate = CDec(98000)
    rates.PI = CDec(3.14)
Else
    rates.TripRate = CDec(125.32) * CDec(782)
End If
```

In `modQuoteLines.AddQuoteLine`:
```vba
If IsMissing(tripRate) Then
    rates.TripRate = CDec(SettingValue("TripRateManual"))
Else
    rates.TripRate = CDec(tripRate)
End If
```

### 4.2 Findings on Freight
1. In the golden test harness (`P3GoldenLine`), `rates.TripRate` is explicitly initialized to $125.32 \times 782 = \text{₹}98,000.24$.
2. In production (`modQuoteLines.AddQuoteLine`), if `tripRate` is missing, it safely falls back to `TripRateManual` (₹98,000).
3. In both Current (defective) and Normal (target) calculation runs, freight is **identical**:
   - Line 1: 42 trailers / 511.511 MT $\rightarrow$ `freight` = ₹8,147/MT $\rightarrow$ `FreightPerM` = ₹2,084/m
   - Line 2: 8 trailers / 94.880 MT $\rightarrow$ `freight` = ₹8,364/MT $\rightarrow$ `FreightPerM` = ₹2,268/m
   - Line 3: 11 trailers / 140.721 MT $\rightarrow$ `freight` = ₹7,761/MT $\rightarrow$ `FreightPerM` = ₹2,185/m
   - Line 4: 92 trailers / 1283.923 MT $\rightarrow$ `freight` = ₹7,123/MT $\rightarrow$ `FreightPerM` = ₹2,079/m
4. **Conclusion on Freight:** Freight is functioning exactly as specified and does **not** default to zero. Freight contributes ₹0 to the ₹24,306,200 discrepancy.

---

## 5. Line-by-Line Calculation Audit

### 5.1 Input Data for Golden Lines

All four golden lines use:
- Basis: `"ID"`, Wall Thickness: $8.4\text{ mm}$, Length per Pipe: $12\text{ m}$, Pipes per Trailer: $4$
- Commercial Rates: Steel = ₹61,000, Yield = 2%, Conversion = ₹5,000, ExtRate = ₹500/$\text{m}^2$, IntRate = ₹275/$\text{m}^2$
- Commercial Terms: Financial = 1.5%, GST = 18%, Handling = ₹100/MT, Trip Rate = ₹98,000.24 ($125.32 \times 782$)
- $PreCost = 61000 \times 1.02 + 5000 = \text{₹}67,220/\text{MT}$
- Financial factor $k = 0.015 \times (1 + 0.18) = 0.0177 \rightarrow k / (1 - k) = 0.0177 / 0.9823 \approx 0.018018935$

Geometry derivation:
- Line 1: ID = 1226 mm $\rightarrow$ OD = $\lceil 1226 + 16.8 \rceil = 1243\text{ mm}$ $\rightarrow$ Mass = $255.75553836\text{ kg/m}$ ($0.255755538\text{ MT/m}$)
- Line 2: ID = 1300 mm $\rightarrow$ OD = $\lceil 1300 + 16.8 \rceil = 1317\text{ mm}$ $\rightarrow$ Mass = $271.08512676\text{ kg/m}$ ($0.271085127\text{ MT/m}$)
- Line 3: ID = 1350 mm $\rightarrow$ OD = $\lceil 1350 + 16.8 \rceil = 1367\text{ mm}$ $\rightarrow$ Mass = $281.44295676\text{ kg/m}$ ($0.281442957\text{ MT/m}$)
- Line 4: ID = 1400 mm $\rightarrow$ OD = $\lceil 1400 + 16.8 \rceil = 1417\text{ mm}$ $\rightarrow$ Mass = $291.80078676\text{ kg/m}$ ($0.291800787\text{ MT/m}$)

---

### 5.2 Line 1: ID 1226 mm, WT 8.4 mm, Qty 2,000 m (511.511 MT, 167 Pipes, 42 Trailers)

| Calculation Parameter | Current Defective | Repaired Normal ($\pi=3.14$) | Repaired ($\pi=\text{Pi()}$) | Unit |
| :--- | :--- | :--- | :--- | :--- |
| External Coating per MT (`ext`) | 7.6342 | **7630.3724** | 7634.2426 | ₹/MT |
| Internal Coating per MT (`intc`) | 4.1414 | **4139.3082** | 4141.4077 | ₹/MT |
| Freight per MT (`freight`) | 8147.0000 | 8147.0000 | 8147.0000 | ₹/MT |
| Financial Base Cost (`baseCost`) | 75378.7756 | 87136.6806 | 87142.6503 | ₹/MT |
| Financial Cost per MT (`financial`)| 1358.2452 | **1570.1102** | 1570.2178 | ₹/MT |
| **Bare Selling Price / m (`BarePerM`)** | 17540 | **17594** | 17594 | ₹/m |
| **External Coating / m (`ExtPerM`)** | 2 | **1952** | 1953 | ₹/m |
| **Internal Coating / m (`IntPerM`)** | 2 | **1059** | 1060 | ₹/m |
| **Freight Price / m (`FreightPerM`)** | 2084 | **2084** | 2084 | ₹/m |
| **Total Price / m (`TotalPerM`)** | **19,628** | **22,689** | 22,691 | ₹/m |
| **Line Value (2,000 m)** | **₹39,256,000** | **₹45,378,000** | ₹45,382,000 | ₹ |
| **Difference vs Target (₹22,689 / ₹45,378,000)** | **-₹3,061 / -₹6,122,000** | **0 / 0 (EXACT)** | +₹2 / +₹4,000 | |

---

### 5.3 Line 2: ID 1300 mm, WT 8.4 mm, Qty 350 m (94.880 MT, 30 Pipes, 8 Trailers)

| Calculation Parameter | Current Defective | Repaired Normal ($\pi=3.14$) | Repaired ($\pi=\text{Pi()}$) | Unit |
| :--- | :--- | :--- | :--- | :--- |
| External Coating per MT (`ext`) | 7.6313 | **7627.4565** | 7631.3252 | ₹/MT |
| Internal Coating per MT (`intc`) | 4.1431 | **4140.9502** | 4143.0505 | ₹/MT |
| Freight per MT (`freight`) | 8364.0000 | 8364.0000 | 8364.0000 | ₹/MT |
| Financial Base Cost (`baseCost`) | 75595.7744 | 87352.4067 | 87358.3757 | ₹/MT |
| Financial Cost per MT (`financial`)| 1362.1552 | **1573.9974** | 1574.1049 | ₹/MT |
| **Bare Selling Price / m (`BarePerM`)** | 18592 | **18650** | 18650 | ₹/m |
| **External Coating / m (`ExtPerM`)** | 3 | **2068** | 2069 | ₹/m |
| **Internal Coating / m (`IntPerM`)** | 2 | **1123** | 1124 | ₹/m |
| **Freight Price / m (`FreightPerM`)** | 2268 | **2268** | 2268 | ₹/m |
| **Total Price / m (`TotalPerM`)** | **20,865** | **24,109** | 24,111 | ₹/m |
| **Line Value (350 m)** | **₹7,302,750** | **₹8,438,150** | ₹8,438,850 | ₹ |
| **Difference vs Target (₹24,109 / ₹8,438,150)** | **-₹3,244 / -₹1,135,400** | **0 / 0 (EXACT)** | +₹2 / +₹700 | |

---

### 5.4 Line 3: ID 1350 mm, WT 8.4 mm, Qty 500 m (140.721 MT, 42 Pipes, 11 Trailers)

| Calculation Parameter | Current Defective | Repaired Normal ($\pi=3.14$) | Repaired ($\pi=\text{Pi()}$) | Unit |
| :--- | :--- | :--- | :--- | :--- |
| External Coating per MT (`ext`) | 7.6295 | **7625.6660** | 7629.5339 | ₹/MT |
| Internal Coating per MT (`intc`) | 4.1441 | **4141.9583** | 4144.0592 | ₹/MT |
| Freight per MT (`freight`) | 7761.0000 | 7761.0000 | 7761.0000 | ₹/MT |
| Financial Base Cost (`baseCost`) | 74992.7736 | 86748.6243 | 86754.5931 | ₹/MT |
| Financial Cost per MT (`financial`)| 1351.2898 | **1563.1178** | 1563.2254 | ₹/MT |
| **Bare Selling Price / m (`BarePerM`)** | 19299 | **19359** | 19359 | ₹/m |
| **External Coating / m (`ExtPerM`)** | 3 | **2147** | 2148 | ₹/m |
| **Internal Coating / m (`IntPerM`)** | 2 | **1166** | 1167 | ₹/m |
| **Freight Price / m (`FreightPerM`)** | 2185 | **2185** | 2185 | ₹/m |
| **Total Price / m (`TotalPerM`)** | **21,489** | **24,857** | 24,859 | ₹/m |
| **Line Value (500 m)** | **₹10,744,500** | **₹12,428,500** | ₹12,429,500 | ₹ |
| **Difference vs Target (₹24,857 / ₹12,428,500)** | **-₹3,368 / -₹1,684,000** | **0 / 0 (EXACT)** | +₹2 / +₹1,000 | |

---

### 5.5 Line 4: ID 1400 mm, WT 8.4 mm, Qty 4,400 m (1283.923 MT, 367 Pipes, 92 Trailers)

| Calculation Parameter | Current Defective | Repaired Normal ($\pi=3.14$) | Repaired ($\pi=3.14026$) | Repaired ($\pi=\text{Pi()}$) | Unit |
| :--- | :--- | :--- | :--- | :--- | :--- |
| External Coating per MT (`ext`) | 7.6240 | **7624.0027** | 7624.6340 | 7627.8698 | ₹/MT |
| Internal Coating per MT (`intc`) | 4.1429 | **4142.8949** | 4143.2380 | 4144.9963 | ₹/MT |
| Freight per MT (`freight`) | 7123.0000 | 7123.0000 | 7123.0000 | 7123.0000 | ₹/MT |
| Financial Base Cost (`baseCost`) | 74354.7669 | 86109.8976 | 86110.8720 | 86115.8661 | ₹/MT |
| Financial Cost per MT (`financial`)| 1339.7937 | **1551.6087** | 1551.6262 | 1551.7162 | ₹/MT |
| **Bare Selling Price / m (`BarePerM`)** | 20006 | **20068** | **20068** | 20068 | ₹/m |
| **External Coating / m (`ExtPerM`)** | 3 | **2225** | **2225** | 2226 | ₹/m |
| **Internal Coating / m (`IntPerM`)** | 2 | **1209** | **1210** | 1210 | ₹/m |
| **Freight Price / m (`FreightPerM`)** | 2079 | **2079** | **2079** | 2079 | ₹/m |
| **Total Price / m (`TotalPerM`)** | **22,090** | **25,581** | **25,582** | 25,583 | ₹/m |
| **Line Value (4,400 m)** | **₹97,196,000** | **₹112,556,400** | **₹112,560,800** | ₹112,565,200 | ₹ |
| **Difference vs Target (₹25,582 / ₹112,560,800)** | **-₹3,492 / -₹15,364,800** | **-₹1 / -₹4,400** | **0 / 0 (EXACT)** | +₹1 / +₹4,400 | |

---

### 5.6 Grand Total Comparison

| Mode / Configuration | Line 1 Value | Line 2 Value | Line 3 Value | Line 4 Value | Grand Total | Diff vs Required Total |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Current Defective Code** | ₹39,256,000 | ₹7,302,750 | ₹10,744,500 | ₹97,196,000 | **₹154,499,250** | -₹24,306,200 |
| **Repaired Normal ($\pi=3.14$)** | ₹45,378,000 | ₹8,438,150 | ₹12,428,500 | ₹112,556,400 | **₹178,801,050** | -₹4,400 |
| **Repaired Normal ($\pi=3.14026$)** | ₹45,378,000 | ₹8,438,150 | ₹12,428,500 | ₹112,560,800 | **₹178,805,450** | **₹0 (EXACT)** |
| **Repaired Normal ($\pi=\text{Pi()}$)**| ₹45,382,000 | ₹8,438,850 | ₹12,429,500 | ₹112,565,200 | **₹178,815,550** | +₹10,100 |
| **Authoritative Target** | **₹45,378,000** | **₹8,438,150** | **₹12,428,500** | **₹112,560,800** | **₹178,805,450** | **Reference Target** |

---

## 6. Investigation of the ₹24,306,200 Difference (Task 6)

The exact deficit between the Current Defective total (₹154,499,250) and the Required Normal total (₹178,805,450) is **₹24,306,200**.

### 6.1 Line-by-Line Breakdown

$$\begin{aligned}
\text{Line 1 (2,000 m): } & (22689 - 19628) \times 2000 = 3061 \times 2000 = \text{₹}6,122,000 \\
\text{Line 2 (350 m): } & (24109 - 20865) \times 350 = 3244 \times 350 = \text{₹}1,135,400 \\
\text{Line 3 (500 m): } & (24857 - 21489) \times 500 = 3368 \times 500 = \text{₹}1,684,000 \\
\text{Line 4 (4,400 m): } & (25582 - 22090) \times 4400 = 3492 \times 4400 = \text{₹}15,364,800 \\
\hline
\textbf{Total: } & \mathbf{6,122,000 + 1,135,400 + 1,684,000 + 15,364,800 = ₹24,306,200}
\end{aligned}$$

### 6.2 Component-by-Component Cause Breakdown

Every rupee of the deficit per metre is completely accounted for:

1. **External Coating per Metre Deficit:**
   - Line 1: $1952 - 2 = \text{₹}1,950/\text{m}$
   - Line 2: $2068 - 3 = \text{₹}2,065/\text{m}$
   - Line 3: $2147 - 3 = \text{₹}2,144/\text{m}$
   - Line 4: $2225 - 3 = \text{₹}2,222/\text{m}$
2. **Internal Coating per Metre Deficit:**
   - Line 1: $1059 - 2 = \text{₹}1,057/\text{m}$
   - Line 2: $1123 - 2 = \text{₹}1,121/\text{m}$
   - Line 3: $1166 - 2 = \text{₹}1,164/\text{m}$
   - Line 4: $1209 - 2 = \text{₹}1,207/\text{m}$ (or ₹1,208/m if Int=1210)
3. **Bare Pipe Price per Metre Deficit (Financial Cost Suppression):**
   - Line 1: $17594 - 17540 = \text{₹}54/\text{m}$
   - Line 2: $18650 - 18592 = \text{₹}58/\text{m}$
   - Line 3: $19359 - 19299 = \text{₹}60/\text{m}$
   - Line 4: $20068 - 20006 = \text{₹}62/\text{m}$
4. **Freight Price per Metre:**
   - Lines 1–4: **₹0/m difference** (100% identical in both modes).

Summing the components:
- Line 1: $1950 + 1057 + 54 + 0 = \mathbf{3061}$
- Line 2: $2065 + 1121 + 58 + 0 = \mathbf{3244}$
- Line 3: $2144 + 1164 + 60 + 0 = \mathbf{3368}$
- Line 4: $2222 + 1208 + 62 + 0 = \mathbf{3492}$

---

## 7. Deep Analysis of Line 4 Single-Rupee Sensitivity

When normal-mode coating is repaired to divide by `massMtPerM`:
- Lines 1, 2, and 3 match the required golden prices **₹22,689, ₹24,109, and ₹24,857 with 0.00 difference**.
- Line 4 produces **₹25,581** under standard $\pi = 3.14$, whereas the target table lists **₹25,582**.

### 7.1 Unrounded Components of Line 4

Under Decimal arithmetic with $\pi = 3.14$:
- $BareUnrounded = (67220 + 1551.60866) \times 0.29180078676 = 20,067.609514 \rightarrow \lceil \cdot \rceil = 20,068$
- $ExtUnrounded = 500 \times 3.14 \times 1.417 = 2,224.690000 \rightarrow \lceil \cdot \rceil = 2,225$
- $IntUnrounded = 275 \times 3.14 \times 1.400 = 1,208.900000 \rightarrow \lceil \cdot \rceil = 1,209$
- $FreightUnrounded = 7123 \times 0.29180078676 = 2,078.497004 \rightarrow \lceil \cdot \rceil = 2,079$
- Component Sum: $20068 + 2225 + 1209 + 2079 = \mathbf{25,581}$

### 7.2 The $\pi$ Rounding Threshold Window

Why would Line 4 be ₹25,582?
Notice that $IntUnrounded$ is $1,208.900000$. The distance to the next integer ($1,210.00$) is only $0.100000$:
$$\frac{1209.000000}{275 \times 1.400} = \frac{1209}{385} \approx 3.14025974$$

For all $\pi > 3.14025974$:
- Line 4 `IntPerM` evaluates to $> 1209.00$, rounding up to **1,210**.
- Line 4 total becomes $20068 + 2225 + 1210 + 2079 = \mathbf{25,582}$!

Crucially, in the window $3.14025974 \le \pi \le 3.140470$:
- Line 1 Ext ($1951.51 \rightarrow 1952$) and Int ($1058.65 \rightarrow 1059$) **do not cross**. Total remains **22,689**.
- Line 2 Ext ($2067.69 \rightarrow 2068$) and Int ($1122.55 \rightarrow 1123$) **do not cross**. Total remains **24,109**.
- Line 3 Ext ($2146.19 \rightarrow 2147$) and Int ($1165.73 \rightarrow 1166$) **do not cross**. Total remains **24,857**.
- Line 4 Int crosses to **1,210**, bringing Line 4 to **25,582** and Grand Total to **₹178,805,450**.

### 7.3 Forensic Discovery from Backup Modules

In `modSelfTest.embedded.bak.20261008_130143` (at 13:01:43 on 2026-10-08):
- `normalExpected = CDec(178805450)`
- `expectedPerM = Choose(index + 1, 22689, 24109, 24857, 25582)`

Just 29 seconds later, in `modSelfTest.embedded.bak.20261008_130212`:
- A previous developer changed `currentExpected = CDec(154499250)` and `expectedPerM = Choose(index + 1, 19628, 20865, 21489, 22090)` to mask the test failure rather than diagnosing the coating unit bug!

---

## 8. Recommended Smallest Required Code Change

### 8.1 Proposed Modification to `modPricing.bas`

In [modPricing.bas](file:///C:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modPricing.bas#L344-L380):

#### Before (Lines 344–350 and 373–379):
```vba
        If rates.legacyMode Then
            ext = _
                CDec(rates.ExtRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        Else
            ext = _
                CDec(rates.ExtRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massKgPerM
        End If
...
        If rates.legacyMode Then
            intc = _
                CDec(rates.IntRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        Else
            intc = _
                CDec(rates.IntRate) * _
                CDec(rates.PI) * _
                idd / _
                CDec(1000) / _
                massKgPerM
        End If
```

#### After:
```vba
        If rates.legacyMode Then
            ext = _
                CDec(rates.ExtRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        Else
            ext = _
                CDec(rates.ExtRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        End If
...
        If rates.legacyMode Then
            intc = _
                CDec(rates.IntRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        Else
            intc = _
                CDec(rates.IntRate) * _
                CDec(rates.PI) * _
                idd / _
                CDec(1000) / _
                massMtPerM
        End If
```

*Note: Since both branches now divide by `massMtPerM`, the `If rates.legacyMode` branch can also be simplified so external coating uses `od / 1000 / massMtPerM` unconditionally, and internal coating uses `IIf(rates.legacyMode, od, idd) / 1000 / massMtPerM`.*

### 8.2 Expected Impact of Proposed Change

1. `ext` and `intc` will now correctly represent costs in **₹/MT**.
2. `result.ExtPerM` and `result.IntPerM` will recover from ₹2/m to ₹1,952/m, ₹2,068/m, ₹2,147/m, ₹2,225/m, and ₹1,059/m, ₹1,123/m, ₹1,166/m, ₹1,209/m.
3. `baseCost` will accurately reflect total production cost per MT, restoring full financial cost per MT (~₹1,570/MT).
4. Normal-mode quotation total will immediately recover from ₹154,499,250 to **₹178,801,050** (recovering ₹24,301,800 of the ₹24,306,200 gap).
5. Legacy mode behavior is 100% preserved.

---

## 9. Conclusion

Required total not reproducible; missing or conflicting business assumption identified.

### Detailed Engineering Justification for Conclusion

1. Under the strict business settings documented in `tblSettings` (`PIPct = 3.14`), the mathematically exact repair of the coating unit defect reproduces Lines 1, 2, and 3 with zero error (₹22,689, ₹24,109, ₹24,857), but produces **₹25,581** on Line 4 (a 1-rupee variance against target ₹25,582), resulting in a normal-mode total of **₹178,801,050** (a ₹4,400 variance against target ₹178,805,450).
2. The exact authoritative target of **₹178,805,450** and Line 4 price of **₹25,582** requires either:
   - A business agreement that $\pi$ for Line 4 internal coating carries higher precision ($\pi \ge 3.14026$), or
   - Clarification of whether the target was derived using unrounded component summation or an independent ceiling rule on Line 4.
3. In accordance with safety rules, no changes have been applied to production VBA or the DEV workbook pending user review of this audit.

