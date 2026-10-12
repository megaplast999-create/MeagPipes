# Mega Pipes Quotation System — Line 4 Specification & Calculation Review

**Date:** 2026-10-12  
**Author:** Senior VBA Engineer, Financial Calculation Auditor & Software Test Engineer  
**Status:** Audit Complete — No Workbook or VBA Modifications Performed  
**Target File:** `TESTS\ANTIGRAVITY_LINE4_SPECIFICATION_REVIEW.md`  

---

## 1. Executive Summary

Following the successful application of the approved coating unit correction in [modPricing.bas](file:///C:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modPricing.bas) and verification in Microsoft Excel 16.0:

1. **Legacy Mode Total:** ₹178,657,900 — **Exact Match** (Difference = 0).
2. **Normal Mode Lines 1–3:**
   - Line 1 (2,000 m): **₹22,689/m** — **Exact Match** (Difference = 0).
   - Line 2 (350 m): **₹24,109/m** — **Exact Match** (Difference = 0).
   - Line 3 (500 m): **₹24,857/m** — **Exact Match** (Difference = 0).
3. **Normal Mode Line 4:**
   - Actual Calculated Price: **₹25,581/m**
   - Required Target Price: **₹25,582/m**
   - Variance: **-₹1/m** over 4,400 metres = **-₹4,400**.
4. **Normal Mode Grand Total:**
   - Actual Calculated Total: **₹178,801,050**
   - Required Target Total: **₹178,805,450**
   - Variance: **-₹4,400** (a 0.00246% variance, fully localized to Line 4).

### Key Investigation Finding

A rigorous forensic audit of all available specification documents, historical spreadsheets, git commit logs, and backup files demonstrates that:

1. **Under the authoritative documented business inputs and formulas (`tblSettings.PIPct = 3.14`), the mathematically exact selling price for Line 4 is ₹25,581/m, and the exact grand total is ₹178,801,050.**
2. There is **no documented specification, engineering drawing, or quotation sheet** in the project repository that derives or justifies ₹25,582/m for Line 4.
3. The number **₹25,582/m** originated solely as an unverified test fixture entry in `modSelfTest.bas` during development on 2026-10-08, derived by dividing a preconceived grand total target of ₹178,805,450 ($178,805,450 - 66,244,650 = 112,560,800 \div 4,400 = 25,582$).
4. Because the authoritative specification does not provide a formulaic basis for the additional 1 rupee on Line 4 without introducing arbitrary constants (such as selecting $\pi \in [3.14026, 3.14047]$ solely to force a match), this audit identifies the **missing business decision** and stops without modifying production VBA or workbook data.

---

## 2. Forensic Trace of Authoritative Specifications & Historical Evidence

### 2.1 Inspection of `MASTER_PROMPT_MANUS.md`

- **Search Scope:** Recursively scanned the project root (`C:\Users\Administrator\Desktop\Mega Pipes\Quotation System`), the parent directory (`C:\Users\Administrator\Desktop\Mega Pipes`), and the user environment.
- **Result:** Confirmed **not present**. No file named `MASTER_PROMPT_MANUS.md` exists in the project workspace.
- **Context:** The name references an earlier AI engineering run (Manus) whose pre-run state was preserved in `BACKUP\Mega_Pipes_Quotation_DEV_pre_Manus_20261012_110421.xlsm`.

### 2.2 Inspection of the Primary Historical Commercial Reference

The only external, commercial quotation artifact available is the original customer enquiry workbook:
`C:\Users\Administrator\Desktop\automation\reference\MKTGENQ26-27-594_-_Unipro_CS.xlsx` (dated 15/09/2026 for Unipro Techno Infrastructure Pvt. Ltd.).

Inspection of all sheets in this reference workbook reveals:

| Sheet Name | Content & Purpose | Normal-Mode Target Present? |
| :--- | :--- | :--- |
| `Synopsis` | Technical pipe data (OD, ID, WT, Qty, Coating m²) | No (Technical inputs only) |
| `CS Stacking 10-09-26` | Detailed cost build-up (Steel, Conversion, FC, Coating, Freight) | **No** (Contains only the **LEGACY** numbers) |
| `Sheet1` | Official commercial quotation offer letter | **No** (Contains only the **LEGACY** offer: ₹22,681, ₹23,914, ₹24,748, ₹25,580; Total: ₹178,657,900) |

**Conclusion from Reference Workbook:**  
The reference workbook contains the exact historical commercial figures for **LEGACY mode only**. It does not contain any calculations or references to ₹22,689, ₹24,109, ₹24,857, ₹25,582, or ₹178,805,450.

### 2.3 Inspection of Git Repository History

Inspection of the git commit log in `Quotation System` (`git log -p` across commits `f32f722`, `6da0747`, `0e9f36f`, `7f2308b`, `382f2d1`, `43865ec`):

- **Commit `f32f722` (2026-10-08 10:01:34):** Established pre-compile baseline. `PIPct` in `tblSettings` was seeded as `3.14`.
- **Commit `43865ec` (2026-10-08 15:13:48):** Completed quote-line calculation orchestration.
  In `DOCS\CALCULATION_ENGINE.md`, the developer recorded:
  ```markdown
  The diagnostic golden run agrees with the reference workbook:
  - LEGACY per-metre values: 22681, 23914, 24748, 25580
  - LEGACY total: 178657900
  - CURRENT contract total: 154499250
  ```
  Neither the commit diffs nor the documentation ever committed the target ₹178,805,450 or the line price ₹25,582 to the repository.

### 2.4 Forensic Inspection of Embedded Backup Modules

The target values ₹25,582 and ₹178,805,450 were discovered exclusively in historical uncommitted backup files created on 2026-10-08 between 12:23 and 13:02:

1. **`modSelfTest.bas.embedded.bak.20261008_122310` (12:23:10):**
   ```vba
   Print #fileNumber, "Legacy total actual=" & CStr(legacyTotal) & " expected=178657900"
   Print #fileNumber, "Normal total actual=" & CStr(normalTotal) & " expected=178805450"
   ```
   *Observation:* The figure **178805450** appeared here for the first time as an expected variable.
2. **`modSelfTest.embedded.bak.20261008_130143` (13:01:43):**
   ```vba
   normalExpected = CDec(178805450)
   ...
   expectedPerM = Choose(index + 1, 22689, 24109, 24857, 25582)
   ```
   *Observation:* The per-metre breakdown `(22689, 24109, 24857, 25582)` appeared here.
3. **`modSelfTest.embedded.bak.20261008_130212` (13:02:12 — exactly 29 seconds later):**
   ```vba
   currentExpected = CDec(154499250)
   ...
   expectedPerM = Choose(index + 1, 19628, 20865, 21489, 22090)
   ```
   *Observation:* Because the defective coating formula produced ₹154,499,250, the developer modified the test expectation to match the failing actual output rather than diagnosing the code.

---

## 3. Independent Line 4 Component Calculation

To determine the mathematically correct Line 4 values, we perform an independent calculation from first principles using standard Decimal arithmetic and the documented settings in `tblSettings` (`PIPct = 3.14`, `DefaultSteelRate = 61000`, `DefaultYieldPct = 2`, `DefaultConversion = 5000`, `DefaultFinancialPct = 1.5`, `GSTPct = 18`, `RatePerKm = 125.32`, `HandlingRsPerMT = 100`, `ReferenceDistanceKm = 782`).

### 3.1 Geometry, Mass, and Logistics

- **Inner Diameter ($ID$):** $1400\text{ mm}$ (Input)
- **Wall Thickness ($WT$):** $8.4\text{ mm}$ (Input)
- **Outer Diameter ($OD$):**
  $$OD = \text{RoundUp}(ID + 2 \times WT, 0) = \text{RoundUp}(1400 + 16.8, 0) = \mathbf{1417\text{ mm}}$$
- **Mass per Metre ($kg/m$):**
  $$massKgPerM = (OD - WT) \times WT \times 0.0246615 = (1417 - 8.4) \times 8.4 \times 0.0246615 = \mathbf{291.80078676\text{ kg/m}}$$
- **Mass per Metre ($MT/m$):**
  $$massMtPerM = \frac{massKgPerM}{1000} = \mathbf{0.29180078676\text{ MT/m}}$$
- **Quantity:** $4400\text{ metres}$
- **Total Tonnage:**
  $$MT = 4400 \times 0.29180078676 = \mathbf{1283.923461744\text{ MT}}$$
- **Logistics (Discrete Truckloads):**
  $$Pipes = \text{RoundUp}\left(\frac{4400}{12}, 0\right) = \text{RoundUp}(366.667, 0) = \mathbf{367\text{ pipes}}$$
  $$Trailers = \text{RoundUp}\left(\frac{367}{4}, 0\right) = \text{RoundUp}(91.75, 0) = \mathbf{92\text{ trailers}}$$

---

### 3.2 Cost per Metric Ton (₹/MT)

#### 1. Pre-Cost (`pre`)
$$\begin{aligned}
SteelCost & = 61000 + (61000 \times 0.02) = 61000 + 1220 = 62220\text{ ₹/MT} \\
PreCost & = 62220 + 5000 = \mathbf{67,220.00\text{ ₹/MT}}
\end{aligned}$$

#### 2. External Coating per MT (`ext`)
Using $Rate = 500\text{ ₹/m}^2$, $\pi = 3.14$, $OD = 1417\text{ mm}$:
$$ext = \frac{500 \times 3.14 \times (1417 / 1000)}{0.29180078676} = \frac{2224.69}{0.29180078676} = \mathbf{7624.002747565\text{ ₹/MT}}$$

#### 3. Internal Coating per MT (`intc`)
Using $Rate = 275\text{ ₹/m}^2$, $\pi = 3.14$, $ID = 1400\text{ mm}$ (normal mode uses inner diameter):
$$intc = \frac{275 \times 3.14 \times (1400 / 1000)}{0.29180078676} = \frac{1208.90}{0.29180078676} = \mathbf{4142.894929870\text{ ₹/MT}}$$

#### 4. Freight per MT (`freight`)
Using $TripRate = 125.32 \times 782 = 98000.24$, $Trailers = 92$, $MT = 1283.923461744$:
$$\begin{aligned}
FreightRaw & = \frac{98000.24 \times 92}{1283.923461744} = \frac{9016022.08}{1283.923461744} \approx 7022.24264\text{ ₹/MT} \\
FreightPerMT & = \text{RoundUp}(7022.24264, 0) + 100 = 7023 + 100 = \mathbf{7123.00\text{ ₹/MT}}
\end{aligned}$$

#### 5. Base Cost & Financial Cost (`financial`)
$$\begin{aligned}
baseCost & = pre + ext + intc + freight \\
& = 67220 + 7624.002747565 + 4142.894929870 + 7123 = \mathbf{86,109.897677435\text{ ₹/MT}}
\end{aligned}$$
Using $FinancialPct = 1.5\%$, $GSTPct = 18\%$:
$$k = 0.015 \times (1 + 0.18) = 0.0177$$
$$1 - k = 1 - 0.0177 = 0.9823$$
$$\frac{k}{1 - k} = \frac{0.0177}{0.9823} \approx 0.0180189351522$$
$$financial = 86109.897677435 \times \frac{0.0177}{0.9823} = \mathbf{1551.608662212\text{ ₹/MT}}$$

---

### 3.3 Selling Price per Metre Components (Before and After Rounding)

In [modPricing.bas](file:///C:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modPricing.bas#L482-L512), each component is converted to ₹/metre by multiplying its ₹/MT value by `massMtPerM` and applying `RoundUpD(..., 0)`:

| Component | Formula | Exact Unrounded Value (₹/m) | Rounding Rule | Rounded Value (₹/m) |
| :--- | :--- | :--- | :--- | :--- |
| **Bare Pipe Price** | `(pre + financial) * massMtPerM` | $(67220 + 1551.608662) \times 0.29180078676 = \mathbf{20067.609514}$ | $\lceil 20067.609514 \rceil$ | **₹20,068** |
| **External Coating** | `ext * massMtPerM` | $7624.002748 \times 0.29180078676 = 500 \times 3.14 \times 1.417 = \mathbf{2224.690000}$ | $\lceil 2224.690000 \rceil$ | **₹2,225** |
| **Internal Coating** | `intc * massMtPerM` | $4142.894930 \times 0.29180078676 = 275 \times 3.14 \times 1.400 = \mathbf{1208.900000}$ | $\lceil 1208.900000 \rceil$ | **₹1,209** |
| **Freight** | `freight * massMtPerM` | $7123 \times 0.29180078676 = \mathbf{2078.497004}$ | $\lceil 2078.497004 \rceil$ | **₹2,079** |
| **Total Price / m** | `Bare + Ext + Int + Freight` | Sum of rounded components: $20068 + 2225 + 1209 + 2079$ | Exact Whole Rupee | **₹25,581** |

$$\text{Line 4 Commercial Value} = 25581 \times 4400 = \mathbf{₹112,556,400}$$

$$\begin{aligned}
\text{Grand Total} & = \text{Line 1} + \text{Line 2} + \text{Line 3} + \text{Line 4} \\
& = 45,378,000 + 8,438,150 + 12,428,500 + 112,556,400 = \mathbf{₹178,801,050}
\end{aligned}$$

---

## 4. Mathematical Decomposition of the Transition: Legacy vs. Normal

Comparing Legacy mode against Normal mode across all four lines reveals why Lines 1–3 match their targets while Line 4 yields ₹25,581 instead of ₹25,582:

### 4.1 The Three Business Modernizations
1. **Dynamic Financial Cost:** Replaced manual array $(1542, 1534, 1529, 1525)$ with formulaic calculation (~₹1,570 to ~₹1,551/MT).
   - Line 1 Bare: $17587 \rightarrow 17594$ (**+7**)
   - Line 2 Bare: $18639 \rightarrow 18650$ (**+11**)
   - Line 3 Bare: $19349 \rightarrow 19359$ (**+10**)
   - Line 4 Bare: $20060 \rightarrow 20068$ (**+8**)
2. **Discrete Truckload Freight:** Replaced fractional trailers with $\lceil pipes / 4 \rceil$.
   - Line 1 Freight: $2068 \rightarrow 2084$ (**+16**)
   - Line 2 Freight: $2069 \rightarrow 2268$ (**+199**)
   - Line 3 Freight: $2071 \rightarrow 2185$ (**+114**)
   - Line 4 Freight: $2071 \rightarrow 2079$ (**+8**)
3. **Internal Diameter Coating:** Replaced outer diameter ($OD$) with inner diameter ($ID$).
   - Line 1 Int: $1074 \rightarrow 1059$ (**-15**)
   - Line 2 Int: $1138 \rightarrow 1123$ (**-15**)
   - Line 3 Int: $1181 \rightarrow 1166$ (**-15**)
   - Line 4 Int: $1224 \rightarrow 1209$ (**-15**)
4. **External Coating:** Uses $OD$ in both modes (Delta = **0** across all lines).

### 4.2 Net Delta by Line

| Line | Legacy Price | Bare Delta | Ext Delta | Int Delta | Freight Delta | Net Delta | Normal Price | Target Price | Target Delta |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Line 1** | ₹22,681 | +7 | 0 | -15 | +16 | **+8** | **₹22,689** | ₹22,689 | **+8 (EXACT)** |
| **Line 2** | ₹23,914 | +11 | 0 | -15 | +195 | **+195** | **₹24,109** | ₹24,109 | **+195 (EXACT)** |
| **Line 3** | ₹24,748 | +10 | 0 | -15 | +114 | **+109** | **₹24,857** | ₹24,857 | **+109 (EXACT)** |
| **Line 4** | ₹25,580 | +8 | 0 | -15 | +8 | **+1** | **₹25,581** | ₹25,582 | **+2 (+1 variance)** |

Notice the symmetry:
- On Line 4, Bare added **+8**, Freight added **+8**, and Internal Coating subtracted **-15**.
- Net change from legacy: $+8 + 8 - 15 = \mathbf{+1}$.
- Legacy Line 4 was ₹25,580.
- Therefore, $25580 + 1 = \mathbf{25,581}$.
- For Line 4 to be ₹25,582, the net change would have to be **+2**, which requires an extra rupee that does not exist in any documented component.

---

## 5. Evaluation of Hypotheses vs. Documented Evidence

| Hypothesis | Mathematical Behavior | Evaluation Against Documented Evidence |
| :--- | :--- | :--- |
| **A. High-Precision $\pi$ ($\pi = \text{math.pi} \approx 3.14159265...$)** | Pushes Line 4 Ext to 2,226 (+1) and Int to 1,210 (+1), yielding Line 4 = ₹25,583 (+1 vs target). | **DISPROVED.** Using full $\pi$ increases Lines 1–3 to ₹22,691 (+2), ₹24,111 (+2), and ₹24,859 (+2), breaking the verified matches on Lines 1–3 and yielding a grand total of ₹178,815,550 (+₹10,100). |
| **B. Artificial Transition Window $\pi \in [3.14026, 3.14047]$** | At $\pi \approx 3.14026$, Line 4 Int crosses from 1,209 to 1,210 ($275 \times \pi \times 1.4 > 1209.00$), giving Line 4 = ₹25,582, while Lines 1–3 remain unchanged. | **NO DOCUMENTED EVIDENCE.** No engineering specification or setting specifies $\pi = 3.14026$. Selecting this constant solely to force a match violates Rule 8. |
| **C. Unrounded Summation Before Final Ceiling** | $20067.61 + 2224.69 + 1208.90 + 2078.50 = 25579.70 \rightarrow \lceil 25579.70 \rceil = 25,580$. | **DISPROVED.** Unrounded summation produces ₹25,580, not ₹25,582. |
| **D. Total Price per MT Converted to per Metre** | $TotalPerMT = 68772 + 7625 + 4143 + 7123 = 87663$. $87663 \times 0.29180078676 = 25580.13 \rightarrow \lceil 25580.13 \rceil = 25,581$. | **DISPROVED.** Converting $TotalPerMT$ produces ₹25,581, not ₹25,582. |
| **E. Manual Author Rounding / Approximation Slip** | The author of the October 8 test fixture manually estimated Line 4 as: $\text{Legacy } 25580 + 8 (\text{Bare}) + 8 (\text{Freight}) - 14 (\text{estimated Int savings}) = 25582$. | **HIGHLY PROBABLE HYPOTHESIS.** The unrounded internal coating savings is $1223.58 - 1208.90 = 14.68\text{ ₹/m}$. If rounded down to ₹14 rather than evaluated as discrete component ceilings, the net delta is $+2$, creating $25580 + 2 = 25582$. |

---

## 6. Root Cause of the Discrepancy

1. **Origin of the Target:**  
   The target values **₹25,582/m** and **₹178,805,450** were not derived from an engineering calculation sheet or customer quote. They were entered directly into test code on 2026-10-08 as an assumed target.
2. **Nature of the Discrepancy:**  
   The discrepancy is not an implementation bug in the pricing engine. Under the verified formulas:
   - `RoundUpD((pre + financial) * massMtPerM, 0) = 20068`
   - `RoundUpD(ext * massMtPerM, 0) = 2225`
   - `RoundUpD(intc * massMtPerM, 0) = 1209`
   - `RoundUpD(freight * massMtPerM, 0) = 2079`
   - $20068 + 2225 + 1209 + 2079 = \mathbf{25,581}$
   Every component is mathematically exact and deterministically rounded in accordance with the documented business rules.
3. **The ₹4,400 Variance:**  
   $$(25582 - 25581) \times 4400 = 1 \times 4400 = \mathbf{₹4,400}$$
   The entire grand total difference of ₹4,400 is attributable solely to this single rupee on Line 4.

---

## 7. Missing Business Decision & Actionable Options

In accordance with **Rule 9** (*"If the authoritative specification cannot resolve the discrepancy, report the missing business decision and stop"*), the following decision is submitted to business and technical leadership:

### Decision Option 1 (Recommended — Align Golden Target to Mathematical Truth)
- **Action:** Accept **₹25,581/m** for Line 4 and **₹178,801,050** as the authoritative normal-mode contract under `PIPct = 3.14`.
- **Rationale:** Preserves complete calculation integrity, avoids introducing arbitrary constants, perfectly aligns with discrete component ceiling rules, and matches Lines 1, 2, and 3 with zero error.
- **Implementation:** Update `modSelfTest.bas` line 20 to `normalExpected = CDec(178801050)` and line 76 to `Choose(index + 1, 22689, 24109, 24857, 25581)`.

### Decision Option 2 (Business Directive to Force Target)
- **Action:** If business commercial management mandates that Line 4 must strictly sell at ₹25,582/m and the quotation grand total must be ₹178,805,450, a formal policy decision must specify which component carries the additional rupee:
  - *Option 2A:* Authorize an internal coating rounding override or higher precision constant specifically for large-diameter pipes.
  - *Option 2B:* Adjust Line 4 freight or handling inputs to yield a rounded freight price of ₹2,080/m.
  - *Option 2C:* Adjust Line 4 base financial rate to yield a rounded bare price of ₹20,069/m.

---

## 8. Investigation Status & Next Steps

- **Workbook State:** `DEV\Mega_Pipes_Quotation_DEV.xlsm` is clean and compiled in Excel 16.0.
- **Source State:** `VBA\Modules\modPricing.bas` contains only the verified denominator correction (`massMtPerM`).
- **Current Position:** Stopped and awaiting business decision on Option 1 vs Option 2. No code or workbook modifications have been performed during this review.

