# MEGA PIPES — COMMERCIAL POLICY DECISION & SIGN-OFF PACKAGE

**Document Reference:** `DOCS/COMMERCIAL_POLICY_SIGNOFF.md`  
**Preparation Date:** 19 October 2026  
**System Target:** Mega Pipes Quotation System (DEV Environment)  
**Baseline Verification Status:** 120/120 Python tests passed, 25/25 Real-Excel COM self-tests passed  
**Package Status:** `PENDING EXECUTIVE SIGN-OFF` (No commercial policy has been changed)

---

## 1. Executive Summary & Purpose

This package presents four key commercial policy decisions required to finalize the enterprise pricing configuration of the Mega Pipes Quotation System. 

In strict accordance with project governance boundaries:
- The DEV workbook (`DEV/Mega_Pipes_Quotation_DEV.xlsm`, SHA-256: `93FAFB501B24B74278892BAD13AD5C30F3A8DD637C1EC3F106CC2ACB7F6C98D9`) and its verified backup (`BACKUP/cuat_20261019_1005/Mega_Pipes_Quotation_DEV.xlsm`) remain frozen.
- The LIVE directory remains completely untouched (timestamp `07-10-2026 14:38:12`).
- All proposed policies are marked **`PENDING`** and await formal executive sign-off before implementation.

---

## 2. Commercial Policy Decision Records

```
====================================================================================================
DECISION RECORD 1: FINANCE COST FORMULATION (NEW ENQUIRIES VS. HISTORICAL REPLICATION)
====================================================================================================
```

### Context & Problem Statement
In historical estimating practice, the approved Unipro reference workbook (`CS Stacking 10-09-26`) recorded two distinct financial cost figures:
1. **Row 32 (Pre-Cost):** Hardcoded manual entries (`1024, 1014, 1011, 1008`) included inside bare-pipe pre-cost `C33 = SUM(C27:C32)`. The approved selling prices (`₹22,351`, `₹23,570`, `₹24,396`, `₹25,220`) directly derive from these Row 32 amounts.
2. **Row 41 (Post-Freight):** A formulaic calculation `=C40 * $B$41` (`1031.15, 1025.93, 1022.76, 1019.80`), representing 1.0% of landed cost including 18% GST. This formula was purely informational in that historical sheet and was not added downstream.

### Current System Behavior in DEV
The engine implements **Dynamic Dual-Resolution**:
- If `tblQuoteLines[ManualFC] > 0`, it uses the exact manual financial cost per MT (enabling 100% exact replication of historical or negotiated tenders to zero paisa).
- If `tblQuoteLines[ManualFC] = 0` (standard for new enquiries), it computes formulaic landed finance cost:
  $$\text{Finance Cost/MT} = \text{Round}\Big(\big(\text{PreCost} + \text{Ext} + \text{Int} + \text{Freight}\big) \times 1.18 \times \frac{\text{Fin}\%}{100}, 0\Big)$$

### Proposed Enterprise Policy
1. **New Enquiries:** Standardize all new commercial quotations on the formulaic landed finance cost. Finance cost will automatically cascade based on raw material coil rates, coating specifications, and delivery distances without manual estimator entry.
2. **Historical / Audit Replication:** Retain the `ManualFC` column in `tblQuoteLines` exclusively for replicating historical contracts or entering pre-approved customer credit agreements.

### Commercial Impact & Example
- **Unipro Line 1 (1226 ID x 8.4mm WT):**
  - Row 32 Manual FC: ₹1,024.00/MT -> Component: ₹262/m -> Line Price: ₹22,351/m.
  - Formulaic Landed FC: ₹1,019.00/MT -> Component: ₹261/m -> Line Price: ₹22,350/m.
  - Delta: -₹1.00/m (-0.004%). Across 2,000m, variance is ₹2,000 on a ₹4.47 Crore order.
- **Advantage:** Eliminates manual data-entry errors; dynamically reflects working capital carrying costs across diverse supply chains.

### Alternatives & Risks
- *Alternative A (Fixed Pre-Cost Markup):* Add a fixed rupee markup to steel conversion. *Risk:* Does not account for financing external coating or long-distance transport costs.
- *Alternative B (Mandatory Manual FC):* Require manual entry on every enquiry. *Risk:* Estimator inconsistency and quote turnaround delays.

### Sign-Off Information
- **Approving Roles:** Chief Financial Officer (CFO) & Commercial Director
- **Status:** **`PENDING`**
- **Decision Date:** `____________________`
- **Approver Name & Title:** `________________________________________`
- **Written Rationale / Conditions:**  
  `__________________________________________________________________________________________________`  
  `__________________________________________________________________________________________________`

---

```
====================================================================================================
DECISION RECORD 2: INTERNAL-COATING SURFACE-AREA CONVENTION (OD VS. ID)
====================================================================================================
```

### Context & Problem Statement
In historical spreadsheets (e.g. Unipro `CS Stacking 10-09-26` Row 20), internal coating cost per MT was calculated using the **Outside Diameter** surface area (`3.14 * OD / 1000 = 3.90302 m²/m`), yielding:
$$\text{Cost/MT} = \frac{₹250/\text{m}^2 \times 3.90302\,\text{m}^2/\text{m}}{0.255756\,\text{MT/m}} = ₹3,815.19/\text{MT}$$
However, rigorous physical geometry dictates that internal lining covers the **Inside Diameter** surface area (`3.14 * ID / 1000 = 3.84964 m²/m`), which physically yields:
$$\text{Cost/MT} = \frac{₹250/\text{m}^2 \times 3.84964\,\text{m}^2/\text{m}}{0.255756\,\text{MT/m}} = ₹3,762.99/\text{MT}$$

### Current System Behavior in DEV
- When `ApprovedMode` is active (or legacy mode), the system applies **OD surface area**, matching the reference CostSheet to zero paisa.
- When standard `Normal` mode is active, the system applies **ID surface area**, reflecting physical paint shop coverage.

### Proposed Enterprise Policy
1. Introduce an explicit configuration in `tblSettings` (`InternalCoatingAreaBasis`: `"OD"` vs `"ID"`).
2. **Default Policy:** Set default to `"OD"` for tender parity with historical customer cost models.
3. Allow project-specific selection of `"ID"` for competitive cost-plus audits or heavy-wall pipes where ID vs OD surface area divergence is significant.

### Commercial Impact & Example
- **Unipro Line 1 (1226 ID, 1243 OD, 8.4mm WT):**
  - OD Basis: ₹3,815.19/MT -> Component: ₹976/m -> Line Price: ₹22,351/m.
  - ID Basis: ₹3,762.99/MT -> Component: ₹963/m -> Line Price: ₹22,338/m (-₹13/m).
  - Across 2,000m, variance is ₹26,000. Across the 7,250m tender, total difference is ~₹95,000 (~0.05% of subtotal).
- **Heavy-Wall Pipeline Example (Lodha 676 ID x 12mm WT, 700 OD):**
  - Surface area difference is 3.5%. On large projects, ID basis can provide a ₹150–200/MT competitive pricing edge.

### Alternatives & Risks
- *Alternative A (Mandate Physical ID Universally):* Enforce ID everywhere. *Risk:* Deviates from approved Unipro baseline (Subtotal would calculate to ₹176,022,500 instead of ₹176,117,500).
- *Alternative B (Mandate OD Universally):* Keep OD everywhere. *Risk:* Vulnerable to technical audit challenge from EPCs or government water authorities who verify physical surface area.

### Sign-Off Information
- **Approving Roles:** Estimating Head & Plant Operations Head
- **Status:** **`PENDING`**
- **Decision Date:** `____________________`
- **Approver Name & Title:** `________________________________________`
- **Written Rationale / Conditions:**  
  `__________________________________________________________________________________________________`  
  `__________________________________________________________________________________________________`

---

```
====================================================================================================
DECISION RECORD 3: MANUAL SELLING-PRICE OVERRIDE UNITS & CONVERSION RULES
====================================================================================================
```

### Context & Problem Statement
Sales estimators and commercial managers negotiate with customers in different commercial units depending on tender terms:
- Water supply transmission lines are typically negotiated in **₹ per Metre** (`INR/MTR`).
- Bulk oil & gas or raw pipe supply tenders are often negotiated in **₹ per Metric Tonne** (`INR/MT`).

### Current System Behavior in DEV
- `UI_ManualPriceOverride` prompts in the line's active `QuantityUnit`.
- It records `ManualSellingPrice`, preserves `CalculatedUnitPrice`, tags `PriceSource = "MANUAL_OVERRIDE"`, computes `PriceVariance = Manual - Calc`, mandates a reason, and updates downstream `LineValue = Manual * Quantity`.
- Underlying stage costs (raw material, conversion, coating, freight) are **never** altered or back-calculated.

### Proposed Enterprise Policy
1. **Context-Aware Unit Display:** The override dialog will present both the system-calculated ₹/m and equivalent ₹/MT prices.
2. **Dual-Unit Input:** Estimators may input the manual price in either ₹/m or ₹/MT.
3. **Conversion Rule:** If entered in ₹/MT, the per-metre price is derived using:
   $$\text{Price/m} = \text{RoundUp}(\text{Price/MT} \times \text{Mass MT/m}, 0)$$
4. **Permanent Preservation:** The system-calculated unit price and original staged breakdown remain permanently recorded in `tblQuoteLines` for audit history.

### Commercial Impact & Example
- An estimator negotiating an enquiry calculated at ₹22,351/m (equivalent to ₹87,392/MT) agrees to a rounded customer rate of ₹88,000/MT.
- The system derives `RoundUp(88000 * 0.2557555, 0) = ₹22,507/m`.
- `PriceVariance` is recorded as `+₹156/m` (+₹608/MT), tagging the line as `MANUAL_OVERRIDE` with mandatory justification logged.

### Alternatives & Risks
- *Alternative A (Per-Metre Only):* Restrict overrides strictly to ₹/m. *Risk:* Estimators must perform manual calculator conversions when customers negotiate in ₹/MT, introducing rounding discrepancies.
- *Alternative B (Allow Back-Calculation of Costs):* Modify steel rate to match override price. *Risk:* HIGH RISK. Distorts inventory accounting, raw material purchasing budgets, and plant conversion reports.

### Sign-Off Information
- **Approving Roles:** Sales & Marketing Operations Head & Commercial Director
- **Status:** **`PENDING`**
- **Decision Date:** `____________________`
- **Approver Name & Title:** `________________________________________`
- **Written Rationale / Conditions:**  
  `__________________________________________________________________________________________________`  
  `__________________________________________________________________________________________________`

---

```
====================================================================================================
DECISION RECORD 4: BELOW-FLOOR PRICE WARNINGS, AUDIT LOGGING & APPROVAL RESTRICTIONS
====================================================================================================
```

### Context & Problem Statement
When competitive pressures require pricing below the fully loaded price, the system must distinguish between:
- Selling with reduced margin (commercial concession).
- Selling below the factory floor manufacturing cost (cash-negative order).

### Current System Behavior in DEV
- The system permits manual overrides, logs variance, tags `PriceSource`, and invalidates approval to `"Calculated"`.
- However, there is currently no hard system block preventing an estimator from entering an override that yields negative gross margin.

### Proposed Enterprise Policy
1. **Floor Price Definition:**
   $$\text{Floor Price/MT} = \text{Raw Material (with Yield)} + \text{Conversion} + \text{External Coating} + \text{Internal Coating} + \text{Landed Freight}$$
   *(Represents zero gross margin and zero financing recovery).*
2. **Level 1 Warning Dialog:** If `ManualSellingPrice < FloorPrice`, display a modal warning:
   > *"CRITICAL WARNING: The entered selling price of ₹X/m is below the baseline manufacturing cost of ₹Y/m. This represents a negative gross margin of -₹Z/m (-₹W/MT). A mandatory executive escalation reason is required."*
3. **Audit Log Flag:** System writes an `UNDER_FLOOR_OVERRIDE` event to `tblQuoteLog`.
4. **Approval Block Gate:** The quotation is tagged `BELOW_FLOOR`. Status transition from `Review` to `Approved` or `Issued` is **STRICTLY BLOCKED** unless authorized with `MEGA_ADMIN_PWD`.
5. **Customer Export Block:** Unapproved below-floor quotes cannot generate official customer PDFs (`ExportPDF` disabled until approved).

### Commercial Impact & Example
- An estimator cannot accidentally issue a quotation that results in a cash loss.
- Strategic decisions to take loss-leader orders (e.g. entry into a new state water board) remain possible, but require logged Managing Director sign-off.

### Alternatives & Risks
- *Alternative A (Hard System Lockout):* Absolutely forbid any quote below floor. *Risk:* Eliminates commercial executive discretion during crisis negotiations.
- *Alternative B (Unrestricted Warning):* Show a warning without an admin password block. *Risk:* HIGH RISK of unauthorized sales reps submitting below-cost tenders.

### Sign-Off Information
- **Approving Roles:** Managing Director & Commercial Director
- **Status:** **`PENDING`**
- **Decision Date:** `____________________`
- **Approver Name & Title:** `________________________________________`
- **Written Rationale / Conditions:**  
  `__________________________________________________________________________________________________`  
  `__________________________________________________________________________________________________`

---

## 3. Technical Confirmation Checklist

The technical engine has been validated against all reference standards. Stakeholders can confirm the following baseline verification items:

| Verification Item | Specification / Benchmark | DEV Verified Result | Evidence File / Test | Verification Status |
|---|---|---|---|---|
| **1. Authentic Unipro Sizing** | Line 1: 1226 ID x 8.4mm (2000m)<br>Line 2: 1300 ID x 8.4mm (350m)<br>Line 3: 1350 ID x 8.4mm (500m)<br>Line 4: 1400 ID x 8.4mm (4400m) | All 4 lines have uniform 8.4mm WT; Total: 7,250m, 2,031.036 MT | [`TESTS/test_commercial_user_acceptance.py`](file:///c:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/TESTS/test_commercial_user_acceptance.py): `test_02` | **VERIFIED (0 Delta)** |
| **2. Approved Baseline Pricing** | Line 1: ₹22,351/m<br>Line 2: ₹23,570/m<br>Line 3: ₹24,396/m<br>Line 4: ₹25,220/m | Exact match to `CS Stacking 10-09-26` Row 48 | [`TESTS/test_pricing_golden_baselines.py`](file:///c:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/TESTS/test_pricing_golden_baselines.py): `TestApprovedCompanyCostSheetBaselines` | **VERIFIED (0 Delta)** |
| **3. Approved Financial Totals** | Subtotal: ₹176,117,500.00<br>GST (18%): ₹31,701,150.00<br>Grand Total: ₹207,818,650.00 | Exact match to approved Unipro offer | `RunGoldenApproved` (Real Excel COM)<br>`test_03_enquiry_1_approved_unipro_line_by_line` | **VERIFIED (0 Delta)** |
| **4. Historical Legacy Baseline** | Subtotal: ₹178,657,900.00<br>Line prices: 22681, 23914, 24748, 25580 | Exact match to historical legacy stacking | `RunGoldenLegacy` (Excel COM)<br>`test_05_enquiry_3_historical_legacy_contract` | **VERIFIED (0 Delta)** |
| **5. Historical Normal Baseline** | Subtotal: ₹178,801,050.00<br>Line prices: 22689, 24109, 24857, 25581 | Exact match to Option 1 contract specification | `RunGoldenNormal` (Excel COM)<br>`test_06_enquiry_4_option1_normal_contract` | **VERIFIED (0 Delta)** |
| **6. Lodha Physical Geometry** | 25,700m, 8,057.338 MT, 2,144 dispatch pipes, 2,141.67 frac pipes, 86,540m² ext | Exact match to `MKTGENQ26-27610 - Lodha 1.xlsx` | `test_04_enquiry_2_lodha_transmission_pipeline`<br>`TESTS/test_lodha_five_line_review.py` | **VERIFIED (Geometry)**<br>*(Pricing `PENDING`)* |
| **7. Export Sanitization** | Customer Excel contains ONLY `OFFER` and `Synopsis`; 0 formulas; 0 rate leaks | 100% values-only; confidential sheets and keywords blocked | `test_08_customer_export_leakage_and_sanitization`<br>`ValidateStagedCustomerExcel` | **VERIFIED (Zero Leak)** |

---

## 4. Management Sign-Off Circulation Sheet

This package is submitted to executive management for review and decision sign-off.

| Role | Designee / Title | Key Decisions to Review | Sign-Off Status | Signature | Date |
|---|---|---|---|---|---|
| **Commercial Director** | Commercial Operations Head | Decisions 1, 3, 4 (Finance, Overrides, Floor Policy) | **`PENDING`** | `_____________________` | `__________` |
| **Chief Financial Officer (CFO)** | Finance & Accounts Head | Decision 1 (Finance Cost Formulation & Working Capital) | **`PENDING`** | `_____________________` | `__________` |
| **Estimating Head** | Chief Tendering & Costing Engineer | Decision 2 (Internal Coating Area Basis) & Sizing Audit | **`PENDING`** | `_____________________` | `__________` |
| **Plant Operations Head** | General Manager - Jabalpur Plant | Decision 2 (Coating Consumption & Manufacturing Costs) | **`PENDING`** | `_____________________` | `__________` |
| **Sales & Marketing Operations** | Regional Sales Head | Decision 3 (Manual Selling-Price Override Units) | **`PENDING`** | `_____________________` | `__________` |
| **Managing Director** | Executive Leadership | Decision 4 (Below-Floor Governance & Final Deployment) | **`PENDING`** | `_____________________` | `__________` |

---

## 5. Governance & Next Steps

1. **Workbook Status:** The DEV workbook remains frozen at SHA-256 `93FAFB501B24B74278892BAD13AD5C30F3A8DD637C1EC3F106CC2ACB7F6C98D9`.
2. **Deployment Guard:** No files in `LIVE` will be modified, opened, or published.
3. **Action Required:** Executive management must review the four decision records and provide written decisions. Once approved, the corresponding configuration parameters will be activated in `DEV`, followed by final regression testing and production cutover authorization.
