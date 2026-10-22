# MEGA PIPES QUOTATION SYSTEM — COMMERCIAL POLICY DECISION REGISTER

**Document Reference:** `DOCS/COMMERCIAL_POLICY_DECISION_REGISTER.md`  
**Parent Package:** `DOCS/COMMERCIAL_POLICY_SIGNOFF.md`  
**Registration Date:** 21 October 2026  
**System Target:** Mega Pipes Quotation System  
**Executive Gate Status:** `BLOCKED PENDING EXECUTIVE SIGN-OFF`  
**Governing Principle:** Zero Commercial Policies Modified Without Written Management Authorization  

---

## 1. Executive Summary & Purpose

This Decision Register provides a rigorous, transparent inventory of all pending commercial pricing policies for the Mega Pipes Quotation System. 

In strict adherence to governance boundaries:
1. No commercial formulas, rate tables, margins, freight models, coating rules, or golden baselines have been silently altered.
2. The current implementation in `DEV/Mega_Pipes_Quotation_DEV.xlsm` is maintained with its existing, fully validated baseline behavior.
3. Every unresolved policy is explicitly classified as an **Executive Release Blocker** for production cutover until formal written approval is executed by the designated corporate authority.

---

## 2. Policy Decision Register

```
====================================================================================================
DECISION ITEM 1: FINANCE COST FORMULATION (NEW ENQUIRIES VS. HISTORICAL REPLICATION)
====================================================================================================
```
- **Exact Unresolved Question:**  
  How should finance cost be derived on new commercial quotations: as a formulaic percentage of total landed cost including GST, or as an estimator-supplied manual rupee markup inside bare-pipe pre-cost?
- **Historical Context:**  
  In the reference workbook `MKTGENQ26-27-594 - Unipro 2.xlsx` sheet `CS Stacking 10-09-26`, two conflicting formulations exist:
  - *Row 32 (Pre-Cost):* Hardcoded manual amounts (`1024, 1014, 1011, 1008` ₹/MT) included inside bare-pipe pre-cost.
  - *Row 41 (Post-Freight):* Formulaic 1.0% of landed cost including 18% GST: `=C40 * $B$41` (`1031.15, 1025.93, 1022.76, 1019.80` ₹/MT), which was informational and not added downstream.
- **Current System Behavior in DEV:**  
  - Dynamic Dual-Resolution:
    - If `tblQuoteLines[ManualFC] > 0`: System applies the exact manual financial cost per MT, replicating historical contracts to ₹0.00 variance.
    - If `tblQuoteLines[ManualFC] = 0` (standard for new enquiries): System computes landed formulaic finance cost:
      $$\text{Finance Cost/MT} = \text{Round}\Big(\big(\text{PreCost} + \text{Ext} + \text{Int} + \text{Freight}\big) \times 1.18 \times \frac{\text{Fin}\%}{100}, 0\Big)$$
- **Commercial Impact:**  
  - On Unipro Line 1 (1226 ID × 8.4mm): Manual FC = ₹1,024/MT (Line price ₹22,351/m); Formulaic FC = ₹1,019/MT (Line price ₹22,350/m). Difference is -₹1.00/m (-0.004%), or ₹2,000 on a ₹4.47 Crore order.
  - Standardizing formulaic finance cost prevents estimator omission of financing on long-haul transport or expensive coating specifications.
- **Relevant Test Coverage:**  
  - `TESTS/test_pricing_golden_baselines.py` (`TestApprovedCompanyCostSheetBaselines`, `test_approved_formulaic_without_manual_override`)
  - `TESTS/test_commercial_user_acceptance.py` (`test_03_enquiry_1_approved_unipro_line_by_line`)
- **Documented Decision Options:**  
  - **Option A (Recommended):** Standardize formulaic landed finance cost for all new enquiries; preserve `ManualFC` column in `tblQuoteLines` strictly for legacy replication or special customer credit terms.
  - **Option B:** Revert to mandatory manual pre-cost input for all quotes.
- **Authorized Decision-Maker:** Chief Financial Officer (CFO) & Commercial Director
- **Approval Status & Evidence:** **`PENDING EXECUTIVE SIGN-OFF`** (Awaiting formal sign-off in `DOCS/COMMERCIAL_POLICY_SIGNOFF.md`).
- **Release Status:** **RELEASE BLOCKER**

---

```
====================================================================================================
DECISION ITEM 2: INTERNAL COATING SURFACE-AREA BASIS (OD VS. ID CONVENTION)
====================================================================================================
```
- **Exact Unresolved Question:**  
  Should internal lining application cost per MT be calculated on Outside Diameter (OD) surface area (historical commercial practice) or Inside Diameter (ID) surface area (rigorous physical geometry)?
- **Historical Context:**  
  - Historical spreadsheets (Unipro Row 20) calculated internal coating cost using OD area: `3.14 * OD / 1000 = 3.90302 m²/m`, yielding ₹3,815.19/MT.
  - Physical geometry dictates that internal paint coats the internal wall surface: `3.14 * ID / 1000 = 3.84964 m²/m`, yielding ₹3,762.99/MT.
- **Current System Behavior in DEV:**  
  - When `ApprovedMode = True` (or Legacy Mode): System applies **OD surface area**, matching the reference CostSheet to zero paisa.
  - When standard `Normal` mode is active: System applies **ID surface area**, matching actual physical paint consumption.
- **Commercial Impact:**  
  - On Unipro Line 1 (1226 ID, 1243 OD, 8.4mm WT): OD basis yields ₹976/m; ID basis yields ₹963/m (-₹13/m). Across the 7,250m tender, variance is ~₹95,000 (~0.05% of subtotal).
  - On heavy-wall projects (e.g. Lodha 676 ID × 12mm WT, 700 OD), OD vs ID surface area divergence is 3.5%, offering a ₹150–200/MT pricing flexibility edge.
- **Relevant Test Coverage:**  
  - `TESTS/test_pricing_variance_regression.py` (`test_internal_coating_diameter_effect`)
  - `TESTS/test_pricing_golden_baselines.py` (`TestApprovedCompanyCostSheetBaselines`)
- **Documented Decision Options:**  
  - **Option A (Recommended):** Introduce explicit configuration `InternalCoatingAreaBasis` in `tblSettings` with default `"OD"` (for historical tender alignment) and optional `"ID"` (for engineering audits and competitive cost-plus bidding).
  - **Option B:** Mandate physical ID area universally across all modes.
  - **Option C:** Mandate OD area universally across all modes.
- **Authorized Decision-Maker:** Estimating Head & Plant Operations Head
- **Approval Status & Evidence:** **`PENDING EXECUTIVE SIGN-OFF`**
- **Release Status:** **RELEASE BLOCKER**

---

```
====================================================================================================
DECISION ITEM 3: MANUAL SELLING-PRICE OVERRIDE UNITS & CONVERSION RULES
====================================================================================================
```
- **Exact Unresolved Question:**  
  When sales estimators enter an optional manual selling price override on a quotation line, how should the system handle negotiations conducted in ₹/MT versus ₹/m?
- **Historical Context:**  
  Water transmission pipelines are quoted in ₹/m, while bulk structural and oil/gas pipes are negotiated in ₹/MT. Previous spreadsheets forced estimators to perform offline calculator conversions.
- **Current System Behavior in DEV:**  
  - The manual price override field on the `QUOTE` sheet (`B32`) accepts numeric input.
  - Cell `C32` dynamically displays `₹/m` or `₹/MT` based on the quotation header's active `PriceBasis` (`B16`).
  - Underlying cost breakdown stages (steel, conversion, freight, coating) are permanently preserved and never back-calculated.
- **Commercial Impact:**  
  - Ensures accurate valuation regardless of customer contracting units. Eliminates manual rounding errors during bid submission.
- **Relevant Test Coverage:**  
  - `TESTS/test_optional_price_override_ui.py` (`test_12_quote_in_metres_with_mt_price_basis_prices_by_mt`, `test_13_quote_in_metres_with_metre_price_basis_prices_by_metre`)
  - `TESTS/test_commercial_user_acceptance.py` (`test_07_enquiry_5_dynamic_staged_costing_and_manual_override`)
- **Documented Decision Options:**  
  - **Option A (Recommended):** Context-Aware Dynamic Basis. Override input unit is governed by quotation `PriceBasis`. Dual display of ₹/m and ₹/MT is shown for estimator reference. Price conversion applies `RoundUp(PricePerMT * MassMTperM, 0)`.
  - **Option B:** Restrict manual overrides strictly to ₹/m.
- **Authorized Decision-Maker:** Sales & Marketing Operations Head & Commercial Director
- **Approval Status & Evidence:** **`PENDING EXECUTIVE SIGN-OFF`**
- **Release Status:** **RELEASE BLOCKER**

---

```
====================================================================================================
DECISION ITEM 4: BELOW-FLOOR PRICE GOVERNANCE, ESCALATION & APPROVAL GATES
====================================================================================================
```
- **Exact Unresolved Question:**  
  What governance controls, audit logging, and authorization gates must be enforced when an estimator enters a manual selling price below factory floor cost (Pre-Cost + Landed Freight)?
- **Historical Context:**  
  Unrestricted manual override in legacy spreadsheets created vulnerability to negative-margin tender submissions without management awareness.
- **Current System Behavior in DEV:**  
  - The system records `ManualSellingPrice`, logs `PriceVariance`, tags `PriceSource = "MANUAL_OVERRIDE"`, mandates an audit reason, and automatically invalidates quote state from `Approved` back to `Calculated`.
- **Commercial Impact:**  
  - Protects company against accidental cash-negative contract commitments while preserving executive flexibility to accept strategic loss-leader tenders.
- **Relevant Test Coverage:**  
  - `TESTS/test_commercial_user_acceptance.py` (`test_07`)
  - `TESTS/test_pricing_golden_baselines.py` (`TestTamperDetectionAndGateEnforcement`)
  - `modApproval.RunApprovalSelfTests`
- **Documented Decision Options:**  
  - **Option A (Recommended):** High-Priority Modal Warning + Administrative Gate. Present negative-margin alert dialog, log `UNDER_FLOOR_OVERRIDE` in `tblQuoteLog`, tag quote `BELOW_FLOOR`, and strictly block state transition to `Approved` or customer PDF export without `MEGA_ADMIN_PWD` authorization.
  - **Option B:** Hard System Lockout. Completely prohibit any manual price entry below floor cost.
  - **Option C:** Advisory Warning Only. Display warning but allow normal quote approval.
- **Authorized Decision-Maker:** Managing Director & Commercial Director
- **Approval Status & Evidence:** **`PENDING EXECUTIVE SIGN-OFF`**
- **Release Status:** **RELEASE BLOCKER**

---

## 3. Summary Decision Register Table

| Decision Record | Policy Item | Current DEV Behavior | Documented Recommendation | Approving Authority | Current Status | Release Impact |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Record 1** | Finance Cost Formulation | Dual Mode (ManualFC if >0; Landed formula if 0) | Standardize formulaic landed for new quotes; retain ManualFC for legacy replication | CFO & Commercial Director | **PENDING** | **RELEASE BLOCKER** |
| **Record 2** | Internal Coating Area | OD in Approved/Legacy Mode; ID in Normal Mode | Configurable setting in `tblSettings` (`"OD"` default, `"ID"` optional) | Estimating Head & Plant Operations Head | **PENDING** | **RELEASE BLOCKER** |
| **Record 3** | Manual Override Units | Dynamic unit based on quotation PriceBasis | Context-aware unit prompt + dual display (₹/m & ₹/MT) + RoundUp conversion | Sales Head & Commercial Director | **PENDING** | **RELEASE BLOCKER** |
| **Record 4** | Below-Floor Governance | Invalidation to Calculated + mandatory audit reason | Modal warning + `UNDER_FLOOR_OVERRIDE` log + `MEGA_ADMIN_PWD` approval gate | Managing Director & Commercial Director | **PENDING** | **RELEASE BLOCKER** |

---

## 4. Formal Recommendation to Executive Leadership

1. Maintain current DEV implementation unchanged.
2. Convene executive review to execute sign-off signatures in [`DOCS/COMMERCIAL_POLICY_SIGNOFF.md`](file:///c:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/DOCS/COMMERCIAL_POLICY_SIGNOFF.md).
3. Upon executive sign-off, activate policy options in DEV settings without modifying core calculation engines.
