# Mega Pipes Quotation System — Human Action Required Items (`NEEDS_HUMAN`)

This catalog records all enterprise requirements, secrets, operational configurations, and managerial approvals that strictly require human action prior to live production deployment, in accordance with Section 15 of `ANTIGRAVITY_MASTER_PROMPT.md`.

---

## 1. OpenRouteService Live API Key (`ORS_API_KEY`)
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 9 (ORS Integration) & Phase 24 (Unified Self-Test)
- **Reason:** Private external API credentials cannot be committed to version control or hardcoded in templates.
- **Action Required:**
  Set the OpenRouteService API key in the Windows environment:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("ORS_API_KEY", "your_ors_api_key_here", "User")
  ```
- **Fallback Activated:**
  - Mocked response tests and cache tests: `PASS`
  - Seeded distance cache from Unipro reference workbook: Active (`782 km` for Jabalpur, `1821 km` for Khopoli to Bagodar 825322).
  - Manual distance fallback: Fully operational.
- **Safety Guarantee:** No fake keys used; live check skipped safely without blocking downstream phases.

---

## 2. Passwords: `MEGA_SHEET_PWD` & `MEGA_ADMIN_PWD`
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 21 (Approval Workflow), Phase 23 (Workbook Protection), Phase 24 (Self-Test)
- **Reason:** Passwords are not hardcoded in source code and require configuration in the Windows environment for production deployments.
- **Action Required:**
  Set production passwords in the Windows environment:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("MEGA_SHEET_PWD", "your_sheet_protection_password", "User")
  ```
- **Fallback Activated:**
  - When unset, development macros proceed with empty-string password validation.
  - Full structural and cell locking protection remains verified.

---

## 3. Production Shared Network Folder UNC Path & Permissions (`MEGA_SHARED_FOLDER`)
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 13 (Sequence Numbering), Phase 16 (Multi-User), Phase 27 (Shared Storage Validation)
- **Reason:** Corporate network file server path and NTFS permissions must be provisioned by IT network administrators.
- **Action Required:**
  1. Create UNC share (e.g., `\\fileserver\MegaPipes\QuotationShared\`).
  2. Grant Read/Write/Modify permissions to Quoting Users, Full Control to Administrators.
  3. Deploy directory structure: `Quotes\`, `Sequence\`, `Masters\`.
  4. Configure environment variable:
     ```powershell
     [System.Environment]::SetEnvironmentVariable("MEGA_SHARED_FOLDER", "\\fileserver\MegaPipes\QuotationShared", "User")
     ```
- **Fallback Activated:**
  - Local workstation `Quotes\` and `Sequence\` directory fallback.
  - Offline sequential generation (`TEMP-...`) operational if share is disconnected.

---

## 4. Manual Selling-Price Override Unit and Conversion Rules
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 5 (Manual Price Override), Phase 6 (Acceptance Audit)
- **Business Question:**
  Currently, the manual selling-price override macro (`UI_ManualPriceOverride`) prompts and captures the override price in the quotation line's active `QuantityUnit` (typically ₹ per Metre). If an enquiry requests a quotation on a per-MT basis (`QuantityUnit = "MT"`), how should the override behave?
- **Options for Management Decision:**
  - **Option A (Recommended):** Context-Aware Unit. The override dialog prompts in the enquiry's active unit (`Metre` or `MT`). If per-MT, derive the line's per-metre selling price using `RoundUp(PricePerMT * MassMTperM, 0)`.
  - **Option B:** Strict Dual-Entry. Require the user to specify both per-Metre and per-MT prices explicitly.
- **Safe Interim Default:** Active unit context is used. The prompt displays both calculated per-Metre and calculated per-MT prices for reference.

---

## 5. Below-Floor Selling Price Warning and Commercial Approval Policy
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 5 (Manual Price Override), Phase 21 (Approval Workflow)
- **Business Question:**
  If an authorized user enters a manual selling price that falls below the baseline manufacturing cost (e.g. `PreCost + Freight` per MT, representing negative gross margin):
- **Options for Management Decision:**
  - **Option A (Recommended):** Warning Confirmation + Admin Escalation. Present a high-priority warning dialog displaying the negative margin amount, log an `UNDER_MARGIN_OVERRIDE` alert in `tblQuoteLog`, and require `MEGA_ADMIN_PWD` approval before transitioning from `Review` to `Approved`.
  - **Option B:** Hard Floor Block. Completely prohibit any manual selling price override below the pre-cost floor.
- **Safe Interim Default:** The override is recorded, approval state is automatically invalidated to `Calculated`, and reason is mandated for audit review.

---

---

## 6. Financial Cost Formulation Ambiguity in Historical Reference CostSheets
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 1 (Audit), Phase 3 (Staged Costing Engine)
- **Business Question:**
  The approved Unipro reference workbook (`CS Stacking 10-09-26`) exhibits two different financial cost locations:
  1. Row 32 (Pre-Cost): Hardcoded manual values (`1024, 1014, 1011, 1008`) included inside bare-pipe pre-cost `C33 = SUM(C27:C32)`.
  2. Row 41 (Post-Freight): Formulaic 1% of total landed cost including 18% GST: `=C40 * $B$41` (`1031.15, 1025.93, 1022.76, 1019.80`), which was not added downstream in that specific historical sheet.
- **Options for Management Decision:**
  - **Option A (Recommended):** Formulaic Standard for New Enquiries. Derive financial cost dynamically as `ROUND((PreCost + Ext + Int + Freight) * 1.18 * (FinPct / 100), 0)`, while preserving the `ManualFC` column in `tblQuoteLines` to allow exact replication of legacy or negotiated contracts.
  - **Option B:** Pre-Cost Insertion. Always add financial cost as a line item inside Pre-Cost before coating and freight.
- **Safe Interim Default:** Option A is implemented. When `ManualFC > 0`, the manual amount is used; when `ManualFC = 0`, the dynamic landed formula is computed.

---

## 7. Internal-Coating Surface Area Convention
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 2 (Coating Engine), Phase 6 (Acceptance Audit)
- **Business Question:**
  In historical estimating practice, the approved Unipro reference workbook (`MKTGENQ26-27-594 - Unipro 2.xlsx`) computes internal coating area using Outer Diameter (`3.14 * OD / 1000 = 3.90302 m²/m`), yielding `(250 * 3.90302) / 0.2557555 = ₹3,815.19/MT`. However, rigorous physical geometry dictates that internal lining covers Inner Diameter (`3.14 * ID / 1000 = 3.84964 m²/m`), which would yield `₹3,762.99/MT` (a ₹52.20/MT lower cost).
- **Options for Management Decision:**
  - **Option A (Recommended):** Configurable Setting in `tblSettings` (`InternalCoatingAreaBasis`: `"OD"` vs `"ID"`). Default to `"OD"` when in approved reference mode or replicating legacy CostSheets, and default to `"ID"` (physical surface area) for new standard quotes unless commercial management specifies OD basis.
  - **Option B:** Strict Physical ID. Mandate ID basis across all new quotations.
- **Safe Interim Default:** The engine supports both conventions. When `ApprovedMode` is active or matching the Unipro baseline, OD surface area is applied; in standard normal mode, geometric ID surface area is applied.

---

## 8. Real Commercial Rates & Financial Rules Verification
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 4 (Calculation Engine), Phase 7 (Freight Engine), Phase 11 (Costing)
- **Reason:** Purchasing contracts, carrier freight rates, raw material coil tariffs, and commercial margins fluctuate dynamically and must be verified by the Commercial Director.
- **Action Required:**
  Commercial Operations team must review and verify:
  - Base HR Coil purchase prices per grade in `SETTINGS`
  - Scrap recovery credit rate per MT
  - 3LPE, 3LPP, FBE, and Zinc Rich Epoxy coating rates per m²
  - Freight trip rate matrices (Method B trailer rates)
  - Commercial floor margin percentages
  - Financial cost / interest rules per customer credit term
- **Fallback Activated:**
  - Verified baseline rates from Unipro reference workbook `MKTGENQ26-27-594 - Unipro 2.xlsx`.

---

## 8. Multi-User Physical Workstation LAN/WAN Trial
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 16 & Phase 27 (Multi-User Trial)
- **Reason:** Validating SMB locking across distinct physical PCs, varying network latencies, and Windows network redirectors requires human operators across multiple physical machines.
- **Action Required:**
  Conduct a 2-day pilot trial with 3-5 sales estimators working from separate physical PCs against the enterprise network share.
- **Fallback Activated:**
  - Automated 50-concurrency thread simulations on single-machine Windows host passed with 0 duplicates and 0 collisions.

---

## 9. Parallel Quoting Trial (3 to 5 Real Historic Enquiries)
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 26 (Golden Quote Validation) & Phase 30 (Go-Live Readiness)
- **Reason:** Validating 3–5 diverse historical commercial enquiries across different diameters, coatings, and delivery distances ensures commercial estimators and management have complete confidence before deprecating legacy spreadsheets.
- **Action Required:**
  Run 3–5 recent real commercial enquiries in parallel through both the new quotation system and legacy sheets, cross-checking line rates and final offer totals.
- **Fallback Activated:**
  - Full Unipro reference quote (7,250 MTR, ₹176,117,500 Approved CostSheet / ₹178,657,900 Legacy / ₹178,801,050 Option 1 Normal) validated with 0 variance.

---

## 10. Training & Final Production Deployment Decision
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 29 & Phase 30 (Production Cutover)
- **Reason:** Formally authorizing live production deployment, scheduling training sessions, and officially retiring the legacy workbook are exclusive management executive responsibilities.
- **Action Required:**
  Commercial Leadership conducts training sessions and issues formal Go-Live Cutover Authorization.
- **Fallback Activated:**
  - Production release candidate packaged and staged under `RELEASE_CANDIDATE/` with complete operator manuals and rollback procedures. `LIVE` workbook remains untouched.
