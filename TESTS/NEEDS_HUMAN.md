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
  [System.Environment]::SetEnvironmentVariable("MEGA_ADMIN_PWD", "your_admin_approval_password", "User")
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

## 4. Real Commercial Rates & Financial Rules
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
  - Verified baseline rates from Unipro reference workbook `MKTGENQ26-27-594_-_Unipro_CS.xlsx`.

---

## 5. Plant Exact Coordinates & Physical Addresses
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 8 (Plant Engine), Phase 9 (ORS Routing)
- **Reason:** Exact dispatch gate addresses and postal codes for Jabalpur and Khopoli facilities must be confirmed with plant logistics management.
- **Action Required:**
  Verify plant postal codes in `MASTERS` sheet table `tblPlants`.
- **Fallback Activated:**
  - Seeded standard coordinates and postal codes for Jabalpur (MP) and Khopoli (MH).

---

## 6. Competitor Pricing Intelligence per Quote
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 10 (Competitor Engine)
- **Reason:** Tender-specific competitor bids cannot be automated without human intelligence inputs.
- **Action Required:**
  Commercial estimators must populate latest competitor intelligence into the `COMPETITORS` sheet on a per-quote basis.
- **Fallback Activated:**
  - Built-in competitor landed price model and floor price threshold protection.

---

## 7. Multi-User Physical Workstation LAN/WAN Trial
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 16 & Phase 27 (Multi-User Trial)
- **Reason:** Validating SMB locking across distinct physical PCs, varying network latencies, and Windows network redirectors requires human operators across multiple physical machines.
- **Action Required:**
  Conduct a 2-day pilot trial with 3-5 sales estimators working from separate physical PCs against the enterprise network share.
- **Fallback Activated:**
  - Automated 50-concurrency thread simulations on single-machine Windows host passed with 0 duplicates and 0 collisions.

---

## 8. Parallel Quoting Trial (3 to 5 Real Historic Enquiries)
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 26 (Golden Quote Validation) & Phase 30 (Go-Live Readiness)
- **Reason:** Validating 3–5 diverse historical commercial enquiries across different diameters, coatings, and delivery distances ensures commercial estimators and management have complete confidence before deprecating legacy spreadsheets.
- **Action Required:**
  Run 3–5 recent real commercial enquiries in parallel through both the new quotation system and legacy sheets, cross-checking line rates and final offer totals.
- **Fallback Activated:**
  - Full Unipro reference quote (7,250 MTR, ₹178,657,900 Legacy / ₹178,801,050 Option 1 Normal) validated with 0 variance.

---

## 9. Visual Approval of Layout & Offer Letter Wording
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 12 (Offer Sheet), Phase 18 (Export Engine)
- **Reason:** Corporate branding, disclaimer phrasing, legal boilerplate terms, and signature authorizations require visual sign-off by Marketing & Legal leadership.
- **Action Required:**
  Executive review of generated PDF offer letters and customer Excel files.
- **Fallback Activated:**
  - Word-for-word reproduction of 27 commercial terms from Unipro reference offer letter.

---

## 10. Training & Final Production Deployment Decision
- **Status:** `NEEDS_HUMAN`
- **Applicable Phase:** Phase 29 & Phase 30 (Production Cutover)
- **Reason:** Formally authorizing live production deployment, scheduling training sessions, and officially retiring the legacy workbook are exclusive management executive responsibilities.
- **Action Required:**
  Commercial Leadership conducts training sessions and issues formal Go-Live Cutover Authorization.
- **Fallback Activated:**
  - Production release candidate packaged and staged under `RELEASE_CANDIDATE/` with complete operator manuals and rollback procedures. LIVE workbook remains untouched.
