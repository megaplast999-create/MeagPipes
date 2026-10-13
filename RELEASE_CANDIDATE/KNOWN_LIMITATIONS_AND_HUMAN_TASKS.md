# Mega Pipes Quotation System — Known Limitations & Human Action Required Preconditions

**Document Reference:** SPEC-LIMIT-P29-001  
**Release Version:** v1.0.0-RC  
**Target Status:** Pre-Go-Live Operational Checklist (`NEEDS_HUMAN`)

---

## 1. Overview of Automated Capabilities vs Human Prerequisites

The core calculation engine, business rules, data structures, multi-user sequence locking, and workbook automation have been **100% verified and validated** with zero errors across automated regression suites, golden contract baselines, and performance stress tests.

However, in accordance with the system specification (Section 15 of `ANTIGRAVITY_MASTER_PROMPT.md`), certain enterprise environmental parameters, live API keys, operational credentials, and commercial rate decisions are strictly governed by enterprise policy and require human intervention before live production cutover.

This document serves as the authoritative catalog of these preconditions.

---

## 2. Complete Inventory of Human Action Items (`NEEDS_HUMAN`)

### 1. OpenRouteService Live API Key (`ORS_API_KEY`)
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** The system executes mocked unit tests and checks the local distance cache (`tblDistanceCache`). When an un-cached destination pincode is entered and no API key is set, the system falls back gracefully to manual distance entry.
- **Why Human is Required:** OpenRouteService requires an account registration and a private API key token (`https://openrouteservice.org`). Storing or hardcoding private API keys in source code violates enterprise security policy.
- **Human Action Required:**
  Register for an enterprise ORS account and configure the key on all user workstations via PowerShell:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("ORS_API_KEY", "your_ors_api_key_here", "User")
  ```

---

### 2. Operational Protection Passwords (`MEGA_SHEET_PWD` & `MEGA_ADMIN_PWD`)
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** In development, empty-string password verification is accepted, allowing automated headless testing.
- **Why Human is Required:** Production sheet protection and commercial administrative approvals require secret passwords known only to authorized administrators and managers.
- **Human Action Required:**
  Define enterprise-grade passwords and set them via PowerShell:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("MEGA_SHEET_PWD", "your_production_sheet_password", "User")
  [System.Environment]::SetEnvironmentVariable("MEGA_ADMIN_PWD", "your_production_admin_approval_password", "User")
  ```

---

### 3. Corporate Network Shared Folder Deployment (`MEGA_SHARED_FOLDER`)
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Concurrency testing verified atomic locking (`sequence.lock`) and isolated snapshot storage using local test directory structures.
- **Why Human is Required:** Physical enterprise file server UNC path and Active Directory security groups must be provisioned by IT network infrastructure personnel.
- **Human Action Required:**
  1. Provision enterprise UNC path (e.g., `\\fileserver\MegaPipes\QuotationShared\`).
  2. Set NTFS permissions: Read/Write/Modify for Estimators, Full Control for Administrators.
  3. Deploy subdirectories: `Quotes\`, `Sequence\`, `Masters\`.
  4. Configure environment variable `MEGA_SHARED_FOLDER` or configure path in `SETTINGS` sheet.

---

### 4. Real Commercial Rates & Financial Rules
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Packaged with reference data derived from `MKTGENQ26-27-594_-_Unipro_CS.xlsx` (Steel base rate ₹68,500/MT, standard scrap recovery credit, 3LPE rate ₹720/m²).
- **Why Human is Required:** Commercial steel coil purchasing contracts, carrier freight rate agreements, and commercial margins fluctuate dynamically and must be verified by the Commercial Director.
- **Human Action Required:**
  Commercial Operations team must review and update:
  - Base HR Coil purchase prices per grade in `SETTINGS`
  - Scrap recovery rate per MT
  - 3LPE, 3LPP, FBE, and Zinc Rich Epoxy coating rates per m²
  - Commercial floor margin percentages
  - Financial cost / interest rules per customer credit term

---

### 5. Manufacturing Plant Physical Coordinates & Addresses
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Seeded with standard coordinates and plant addresses for Jabalpur (Madhya Pradesh) and Khopoli (Maharashtra).
- **Why Human is Required:** Physical gate dispatch addresses and dispatch pin codes must be confirmed with plant logistics management.
- **Human Action Required:**
  Verify plant dispatch postal codes in `MASTERS` sheet table `tblPlants`.

---

### 6. Competitor Pricing Intelligence
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Includes competitive benchmark calculation and price floor protection models.
- **Why Human is Required:** Market intelligence regarding competitor quotes is dynamic per tender/project.
- **Human Action Required:**
  Commercial estimators must populate latest competitor intelligence into the `COMPETITORS` sheet on a per-quote basis.

---

### 7. Multi-User Physical Workstation LAN/WAN Trial
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Phase 27 verified 50 concurrent allocations, 50 simultaneous snapshots, and lock contention handling via automated simulated multi-threaded processes on a single host.
- **Why Human is Required:** Real-world SMB file locking across distinct physical PC endpoints, differing network latencies, and Windows network redirectors must be verified with human operators.
- **Human Action Required:**
  Conduct a 2-day pilot trial with 3-5 sales estimators working from separate physical PCs against the enterprise network share.

---

### 8. Parallel Quoting Trial (3 to 5 Real Historic Enquiries)
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Verified against Unipro golden quotation (7,250 MTR, ₹178,657,900 Legacy / ₹178,801,050 Option 1 Normal) with 0 variance.
- **Why Human is Required:** Validating 3–5 diverse historical commercial enquiries across different diameters, coatings, and delivery distances ensures end-user confidence before deprecating legacy spreadsheets.
- **Human Action Required:**
  Run 3–5 recent real commercial enquiries in parallel through both the new quotation system and legacy sheets, cross-checking line rates and final offer totals.

---

### 9. Visual Approval of Layout & Offer Letter Wording
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** Offer letter reproduces standard terms and layout extracted from `MKTGENQ26-27-594_-_Unipro_CS.xlsx`.
- **Why Human is Required:** Commercial legal wording, signature blocks, corporate logos, and formal disclaimer wording require visual inspection and sign-off by Marketing & Legal leadership.
- **Human Action Required:**
  Executive review of generated PDF offer letters and customer Excel files.

---

### 10. Training & Final Production Deployment Decision
- **Status:** `NEEDS_HUMAN`
- **Current Behavior in RC:** System is packaged as `Mega_Pipes_Quotation_v1.0.0_RC.xlsm` in `RELEASE_CANDIDATE\`. DEV remains frozen. LIVE remains untouched.
- **Why Human is Required:** Authorizing live production cutover, retiring legacy spreadsheets, and conducting sales user training are management executive responsibilities.
- **Human Action Required:**
  Commercial Leadership conducts training sessions and issues formal Go-Live Cutover Authorization.
