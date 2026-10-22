# MEGA PIPES QUOTATION SYSTEM — CONTROLLED SALES-USER PILOT PACKAGE
**Audited Baseline Commit:** `e1b201c`  
**Release Tag:** `v2.0.0-multi-size-synopsis`  
**Target Environment:** DEV Only (`DEV\Mega_Pipes_Quotation_DEV.xlsm`)  
**Status:** Ready for Pilot Execution  

---

## 1. DEV BUILD IDENTITY & PRE-PILOT INTEGRITY VERIFICATION

| Verification Item | Specification / Expected Value | Pilot Verification Value | Status |
|---|---|---|:---:|
| **Target Workbook File** | `DEV\Mega_Pipes_Quotation_DEV.xlsm` | `DEV\Mega_Pipes_Quotation_DEV.xlsm` | [X] |
| **File Size** | 518,243 bytes | 518,243 bytes | [X] |
| **SHA-256 Checksum** | `BEC676B07ABE7786A50D1F824B2B5F834B4A7BC0E2B91CAF42D1038CB3EB3005` | `BEC676B07ABE7786A50D1F824B2B5F834B4A7BC0E2B91CAF42D1038CB3EB3005` | [X] |
| **Audited Git Commit** | `e1b201c36eb3d05477e4df1bc9d914cb8e1d6ed6` (+ UI Repair) | Source tree synchronized & compiled | [X] |
| **Pre-Implementation Backup** | `BACKUP\Mega_Pipes_Quotation_DEV_pre_ui_repair_20261015_155238.xlsm` | Present in `BACKUP/` | [X] |
| **Pre-Pilot Safety Copy** | `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm` | Generated as dedicated pilot copy | [X] |
| **LIVE Isolation Gate** | `LIVE\` directory completely untouched and unreferenced | Confirmed untouched | [X] |

---

## 2. PILOT-MACHINE CONFIGURATION & PERMISSIONS

| Requirement | Minimum Standard | Validation Instruction | Sign-Off |
|---|---|---|:---:|
| **Operating System** | Windows 10 or 11 (64-bit) | Run `winver` | [ ] |
| **Microsoft Excel** | Excel 2016, 2019, 2021, or Microsoft 365 (Desktop App) | File > Account > About Excel | [ ] |
| **Macro Security** | "Disable VBA macros with notification" or "Enable all macros" | Trust Center > Macro Settings | [ ] |
| **VBA Object Model** | "Trust access to the VBA project object model" ENABLED | Trust Center > Macro Settings | [ ] |
| **Trusted Location** | Pilot folder added to Excel Trusted Locations | Trust Center > Trusted Locations | [ ] |
| **Filesystem Permissions** | Full Read/Write/Delete on workbook directory and subfolders | Verify write access to `Quotes/` and `Exports/` | [ ] |
| **Display Resolution** | 1920 x 1080 (100% or 125% DPI scale recommended) | Display Settings > Scale & layout | [ ] |
| **Print Subsystem** | Microsoft Print to PDF or local printer driver installed | Devices and Printers | [ ] |

---

## 3. SHARED-FOLDER & DOCUMENT NUMBERING CHECKS

| Check Item | Test Procedure | Expected Result | Result |
|---|---|---|:---:|
| **Local Standalone Numbering** | Open quote without `MEGA_SHARED_FOLDER` configured. Click **NEW QUOTATION**. | Generates official sequence (e.g. `MKTG/ENQ/26-27/NNN`) from `tblSettings`. | [ ] |
| **Reopening Idempotency** | Open an existing quotation (e.g. `MKTG/ENQ/26-27/594`). Verify cell `B4`. | Number `MKTG/ENQ/26-27/594` is retained. Counter in `tblSettings` does NOT increment. | [ ] |
| **Click Debounce Protection** | Double-click **NEW QUOTATION** rapidly (<1 second). | Exactly ONE number is allocated; second click is ignored. | [ ] |
| **Shared Folder Network Lock** | Configure `MEGA_SHARED_FOLDER` environment variable to shared UNC path. | Lock file `sequence.lock` created during allocation, held <50ms, then released cleanly. | [ ] |
| **Lock Contention Handling** | Simulate concurrent access or file lock on `sequence.lock`. | System performs up to 20 retries (50ms pauses); if lock persists, falls back gracefully to `TEMP-YYYYMMDD-USER-NNNN`. | [ ] |

---

## 4. OFFLINE & LIVE ROUTING BEHAVIOR

| Feature | Condition | Expected System Behavior | Result |
|---|---|---|:---:|
| **Cached Road Distance** | PIN present in `tblDistanceCache` | Instantaneous distance retrieval (<5ms); zero network calls. | [ ] |
| **Live ORS Network Routing** | `ORS_API_KEY` present & valid | Queries OpenRouteService API; retrieves driving distance; caches into `tblDistanceCache`. | [ ] |
| **Offline / No API Key** | `ORS_API_KEY` absent or network down | System displays manual distance input prompt without throwing unhandled exceptions. | [ ] |
| **Invalid / Unreachable PIN** | PIN code not found or route blocked | Graceful fallback to nearest regional plant distance or manual prompt with audit logging. | [ ] |
| **Offline Quoting Workflow** | Laptop disconnected from office network | Generates `TEMP` quote; calculates complete costing; saves local snapshot. | [ ] |
| **Reconnection Reconciliation** | Reconnect to network and run reconcile | Converts `TEMP` identifier to official sequential number; updates cell `B4` and audit log. | [ ] |

---

## 5. MULTI-SIZE QUOTATION CREATION & EDITING

| Action | Operational Workflow | Expected Validation Result | Result |
|---|---|---|:---:|
| **Standard Pipe Addition** | Select Standard, Family, Basis (`ID`/`OD`), Size, WT, Grade, Qty, Unit. Click **+ ADD ITEM TO QUOTE**. | Line added to Section 3 table; calculates OD, Mass, Pipes, Trailers, Unit Price, and Line Total. | [ ] |
| **Custom Size Geometry** | Click **CUSTOM PIPE SIZE**. Enter non-standard OD (e.g. 1320mm) and WT (e.g. 9.0mm). Click OK. | Accepts size if $OD \in [100, 3500]$, $WT \in [3.5, 60]$, and $OD > 2 \times WT$. Rejects if $OD \le 2 \times WT$. | [ ] |
| **10-Size Capacity Limit** | Add 10 distinct pipe sizes to quotation. Attempt to add 11th size. | System cleanly blocks 11th line with error dialog: *"Quotation capacity reached: Maximum of 10 product-size lines are supported per quotation."* Table remains undamaged with 10 lines. | [ ] |
| **Line Editing & In-Place Update** | Select Line 4 in table. Click **EDIT LINE**. Modify quantity or specifications in form. Click **UPDATE LINE (L0004)**. | Form loads selected line specs; button turns amber `UPDATE LINE (Lxxxx)`; saving updates existing line in place with zero duplicates; line counter and totals recalculate immediately. | [ ] |
| **Line Duplication** | Select Line 5. Click **DUPLICATE ITEM**. | Creates new independent line with matching dimensions. Changing quantity on duplicate does NOT alter original line. | [ ] |
| **Line Removal & Renumbering** | Select Line 3 out of 5 lines. Click **- REMOVE LINE**. | Line 3 removed; remaining lines renumbered contiguously (`L0001`, `L0002`, `L0003`, `L0004`). Totals recalculate immediately. | [ ] |
| **Clear Line Inputs** | Enter partial data into Section 2 inputs. Click **CLEAR LINE INPUTS**. | Inputs `B20:B29` cleared. Existing lines in table remain untouched. | [ ] |
| **Clear All Lines** | Click **CLEAR ALL LINES**. Confirm prompt. | Wipes all table rows; resets Live Summary to 0; resets capacity counter to `0 of 10 sizes added`. | [ ] |

---

## 6. SAVE, CLOSE & REOPEN WORKFLOWS

| Step | Action | Expected Behavior | Result |
|---|---|---|:---:|
| **1. Save Draft** | On populated quotation, click **SAVE DRAFT**. | Creates immutable snapshot `Quotes/<FY>/<Enquiry>_R0_<User>.xlsx`. Logs `SAVE_QUOTE` in `tblQuoteLog`. Status displays *"Quotation saved successfully"*. | [ ] |
| **2. Active Modification** | Change quantity on Line 1. Click **CALCULATE QUOTATION**. | Grand total updates; revision increments from R0 to R1. | [ ] |
| **3. Save Revision** | Click **SAVE DRAFT**. | Creates `Quotes/<FY>/<Enquiry>_R1_<User>.xlsx`. | [ ] |
| **4. Close & Clear** | Click **RESET DRAFT** or **NEW QUOTATION**. | Console cleared to blank state. | [ ] |
| **5. Reopen via History** | Click **QUOTATION HISTORY** (or **OPEN EXISTING QUOTATION**). Select saved quote. | Reconstructs exact lines, dimensions, unit prices, and grand total with **₹0.00 delta**. | [ ] |

---

## 7. OFFER & SYNOPSIS PREVIEWS & EXPORTS

| Function | Test Action | Acceptance Criteria | Result |
|---|---|---|:---:|
| **Preview Offer** | Click **PREVIEW OFFER**. | Switches to `OFFER` sheet; displays header details, items table, 27 Unipro standard commercial terms, and Indian numbering words. Values-only (zero formulas). | [ ] |
| **Preview Synopsis** | Click **PREVIEW SYNOPSIS**. | Switches to `SYNOPSIS` sheet; displays Executive Summary, Plant Source, Cost Components, Freight Breakdown, and Margin Analysis. Replicates reference layout. | [ ] |
| **Export PDF** | Click **EXPORT PDF**. | Generates client-ready PDF under `DEV/Exports/`. Formatted to clean 1-page width; professional typography; zero chopped borders. | [ ] |
| **Export Sanitized Excel** | Click **EXPORT EXCEL**. | Generates customer-facing `.xlsx` file. **STRICT SANITIZATION AUDIT:** Contains ONLY `OFFER` and `SYNOPSIS` sheets. Does NOT contain `COSTING`, internal formulas, steel buying rates, freight slabs, or margin matrices. Values-only guarantee. | [ ] |

---

## 8. HISTORICAL QUOTATION BENCHMARK COMPARISONS

Execute these 5 real-world scenarios on the pilot copy and verify mathematical agreement:

### Case 1: Unipro Golden Reference Quotation (HSAW Coated)
- **Parameters:** Quote `MKTG/ENQ/26-27/594`, Customer `UNIPRO`, Delivery `Bagodar, Jharkhand`, Plant `P02 (Jabalpur)`.
- **Lines (7,250 MTR Total):**
  - Line 1: HSAW ID 1226mm x 8.4mm WT, Fe 410, PU 1000µ Ext / FGE 406µ Int, Qty: 2,000 MTR.
  - Line 2: HSAW ID 1300mm x 8.4mm WT, Fe 410, PU 1000µ Ext / FGE 406µ Int, Qty: 350 MTR.
  - Line 3: HSAW ID 1350mm x 8.4mm WT, Fe 410, PU 1000µ Ext / FGE 406µ Int, Qty: 500 MTR.
  - Line 4: HSAW ID 1400mm x 8.4mm WT, Fe 410, PU 1000µ Ext / FGE 406µ Int, Qty: 4,400 MTR.
- **Expected Benchmark Values:**
  - **Option 1 Normal Total:** **₹178,801,050** (L1: ₹22,689/m, L2: ₹24,109/m, L3: ₹24,857/m, L4: ₹25,581/m).
  - **Legacy Mode Total:** **₹178,657,900** (L1: ₹22,681/m, L2: ₹23,914/m, L3: ₹24,748/m, L4: ₹25,580/m).
  - *Commercial Sign-Off Item:* Unresolved difference of +₹4,400 against older baseline total ₹178,805,450.
- **Pilot Actual Result:** Grand Total: `₹______________` | Variance: `₹______` | Status: `[PASS / FAIL]`

### Case 2: Multi-Size Water Transmission Project (8 HSAW Sizes)
- **Parameters:** Customer `State Water Board`, 8 sizes ranging from 600mm ID to 1600mm ID, Total 14,200 MTR, Plant `P02`.
- **Validation Goal:** Verify system stability, live summary updates, trailer stacking calculations across 8 simultaneous sizes, and correct Grand Total summation.
- **Pilot Actual Result:** Grand Total: `₹______________` | Line Count: `8 of 10` | Status: `[PASS / FAIL]`

### Case 3: Single-Size Emergency High-Volume Dispatch
- **Parameters:** Customer `National Gas Grid`, 1 line: HSAW OD 1422mm x 12.0mm WT, Fe 500, Bare, Qty: 5,000 MTR, Plant `P01 (Ghaziabad)`.
- **Validation Goal:** Verify bare pipe calculation, Fe 500 grade extra, single-line trailer count, and prompt generation.
- **Pilot Actual Result:** Unit Price: `₹__________/m` | Grand Total: `₹______________` | Status: `[PASS / FAIL]`

### Case 4: Heavy-Wall Custom Geometry Tender (Non-Standard Sizes)
- **Parameters:** 3 custom lines: 1120mm ID x 14.2mm WT, 1370mm ID x 16.0mm WT, 1750mm ID x 18.0mm WT, Method C freight.
- **Validation Goal:** Verify `CUSTOM PIPE SIZE` input, geometric clearance checks ($OD > 2 \times WT$), and non-standard tooling rate calculation.
- **Pilot Actual Result:** Grand Total: `₹______________` | Geometry Passed: `[YES / NO]` | Status: `[PASS / FAIL]`

### Case 5: Competitive Defense Bid (Competitor C01 Active)
- **Parameters:** Customer `EPC Contractor`, HSAW 1200mm OD x 9.5mm WT, 3,000 MTR, Competitor C01 price present.
- **Validation Goal:** Verify competitor landed price calculation, floor price protection (prevents quoting below raw cost + min margin), and yellow highlight audit.
- **Pilot Actual Result:** Offer Price: `₹__________/m` | Floor Protected: `[YES / NO]` | Status: `[PASS / FAIL]`

---

## 9. SALES-USER FEEDBACK RECORDS

### Feedback Record 1: Senior Commercial Estimator
- **Tester Name:** __________________________________ | **Date:** ______________
- **Quotation Workflow Usability (1–5):** `[ ]`
- **Console Layout & Button Clarity (1–5):** `[ ]`
- **Feedback Comments:**
  __________________________________________________________________________________
  __________________________________________________________________________________
- **Identified Friction Points:**
  __________________________________________________________________________________

### Feedback Record 2: Regional Sales Manager
- **Tester Name:** __________________________________ | **Date:** ______________
- **Offer Letter & Commercial Terms Completeness (1–5):** `[ ]`
- **Sanitized Customer Excel Confidentiality (1–5):** `[ ]`
- **Feedback Comments:**
  __________________________________________________________________________________
  __________________________________________________________________________________
- **Identified Friction Points:**
  __________________________________________________________________________________

### Feedback Record 3: Plant Logistics & Dispatch Coordinator
- **Tester Name:** __________________________________ | **Date:** ______________
- **Freight Calculation & Trailer Rounding (1–5):** `[ ]`
- **Multi-Plant Comparison & Selection (1–5):** `[ ]`
- **Feedback Comments:**
  __________________________________________________________________________________
  __________________________________________________________________________________
- **Identified Friction Points:**
  __________________________________________________________________________________

### Feedback Record 4: Back-Office Quotation Clerk
- **Tester Name:** __________________________________ | **Date:** ______________
- **Quotation Search, Reload & History (1–5):** `[ ]`
- **Line Addition, Duplication & Renumbering (1–5):** `[ ]`
- **Feedback Comments:**
  __________________________________________________________________________________
  __________________________________________________________________________________
- **Identified Friction Points:**
  __________________________________________________________________________________

---

## 10. DEFECT LOGGING & ACCEPTANCE CRITERIA

### Severity Classification Guide
- **Severity 1 (Blocker):** Calculation error in pricing/totals; data corruption; application crash; failure to enforce 10-line limit; leaking confidential costing data in customer export.
- **Severity 2 (Major):** Failure of core button action (Add, Duplicate, Remove, Save, Export); incorrect dropdown values; broken document numbering.
- **Severity 3 (Minor):** Non-blocking visual defect; formatting or alignment inconsistency; confusing warning message; non-critical printer driver delay.
- **Severity 4 (Cosmetic):** Label wording preference; color hue suggestion; optional UX enhancement request.

### Pilot Defect Log Sheet

| Defect ID | Date | Reporter | Severity | Sheet / Function | Steps to Reproduce | Expected Behavior | Actual Behavior | Resolution / Retest Status |
|:---:|:---:|:---:|:---:|:---:|---|---|---|:---:|
| **D-01** | | | | | | | | |
| **D-02** | | | | | | | | |
| **D-03** | | | | | | | | |
| **D-04** | | | | | | | | |

### Pilot Exit & Acceptance Gate Criteria
- [ ] **1. Zero Severity 1 & Severity 2 Defects:** No open critical calculation or functional bugs.
- [ ] **2. 100% Golden Agreement:** Case 1 reproduces exact Golden Normal (₹178,801,050) and Legacy (₹178,657,900) totals with ₹0 delta.
- [ ] **3. Commercial Sign-Off:** Written management approval for the Option 1 calculation rule.
- [ ] **4. User Satisfaction Threshold:** Minimum average rating of $\ge 4.0 / 5.0$ across all sales tester feedback records.
- [ ] **5. Sanitization Verification:** 100% of exported customer Excel files verified free of internal costing formulas or sheets.

---

## 11. OUTSTANDING APPROVALS & OPERATIONAL PRECONDITIONS

| Approval Domain | Description / Requirement | Authority | Status | Sign-Off Date |
|---|---|---|:---:|:---:|
| **Commercial Baseline** | Authorize Option 1 Normal total ₹178,801,050 (Line 4 = ₹25,581/m) vs historical ₹178,805,450. | Commercial Director / VP Sales | `[PENDING]` | |
| **Network Infrastructure** | Provision corporate shared UNC path (`MEGA_SHARED_FOLDER`) with read/write access for all sales users. | IT Infrastructure Lead | `[PENDING]` | |
| **Routing API Key** | Provision production `ORS_API_KEY` for live road distance calculation (or approve cached-only mode). | IT / Commercial Operations | `[PENDING]` | |
| **Security Passwords** | Securely inject `MEGA_SHEET_PWD` and `MEGA_ADMIN_PWD` on all production workstation environments. | Systems Administrator | `[PENDING]` | |
| **Production Cutover** | Formal authorization to copy audited DEV build into LIVE production directory. | Executive Sponsor | `[PENDING]` | |

---

## 12. PILOT GO / NO-GO DECISION TEMPLATE

```text
================================================================================
                    MEGA PIPES QUOTATION SYSTEM v2.0
                  CONTROLLED PILOT GATE DECISION RECORD
================================================================================

PILOT EVALUATION DATE: _______________________
AUDITED DEV BUILD:     e1b201c (SHA256: 8278673E96C14F7B89488207E9DE7E2042F1EC518C83997451D6D49C1B38B78D)

CRITERIA EVALUATION:
1. Critical Defects (Sev 1 / Sev 2):     [ ] 0 Open       [ ] Fail (Count: ____)
2. Golden Baseline Verification:         [ ] 100% Match   [ ] Mismatch Detected
3. Sales-User Usability Score:           [ ] >= 4.0 / 5.0 [ ] Below Threshold (Score: ____)
4. Customer Export Sanitization:         [ ] Verified 0 Leakages [ ] Leaks Found
5. Commercial Sign-Off on Option 1:      [ ] Approved     [ ] Pending Sign-Off
6. Infrastructure Preconditions (UNC/PWD):[ ] Ready       [ ] Incomplete

FINAL DECISION:
[ ] GO          — PROCEED TO CONTROLLED PRODUCTION DEPLOYMENT
[ ] CONDITIONAL — PROCEED SUBJECT TO COMPLETION OF ACTION ITEMS:
                  1. ___________________________________________________________
                  2. ___________________________________________________________
[ ] NO-GO       — REMAIN IN DEV; RESOLVE OUTSTANDING DEFECTS AND RE-AUDIT

DECISION AUTHORITIES:

Commercial Operations Lead:  ____________________  Signature: ________________  Date: ______
Head of Sales / Estimations: ____________________  Signature: ________________  Date: ______
Lead QA & Release Auditor:   ____________________  Signature: ________________  Date: ______
================================================================================
```
