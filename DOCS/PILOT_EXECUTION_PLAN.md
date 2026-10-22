# MEGA PIPES QUOTATION SYSTEM — CONTROLLED SALES PILOT EXECUTION PLAN
**Phase:** Controlled Sales Pilot — Phase 1  
**Project Root:** `C:\Users\Administrator\Desktop\Mega Pipes\Quotation System`  
**Pilot Workbook File:** `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`  
**Target Reference Checklist:** `DOCS\CONTROLLED_PILOT_CHECKLIST.md`  
**Distribution Manifest:** `DOCS\PILOT_DISTRIBUTION_MANIFEST.md`  
**Audited Baseline Commit:** `e1b201c36eb3d05477e4df1bc9d914cb8e1d6ed6` (+ UI Repair)  
**Status:** Approved for Supervised Execution  

---

## 1. PURPOSE & BOUNDARY RULES

### 1.1 Purpose
This execution plan defines the supervised testing protocol for validating the Mega Pipes Quotation System with pilot sales users. The objective is to verify operational ergonomics, multi-size quotation workflows (up to 10 sizes), in-place line editing, persistence, export fidelity, and commercial accuracy in a controlled development environment.

### 1.2 Absolute Safety Boundaries
1. **DEV Only:** Testing is conducted exclusively on `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`.
2. **Untouched Baseline:** The primary development workbook (`DEV\Mega_Pipes_Quotation_DEV.xlsm`) and all files in `BACKUP\` must remain unopened and unmodified throughout testing.
3. **LIVE Isolation Gate:** Never open, modify, reference, or deploy files to `LIVE\`. No production access or cutover is authorized.
4. **Immediate Stop Condition:** If any commercial calculation discrepancy (even ₹1), data loss, workbook crash, or access-control error occurs, stop testing immediately, log the defect, and escalate. Do not bypass errors.

---

## 2. PILOT WORKBOOK IDENTITY & HARDWARE PREREQUISITES

### 2.1 File Identity
- **File Path:** `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`
- **File Size:** `518,243 bytes`
- **SHA-256 Checksum:** `BEC676B07ABE7786A50D1F824B2B5F834B4A7BC0E2B91CAF42D1038CB3EB3005`
- **VBA Compilation:** Clean compile (CommandBar ID 578 verified)

### 2.2 System & Excel Requirements
- **OS:** Windows 10 or 11 (64-bit)
- **Excel:** Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365 Desktop
- **Display Resolution:** 1920 x 1080 (100% or 125% DPI scaling recommended)
- **Macro Security:** Macro Settings set to "Disable VBA macros with notification" or "Enable all macros"
- **Trusted Location:** Project directory added to Excel Trusted Locations (`Trust Center > Trusted Locations`)
- **Print Subsystem:** Microsoft Print to PDF installed and set as accessible

### 2.3 Environmental Prerequisites (NEEDS_HUMAN)
- **Shared Network Sequence Folder (`MEGA_SHARED_FOLDER`):**
  - *Status for Pilot:* Standalone local sequence generator active (`MKTG/ENQ/26-27/NNN`).
  - *Prerequisite for Production:* Corporate UNC path required for multi-user file locking.
- **Production Sheet/Admin Passwords (`MEGA_SHEET_PWD` / `MEGA_ADMIN_PWD`):**
  - *Status for Pilot:* Unlocked in pilot copy to allow interactive testing without password prompts.
  - *Prerequisite for Production:* Passwords required for permanent structure lockdown prior to production cutover.

---

## 3. ROLES & RESPONSIBILITIES

| Role | Assigned Personnel | Responsibilities |
|---|---|---|
| **Test Supervisor** | Senior Automation / QA Engineer | Conducts pre-flight checks, observes workflows, enforces safety gates, records timings, logs defects. |
| **Sales Pilot User** | Commercial Sales Estimator | Operates the UI as in daily business, enters enquiries, validates usability, reports ergonomics. |
| **Calculation Auditor** | Commercial / Pricing Analyst | Independently audits unit rates, freight, GST, and totals against approved reference standards. |

---

## 4. SUPERVISED TEST SEQUENCE (PHASE 1)

Each test case must be executed in order, with evidence saved to `TESTS\`:

```
[Pre-Flight Audit] ──> [P-TC01: Empty State] ──> [P-TC02: 1-Line Entry]
                              │
                              ▼
[P-TC03: 4-Line Golden] ──> [P-TC04: 10-Line Capacity] ──> [P-TC05: 11-Line Rejection]
                              │
                              ▼
[P-TC06: 2-Way Edit] ──> [P-TC07: Duplicate & Remove] ──> [P-TC08: Form/Table Clears]
                              │
                              ▼
[P-TC09: Save & Reopen] ──> [P-TC10: Exports & Layout] ──> [Session Sign-Off]
```

### P-TC01: Initial Form & Empty State Verification
- **Action:** Open `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`. Click **NEW QUOTATION**.
- **Expected Result:**
  - Header defaults populated (Customer: Unipro, Plant: P02, Mode: NORMAL).
  - Cell `D21:I21` merged with prompt: `"(No quotation lines added yet. Select product specifications on left and click '+ ADD ITEM TO QUOTE')"`.
  - Section 3 totals and Section 1 Live Summary reflect ₹0.00.
  - Capacity console (`F14:I17`) displays: `"QUOTATION CAPACITY: 0 of 10 sizes added (10 sizes remaining)"`.
- **Target Timing:** < 5 seconds.

### P-TC02: Single-Line Quotation Workflow
- **Action:** Enter standard pipe parameters into `B20:B30`:
  - Standard: IS-3589-2001 | Family: HSAW | Basis: ID | Size: 1226mm | WT: 8.4mm | Grade: Fe 410
  - External Coating: PU 1000 micron | Internal Coating: Food grade epoxy 406 micron
  - Quantity: 2,000 | Unit: Metres | Length: 12m
  - Click **+ ADD ITEM TO QUOTE**.
- **Expected Result:**
  - Line `L0001` appears in Row 21: `L0001: HSAW 1226mm (ID) x 8.4mm HSAW-002`.
  - Coating `PU 1000 micron / Food grade epoxy 406 micron` cleanly left-aligned in Column E.
  - Rate: `₹22,689` | Line Total: `₹4,53,78,000`.
  - Dynamic Table Total row appears at Row 22: `TABLE TOTAL (1 of 10 sizes): 2,000 Metres 167 pcs (511.51 MT) ₹4,53,78,000`.
  - Quantity cell `B28` resets to 0. Capacity badge increments to `1 of 10 sizes added`.
- **Target Timing:** < 30 seconds.

### P-TC03: Four-Line Golden Reference Quotation
- **Action:** Add the 3 remaining reference project sizes:
  - Line 2: 1300mm ID x 8.4mm, 350 MTR
  - Line 3: 1350mm ID x 8.4mm, 500 MTR
  - Line 4: 1400mm ID x 8.4mm, 4,400 MTR
  - Click **CALCULATE QUOTATION**.
- **Expected Result:**
  - Rows 21–24 populated contiguously with zero overlapping shapes or text bleeding.
  - Unit rates: `₹22,689` / `₹24,109` / `₹24,857` / `₹25,581`.
  - Dynamic Total row sits at Row 25: `TABLE TOTAL (4 of 10 sizes): 7,250 Metres 606 pcs (2,031.04 MT) ₹17,88,01,050`.
  - Live Summary (`D8:E10`): Subtotal = `17,88,01,050.00`, GST @ 18% = `3,21,84,189.00`, Grand Total = `21,09,85,239.00`.
  - Capacity badge: `4 of 10 sizes added (6 sizes remaining)`.
- **Target Timing:** < 60 seconds.

### P-TC04: Ten-Line Full Capacity Stress Test
- **Action:** Add 6 additional valid pipe sizes (e.g. 1050mm, 900mm, 800mm, 700mm, 600mm, 500mm).
- **Expected Result:**
  - All 10 lines display cleanly in Rows 21–30.
  - Table Total row moves smoothly to Row 31: `TABLE TOTAL (10 of 10 sizes)`.
  - Capacity badge displays amber/red alert: `"QUOTATION CAPACITY: 10 of 10 sizes added (MAXIMUM CAPACITY REACHED)"`.
  - Zero performance lag (<500ms recalculation).
- **Target Timing:** < 90 seconds.

### P-TC05: Eleven-Line Boundary Rejection
- **Action:** Attempt to add an 11th pipe size (e.g. 400mm x 5.0mm, 100m).
- **Expected Result:**
  - Action safely blocked.
  - Modal dialog appears: *"Quotation capacity reached: Maximum of 10 product-size lines are supported per quotation."*
  - Banner alert in `A34` and `D32` turns red.
  - Quotation table remains undamaged with exactly 10 lines.
- **Target Timing:** Immediate (< 2 seconds).

### P-TC06: Interactive Two-Way Line Editing
- **Action:**
  1. Click on Row 24 (Line `L0004`: 1400mm x 8.4mm, 4400m).
  2. Click **EDIT LINE**.
  3. Verify form `B21:B30` loads Line 4's exact dimensions.
  4. Verify button transforms to amber: `UPDATE LINE (L0004)`.
  5. Change quantity in `B28` to `4500`.
  6. Click **UPDATE LINE (L0004)**.
- **Expected Result:**
  - Line `L0004` updates in place at Row 24 (`4,500 Metres`).
  - Button reverts to green `+ ADD ITEM TO QUOTE`.
  - Line count remains 10 (zero duplicate created).
  - Totals recalculate immediately.
- **Target Timing:** < 20 seconds.

### P-TC07: Line Duplication & Removal
- **Action:**
  1. Remove lines 5–10 to leave 4 lines.
  2. Select Row 22 (`L0002`: 1300mm x 8.4mm). Click **DUPLICATE**.
  3. Verify Line `L0005` is created with identical specs.
  4. Select Row 25 (`L0005`). Click **- REMOVE**.
- **Expected Result:**
  - Duplication creates exactly one new line; total row updates to 5 sizes.
  - Removal prompts confirmation dialog; upon confirmation, Line 5 is removed.
  - Remaining lines renumber contiguously; Total row repositions to Row 25.
- **Target Timing:** < 30 seconds.

### P-TC08: Form and Table Clear Operations
- **Action:**
  1. Enter test quantity `9999` in `B28`. Click **CLEAR INPUTS**.
  2. Click **CLEAR ALL LINES**. Confirm dialog.
- **Expected Result:**
  - `CLEAR INPUTS`: `B28` resets to 0; all 4 table lines remain intact.
  - `CLEAR ALL LINES`: All table lines removed; cell `D21` resets to empty prompt; Live Summary resets to ₹0.00; capacity badge resets to `0 of 10 sizes added`.
- **Target Timing:** < 15 seconds.

### P-TC09: Save Draft, Close & Reopen Persistence
- **Action:**
  1. Restore 4 reference lines. Click **CALCULATE QUOTATION**.
  2. Click **SAVE DRAFT**.
  3. Close Excel completely.
  4. Reopen `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`.
  5. Click **OPEN EXISTING QUOTATION** (or review active state).
- **Expected Result:**
  - Immutable snapshot created under `DEV\Quotes\26-27\<Enquiry>_R0_<User>.xlsx`.
  - Reopened workbook restores exact quotation specifications, line IDs, unit rates, and totals with **₹0.00 variance**.
- **Target Timing:** < 45 seconds.

### P-TC10: Previews, PDF & Sanitized Customer Excel Export
- **Action:**
  1. Click **PREVIEW OFFER**. Validate 27 standard Unipro terms and Indian words.
  2. Click **PREVIEW SYNOPSIS**. Validate executive summary and unrounded Columns M/P.
  3. Click **EXPORT PDF**. Inspect generated PDF.
  4. Click **EXPORT CUSTOMER EXCEL** (or **EXPORT EXCEL** on Quote screen).
  5. Select local download folder when prompted.
- **Expected Result:**
  - `OFFER` and `Synopsis` render cleanly with zero `#REF!` or formula errors.
  - PDF generated under `EXPORTS\` formatted to crisp single-page landscape width.
  - Customer Excel file contains strictly values-only `OFFER` and `Synopsis` sheets (no internal costing, formulas, or raw steel buying rates).
  - Central UNC archive and local copy are created and verified with 100% byte-for-byte binary match.
  - User receives a structured summary dialog detailing Central Archive and Local Copy status independently.
- **Target Timing:** < 45 seconds.

### P-TC11: Central UNC Archive & Local Copy Resilience (Offline & Retry)
- **Action:**
  1. Test export when corporate central share is unavailable or workstation is offline.
  2. Verify local copy succeeds while central export is reported as `PENDING`.
  3. Verify staged export is safely preserved in `<ProjectRoot>\EXPORTS\Pending_Central_Archive\`.
  4. Reconnect network / restore share access.
  5. Run **Retry Pending Central Uploads** (`modUI.UI_RetryCentralExport`).
- **Expected Result:**
  - Central failure does not block local copy; clear notification that central archiving is pending.
  - Pending file safely uploaded to central directory with binary verification once connection is restored.
  - Queue cleaned up automatically; retry logged in `tblQuoteLog`.
  - Zero quotation duplication or sequence number re-allocation.
- **Target Timing:** < 30 seconds.

---

## 5. INDEPENDENT CALCULATION AUDIT CHECKLIST

Before marking any test session complete, the Calculation Auditor must verify:

1. **Outer Diameter (OD) Ceiling Rule:**
   $$\text{For ID Basis: } OD = \lceil ID + 2 \times WT \rceil$$
   *(e.g., $1226 + 2 \times 8.4 = 1242.8 \implies 1243\text{ mm}$)*
2. **Pipe Mass per Metre:**
   $$\text{Mass (kg/m)} = (OD - WT) \times WT \times 0.0246615$$
3. **Trailers & Freight:**
   - Normal mode: Whole trailers ceiling allocation.
   - Legacy mode: Fractional trailers.
4. **GST Calculation:**
   $$\text{GST} = \text{Subtotal} \times 18\%$$
5. **Golden Totals Reconciliation:**
   - Normal Golden: **₹178,801,050.00**
   - Legacy Golden: **₹178,657,900.00**

---

## 6. ESCALATION & STOP PROTOCOL

| Trigger Condition | Immediate Action | Escalation Pathway |
|---|---|---|
| **Calculation Mismatch** | Stop testing immediately. Do not adjust settings to force match. | Log exact line, rate, and difference in `PILOT_EXECUTION_LOG.md`. Escalate to Commercial Pricing Lead. |
| **Data Loss / Overwrite** | Preserve file state. Do not save over workbook. | Capture screenshot and Excel error. Escalate to Lead QA Architect. |
| **Excel Crash / VBE Error** | Note exact user action and input value. | Check Windows Event Log. Restore from `Mega_Pipes_Quotation_DEV_pre_ui_repair_20261015_155238.xlsm`. |
| **Access Control / Password** | Do not attempt brute force. | Confirm environment is pilot-unlocked. Escalate to Security Administrator. |

---

## 7. POST-PILOT ACCEPTANCE CRITERIA

Phase 1 pilot is accepted for sales-team sign-off when:
1. All 10 test cases (P-TC01 through P-TC10) are executed with **100% PASS**.
2. Sales-user qualitative rating achieves $\ge 4.0 / 5.0$ across usability categories.
3. Average time to create a 4-line quotation is under 3 minutes.
4. Client-facing exports are verified 100% free of internal sensitive pricing data.
5. All completed session scorecards are logged in `TESTS\PILOT_EXECUTION_LOG.md`.
