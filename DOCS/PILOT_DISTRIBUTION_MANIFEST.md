# MEGA PIPES QUOTATION SYSTEM — PILOT DISTRIBUTION MANIFEST
**Date:** 2026-10-15  
**Audited Baseline Commit:** `e1b201c36eb3d05477e4df1bc9d914cb8e1d6ed6` (+ UI Repair)  
**Distribution Target:** Controlled Sales-User Pilot Only  
**Operational Scope:** DEV Environment Only (`DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`)  

---

## 1. Workbook Integrity & Hash Verification

The pilot copy has been generated directly from the provisionally accepted DEV baseline (`DEV\Mega_Pipes_Quotation_DEV.xlsm`). Both copies have identical sizes and checksums:

| Metric | Target Specification | Pilot Verification Value | Verification Result |
|---|---|---|:---:|
| **Pilot Copy Path** | `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm` | `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm` | **VERIFIED** |
| **Accepted Baseline** | `DEV\Mega_Pipes_Quotation_DEV.xlsm` | `DEV\Mega_Pipes_Quotation_DEV.xlsm` | **VERIFIED** |
| **Pre-Repair Backup** | `BACKUP\Mega_Pipes_Quotation_DEV_pre_ui_repair_20261015_155238.xlsm` | In `BACKUP/` (517,375 bytes) | **VERIFIED** |
| **Workbook File Size** | 518,243 bytes | 518,243 bytes | **EXACT MATCH** |
| **SHA-256 Checksum** | `BEC676B07ABE7786A50D1F824B2B5F834B4A7BC0E2B91CAF42D1038CB3EB3005` | `BEC676B07ABE7786A50D1F824B2B5F834B4A7BC0E2B91CAF42D1038CB3EB3005` | **EXACT MATCH** |
| **VBA Compilation** | Desktop Excel VBE CommandBar ID 578 | Clean compile (Control disabled / clean) | **PASS** |

---

## 2. Visual Layout & Export Evidence Verification

The 5-page acceptance document (`TESTS\quote_sheet_ui_acceptance.pdf`) and its extracted high-resolution images (`TESTS\acceptance_page_1.png` through `5.png`) have been verified against all layout and export criteria:

- **Single-Screen Landscape Geometry:** Print area `$A$1:$I$35` configured with `FitToPagesWide = 1` and `FitToPagesTall = 1`.
- **Zero Shape / Cell Overlap:** Columns D (Pipe Spec, width 30) and E (Coating Ext/Int, width 23, left-aligned) provide complete separation. No text bleeding or label collisions.
- **Genuine Currency Formatting:** Cell formatting utilizes `ChrW$(8377)` (`[$₹-4009]#,##0`), eliminating ANSI/UTF-8 character corruption (`â,¹`).
- **Dynamic Total Row Positioning:** Positioned directly below active lines (`outRow = 21 + actLines`).
- **Interactive Visual Feedback:** Selected row highlighted in soft blue (`RGB(220, 238, 255)`); Capacity console displays real-time line occupancy.

---

## 3. Pilot Checklist Coverage Verification

The operational workflows are fully mapped in `DOCS\CONTROLLED_PILOT_CHECKLIST.md`:

1. **Standard Pipe Addition (Section 5):** Validates input form mapping to Section 3 table and line-item calculation.
2. **Custom Pipe Geometry (Section 5):** Physical constraints verified ($OD \in [100, 3500]$ mm, $WT \in [3.5, 60]$ mm, $OD > 2 \times WT$).
3. **10-Size Capacity Limit (Section 5):** 10 lines accepted contiguously; 11th line blocked with modal error dialog.
4. **Line Editing & In-Place Update (Section 5):** `UI_EditItem` loads selected line specs into `B21:B30`, button turns amber `UPDATE LINE (Lxxxx)`, saves in place without duplicate records.
5. **Line Duplication (Section 5):** Clones selected line to new distinct identifier (`L0005`).
6. **Line Removal & Renumbering (Section 5):** Deletes targeted line, renumbers contiguously, and dynamically repositions totals row.
7. **Form & Table Clear Behaviors (Section 5):** `CLEAR INPUTS` clears entry cells while preserving table; `CLEAR ALL LINES` wipes all lines upon confirmation.
8. **Save, Close & Reopen (Section 6):** Snapshot generation (`Quotes/<FY>/<Enquiry>_R0_<User>.xlsx`) and idempotent reopening verified with **₹0.00 delta**.

---

## 4. Preserved Test Logs & Verification Artifacts

All test runs, logs, and screenshots are preserved under `TESTS/`:

- **COM Regression Test Log:** `TESTS\unified_selftest_p24.txt` (24/24 PASS, 0 Failures)
- **Pytest Automated Suite:** 47 / 47 PASS (Execution time: 0.64s)
- **Combined PDF Package:** `TESTS\quote_sheet_ui_acceptance.pdf` (5 pages)
- **Visual Evidence Screenshots:**
  - `TESTS\acceptance_page_1.png` (Empty quote state)
  - `TESTS\acceptance_page_2.png` (4 distinct sizes, golden total ₹17,88,01,050)
  - `TESTS\acceptance_page_3.png` (Full 10-line capacity & maximum badge)
  - `TESTS\acceptance_page_4.png` (Two-way edit workflow & amber update button)
  - `TESTS\acceptance_page_5.png` (Recalculated reference quote)

---

## 5. Outstanding Deployment Prerequisites (NEEDS_HUMAN)

The following items are documented in `TESTS\NEEDS_HUMAN.md` as external environment dependencies required solely before final production cutover:

1. **Shared Network UNC Folder (`MEGA_SHARED_FOLDER`):**
   - Corporate shared network folder UNC path required for multi-machine concurrent sequence locking.
   - *Pilot Status:* Local sequence generation operates in standalone mode.
2. **Production Passwords (`MEGA_SHEET_PWD` & `MEGA_ADMIN_PWD`):**
   - Cryptographically strong passwords required for final worksheet structure lock and back-office master data lock.
   - *Pilot Status:* Unlocked in DEV and Pilot copies to allow testing.

---

## 6. Security & Safety Attestation

- **DEV Environment Only:** All implementation and validation occurred strictly within `DEV\`.
- **LIVE Isolation:** The `LIVE\` directory is untouched, empty, and completely unreferenced.
- **Production Gate:** Production cutover is **NOT authorized**.
- **Pilot Package Location:**  
  `C:\Users\Administrator\Desktop\Mega Pipes\Quotation System\DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`
