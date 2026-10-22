# MEGA PIPES QUOTATION SYSTEM — CENTRALIZED EXPORT & LOCAL COPY ARCHITECTURE

**Document Version:** 1.0  
**Effective Date:** 2026-10-16  
**Status:** Validated in Pilot Environment (`DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`)  
**Applicable Modules:** `VBA\Modules\modExport.bas`, `VBA\Modules\modUI.bas`, `VBA\Modules\modLog.bas`  

---

## 1. Executive Summary

The Mega Pipes Quotation System provides a dual-distribution export architecture for customer-facing Excel workbooks (`.xlsx`). When a quotation is exported, the system:
1. Generates a fully sanitized, values-only customer workbook in a secure local staging folder.
2. Validates the staged file before distribution (verifying approved sheets, zero formulas, zero defined names, and absence of internal commercial secrets).
3. Copies the staged file to a centralized enterprise UNC network export directory and verifies byte-for-byte binary integrity.
4. Generates a second verified copy in a configured or user-selected local directory.
5. In the event of network disruption or share unavailability, safely preserves the staged file in a persistent local pending archive (`EXPORTS\Pending_Central_Archive\`) and provides a decoupled retry workflow without regenerating or duplicating the quotation.

---

## 2. Directory Resolution Hierarchy

### 2.1 Central UNC Export Directory (`GetCentralExportsDir`)
The system resolves the centralized corporate repository using the following strict priority sequence:
1. **Explicit Parameter Override:** `overrideCentralFolder` (used for automated test isolation and explicit automation hooks).
2. **Settings Worksheet (`tblSettings`):** Setting key `CentralExportFolder` on worksheet `SETTINGS` (e.g., `\\fileserver\MegaPipes\Exports`).
3. **Environment Variable:** `MEGA_CENTRAL_EXPORTS` (set at User or System environment level).
4. **Shared Environment Variable:** `Environ$("MEGA_SHARED_FOLDER") & "\Exports\"` (inherits from core multi-user shared root).
5. **Standalone Fallback:** `<ProjectRoot>\EXPORTS\Central\` (ensures non-blocking standalone development and local pilot testing).

### 2.2 Local User Copy Directory (`GetLocalExportsDir`)
The local user destination is resolved via:
1. **Explicit Parameter Override:** `overrideLocalFolder` (used for automated test suites).
2. **Interactive Folder Picker:** When interactive (`Application.UserControl` and `Application.DisplayAlerts` are True), opens `Application.FileDialog(msoFileDialogFolderPicker)` allowing user selection. If cancelled, local copy is safely marked `SKIPPED` without invalidating the central archive.
3. **Settings Worksheet (`tblSettings`):** Setting key `LocalExportFolder` on worksheet `SETTINGS`.
4. **User Downloads Folder:** `Environ$("USERPROFILE") & "\Downloads\"`.
5. **Standalone Fallback:** `<ProjectRoot>\EXPORTS\Local\`.

---

## 3. Staging & Pre-Flight Validation Engine

### 3.1 Local Staging Directory
- **Path:** `%TEMP%\MegaPipes_Staging\`
- **Rationale:** Customer workbooks are never assembled directly over network shares or in final user directories. Local staging guarantees fast, uninterrupted generation, atomic disk operations, and zero risk of partially written files on remote servers.

### 3.2 Pre-Flight Validation Protocol (`ValidateStagedCustomerExcel`)
Before any distribution or copying occurs, the staged file is inspected via a background Excel instance:
1. **File Existence & Minimum Size:** Verifies file exists on disk and `FileLen >= 500` bytes.
2. **Worksheet Whitelist:** Exactly 2 worksheets permitted: `OFFER` and `Synopsis`. All other sheets (`QUOTE`, `QUOTE_LINES`, `COSTING`, `SETTINGS`, `QUOTE_LOG`, `PRICING_MASTER`) are strictly prohibited.
3. **Values-Only Guarantee:** Iterates through every cell in `UsedRange` of both sheets; any cell where `HasFormula = True` triggers an immediate validation failure and halts distribution.
4. **Defined Names Scrubbing:** Verifies `Names.Count = 0`. All internal workbook names, scoped formulas, and print ranges are stripped.
5. **Confidentiality Keyword Scan:** Scans text across all cells for internal metadata keywords (e.g., `tblSettings`, `tblQuoteLines`, `DefaultSteelRate`, `MinMarginPct`, `UndercutRsPerMT`, `ManualFC`).

---

## 4. Verification Protocol (Byte-for-Byte Binary Match)

### 4.1 Verification Method (`FilesMatch`)
1. **Size Verification:** Confirms `FileLen(Source) = FileLen(Destination)` and `FileLen > 0`.
2. **Binary Chunk Comparison:** Opens both files in binary access mode (`Open For Binary Access Read`) and reads in 64 KB buffers (`65,536 bytes`).
3. **Byte-Level Identity:** Compares byte arrays chunk-by-chunk. If any byte differs or read lengths diverge, returns `False`.
4. **Mathematical Identity:** Provides 100% cryptographic certainty that destination files are exact replicas of the validated staged file, without relying on external utilities or third-party DLLs.

---

## 5. Collision Prevention & Filename Conventions

### 5.1 Naming Scheme
`[QuoteID]_[Customer]_Customer_[YYYYMMDD_HHNNSS].xlsx`
- Example: `MEGA_OFFER_26-27_594_UNIPRO_Customer_20261016_104316.xlsx`

### 5.2 Same-Second & Duplicate Collision Resolution (`ResolveCollisionFilename`)
If a file with the generated name already exists in the destination:
- Appends `_1.xlsx`, `_2.xlsx`, `_3.xlsx` incrementally.
- Example: `MEGA_OFFER_26-27_594_UNIPRO_Customer_20261016_104316_1.xlsx`
- **Zero-Overwrite Guarantee:** Under no circumstances is an existing file overwritten or truncated.

---

## 6. Fault Tolerance & Decoupled Retry Workflows

### 6.1 Central Storage Unavailable / Access Denied
1. Staged file is preserved safely in `<ProjectRoot>\EXPORTS\Pending_Central_Archive\`.
2. A companion `.meta` file is created containing:
   - `QuoteID`
   - `OriginalFilename`
   - `Timestamp`
   - `User`
   - `TargetCentralFolder`
   - `LastError`
3. Central status is reported as `PENDING` (or `FAILED` with exact OS error).
4. Local user copy proceeds independently and is verified.
5. Overall export status is reported as `PARTIAL`.
6. Logged in `tblQuoteLog` with Status `PENDING`.

### 6.2 Retry Pending Central Uploads (`RetryPendingCentralExports`)
- Scans `Pending_Central_Archive\` for queued files.
- Tests central folder accessibility. If online and accessible, copies each file, performs byte-for-byte binary verification (`FilesMatch`), logs to `tblQuoteLog`, and deletes the staged file and `.meta` file upon verified completion.
- Available via macro `modExport.RetryPendingCentralExports()` and UI button/sub `modUI.UI_RetryCentralExport()`.

### 6.3 Local Copy Failure / Retry Workflow (`RetryLocalExport`)
- If central archiving succeeds but local copy fails (e.g. invalid local path, permission denial, or user cancelled):
- Central archive is never invalidated.
- The user is prompted to select a new local directory.
- `RetryLocalExport(centralPath, newLocalFolder)` copies directly from the verified central archive to the new local destination and verifies the copy byte-for-byte.
- **Sequence Number Decoupling:** Local retry never re-executes pricing, never generates a new quotation, and never increments enquiry numbers.

---

## 7. Security, Confidentiality & Audit Logging

### 7.1 No Embedded Credentials
In strict compliance with corporate security standards, no network passwords, UNC credentials, or sheet passwords are hardcoded in VBA source code. Central authentication relies on standard Windows Kerberos/NTLM network share permissions.

### 7.2 Audit Trail (`tblQuoteLog`)
Every export event is recorded in `tblQuoteLog` with:
- `LogID`: Incremental unique ID (`LOG-00001`).
- `QuoteID`: Quotation enquiry identifier.
- `Timestamp`: Precise date and time.
- `User`: Operator username via `modLog.GetCurrentUser()`.
- `Action`: `EXPORT_CUSTOMER_CENTRAL`, `EXPORT_CUSTOMER_LOCAL`, `EXPORT_CUSTOMER_CENTRAL_RETRY`, `EXPORT_CUSTOMER_LOCAL_RETRY`.
- `Status`: `SUCCESS`, `PENDING`, `FAILED`, `SKIPPED`.
- `Notes`: Full destination path, file size in bytes, and binary verification status (`BinaryMatch: PASS`). Secrets are automatically masked via `MaskSecrets()`.
