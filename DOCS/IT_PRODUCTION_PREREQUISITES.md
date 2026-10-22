# MEGA PIPES QUOTATION SYSTEM — IT PRODUCTION PREREQUISITES & INFRASTRUCTURE READINESS GUIDE

**Document Reference:** `DOCS/IT_PRODUCTION_PREREQUISITES.md`  
**Preparation Date:** 21 October 2026  
**System Target:** Mega Pipes Quotation System  
**Audience:** Enterprise IT Systems Administrators, Infrastructure Engineers, Commercial Operations  
**Status:** `PENDING IT IMPLEMENTATION & VERIFICATION`  

---

## 1. Executive Summary & Purpose

This document provides the authoritative technical specification, security audit, and deployment checklist for the production infrastructure required by the Mega Pipes Quotation System.

The application is architected to operate in two distinct modes:
1. **Local Standalone Mode (DEV / Fallback):** Operates on an isolated single workstation without external infrastructure. Uses local directory trees (`QUOTES\`, `Sequence\`, `EXPORTS\`) and local workbook settings. This mode is 100% verified by automated test suites.
2. **Production Multi-User Enterprise Mode (Target LIVE):** Operates across multiple networked workstations connected to an enterprise corporate SMB/UNC file share. Requires centralized sequence allocation, atomic file locking, centralized export storage, and cryptographic environment-variable protection.

In strict compliance with project safety rules, **no corporate UNC paths, production passwords, or credentials have been invented or applied**. These prerequisites must be provisioned, configured, and verified by authorized enterprise IT staff.

---

## 2. Infrastructure Environment Variables Specification

The system queries three primary Windows environment variables for security and storage (along with two secondary integration variables). Below is their exact specification, module call sites, and missing-value fallback behaviors:

| Setting / Variable Name | Security Scope | Reading Procedures / Code Modules | Primary Purpose & Control Scope | Missing / Unset Fallback Behavior |
| :--- | :--- | :--- | :--- | :--- |
| `MEGA_SHARED_FOLDER` | IT Infrastructure (UNC Path) | `modSequence.AllocateEnquiryNumber`<br>`modAdmin.GetSharedFolder`<br>`modExport.GetCentralExportsDir`<br>`modSelfTest.RunAllSelfTests` (Line 554) | Directs the system to the enterprise network share hosting sequential counters, centralized quote snapshots, and master synchronisation indexes. | Falls back to local workstation folders (`ThisWorkbook.Path & "\QUOTES\"` and local `tblSettings["EnquirySequence"]`). Generates offline `TEMP-...` sequence numbers if locking fails. |
| `MEGA_SHEET_PWD` | Cryptographic Secret (Password) | `modProtection.SheetPassword`<br>`modProtection.ProtectSheet`<br>`modProtection.UnprotectSheet`<br>`modSelfTest.RunAllSelfTests` (Line 574) | Enforces worksheet-level protection across sensitive calculation sheets (`COSTING`, `MASTERS`, `SETTINGS`) to prevent unauthorized end-user cell editing or formula tampering. | Uses empty-string (`""`) protection. Worksheets remain structurally locked to user UI interaction via `UserInterfaceOnly:=True`, but can be unprotected without a password prompt. |
| `MEGA_ADMIN_PWD` | Cryptographic Secret (Password) | `modAdmin.AdminPassword`<br>`modApproval.TransitionQuoteState` (Line 116)<br>`modSelfTest.RunAllSelfTests` (Line 574) | Authorizes commercial management overrides (e.g. approving quotations that fall below commercial floor price, or transitioning quote state from `Review` to `Approved`). | Evaluates to empty string. In DEV/Pilot, state transitions log `"Approved (MEGA_ADMIN_PWD not configured)"` without raising an unauthorized error. |
| `MEGA_CENTRAL_EXPORTS` | IT Infrastructure (UNC Path) | `modExport.GetCentralExportsDir` (Line 1096) | Optional dedicated central repository for generated customer and management quote exports (PDF, Excel). | If unset, defaults to `MEGA_SHARED_FOLDER & "\Exports\"`; if `MEGA_SHARED_FOLDER` is also unset, falls back to `ProjectRoot & "\EXPORTS\Central\"`. |
| `ORS_API_KEY` | External Secret (API Key) | `modORS.GetORSKey` (Line 152)<br>`modSelfTest.RunAllSelfTests` (Line 542) | External REST API authentication token for OpenRouteService live road distance matrix computation. | Disables live HTTP calls; uses pre-seeded SQLite distance cache (e.g. 782 km Jabalpur, 1821 km Khopoli) and allows manual road distance override. |

---

## 3. Security Audit & Credential Hygiene Verification

A comprehensive repository audit was conducted across all VBA source modules (`VBA/Modules/*.bas`), Python test suites (`TESTS/*.py`), and documentation:

1. **Zero Hardcoded Secrets in Source Code:**
   - `modProtection.bas` (Lines 18–21): `SheetPassword()` returns `""` when the environment variable is not defined. No production password string exists in VBA code.
   - `modAdmin.bas` (Lines 13–16): `AdminPassword()` returns `""` when not configured.
   - `modORS.bas` (Lines 150–154): `GetORSKey()` reads exclusively from `Environ$("ORS_API_KEY")`. A hardcoded fallback is explicitly blocked by code guard.
2. **Zero Password Leakage in Logs:**
   - `modLog.bas` incorporates `MaskSecrets()` which redacts any password, key, or sensitive substring to `[REDACTED]` before writing audit records to `tblQuoteLog` or external diagnostic logs.
3. **Zero Secrets in Git History:**
   - Git commit history (`e1b201c` and parent commits) contains no credentials, server IPs, or corporate network shares.

---

## 4. Enterprise IT Setup Checklist for Production Deployment

IT Network Administrators must execute the following setup sequence prior to production cutover:

### Step 1: Corporate UNC Share Provisioning
- [ ] Create an enterprise network shared folder on a high-availability Windows Server or clustered NAS:
  - Example Target UNC: `\\corp-fileserver.megapipes.local\QuotationShared\`
- [ ] Establish standard directory structure on the share:
  ```text
  \\corp-fileserver.megapipes.local\QuotationShared\
  ├── Counter\          (Sequential text counter files: enquiry_26-27.txt)
  ├── Quotes\           (Historical immutable quote snapshots .xlsx organized by FY)
  ├── Masters\          (Central master synchronization workbook: Masters_Central.xlsx)
  ├── Exports\          (Centralized customer PDFs and sanitized Excel sheets)
  └── Logs\             (Consolidated multi-user audit logging)
  ```

### Step 2: NTFS & SMB Share Permissions
Configure strict access control lists (ACLs):
- [ ] **Quoting Sales Estimators (`SG_MegaPipes_Estimators`):**
  - Share Permission: `Change` (Read/Write)
  - NTFS Permissions:
    - `Counter\`: `Read`, `Write`, `Modify` (Required for exclusive byte-range locking and counter increment)
    - `Quotes\`: `Read`, `Write` (Create new snapshot files; deny deletion to enforce immutability)
    - `Exports\`: `Read`, `Write` (Save PDF/Excel offers)
    - `Masters\`: `Read Only` (Estimators cannot alter master coil/coating/freight rates)
- [ ] **Commercial Management & IT Admins (`SG_MegaPipes_Admins`):**
  - Share Permission: `Full Control`
  - NTFS Permissions: `Full Control` across all subdirectories

### Step 3: Workstation Environment Variable Deployment
Deploy environment variables across all authorized sales workstation PCs via Active Directory Group Policy Object (GPO) or Microsoft Intune:

```powershell
# Execute on each target client workstation (User or System scope)
[System.Environment]::SetEnvironmentVariable("MEGA_SHARED_FOLDER", "\\corp-fileserver.megapipes.local\QuotationShared", "Machine")
[System.Environment]::SetEnvironmentVariable("MEGA_SHEET_PWD", "<SECURE_ENTERPRISE_PASSWORD>", "Machine")
[System.Environment]::SetEnvironmentVariable("MEGA_ADMIN_PWD", "<COMMERCIAL_DIRECTOR_PASSWORD>", "Machine")
[System.Environment]::SetEnvironmentVariable("MEGA_CENTRAL_EXPORTS", "\\corp-fileserver.megapipes.local\QuotationShared\Exports", "Machine")
```

### Step 4: Client Excel Security & Trust Center Configuration
- [ ] Add the network share and local client installation directories to Excel Trusted Locations (`Trust Center > Trusted Locations > Add new location`, enable "Allow Trusted Locations on my network").
- [ ] Ensure `VBA Macro Settings` are configured to `"Disable VBA macros with notification"` or digitally signed by Mega Pipes Enterprise CA.
- [ ] Configure default client printer to ensure "Microsoft Print to PDF" is operational.

---

## 5. Multi-Machine Concurrency & Network Share Test Protocol

Prior to production cutover, IT and QA must execute a physical multi-workstation test using two authorized test client PCs (Client A and Client B) mapped to a staging UNC network share:

### Test Hardware & Staging Setup
- **Staging Share:** `\\test-server\MegaPipes_TestShare\`
- **Workstation A:** Host PC 1 (Sales Estimator 1)
- **Workstation B:** Host PC 2 (Sales Estimator 2)
- **Environment:** Both workstations configured with `MEGA_SHARED_FOLDER=\\test-server\MegaPipes_TestShare`

### Verification Test Cases
1. **TC-NET-01: Concurrent Number Allocation (Contention & Atomic Locking)**
   - *Action:* Client A and Client B simultaneously trigger `AllocateEnquiryNumber`.
   - *Verification:* The exclusive lock file (`\Counter\enquiry_26-27.txt`) allows one process to read and increment the serial (e.g. `595 -> 596`), while the second workstation pauses (`Sleep 50ms`), retries, and allocates the subsequent number (`597`).
   - *Pass Criteria:* Zero duplicate enquiry numbers generated; counter file reflects contiguous incrementation without file corruption.
2. **TC-NET-02: Offline Recovery & TEMP Reconciliation**
   - *Action:* Disconnect network interface on Client A. Click `AllocateEnquiryNumber`.
   - *Verification:* Client A assigns `MKTG/ENQ/26-27/TEMP-<timestamp>`. Reconnect network and trigger `ReconcileTempQuote`.
   - *Pass Criteria:* System reconciles the temporary draft to the next official sequential number from the central share and records `OFFLINE_RECONCILE` in `tblQuoteLog`.
3. **TC-NET-03: Centralized Snapshot Isolation**
   - *Action:* Client A creates and saves Quote `MKTG/ENQ/26-27/596`. Client B queries `LoadQuote` for Quote `596`.
   - *Pass Criteria:* Client B loads the snapshot file from the central network share identically, verifying multi-user read access.
4. **TC-NET-04: Centralized Export File Deposition**
   - *Action:* Generate customer offer PDF and sanitized customer Excel from Client A.
   - *Pass Criteria:* Files are deposited simultaneously into client's local Downloads folder AND centralized share `\Exports\` folder with identical checksums.

---

## 6. Distinguishing Local Fallback vs. Genuine Network Tests

| Capability | Local Fallback (DEV Environment) | Genuine Network Share (Production Target) | Current Verification Status |
| :--- | :--- | :--- | :--- |
| **Storage Mechanism** | Single-user local disk path (`ThisWorkbook.Path`) | Multi-user SMB/UNC network share (`\\server\share\`) | Local: **VERIFIED (PASS)**<br>Network: **PENDING IT DEPLOYMENT** |
| **Numbering Concurrency** | Single process thread incrementing `tblSettings` | Exclusive OS byte-range lock (`Open ... Lock Read Write`) | Local: **VERIFIED (PASS)**<br>Network: **PENDING IT TRIAL** |
| **Concurrency Simulation** | 50 concurrent simulated threads on single machine (`test_p16_multiuser_concurrency.py`) | Real physical SMB file redirector across independent physical network interfaces | Simulation: **VERIFIED (PASS)**<br>Physical SMB: **PENDING IT TRIAL** |
| **Worksheet Protection** | Open structure with empty-string bypass (`ENABLE_SHEET_PROTECTION = False`) | Full structural lockdown with `MEGA_SHEET_PWD` hash | Bypass: **VERIFIED (PASS)**<br>Lockdown: **PENDING IT PASSWORDS** |
| **Admin Escalation** | State transition note: `"Approved (MEGA_ADMIN_PWD not configured)"` | Strict password authentication prompt requiring Commercial Director password | Note: **VERIFIED (PASS)**<br>Auth Gate: **PENDING IT PASSWORDS** |

---

## 7. IT Sign-Off Gate & Release Blockers

The following items are **MANDATORY RELEASE BLOCKERS** and must remain marked pending until certified by IT:

1. **UNC Shared Folder Provisioning:** Pending IT ticket resolution for network share creation and ACL assignment.
2. **Group Policy Environment Variable Push:** Pending distribution of `MEGA_SHARED_FOLDER`, `MEGA_SHEET_PWD`, and `MEGA_ADMIN_PWD` to client PCs.
3. **Physical 2-Workstation Concurrency Sign-Off:** Pending execution of TC-NET-01 through TC-NET-04 on physical hardware.
