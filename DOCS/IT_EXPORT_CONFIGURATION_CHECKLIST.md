# MEGA PIPES QUOTATION SYSTEM — IT CENTRAL EXPORT CONFIGURATION CHECKLIST

**Document Version:** 1.0  
**Target Audience:** Enterprise IT Infrastructure, Windows Network Administrators, Systems Engineers  
**Date:** 2026-10-16  
**System:** Mega Pipes Quotation System — Centralized Customer Export Engine  

---

## 1. Overview
The Mega Pipes Quotation System exports sanitized customer-facing Excel quotes (`.xlsx`) directly to a centralized corporate UNC network repository while providing a verified local copy to the user. This document details the exact network infrastructure, permissions, and environment configurations required from Corporate IT.

---

## 2. IT Infrastructure Requirements

### 2.1 File Server UNC Share Setup
- **Recommended UNC Path:** `\\<FileServer>\MegaPipes\Exports\` (or `\\<FileServer>\MegaPipes\QuotationShared\Exports\`)
- **Underlying Protocol:** SMB 3.0 or higher recommended.
- **Latency / Bandwidth:** Recommended LAN/VPN latency < 50ms.
- **Storage Sizing:** Each customer `.xlsx` file is approximately 14 KB to 30 KB.
  - Sizing guidance: ~10,000 quotes/year ≈ 300 MB/year. Quota of 5 GB to 10 GB is sufficient for multi-year operations.

### 2.2 NTFS File & Share Permissions (ACLs)

| Principal / Security Group | Share Permission | NTFS File Permission | Purpose |
|---|---|---|---|
| **MegaPipes-QuotingUsers** (Sales Estimators) | Change & Read | Modify, Read & Execute, List Folder Contents, Read, Write | Export quotation workbooks, read back for verification. |
| **MegaPipes-Admins** (Pricing Managers / IT) | Full Control | Full Control | Manage archive structure, run audits, purge obsolete archives. |
| **Everyone / Domain Users** | None | No Access | Block unauthenticated access to commercial quotes. |

> [!IMPORTANT]
> The quoting user's Windows process writes a temporary test token (`.perm_test_<time>_<rnd>.tmp`) during pre-flight checks and immediately deletes it. Ensure that the **Delete** permission is granted on the folder so pre-flight validation succeeds.

---

## 3. Configuration Methods (Select One)

Corporate IT can configure the central export location using any of the following three options, evaluated in order of precedence:

### Option A: `tblSettings` in Workbook (Application Level)
- Open `SETTINGS` worksheet in the quotation workbook.
- In `tblSettings`, locate row:
  - `SettingKey`: `CentralExportFolder`
  - `SettingValue`: `\\fileserver\MegaPipes\Exports\`
- Save the workbook. This path is used across all workstations using that master workbook.

### Option B: Dedicated Environment Variable (Recommended for Enterprise GPO)
Deploy the environment variable across sales laptops/workstations via Group Policy Object (GPO) or PowerShell:
```powershell
[System.Environment]::SetEnvironmentVariable(
    "MEGA_CENTRAL_EXPORTS",
    "\\fileserver\MegaPipes\Exports",
    "Machine"
)
```

### Option C: Multi-User Shared Root
If `MEGA_SHARED_FOLDER` is already configured for sequence numbering and master synchronization (e.g. `\\fileserver\MegaPipes\QuotationShared`), the export engine will automatically use the subfolder:
`\\fileserver\MegaPipes\QuotationShared\Exports\`
(Ensure the `Exports` subfolder is created on the share with appropriate NTFS write permissions).

---

## 4. Workstation Client Pre-Flight Verification

Run the following PowerShell verification script from a quoting user's workstation to confirm readiness before pilot cutover:

```powershell
# 1. Check configured environment variables
$centralVar = [System.Environment]::GetEnvironmentVariable("MEGA_CENTRAL_EXPORTS", "Machine")
if (-not $centralVar) {
    $centralVar = [System.Environment]::GetEnvironmentVariable("MEGA_CENTRAL_EXPORTS", "User")
}
Write-Host "MEGA_CENTRAL_EXPORTS Variable: $centralVar"

# 2. Test directory existence and write permissions
$targetDir = if ($centralVar) { $centralVar } else { "\\fileserver\MegaPipes\Exports" }
Write-Host "Testing access to target directory: $targetDir"

if (Test-Path $targetDir) {
    Write-Host "[PASS] Network path is reachable." -ForegroundColor Green
    $testToken = Join-Path $targetDir ".it_verify_test_token.tmp"
    try {
        [System.IO.File]::WriteAllText($testToken, "MegaPipes IT Verification Token")
        Write-Host "[PASS] Write permission verified." -ForegroundColor Green
        Remove-Item -Force $testToken
        Write-Host "[PASS] Delete permission verified." -ForegroundColor Green
    } catch {
        Write-Host "[FAIL] Write/Delete permission denied: $_" -ForegroundColor Red
    }
} else {
    Write-Host "[FAIL] Network directory not found or unreachable." -ForegroundColor Red
}
```

---

## 5. Offline Queue & IT Monitoring

If central storage is unreachable (e.g., traveling estimator without corporate VPN access):
1. The Quotation System safely stages exports locally in:
   `<ProjectRoot>\EXPORTS\Pending_Central_Archive\`
2. When the workstation reconnects to the corporate LAN or VPN, pending exports can be synchronized automatically or manually:
   - Via Excel UI: Run macro `modUI.UI_RetryCentralExport`
   - Via VBA: `modExport.RetryPendingCentralExports()`
3. **IT Audit Monitoring:** IT can monitor the pending directory size to identify workstations with prolonged offline disconnection.

---

## 6. IT Sign-Off Checklist

| Item | Requirement | Verification Method | Status | Sign-off Date |
|---|---|---|---|---|
| **1. UNC Share Created** | Share provisioned on resilient storage server | `Test-Path \\<server>\<share>` | [ ] | ____________ |
| **2. NTFS Permissions** | Modify rights granted to Quoting Users | Test write/delete file token | [ ] | ____________ |
| **3. Config Deployed** | Variable or `tblSettings` deployed | Value verified on pilot client | [ ] | ____________ |
| **4. Offline Queue Verified**| Local pending directory accessible | Folder exists under `EXPORTS\` | [ ] | ____________ |
| **5. Antivirus Exclusions** | Quoting directory excluded from file-lock scans | AV policy verified | [ ] | ____________ |
| **6. Pilot User Acceptance** | Test quote successfully exported | Binary match verified | [ ] | ____________ |
