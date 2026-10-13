# Mega Pipes Quotation System — Backup, Restore & Rollback Procedures

**Document Code:** SOP-IT-DR-002  
**Target Audience:** Enterprise IT Operations, System Administrators, Commercial Custodians  
**Version:** v1.0.0-RC

---

## 1. Operational Backup Strategy

To guarantee zero data loss and business continuity, the Mega Pipes Quotation System implements multi-tier backup redundancy across local user workstations and corporate network shares.

### 1.1 Automated Snapshot Redundancy
- **Immutable Snapshot Files:** Whenever an operator clicks **Save Snapshot**, the system generates an immutable workbook stored at:
  `\\fileserver\MegaPipes\QuotationShared\Quotes\<FiscalYear>\<EnquiryNumber>_R<Revision>_<Username>.xlsx`
- **Tamper-Resistance:** Snapshots are written as standalone, self-contained Excel workbooks containing complete frozen line items, plant selections, and commercial calculations.
- **Audit Logging:** Every snapshot creation event is appended to `tblQuoteLog` with user ID, timestamp, and quotation checksum.

### 1.2 Enterprise Scheduled Backup Policy
IT Operations must configure automated Windows Volume Shadow Copy (VSS) or file server backup tasks targeting the shared directory:

| Backup Type | Target Path | Frequency | Retention Period |
|---|---|---|---|
| **Incremental Snapshot** | `\\fileserver\...\QuotationShared\Quotes\` | Every 60 Minutes (during business hours) | 30 Days |
| **Daily Full Backup** | Entire `QuotationShared\` share | Nightly at 23:00 | 1 Year |
| **Monthly Archive** | Master tables & completed quote snapshots | Last day of month | 7 Years (statutory audit requirement) |

---

## 2. Disaster Recovery Scenarios & Resolution Protocols

### Scenario A: Stale Sequence Lock File (`sequence.lock`)
**Symptom:** Users receive error: `"Enquiry sequence file is locked by another user. Retrying..."` which eventually times out.  
**Root Cause:** A user workstation crashed, lost network connection, or was forcibly terminated while holding the sequential lock file.  
**Resolution Protocol:**
1. Check the timestamp of `\\fileserver\...\Sequence\sequence.lock`.
2. If the lock file is older than 5 minutes (indicating an orphaned lock):
   ```powershell
   # IT Administrator clears stale lock file
   Remove-Item -Path "\\fileserver\MegaPipes\QuotationShared\Sequence\sequence.lock" -Force
   ```
3. The system's built-in timeout mechanism automatically falls back to temporary enquiry numbering (`TEMP-...`) if lock resolution exceeds 30 seconds, preventing user workflow stoppage.

### Scenario B: Network Disconnection During Quotation (Offline Recovery)
**Symptom:** Estimator loses network connection to the shared drive while creating a quotation.  
**System Behavior:**
1. The workbook automatically transitions to **Offline Mode**.
2. An offline quote number is allocated: `TEMP-YYYYMMDD-XXX`.
3. The user can complete all pricing calculations, engineering selections, and client reviews without interruption.
**Reconciliation Protocol upon Network Restoration:**
1. Once network connectivity is restored, open the offline quote workbook.
2. Click **Reconcile Offline Quote** (or run `ReconcileTempQuote`).
3. The system claims the next official sequence number, updates Cell `B4`, updates the audit log (`tblQuoteLog`), and saves the permanent snapshot.

### Scenario C: Workbook Corruption or Accidental Modification
**Symptom:** User inadvertently alters worksheet layouts or damages Excel binary structure.  
**Resolution Protocol:**
1. Discard the corrupted workbook file.
2. Copy a fresh, clean instance of `Mega_Pipes_Quotation_v1.0.0_RC.xlsm` from the protected software repository.
3. Open the new workbook.
4. If an existing quote needs to be restored: navigate to **QUOTE**, enter the Enquiry Number, and click **Load Snapshot**. The entire line item table and commercial configuration will be cleanly repopulated from the network snapshot archive.

---

## 3. Production Rollback Procedure

If a critical flaw, calculation discrepancy, or blocking incompatibility is discovered after deploying the release to live operations, follow this emergency rollback procedure.

### Step 1: Emergency Freeze & Notification
1. **Broadcast Notice:** Notify commercial estimators to stop entering new quotations in the current release.
2. **Access Revocation:** Change NTFS permissions on `Mega_Pipes_Quotation_v1.0.0_RC.xlsm` to **Read-Only** or move the deployment file to an quarantine directory to prevent further execution.

### Step 2: Data Preservation & Rescue
Before replacing the application:
1. Export a complete copy of the shared directory:
   ```powershell
   Copy-Item -Path "\\fileserver\MegaPipes\QuotationShared" -Destination "C:\Emergency_Archive\QuotationShared_PreRollback_$(Get-Date -Format 'yyyyMMdd_HHmmss')" -Recurse
   ```
2. Verify that all completed snapshots in `Quotes\<FY>\` generated during the active window are intact.

### Step 3: Reinstating Previous Baseline
1. Identify the last verified baseline version:
   - For v1.0.0-RC rollback: reinstate previous stable release or legacy calculation workbook (`MKTGENQ26-27-594_-_Unipro_CS.xlsx`).
2. Restore the previous application binary to the distribution share.
3. Notify users that the system has been rolled back and provide guidance for manual entry of any quotes originated during the affected period.

### Step 4: Root Cause Analysis & Post-Incident Review
1. Retrieve log entries from `tblQuoteLog` and runtime error logs.
2. File an urgent incident report detailing:
   - Triggering incident / error manifestation.
   - Affected quotes and monetary delta.
   - Remediation steps taken and required engineering patches.
