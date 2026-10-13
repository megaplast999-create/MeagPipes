# Multi-User and Shared-Storage Validation Report (P27)

**Execution Date:** 2026-10-13  
**Simulated Storage Root:** `C:\Users\ADMINI~1\AppData\Local\Temp\megapipes_p27_shared_op_2zfzu`  
**DEV Workbook:** `DEV\Mega_Pipes_Quotation_DEV.xlsm`  
**Automated Concurrency Result:** **ALL TESTS PASS (0 COLLISIONS, 0 OVERWRITES)**

---

## 1. Concurrency Verification Results

| Scenario | Threads / Workers | Target Resource | Outcome | Detail |
|---|:---:|---|:---:|---|
| **Concurrent Number Allocation** | 50 threads | `Counter\enquiry_26-27.txt` | **PASS** | 50 completely distinct enquiry numbers allocated (`MKTG/ENQ/26-27/700` through `749`). 0 duplicates. |
| **Simultaneous Snapshot Saves** | 50 workers | `Quotes\26-27\` | **PASS** | 50 independent `.xlsx` snapshots saved without collision using `<Enquiry>_R0_<User>.xlsx` naming convention. |
| **Per-User Append-Only Indexing** | 5 users | `Index\Index_<User>.csv` | **PASS** | 10 records per user file (`Alice`, `Bob`, `Charlie`, `David`, `Emma`). Zero lock contention between users. |
| **Lock Contention & Retry Exhaustion** | 1 worker | Artificial lock file | **PASS** | Retried 25 times at 20ms intervals; gracefully fell back to `TEMP-<timestamp>-<User>` without throwing unhandled exceptions. |
| **Stale Lock Recovery** | 1 worker | Released lock file | **PASS** | Subsequent worker immediately resumed official sequential numbering (`MKTG/ENQ/26-27/750`). |
| **Real Desktop Excel VBA Macros** | Single-user COM | DEV Workbook | **PASS** | `RunSequenceSelfTests` = PASS, `RunSnapshotSelfTests` = PASS, `RunMastersSelfTests` = PASS. |

---

## 2. Environment Distinction: Simulated Concurrency vs Real Network Trial

> [!IMPORTANT]
> The automated tests above verify algorithmic correctness, retry backoff, atomic increments, filename partitioning, and TEMP fallback under high local concurrency (50 threads).
> 
> However, **a real physical SMB network trial across multiple separate workstations and actual human estimators** cannot be automated and remains categorized as **NEEDS_HUMAN** in `TESTS/NEEDS_HUMAN.md`.

### Critical SMB / Network File System Prerequisites for Production:
1. **Mandatory Byte-Range Locking**: Requires standard Microsoft SMB 3.0+ file sharing.
2. **Cloud Storage Incompatibility**: OneDrive, Dropbox, Box, Google Drive, and SharePoint sync clients do **NOT** support real-time byte locking. Running the shared counter on synchronized cloud folders will cause duplicate numbers and sync conflicts.
3. **UNC Path Configuration**: Production shared folder must be deployed to a UNC path (e.g. `\\server\megapipes\shared`) and configured in `tblSettings` or environment variable `MEGA_SHARED_FOLDER`.
