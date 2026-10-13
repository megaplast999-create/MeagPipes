# Mega Pipes Quotation System — Multi-User Architecture Specification

## 1. System Topology & Directory Structure

Each estimator operates their own local copy of the quotation workbook (`Mega_Pipes_Quotation_DEV.xlsm`). A configured Windows SMB network folder (`SharedFolder`) coordinates shared assets, sequence allocation, master synchronisation, and quote snapshots.

```
SharedFolder\
  ├── Masters\
  │   └── Masters_Central.xlsx               — Admin-controlled authoritative master tables
  ├── Counter\
  │   └── enquiry_<FY>.txt                   — Exclusive binary-lock sequence counter file
  ├── Quotes\
  │   └── <FY>\
  │       └── <EnquiryID>_R<rev>_<User>.xlsx — Immutable values-only snapshot workbooks
  ├── Index\
  │   └── Index_<User>.csv                   — Append-only personal quote index per estimator
  ├── Requests\
  │   └── SizeRequest_<Timestamp>_<User>.csv — New pipe size / customer requests for Admin review
  └── Exports\
      └── <OfferID>_<Customer>_R<rev>.pdf    — Generated offer letters and customer workbooks
```

---

## 2. Data Partitioning & Mutability Rules

| Classification | Scope & Location | Description | Mutability |
|---|---|---|---|
| **Local Private** | Local Workbook (`QUOTE`, `QUOTE_LINES`) | In-flight quote draft, working quantities, customer selection | Read/Write by active user |
| **Local Cache** | Local Workbook (`DISTANCE_CACHE`, `SETTINGS`) | Seeded road distances, local configuration keys | Updated on demand / lookup |
| **Shared Counter** | `SharedFolder\Counter\enquiry_<FY>.txt` | Master integer sequence counter for fiscal year enquiries | Incremented under exclusive lock |
| **Central Masters** | `SharedFolder\Masters\Masters_Central.xlsx` | Authoritative steel grades, coating rates, competitor prices | Read-only for users; Admin updates |
| **Immutable Snapshots** | `SharedFolder\Quotes\<FY>\*.xlsx` | Static values-only archive of completed quotes and revisions | Strictly Read-Only once saved |
| **Append-Only Index** | `SharedFolder\Index\Index_<User>.csv` | Log of quotes issued by specific user | Append-only (zero multi-user collisions) |
| **Asynchronous Requests** | `SharedFolder\Requests\*.csv` | Pending size and customer requests submitted to admin | Created by user, archived by admin |

---

## 3. Concurrent Enquiry Numbering & Lock Protocol

Sequential enquiry numbering across concurrent users is governed by an exclusive OS-level file lock mechanism implemented in [`VBA/Modules/modSequence.bas`](file:///c:/Users/Administrator/Desktop/Mega%20Pipes/Quotation%20System/VBA/Modules/modSequence.bas):

```vba
' Exclusive lock acquisition with retry loop
Do While retries < 20
    Err.Clear
    On Error Resume Next
    f = FreeFile
    Open lockFile For Binary Access Read Write Lock Read Write As #f
    If Err.Number = 0 Then
        ' Lock acquired: read serial, increment, write back
        ...
        Close #f
        Exit Do
    Else
        retries = retries + 1
        Sleep 50 ' 50ms sub-second pause via kernel32
    End If
Loop
```

### Contention & Offline Fallback:
1. **Normal Flow**: When lock succeeds within 20 retries (<1 second total contention window), sequential number `MKTG/ENQ/<FY>/<Serial>` is assigned.
2. **Contention / Timeout Fallback**: If 20 retries are exhausted without lock acquisition, a distinct temporary identifier is assigned:
   `MKTG/ENQ/<FY>/TEMP-<Timestamp>`
3. **No Network / Offline Mode**: If `SharedFolder` is unconfigured or unavailable, local DEV sequence counter (`tblSettings.CurrentEnquirySerial`) or `TEMP` identifier is used, allowing uninterrupted quoting.

---

## 4. Collision Prevention Design

1. **Per-User File Naming**: Every saved snapshot incorporates the fiscal year, cleaned enquiry number, revision number, and Windows username:
   `<CleanEnquiryNo>_R<rev>_<User>.xlsx`
   Even if two estimators work simultaneously on related tenders, snapshots cannot overwrite or corrupt each other.
2. **Isolated Index Logs**: Rather than locking a single central index file on every quote creation, each user writes exclusively to `Index_<User>.csv`. The Admin macro `ConsolidateIndex` aggregates them asynchronously.
3. **Values-Only Snapshots**: Snapshots contain zero active formulas or external links, guaranteeing 100% numerical immutability over time regardless of future rate revisions.

---

## 5. Network Filesystem Assumptions & Limitations

- **Protocol**: Requires Microsoft SMB (Server Message Block) / CIFS file sharing that supports mandatory byte-range and file locking (`Lock Read Write`).
- **Cloud Sync Warning**: Asynchronous sync services (OneDrive, Dropbox, SharePoint sync client) do **not** support real-time byte locking; running the shared counter on such services will cause duplicated numbers or sync conflict copies. A true network file share is required for multi-user production.
- **Human Verification**: End-to-end multi-person concurrency testing across multiple physical machines is cataloged as `NEEDS_HUMAN` for Phase 27.
