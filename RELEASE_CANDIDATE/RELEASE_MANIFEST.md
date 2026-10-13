# Mega Pipes Quotation System — Release Candidate Manifest

**Package Version:** v1.0.0-RC  
**Release Date:** 2026-10-13  
**Target Environment:** Windows 10 / Windows 11 / Windows Server, Microsoft Excel 2016 / 2019 / 2021 / Office 365 (32-bit & 64-bit VBA 7.x supported)  
**Git Baseline Commit:** `bfbf626` (Branch `master`, Tag `p28-performance`)  
**Target Tag:** `p29-release-candidate`  
**Security Boundary:** Development Release Staging (LIVE Workbook untouched)

---

## 1. Package File Inventory and Cryptographic Integrity

All files included in this release candidate package are listed below with their cryptographic SHA-256 checksums:

| File Name | Size (Bytes) | SHA-256 Checksum | Description |
|---|:---:|---|---|
| `Mega_Pipes_Quotation_v1.0.0_RC.xlsm` | 492,484 | `5DD5A3DF763970B43A478C2CC6729F264C73506B856BC0F6E6840A0B2B39FA1A` | Core Macro-Enabled Quotation Workbook (Frozen RC) |
| `RELEASE_MANIFEST.md` | — | Documented Herein | Comprehensive Release Manifest & Integrity Bill of Materials |
| `INSTALLATION_AND_CONFIGURATION_GUIDE.md` | 6,162 | `4A0A420667EEB60C4D86FF18D6C06A85DC47874A2684324949CFB53F84FC3207` | Administrator Installation, Trust Center & Network Setup Guide |
| `OPERATOR_USER_MANUAL.md` | 7,579 | `C2541EEF58C29D1A5867B8A830792142EFDD4FB86B7D74F056C35AEC90D186B1` | Commercial Estimator & Sales Operator Standard Operating Procedure |
| `BACKUP_RESTORE_ROLLBACK_PROCEDURES.md` | 5,679 | `DEB92E3A630E1B614A41729B6BEEADF8DFCCC261D2D0629716FE1CA592F4568C` | Operational Backup, Disaster Recovery & Rapid Rollback Procedures |
| `KNOWN_LIMITATIONS_AND_HUMAN_TASKS.md` | 7,518 | `E2EE0F80E6A5510CECE2E96F32EB94D2BCD588516CF6F59447036D9200BCDEDA` | Catalog of Human-Action Preconditions (`NEEDS_HUMAN`) |

---

## 2. Packaged System Architecture & Components

The workbook contains the complete suite of verified modules, sheets, tables, and automation rules:

### 2.1 VBA Modules Included (18 Modules)
1. `modMath`: Decimal arithmetic routines, `RoundUpD` (away-from-zero rounding via `CDec`), hex color decoders, `MassMtPerM` unit calculations.
2. `modPricing`: Core commercial calculation engine, `LineInput`, `RateInput`, `LineResult`, Normal vs Legacy costing algorithms, 3LPE & internal coating calculation with MT denominator alignment.
3. `modFreight`: Multi-modal freight calculator (Method A: Fixed rate/MT, Method B: Per-trailer trip rate with stacking geometry, Method C: ORS distance routing).
4. `modPlants`: Plant eligibility evaluation (Jabalpur, Khopoli), nearest plant locator, automatic plant comparison matrix, plant override authorization with reason logging.
5. `modORS`: OpenRouteService integration, geocoding chain, distance matrix API calls, persistent distance cache (`tblDistanceCache`), and graceful manual distance fallback.
6. `modCompetitors`: Competitor landed price model, competitive price comparison matrix, minimum floor price margin protection, recommended offer pricing.
7. `modValidate`: Multi-tier input validation, geometric range limits, steel grade suitability, coating temperature bounds.
8. `modQuote`: Quotation worksheet management, line addition/modification/deletion, batch calculation (`CalcAllLines`), real-time executive summary updates.
9. `modQuoteHeader`: Quotation header lifecycle, enquiry number allocation, click debouncing protection, quote reopening number preservation.
10. `modExport`: Automated export engine, 1-page PDF offer generation, sanitised customer Excel export (strictly OFFER + Synopsis, zero formulas, zero margin/floor leakage), internal Excel export.
11. `modLog`: Centralized audit logging (`tblQuoteLog`), user identity resolution (`GetCurrentUser`), recursive event protection, credential/secret masking (`[REDACTED]`).
12. `modSequence`: Sequential enquiry numbering engine, atomic lock protocol (`sequence.lock`), offline fallback to `TEMP-...`, and online reconciliation (`ReconcileTempQuote`).
13. `modApproval`: Commercial approval state machine (`Draft` -> `Calculated` -> `Review` -> `Approved` -> `Issued`), transition guards, password-protected approval, modification-triggered invalidation.
14. `modProtection`: Sheet structure protection, selective cell unlocking (user inputs editable, calculated cells locked), macro background execution (`UserInterfaceOnly:=True`).
15. `modAdmin`: Administrative maintenance, master table synchronisation (`Masters_Central.xlsx`), competitor price merge rules (newer date wins), master request logging.
16. `modSelfTest`: Comprehensive unified self-test engine (`RunAllSelfTests`, `SelfTest`), 23 automated and human-cataloged component diagnostics.
17. `modNavigation`: Seamless workbook navigation, sheet activation buttons, jump menus.
18. `modSetup`: Workbook initialisation, seed tables setup, theme formatting.

### 2.2 Worksheet Inventory (9 Worksheets)
1. `START`: Executive landing dashboard, entry navigation buttons, system status, sync timestamps.
2. `QUOTE`: Primary interactive quotation interface, enquiry metadata, line items entry table (`tblQuoteLines`), live totals summary.
3. `COSTING`: Comprehensive engineering cost breakdown, reference layout, detailed row-by-row cost elements, freight analysis.
4. `STACKING`: Pipe stacking geometry calculator, trailer utilization matrix, nested packing rules.
5. `OFFER`: Formal commercial quotation letter, reference terms & conditions, validity dates, payment terms, delivery terms.
6. `COMPETITORS`: Competitive market intelligence matrix, competitor landed pricing, recommended offer benchmarking.
7. `SETTINGS`: Global master configuration parameters, freight tariffs, raw material rates, margin rules.
8. `MASTERS`: Standard pipe sizes catalog, steel grades, coating systems, approved plant capabilities.
9. `QUOTE_LOG`: Tamper-evident audit trail table (`tblQuoteLog`), recording timestamp, user, action, field, old value, and new value.

---

## 3. Verified Golden Contract Baselines

The release candidate has undergone regression testing against the authoritative golden baselines:

1. **Option 1 Normal-Mode Contract Baseline (Authoritatively Approved):**
   - Line 1 (18" OD x 7.92mm WT, 3LPE, Jabalpur): **₹22,689 / m**
   - Line 2 (20" OD x 7.92mm WT, 3LPE, Jabalpur): **₹24,109 / m**
   - Line 3 (22" OD x 7.92mm WT, 3LPE, Jabalpur): **₹24,857 / m**
   - Line 4 (24" OD x 7.92mm WT, 3LPE, Jabalpur): **₹25,581 / m**
   - **Normal Mode Grand Total:** **₹178,801,050**
   - **Variance vs Reference:** **0.00 INR (0.00%)**

2. **Legacy-Mode Contract Baseline (Unipro Reference Reproduction):**
   - Reference File: `MKTGENQ26-27-594_-_Unipro_CS.xlsx` (Total Length: 7,250 MTR)
   - Line 1: **₹22,681 / m**
   - Line 2: **₹23,914 / m**
   - Line 3: **₹24,748 / m**
   - Line 4: **₹25,580 / m**
   - **Legacy Mode Grand Total:** **₹178,657,900**
   - **Variance vs Reference:** **0.00 INR (0.00%)**

3. **Performance Metrics:**
   - Workbook Open: **0.2398 seconds** (Target < 5.0s)
   - Quotation Recalculation: **0.0180 seconds** (Target < 5.0s)
   - Snapshot Write: **0.6081 seconds** (Target < 5.0s)
   - Snapshot Reload: **0.9168 seconds** (Target < 5.0s)
   - Unified Self-Test Suite: **7.0100 seconds** (Target < 10.0s)

---

## 4. Dependencies & Prerequisites

- **Host Operating System:** Microsoft Windows 10, Windows 11, or Windows Server 2016/2019/2022.
- **Spreadsheet Software:** Microsoft Excel 2016, 2019, 2021, or Microsoft 365 Desktop Edition.
- **VBA Environment:** Visual Basic for Applications 7.1 with Microsoft Scripting Runtime (`scrrun.dll`).
- **Network File Share:** SMB 3.0+ UNC share for multi-user enquiry numbering and centralized quote storage.
- **Internet Connectivity:** Optional for OpenRouteService live distance calculations; fully operational offline using local distance cache.
