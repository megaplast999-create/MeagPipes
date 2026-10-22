# MEGA PIPES — UNIT PRICE DISCREPANCY & LOCATION VISIBILITY REPORT

**Document Reference:** `DOCS/UNIT_PRICE_AND_LOCATION_VISIBILITY_REMEDIATION_REPORT.md`  
**Execution Date:** 21 October 2026  
**Primary Pilot Target:** `DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm`  
**Protected Baseline:** `DEV/Mega_Pipes_Quotation_DEV.xlsm` (SHA-256: `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA`, frozen & untouched)  
**Production Status:** STRICTLY ISOLATED (LIVE directory untouched)  
**Evaluation Status:** **VERIFIED & COMPLETED**

---

## 1. Executive Summary

This investigation resolves the two core items requested for the Mega Pipes Quotation System:
1. **Workstream A (Unit Price Discrepancy — ₹22,350 vs ₹22,351):**  
   Conducted forensic tracing down to the single-paisa and integer rounding level. Proved mathematically and empirically that ₹22,350/m is **not** a software bug, calculation failure, or stale cache. It is the exact output of the **Dynamic Landed Formulaic Finance Cost** (₹1,019/MT) applied when lines are added via the user interface console without a manual financial cost override. Conversely, ₹22,351/m is the output of the **Historical Hardcoded Pre-Cost Finance Cost** (Row 32 in `CS Stacking 10-09-26`, ₹1,024/MT). This dual behavior is cataloged under **Decision Record 1** in `DOCS/COMMERCIAL_POLICY_SIGNOFF.md` (`PENDING EXECUTIVE SIGN-OFF`). Per strict safety boundary rules, both pathways are preserved and protected by dedicated regression tests (`TESTS/test_unipro_line1_discrepancy.py`).
2. **Workstream B (Plant, Destination, and Competitor Location Visibility):**  
   Completed a comprehensive gap analysis across the existing quotation interface, master data (`PLANTS`, `CUSTOMERS`, `DISTANCE_CACHE`), and geospatial datasets (`DATA/india_pincode_master.csv`). Identified that the dedicated location intelligence screen (`COMPARE`) was inaccessible from the quotation console due to missing UI navigation triggers, and lacked geographic straight-line distance, marketing territory screening (50 km radius), and explicit missing-data auditing.
3. **Workstream C (Pilot Copy Implementation & Verification):**  
   Implemented complete location and territory visibility strictly within `DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm`:
   - Populated verified plant PINs in `tblPlants` (P01 Jabalpur: `482001`, P02 Khopoli: `410203`).
   - Added ergonomic `btnQuoteCompare` ("PLANT & TERRITORY") navigation button on `QUOTE` command bar without table disruption or clipping.
   - Added `btnCompareReturn` round-trip navigation button on `COMPARE`.
   - Enhanced `COMPARE` dashboard with route metrics, straight-line Haversine distance, 50 km marketing territory radius screening, and a dedicated **Section D Location & Territory Summary**.
   - Verified metric isolation: Road distance (`driving-hgv`) is used strictly for freight costing; Geographic straight-line distance is used strictly for marketing territory screening.
   - All 148 automated Python tests and Excel COM self-tests pass with 100% green status.

---

## 2. Checksum and Integrity Register

| File / Component | Role / Target | SHA-256 Checksum | Status |
| :--- | :--- | :--- | :--- |
| `DEV/Mega_Pipes_Quotation_DEV.xlsm` | Primary Baseline (Protected) | `74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA` | **Bit-for-Bit Untouched** |
| `BACKUP/..._PILOT_COPY_pre_location_fix_...` | Pre-Change Safety Backup | `BFF824EB7BA7602EB5984A0477951854ED25C48F71EFAAD4C5D914E058DFAA2F` | **Verified Backup** |
| `DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm` | Remediated Pilot Workbook | `47AD1649D12BE2C3A8BDC20FDF5EA42AD902D168BA43FC12CF751865BD47F604` | **Updated & Verified** |
| `LIVE/` | Production Environment | *(Not Accessed)* | **Strictly Isolated** |

---

## 3. Workstream A — Forensic Root-Cause Analysis (₹22,350 vs ₹22,351)

### 3.1 Pipe Specification & Tender Context
- **Tender Reference:** `MKTG/ENQ/26-27/594` (Customer: `UNIPRO TECHNO INFRASTRUCTURE PVT LTD`)
- **Line 1 Specification:** HSAW Pipe, Basis `ID`, Diameter `1226.0 mm`, Wall Thickness `8.4 mm`, Grade `HSAW-002`
- **Quantity:** `2,000 Metres` (167 pipes @ 12.0m length)
- **Calculated Pipe Geometry & Mass:**
  $$\text{OD} = 1226.0 + (2 \times 8.4) = 1242.8 \rightarrow 1243.0\,\text{mm}$$
  $$\text{Unit Mass} = (1243.0 - 8.4) \times 8.4 \times 0.0246615 / 1000 = 0.25575635\,\text{MT/m}$$
  $$\text{Total Line Mass} = 2,000 \times 0.25575635 = 511.513\,\text{MT}$$

### 3.2 Component-by-Component Cost Breakdown & Variance Bridge

The approved rates for Unipro are:
- Coil Steel Rate: ₹62,000/MT with 2.0% yield $\rightarrow ₹63,240/\text{MT}$
- Conversion Cost: ₹4,000/MT $\rightarrow$ Raw Pre-Cost Bare: ₹67,240/MT
- External Coating (3LPE @ ₹500/m²): $500 \times 3.14 \times 1.2430 = 1,951.51\,\text{₹/m} \rightarrow ₹1,952/\text{m}$ (₹7,630.34/MT)
- Internal Coating (Liquid Epoxy @ ₹250/m²): $250 \times 3.14 \times 1.2430 = 975.76\,\text{₹/m} \rightarrow ₹976/\text{m}$ (₹3,815.19/MT)
- Freight (Khopoli to Bagodar @ 93,000/trip, 4 pipes/trailer): ₹7,676.45/MT $\rightarrow ₹1,964/\text{m}$

```
+---------------------------------------------------------------------------------------------------+
| COST COMPONENT       | HISTORICAL CS STACKING (Row 32)      | FORMULAIC LANDED FC (Row 41)       |
+----------------------+--------------------------------------+------------------------------------+
| Finance Cost Basis   | Hardcoded Pre-Cost: 1,024.00 ₹/MT    | Formulaic Landed: 1,019.00 ₹/MT   |
| Finance Derivation   | Historical manual estimator entry    | Round((Pre+Ext+Int+Fr) * 1.18 * 1%)|
| Bare Pre-Cost / MT   | 67,240 + 1,024 = 68,264.00 ₹/MT      | 67,240 + 1,019 = 68,259.00 ₹/MT    |
| Bare Price / m (raw) | 68,264.00 * 0.25575635 = 17,458.85   | 68,259.00 * 0.25575635 = 17,457.57 |
| Bare Price / m (int) | Round(17,458.85) = 17,459 ₹/m       | Round(17,457.57) = 17,458 ₹/m     |
| External Coating / m | 1,952 ₹/m                            | 1,952 ₹/m                          |
| Internal Coating / m | 976 ₹/m                              | 976 ₹/m                            |
| Freight / m          | 1,964 ₹/m                            | 1,964 ₹/m                          |
+----------------------+--------------------------------------+------------------------------------+
| TOTAL UNIT PRICE     | ₹22,351 / m                          | ₹22,350 / m                        |
| LINE TOTAL (2,000 m) | ₹44,702,000.00                       | ₹44,700,000.00                     |
| DELTA (Variance)     | Reference Baseline                   | -₹1.00 / m (-₹2,000 on ₹4.47 Cr)   |
+----------------------+--------------------------------------+------------------------------------+
```

### 3.3 Root Cause Confirmation
- When lines are populated via `modSelfTest.PopulateGoldenReferenceQuote`, the fixture passes `manualFC:=1024`, replicating the exact approved CostSheet Row 32 value and producing **₹22,351/m**.
- When an estimator enters the Unipro specification via the `QUOTE` UI Console (`+ ADD ITEM TO QUOTE`), `ManualFC` is not passed (defaults to `0`). The calculation engine evaluates the formulaic landed finance cost (`₹1,019/MT`), producing **₹22,350/m**.
- **Conclusion:** ₹22,350/m is **not** an erroneous or corrupt calculation. It is the legitimate formulaic landed rate under the dynamic finance policy cataloged in `DOCS/COMMERCIAL_POLICY_SIGNOFF.md` (Decision Record 1).
- **Policy Recommendation:** Standardize formulaic landed finance cost on all new commercial tenders, while preserving `tblQuoteLines[ManualFC]` for tender replication or pre-approved commercial credit agreements.

---

## 4. Workstream B — Location Feature Gap Analysis

### 4.1 Feature Audit & Gap Matrix

| Feature / Data Item | Master Source | Exists in Engine | Exists in UI | Identified Gap & Remediation |
| :--- | :--- | :---: | :---: | :--- |
| **Plant Master Coordinates** | `tblPlants` / CSV | Yes (`482001`, `410203`) | Missing | `PIN`, `Latitude`, `Longitude` were `None` in `tblPlants`. Populated verified coordinates. |
| **Delivery Destination PIN** | `QUOTE!B10` / `B11` | Yes (`825322`) | Partial | PIN was extracted for road lookup, but not displayed with geographic distance summary. |
| **Geographic Distance Engine**| `tools/pin_distance_engine.py` | Yes (Haversine) | Missing in VBA | VBA lacked Haversine calculation. Added `HaversineDistanceKm` and geocode lookups in `modCompare.bas`. |
| **Road vs Geo Distance Isolation** | Cache / Engine | Yes | Missing in UI | UI displayed road km only. Enhanced dashboard to display both side-by-side with strict metric separation. |
| **Marketing Territory (50 km)** | Spatial Engine | Yes | Missing in UI | Added automated `<= 50 km` vs `> 50 km` radius screening on `COMPARE` dashboard. |
| **Competitor Geocoding & Audit** | `tblPlants` (C01–C13) | Partial | Missing in UI | 11 competitors have verified PINs; C02 (Bhopal) and C03 (Anjar) missing PINs. Added explicit audit status. |
| **Quotation Screen Navigation** | `QUOTE` Sheet | No | **MISSING** | **Critical UI Gap:** Estimators had no button on `QUOTE` to access `COMPARE`. Added `btnQuoteCompare`. |
| **Dashboard Return Navigation** | `COMPARE` Sheet | No | **MISSING** | **Critical UI Gap:** Estimators had no button to return to `QUOTE`. Added `btnCompareReturn`. |

### 4.2 Architectural Screen Selection
- **Quotations Line Table (`QUOTE` Sheet):** Kept clean and focused on commercial line items, quantities, pipe geometry, coatings, unit prices, and totals. Overcrowding this screen with 13 competitor locations and spatial coordinates was rejected per User Requirement B3.
- **Dedicated Location Screen (`COMPARE` Sheet):** Positioned as the centralized **Competitor Intelligence & Marketing Territory Dashboard**, containing Section A (Our Plants), Section B (Competitors), Section C (Recommended Offer), and Section D (Location & Territory Summary).

---

## 5. Workstream C — UI Verification & Implementation Evidence

### 5.1 Verification Scenarios

1. **Unipro Reference Scenario (Bagodar 825322):**
   - Plant Jabalpur P01 (`482001`): Geographic = `572.4 km`, Road = `782 km`, Territory = `OUTSIDE 50 km`
   - Plant Khopoli P02 (`410203`): Geographic = `1,386.5 km`, Road = `1,821 km`, Territory = `OUTSIDE 50 km`
   - Competitor Ratnamani C01 (`770019`): Geographic = `272.9 km`, Road = `403 km`, Territory = `OUTSIDE 50 km`
2. **Local Marketing Territory Scenario (Within 50 km Radius):**
   - Destination: Khamaria Factory, Jabalpur (`482005`)
   - Anchor Jabalpur P01 (`482001`): Geographic = `3.3 km` ($\le 50\,\text{km}$)
   - Dashboard Evaluation: `IN MARKETING TERRITORY (3.3 km <= 50 km)` — **VERIFIED**
3. **Unresolved PIN / Missing Coordinates Scenario:**
   - Destination: Unknown PIN `999999`
   - Dashboard Evaluation: Geographic Distance = `N/A`, Status = `UNKNOWN (Destination PIN 999999 coordinates unresolved in master)` — **VERIFIED**
4. **Competitor Master Data Audit Scenario:**
   - C01 Ratnamani (`770019`): `PIN: 770019 | Geo: 272.9 km (Outside 50km) | Verified Data`
   - C02 Welspun Bhopal (`None`): `PIN MISSING | Geo: N/A | Missing Data: PIN, Lat, Lon required | Bhopal Plant`
   - C03 MAN Anjar (`None`): `PIN MISSING | Geo: N/A | Missing Data: PIN, Lat, Lon required | Anjar Plant`

### 5.2 Visual Evidence Artifacts

The following visual artifacts have been generated and archived in the project repository and artifact storage:
- `TESTS/pilot_compare_sheet_location_summary.pdf` (Rendered PNG: `pilot_compare_sheet_location_summary.png`)
- `TESTS/pilot_quote_sheet_navigation_verified.pdf` (Rendered PNG: `pilot_quote_sheet_navigation_verified.png`)

```mermaid
graph TD
    A[QUOTE Sheet Console] -->|Click 'PLANT & TERRITORY' (btnQuoteCompare)| B[COMPARE Sheet Dashboard]
    B -->|Click 'RETURN TO QUOTE' (btnCompareReturn)| A
    B --> C[Section A: Our Plants - PIN, Road km, Geo km, Freight]
    B --> D[Section B: Competitor Distances - PIN, Threat Coding, Status]
    B --> E[Section C: Recommended Offer & Floor Price]
    B --> F[Section D: Location & 50 km Marketing Territory Summary]
```

---

## 6. Outstanding Master Data Requirements

To enable 100% geocoding coverage across all competitor facilities, Commercial Operations should obtain and enter verified postal PINs for the following two competitor plants in `tblPlants`:
1. **Competitor C02 (Welspun Corp Limited - Bhopal):**  
   - Currently: `City: Bhopal`, `State: Madhya Pradesh`, `PIN: None`.  
   - Required Data: Postal PIN (e.g. `462001` for Bhopal GPO or industrial area PIN), Latitude, Longitude.
2. **Competitor C03 (MAN / JSW / Ratnamani - Anjar):**  
   - Currently: `City: Anjar`, `State: Gujarat`, `PIN: None`.  
   - Required Data: Postal PIN (e.g. `370110` for Anjar SO or industrial area PIN), Latitude, Longitude.

---

## 7. Automated Test Suite Execution Summary

```text
============================= test session starts =============================
platform win32 -- Python 3.14.8, pytest-9.1.1, pluggy-1.6.0
rootdir: C:\Users\Administrator\Desktop\Mega Pipes\Quotation System
configfile: pytest.ini
testpaths: TESTS
collected 148 items

TESTS\test_calculation_engine.py ..................                      [ 12%]
TESTS\test_commercial_user_acceptance.py ........                        [ 17%]
TESTS\test_dynamic_staged_costing.py ......                              [ 21%]
TESTS\test_location_and_territory_visibility.py ......                   [ 25%]
TESTS\test_lodha_five_line_review.py .....                               [ 29%]
TESTS\test_new_enquiry_acceptance.py .......                             [ 33%]
TESTS\test_optional_price_override_ui.py ..................              [ 45%]
TESTS\test_p16_multiuser_concurrency.py .                                [ 46%]
TESTS\test_p4_line_agreement.py ...................                      [ 59%]
TESTS\test_pin_distance_engine.py ....................                   [ 72%]
TESTS\test_pricing_golden_baselines.py ....................              [ 86%]
TESTS\test_pricing_variance_regression.py .......                        [ 91%]
TESTS\test_synopsis_precision.py .......                                 [ 95%]
TESTS\test_unipro_line1_discrepancy.py ....                              [ 98%]
TESTS\test_workbook_runtime_audit.py ..                                  [100%]

============================ 148 passed in 51.29s =============================
```

- Excel COM Self-Tests (`modCompare.RunCompareSelfTests`): **PASS** (100% green).
- Primary DEV workbook integrity: **Bit-for-Bit Untouched** (`74B5D738FB280961CF89E0312F3BE42E45557EF1B0DD2082E213B4D332EA44EA`).
- Dedicated Pilot Copy (`DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm`): **Successfully updated, tested, and verified**.
