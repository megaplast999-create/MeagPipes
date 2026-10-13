# Real-World Golden Quotation Validation Report (P26)

**Validation Date:** 2026-10-13  
**Reference Workbook:** `MKTGENQ26-27-594_-_Unipro_CS.xlsx`  
**Development Workbook:** `DEV\Mega_Pipes_Quotation_DEV.xlsm`  
**Overall Status:** **100% MATCH (0 DIFFERENCE)**

---

## 1. Legacy Contract Comparison (Reference vs DEV System)

| Line | OD (mm) | WT (mm) | Reference Qty (MTR) | DEV Qty (MTR) | Reference Price (₹/m) | DEV Price (₹/m) | Price Diff | Reference Total (₹) | DEV Total (₹) | Total Diff | Result |
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| 1 | 1243 | 8.4 | 2,000 | 2,000 | ₹22,681 | ₹22,681 | ₹0 | ₹45,362,000 | ₹45,362,000 | ₹0 | **PASS** |
| 2 | 1317 | 8.4 | 350 | 350 | ₹23,914 | ₹23,914 | ₹0 | ₹8,369,900 | ₹8,369,900 | ₹0 | **PASS** |
| 3 | 1367 | 8.4 | 500 | 500 | ₹24,748 | ₹24,748 | ₹0 | ₹12,374,000 | ₹12,374,000 | ₹0 | **PASS** |
| 4 | 1417 | 8.4 | 4,400 | 4,400 | ₹25,580 | ₹25,580 | ₹0 | ₹112,552,000 | ₹112,552,000 | ₹0 | **PASS** |
| **TOTAL** | — | — | **7,250** | **7,250** | — | — | — | **₹178,657,900** | **₹178,657,900** | **₹0** | **PASS** |

---

## 2. Normal Mode Golden Comparison (Approved Option 1 Contract)

| Line | Nom ID (mm) | WT (mm) | Qty (MTR) | Approved Price (₹/m) | Calculated Price (₹/m) | Line Total (₹) | Variance | Status |
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| 1 | 1200 | 8.4 | 2,000 | ₹22,689 | ₹22,689 | ₹45,378,000 | ₹0 | **PASS** |
| 2 | 1400 | 8.4 | 350 | ₹24,109 | ₹24,109 | ₹8,438,150 | ₹0 | **PASS** |
| 3 | 1500 | 8.4 | 500 | ₹24,857 | ₹24,857 | ₹12,428,500 | ₹0 | **PASS** |
| 4 | 1600 | 8.4 | 4,400 | ₹25,581 | ₹25,581 | ₹112,556,400 | ₹0 | **PASS** |
| **TOTAL** | — | — | **7,250** | — | — | **₹178,801,050** | **₹0** | **PASS** |

---

## 3. Discrepancy & Traceability Notes

- **Legacy Contract reproduced exactly**:
  - Total Quantity = **7,250 MTR**
  - Grand Total = **₹178,657,900**
  - Line Prices = **₹22,681 / ₹23,914 / ₹24,748 / ₹25,580**
- **Normal Mode (Option 1) confirmed**:
  - Total Quantity = **7,250 MTR**
  - Grand Total = **₹178,801,050**
  - Line 4 Price = **₹25,581 / m**
  - Historical Discrepancy: The historical draft total of ₹178,805,450 differs by ₹4,400 due to earlier unrounded per-metre aggregation; Option 1 (₹178,801,050) is formally approved and locked as authoritative.
