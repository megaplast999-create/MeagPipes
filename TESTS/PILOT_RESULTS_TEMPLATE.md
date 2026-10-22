# MEGA PIPES QUOTATION SYSTEM — PILOT RESULTS SCORECARD & TEMPLATE
**Document ID:** `TESTS\PILOT_RESULTS_TEMPLATE.md`  
**Target Environment:** DEV / Pilot (`DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`)  
**Reference Plan:** `DOCS\PILOT_EXECUTION_PLAN.md`  

---

## 1. TEST SESSION METADATA

| Metadata Field | Value / Details |
|---|---|
| **Session Identifier** | `PILOT-SESS-YYYYMMDD-01` |
| **Execution Date & Time** | YYYY-MM-DD HH:MM IST |
| **Sales Pilot Operator** | [Name / Department] |
| **Test Supervisor** | [Name / QA Lead] |
| **Calculation Auditor** | [Name / Commercial Pricing] |
| **Target Workbook File** | `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm` |
| **Workbook SHA-256 Checksum** | `BEC676B07ABE7786A50D1F824B2B5F834B4A7BC0E2B91CAF42D1038CB3EB3005` |
| **Operating System** | Windows 10 / 11 64-bit (Build: _________) |
| **Microsoft Excel Version** | Excel 2016 / 2019 / 2021 / 365 (Build: _________) |
| **Display Resolution & Scaling** | 1920 x 1080 @ ____% DPI |
| **Overall Session Duration** | ____ minutes |

---

## 2. PRE-FLIGHT ENVIRONMENT & SAFETY CHECKS

| Check Item | Acceptance Standard | Verification Result | Verified By |
|---|---|:---:|:---:|
| **Target Isolation** | Workbook opened is strictly `DEV\Mega_Pipes_Quotation_PILOT_COPY.xlsm`. | [ ] PASS  [ ] FAIL | |
| **Baseline Untouched** | `DEV\Mega_Pipes_Quotation_DEV.xlsm` remains closed and unmodified. | [ ] PASS  [ ] FAIL | |
| **LIVE Isolation Gate** | `LIVE\` directory is untouched, empty, and unreferenced. | [ ] PASS  [ ] FAIL | |
| **Excel Macro Security** | Macros enabled or prompt allowed; VBA Trusted Access checked. | [ ] PASS  [ ] FAIL | |
| **Clean Initial State** | START sheet displays with zero compile or macro startup errors. | [ ] PASS  [ ] FAIL | |

---

## 3. SUPERVISED TEST EXECUTION SCORECARD

| Test ID | Scenario & Workflow | Expected Outcome | Actual Outcome | Status | Time Taken | Evidence Reference |
|:---:|---|---|---|:---:|:---:|---|
| **P-TC01** | **Empty State Verification**<br>Click *NEW QUOTATION* | Form defaults load; D21 shows empty prompt; Section 3 & Live Summary total = ₹0.00; Capacity = `0 of 10 sizes added`. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `acceptance_page_1.png` |
| **P-TC02** | **Single-Line Entry**<br>Add HSAW 1226mm ID, 8.4mm WT, PU/Epoxy, 2000m | Row 21 displays L0001; Rate = ₹22,689; Total = ₹4,53,78,000; Total row moves to Row 22; B28 resets to 0. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `acceptance_page_2.png` |
| **P-TC03** | **Four-Line Golden Quote**<br>Add 1300mm (350m), 1350mm (500m), 1400mm (4400m) | Rows 21–24 populated contiguously; Total row at D25 = **₹17,88,01,050**; Grand total = **₹21,09,85,239**; zero overlap. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `acceptance_page_2.png` |
| **P-TC04** | **10-Line Full Capacity**<br>Add 6 additional distinct pipe sizes | Rows 21–30 populated; Total row at Row 31; Capacity console = `10 of 10 sizes added (MAXIMUM CAPACITY REACHED)`. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `acceptance_page_3.png` |
| **P-TC05** | **11-Line Hard Block**<br>Attempt 11th pipe size entry | Action blocked; modal warning dialog displayed; banner turns red; table preserved with exactly 10 lines. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | Dialog Screenshot |
| **P-TC06** | **2-Way Line Editing**<br>Select Line 4, click *EDIT LINE*, change to 4500m | Form loads specs; button turns amber `UPDATE LINE (L0004)`; saves in place at Row 24; zero duplicates created. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `acceptance_page_4.png` |
| **P-TC07** | **Line Duplication & Removal**<br>Duplicate L0002 to L0005, then remove L0005 | Duplication creates L0005; removal prompts confirm dialog, deletes line, renumbers contiguously, moves total to D25. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | Row Screenshot |
| **P-TC08** | **Clear Operations**<br>Test *CLEAR INPUTS* then *CLEAR ALL LINES* | `CLEAR INPUTS` clears B28 preserving table; `CLEAR ALL LINES` wipes all lines with confirmation; resets to empty state. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | Console Screenshot |
| **P-TC09** | **Save & Reopen Persistence**<br>Save draft, close Excel, reopen workbook | Snapshot saved in `Quotes/`; reopened file reconstructs exact lines, rates, and totals with **₹0.00 delta**. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `acceptance_page_5.png` |
| **P-TC10** | **Exports & Layout Audit**<br>Preview Offer/Synopsis, Export PDF & Excel | Clean 1-page landscape PDF generated; customer Excel strictly values-only without internal formulas or buying rates. | | [ ] PASS<br>[ ] FAIL<br>[ ] BLOCK | ____s | `quote_sheet_ui_acceptance.pdf` |

---

## 4. INDEPENDENT COMMERCIAL CALCULATION SIGN-OFF

*Audited against approved reference standards for 4-line Golden Quote:*

| Item Description | Expected Baseline Rate | Pilot Recorded Rate | Difference (Delta) | Audit Sign-Off |
|---|---|---|---|:---:|
| **Line 1 (1226mm ID x 8.4mm WT - 2,000m)** | ₹22,689 / m | ₹_______________ | ₹0.00 | [ ] APPROVED |
| **Line 2 (1300mm ID x 8.4mm WT - 350m)** | ₹24,109 / m | ₹_______________ | ₹0.00 | [ ] APPROVED |
| **Line 3 (1350mm ID x 8.4mm WT - 500m)** | ₹24,857 / m | ₹_______________ | ₹0.00 | [ ] APPROVED |
| **Line 4 (1400mm ID x 8.4mm WT - 4,400m)** | ₹25,581 / m | ₹_______________ | ₹0.00 | [ ] APPROVED |
| **Subtotal (Excl. GST)** | **₹178,801,050.00** | ₹_______________ | ₹0.00 | [ ] APPROVED |
| **GST @ 18%** | **₹32,184,189.00** | ₹_______________ | ₹0.00 | [ ] APPROVED |
| **Grand Total (Incl. GST)** | **₹210,985,239.00** | ₹_______________ | ₹0.00 | [ ] APPROVED |

---

## 5. DEFECT LOGGING MATRIX

*(Record any unexpected behavior, layout anomaly, or calculation deviation)*

| Defect ID | Severity<br>*(Critical / High / Med / Low)* | Test ID | Description of Observed Behavior | Steps to Reproduce | Assigned Fix / Status |
|:---:|:---:|:---:|---|---|:---:|
| *Example: D-01* | *Low* | *P-TC02* | *Coating text slightly close to border* | *Add 406 micron epoxy* | *Resolved in modUI* |
| **D-____** | | | | | |
| **D-____** | | | | | |

---

## 6. SALES-USER QUALITATIVE FEEDBACK

*(To be completed by Sales Pilot Estimator at end of session)*

| Usability Dimension | Score (1 to 5)<br>*(1=Poor, 5=Excellent)* | Qualitative Comments / Observations |
|---|:---:|---|
| **1. Form Data Entry Speed** | [ ] 1  [ ] 2  [ ] 3  [ ] 4  [ ] 5 | |
| **2. Clarity of Side-by-Side Layout** | [ ] 1  [ ] 2  [ ] 3  [ ] 4  [ ] 5 | |
| **3. Line Selection & In-Place Editing** | [ ] 1  [ ] 2  [ ] 3  [ ] 4  [ ] 5 | |
| **4. Capacity Badge & Alert Messages** | [ ] 1  [ ] 2  [ ] 3  [ ] 4  [ ] 5 | |
| **5. PDF & Customer Excel Export Quality**| [ ] 1  [ ] 2  [ ] 3  [ ] 4  [ ] 5 | |
| **6. Overall Quotation Workflow Comfort**| [ ] 1  [ ] 2  [ ] 3  [ ] 4  [ ] 5 | |

**Sales Estimator Recommendation:**  
[ ] READY FOR EXPANDED PILOT  
[ ] MINOR ADJUSTMENTS RECOMMENDED  
[ ] BLOCKED (REMEDIATION REQUIRED)  

**Operator Signature:** __________________________________ **Date:** _______________  
**Supervisor Signature:** ________________________________ **Date:** _______________  
