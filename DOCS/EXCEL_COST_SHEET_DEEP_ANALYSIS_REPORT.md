# MEGA PIPES — EXCEL COST SHEET DEEP ANALYSIS & AUTOMATION ARCHITECTURE REPORT

**Document Ref:** `DOCS/EXCEL_COST_SHEET_DEEP_ANALYSIS_REPORT.md`  
**Date:** 2026-10-21  
**Status:** Approved for Architectural Alignment & Implementation Planning  
**Source Files Analyzed:**
1. `C:\Users\Administrator\Downloads\MKTGENQ26-27594 - Unipro.xlsx`
2. `C:\Users\Administrator\Downloads\MKTGENQ26-27-594 - Unipro 2.xlsx`
3. `C:\Users\Administrator\Downloads\MKTGENQ26-27610 - Lodha 1.xlsx`

---

## 1. Executive Summary

This comprehensive investigation provides a cell-by-cell, formula-by-formula reverse engineering of the three live commercial cost sheets provided by marketing. The objective is to define an exact, unified calculation engine and user experience that **eliminates the need for any manual formula entry** while **strictly preserving commercial flexibility for manual inputs and custom price overrides**.

### Key Discoveries:
1. **The Origin of the Golden Baseline (₹22,351 / ₹23,570 / ₹24,396 / ₹25,220):**
   - Discovered in `MKTGENQ26-27-594 - Unipro 2.xlsx` on sheet `CS Stacking 10-09-26`.
   - The Jabalpur plant was selected for supply due to lower logistics distance (782 km vs 1,821 km from Khopoli).
   - Stacking freight was modeled using un-nested trailers (151 trips total @ ₹93,000/trip ODC), producing line-specific freight costs of ₹7,676, ₹7,248, ₹6,985, and ₹6,740 per MT.
   - The official commercial offer (`Sheet1`) links directly to these numbers.
2. **The Finance Cost "Circular Loop" Mystery Solved:**
   - In both `CS NESTING10-09-2026` and `CS Stacking 10-09-26`, Row 32 contains a hardcoded manual value (e.g. ₹1,024, ₹1,014, ₹1,011, ₹1,008 in Stacking; ₹1,480, ₹1,569 in Nesting), while Row 41 calculates `=Grand_Total_Incl_GST * FC%`.
   - Estimators manually approximated the finance cost on delivered price with GST to avoid Excel circular reference warnings (`#CIRCULAR!`). Our software engine solves this cleanly in closed form or via automated pre-cost staging.
3. **Dual Sizing Inputs (ID-First vs OD-First):**
   - In Unipro, pipes are specified by **Inside Diameter (ID)**, calculating $\text{OD} = \text{ROUNDUP}(\text{ID} + 2 \times \text{WT}, 0)$.
   - In Lodha (`MKTGENQ26-27610 - Lodha 1.xlsx`), pipes are mixed: 4 lines are specified by **Outside Diameter (OD)** ($\text{ID} = \text{OD} - 2 \times \text{WT}$) and 1 line by **ID**.
   - Both modes must be seamlessly supported without user formula manipulation.
4. **Nesting vs Stacking Freight Dichotomy:**
   - **Stacking (Un-nested):** Each pipe diameter is transported in its own trailers. Larger diameters have fewer pipes per trailer (e.g. 4 pipes of 1417mm = 14.01 MT payload; 4 pipes of 1243mm = 12.28 MT payload). Therefore, smaller diameters bear higher freight per MT.
   - **Nesting (Telescoped):** Smaller pipes (e.g. 1243mm or 1317mm) are inserted inside larger pipes (1417mm or 1367mm). Total trips drop from 151 down to 103 (31.8% reduction), and freight is pooled across the project as a flat ₹5,070/MT.
5. **RMR Coating Costing:**
   - Raw Material Requirement (RMR) build-up sheets compute true chemical, abrasive, and packaging costs per square metre (₹204.98/m² internal FGE, ₹420.07/m² external FGE). Estimators then round or benchmark these to standard commercial figures (₹250/m² and ₹500/m²).

---

## 2. File-by-File & Sheet-by-Sheet Structural Analysis

### File 1: `MKTGENQ26-27594 - Unipro.xlsx`
- **Sheet Name:** `Sheet1` (Enquiry Synopsis)
- **Role:** Preliminary engineering sizing and mass take-off for Enquiry `MKTG/ENQ/26-27/594` (Konar, Jharkhand Project; Unipro Techno Infrastructure Pvt. Ltd.).
- **Content:** 4 line items of HSAW Coated pipe (Fe 410, 8.4mm WT, total 7,250 metres).
- **Core Formulas:**
  - `D12: =ROUNDUP(B12+2*E12, 0)` $\rightarrow$ Calculates OD (1243 mm) from ID (1226 mm) and WT (8.4 mm).
  - `C12: =ROUND(D12/25.4, 0)` $\rightarrow$ Nominal size in inches (49").
  - `I12: =CEILING(H12/12, 1)` $\rightarrow$ Pipe pieces count (167 Nos for 2,000m).
  - `J12: =(D12-E12)*E12*0.0246615` $\rightarrow$ Weight in kg/m ($255.756$ kg/m).
  - `K12: =H12*J12/1000` $\rightarrow$ Total weight in MT ($511.511$ MT).
  - `L12: =(D12/1000)*(D12/1000)*H12` $\rightarrow$ Bounding box volume ($3,090.10\text{ m}^3$).
  - `M12: =ROUNDUP(L12/K12, 0)` $\rightarrow$ CBM-to-weight ratio ($7\text{ m}^3/\text{MT}$).
  - `N12: =ROUNDUP(D12*H12*(22/7)/1000, 0)` $\rightarrow$ External coating area ($7,814\text{ m}^2$).
  - `O12: =ROUNDUP((22/7)*H12*(D12-(2*E12))/1000, 0)` $\rightarrow$ Internal coating area using ID ($7,708\text{ m}^2$).
  - `P12: =ROUNDUP(12*J12/1000, 0)` $\rightarrow$ Single 12m pipe weight rounded ($4$ MT).

---

### File 2: `MKTGENQ26-27-594 - Unipro 2.xlsx`
This is the master commercial tender evaluation containing 4 comprehensive sheets:

#### Sheet 1: `Synopsis`
Identical to File 1, recording the geometry and surface areas.

#### Sheet 2: `CS NESTING10-09-2026`
- **Plant Base:** Khopoli (`TO BE SUPPLIED FROM MPPL KHOPOLI`).
- **Commercial Base Rates:**
  - Steel BEN + GE: ₹61,000/MT (manual input).
  - Conversion: ₹5,000/MT for Line 1 & 4; ₹10,000/MT for Line 2 & 3.
  - Finance Cost: 1.5% with 25% Advance, Balance against LC.
  - External PU Rate: ₹500/m² (`B34`).
  - Internal FGE Rate: ₹250/m² (`B35`).
- **Freight Modeling (Nesting):**
  - Grouped into 7 trailer configurations nesting 1243mm and 1317mm inside 1417mm and 1367mm pipes.
  - Total trips = **103 trips** (`N103: =N100+N94+N89+N84+N78+N72+N66`).
  - Khopoli freight rate = ₹98,000/trip.
  - Total Freight = $103 \times 98,000 = ₹10,094,000$.
  - Freight per MT: `Q67: =Q66/G19+100` = ₹5,069.88/MT $\rightarrow$ ₹5,070/MT flat across all 4 items.
- **Resulting Prices per Metre:**
  - Item 1 (1243mm): ₹21,796/m
  - Item 2 (1317mm): ₹24,481/m
  - Item 3 (1367mm): ₹25,416/m
  - Item 4 (1417mm): ₹24,865/m
  - Total Project Value (Excl. GST): ₹17,42,74,350.

#### Sheet 3: `CS Stacking 10-09-26` (The Approved Baseline Sheet)
- **Plant Base:** Jabalpur (`TO BE SUPPLIED FROM MPPL Jabalpur`).
- **Commercial Base Rates:**
  - Steel BEN + GE: ₹62,000/MT.
  - Conversion: ₹4,000/MT across all sizes.
  - Yield Loss: 2% (`=C27*C24` = ₹1,240/MT).
  - Finance Cost: 1.0% with 25% Advance, Balance against LC.
  - External PU Rate: ₹500/m² (`B34`).
  - Internal FGE Rate: ₹250/m² (`B35`).
- **Freight Modeling (Stacking without Nesting):**
  - Rows 54–61 define trailer stacking without nesting:
    - 4 pipes per trailer ($2 + 2$ layout) for all 4 sizes.
    - Trailer payloads: 12.28 MT (Line 1), 13.01 MT (Line 2), 13.51 MT (Line 3), 14.01 MT (Line 4).
    - Total Trips = $41.67 + 7.29 + 10.42 + 91.67 = \mathbf{151.04\text{ trips}}$ (`N61`).
    - Jabalpur to Bagodar freight rate = **₹93,000 / trip** (`O57:O60`).
    - Freight per MT formula: `P57: =ROUNDUP(O57*N57/E57, 0) + 100`
      - Line 1: `ROUNDUP(93000 * 41.6667 / 511.511, 0) + 100` = ₹7,576 + 100 = **₹7,676 / MT**
      - Line 2: `ROUNDUP(93000 * 7.2917 / 94.880, 0) + 100` = ₹7,148 + 100 = **₹7,248 / MT**
      - Line 3: `ROUNDUP(93000 * 10.4167 / 140.721, 0) + 100` = ₹6,885 + 100 = **₹6,985 / MT**
      - Line 4: `ROUNDUP(93000 * 91.6667 / 1283.923, 0) + 100` = ₹6,640 + 100 = **₹6,740 / MT**
- **Per-Metre Pricing Build-up:**
  - Bare Pipe/m: `=ROUNDUP(C33*C17, 0)` $\rightarrow$ Line 1: ₹17,459
  - Ext Coating/m: `=ROUNDUP(C34*C17, 0)` $\rightarrow$ Line 1: ₹1,952
  - Int Coating/m: `=ROUNDUP(C35*C17, 0)` $\rightarrow$ Line 1: ₹976
  - Freight/m: `=ROUNDUP(C37*C17, 0)` $\rightarrow$ Line 1: ₹1,964
  - **Total Price/m:** `=SUM(C44:C47)` $\rightarrow$ **₹22,351 / m**
  - Line 2 Total Price/m: $18503 + 2068 + 1034 + 1965 =$ **₹23,570 / m**
  - Line 3 Total Price/m: $19209 + 2147 + 1074 + 1966 =$ **₹24,396 / m**
  - Line 4 Total Price/m: $19915 + 2225 + 1113 + 1967 =$ **₹25,220 / m**
  - Project Value (Excl. GST): **₹17,61,17,500**.
- **Competitor Logistics Analysis:**
  - Rows 12–28, Columns I:M compare 14 competing pipe mills. MPPL Jabalpur (782 km) is closer to Bagodar than MPPL Khopoli (1,821 km), MAN/JSW/Ratnamani Anjar (1,973 km), Welspun Bhopal (1,019 km), AMNS Hazira (1,778 km), etc. Only Ratnamani Sundargarh (403 km) is closer.
- **RMR Coating Cost Analysis:**
  - Rows 66–74: Internal FGE build-up computes chemical cost of ₹163.98/m² + 25% margin = ₹204.98/m².
  - Rows 77–87: External FGE build-up computes chemical + tie + rope cost of ₹336.06/m² + 25% margin = ₹420.07/m².

#### Sheet 4: `Sheet1` (The Commercial Offer)
- **Role:** Customer-facing commercial quotation document.
- **Linkages:**
  - `D12: ='CS Stacking 10-09-26'!C15` (1243 mm)
  - `E12: ='CS Stacking 10-09-26'!C16` (8.4 mm)
  - `F12: ='CS Stacking 10-09-26'!C18` (2000 MTR)
  - `G12: ='CS Stacking 10-09-26'!C48` (**₹22,351 / m**)
  - `H12: =G12*F12` (**₹4,47,02,000**)
  - Total Offer Value: **₹17,61,17,500**.
  - Confirms that Stacking from Jabalpur was the final approved offer presented to the client!

---

### File 3: `MKTGENQ26-27610 - Lodha 1.xlsx`
- **Sheet Name:** `Sheet1` (Enquiry Synopsis)
- **Role:** Technical synopsis for Lodha Enquiry `MKTG/ENQ/26-27/610` (MMR Region, 25,700 metres, 8,057 MT).
- **Core Differences & Lessons:**
  1. **OD-Specified Lines:**
     - Line 1: OD is entered as 1050 mm, WT 12 mm.
       `B12: =D12-2*E12` $\rightarrow$ ID is calculated as 1026 mm.
     - Line 3: OD is entered as 900 mm, WT 12 mm $\rightarrow$ ID = 876 mm (`B14: =D14-2*E14`).
     - Line 4: OD is entered as 700 mm, WT 12 mm $\rightarrow$ ID = 676 mm (`B15: =D15-2*E15`).
     - Line 5: OD is entered as 1000 mm, WT 12 mm $\rightarrow$ ID = 976 mm (`B16: =D16-2*E16`).
  2. **ID-Specified Lines:**
     - Line 2: ID is entered as 1050 mm, WT 12 mm.
       `D13: =B13+2*E13` $\rightarrow$ OD is calculated as 1074 mm.
  3. **Multi-Segment Metres:**
     - Line 2 Metres: `H13: =10700+6600+255+7480` = 25,035 m (estimators combine multiple pipeline stretches directly in the quantity cell).
  4. **Rounding Variations:**
     - Inch column: `=D12/25.4` (exact decimal 41.34" rather than integer 41").
     - Pipe piece count: `=H12/12` (exact decimal 30.42 Nos rather than `CEILING`).
     - Pipe weight: `=(12*J12/1000)` (exact decimal 3.686 MT rather than `ROUNDUP`).

---

## 3. Complete Master Formula Catalogue

| # | Parameter / Field | Source / Master Excel Formula | Purpose & Engineering Notes |
|---|---|---|---|
| **1** | **OD from ID** | `=ROUNDUP(ID + 2*WT, 0)` | Calculates OD when customer specifies ID. (Can also be exact `=ID + 2*WT`). |
| **2** | **ID from OD** | `=OD - 2*WT` | Calculates ID when customer specifies OD. |
| **3** | **Nominal Inch** | `=ROUND(OD/25.4, 0)` or `=OD/25.4` | Converts mm to nominal pipe inches. |
| **4** | **Pipe Mass (kg/m)** | `=(OD - WT) * WT * 0.0246615` | Standard IS 3589 / API 5L steel pipe weight formula. |
| **5** | **Pipe Mass (MT/m)** | `=Mass_kg_per_m / 1000` | Used for per-metre and per-tonne cross-conversions. |
| **6** | **Total Line Weight (MT)** | `=MTRS * Mass_MT_per_m` | Total steel tonnage for raw material planning. |
| **7** | **Pipe Piece Count (Nos)** | `=CEILING(MTRS/12, 1)` or `=MTRS/12` | Total standard 12-metre pipes required. |
| **8** | **Single Pipe Weight (MT)**| `=Mass_MT_per_m * 12` | Weight of one 12m pipe for trailer loading checks. |
| **9** | **Coating Factor ($m^2/m$)**| `=3.14 * (OD / 1000)` | Circumferential surface area per linear metre. |
| **10**| **Ext Coating Area ($m^2$)**| `=ROUNDUP(OD * MTRS * (22/7) / 1000, 0)` | Total outside surface area to be coated. |
| **11**| **Int Coating Area ($m^2$)**| `=ROUNDUP((22/7) * MTRS * ID / 1000, 0)` | Total inside lining surface area. |
| **12**| **Delivered Steel Cost** | `=Steel_BEN_GE + Grade_Extra` | Raw steel base price landed at mill. |
| **13**| **Yield / Scrap Loss Cost** | `=Delivered_Steel * Yield_Loss_Pct` | Cost of coil trimming, scrap, and scale loss (default 2%). |
| **14**| **Finance Cost (Pre-Cost)**| Hardcoded in Excel (e.g. ₹1,024); closed-form in engine | Cost of financing deferred payment / LC terms. |
| **15**| **Bare Pipe Price (₹/MT)**| `=Steel + Yield + Conversion + Testing + Margin + FC` | Total bare pipe manufacturing cost per tonne. |
| **16**| **Ext Coating Cost (₹/MT)**| `=($B$34 * Coating_Factor) / Mass_MT_per_m` | External coating rate converted to ₹/MT. |
| **17**| **Int Coating Cost (₹/MT)**| `=($B$35 * Coating_Factor) / Mass_MT_per_m` | Internal coating rate converted to ₹/MT. |
| **18**| **Ex-Mill Price (₹/MT)** | `=Bare_Price_MT + Ext_Cost_MT + Int_Cost_MT` | Total price at factory gate. |
| **19**| **Freight (Stacking ₹/MT)**| `=ROUNDUP(Trip_Rate * Trips / Total_MT, 0) + 100` | ODC freight per tonne including ₹100 handling. |
| **20**| **Freight (Nesting ₹/MT)**| `=ROUNDUP(Total_Nesting_Freight / Project_MT, 0) + 100` | Blended freight across nested trailer groups. |
| **21**| **Subtotal / Delivered** | `=Ex_Mill_Price_MT + Freight_MT` | Landed price before taxes. |
| **22**| **GST @ 18%** | `=Subtotal_MT * 18%` | Statutory tax per tonne. |
| **23**| **Grand Total (Incl GST)** | `=Subtotal_MT + GST_MT` | Total customer billing rate per tonne. |
| **24**| **Bare Pipe Price / metre** | `=ROUNDUP(Bare_Price_MT * Mass_MT_per_m, 0)` | Commercial rate for bare pipe per metre. |
| **25**| **Ext Coating / metre** | `=ROUNDUP(Ext_Cost_MT * Mass_MT_per_m, 0)` | External coating charge per metre. |
| **26**| **Int Coating / metre** | `=ROUNDUP(Int_Cost_MT * Mass_MT_per_m, 0)` | Internal coating charge per metre. |
| **27**| **Freight / metre** | `=ROUNDUP(Freight_MT * Mass_MT_per_m, 0)` | Delivery freight per metre. |
| **28**| **Total Price / metre** | `=Bare_m + Ext_m + Int_m + Freight_m` | **Final calculated baseline price per metre.** |
| **29**| **Effective Price / metre**| `=IF(ISBLANK(Override_m), Calc_m, Override_m)` | Allows commercial manual selling price customization! |
| **30**| **Line Total Value** | `=Effective_Price_m * MTRS` | Total billing amount per line item (excl. GST). |

---

## 4. Balanced Architecture: What Must Be Manual vs What Must Be Automated

The user's core requirement is:
> *"make this formulas save keep what is manual and do what can be automation there is no need to be ebery thing automation cause for example the coustomized prize is costmizable it cant be save for all the quesrys same so let it be what can be manually and do what can be automation i need this method and formauls calculation the thing i dont want to enter fomaulas manually"*

### A. What MUST Be Manual (User Input Fields — No Formulas to Type!)
These fields represent commercial judgment, market rates, and customer enquiry specifics that vary from enquiry to enquiry:
1. **Header / Project Information:**
   - Customer Name, Project / Tender Name, Enquiry Reference, Delivery Destination (PIN code / Location).
   - Selected Manufacturing Plant: `Jabalpur` vs `Khopoli`.
2. **Line Item Sizing Inputs:**
   - Sizing Selection: `Specify ID` or `Specify OD`.
   - Dimension: ID (mm) or OD (mm).
   - Wall Thickness (WT in mm).
   - Standard & Grade (e.g. `IS 3589 Fe 410`).
   - Pipe Type: `HSAW Coated` or `HSAW Bare`.
   - Length (MTRS).
3. **Commercial Base Rates (Pre-filled with company defaults, overridable per quote):**
   - Base Steel Coil Rate (`BEN + GE`): e.g. ₹62,000/MT (changes with market).
   - Grade Extra: e.g. ₹0 or ₹1,500/MT.
   - Conversion Charge: e.g. ₹4,000/MT.
   - Profit Margin: e.g. ₹0 or ₹2,000/MT.
   - Payment Terms / Finance Cost Rate: e.g. 1.0% (for 25% Adv + Bal LC) or 1.5%.
   - External Coating Rate: e.g. ₹500/m² (PU 1000µ) or ₹0 (if Bare).
   - Internal Coating Rate: e.g. ₹250/m² (FGE 406µ) or ₹0 (if Bare).
   - Freight Mode: `Stacking (Standard)` or `Nesting` or `Direct Rate per MT`.
   - Trailer Trip Freight Rate: e.g. ₹93,000/trip to Bagodar.
4. **Commercial Selling Price Override (Customizable per Line):**
   - **Manual Selling Price Override (₹/m):** The user can type any custom agreed price (e.g. ₹22,350 or ₹22,500). If left blank, the system uses the automated calculated price (₹22,351).

### B. What MUST Be Fully Automated (Calculated Behind the Scenes — Zero Formula Entry!)
The user never types an Excel formula. The software automatically computes:
1. **Geometric Calculations:**
   - Automatically computes OD (if ID given) or ID (if OD given).
   - Automatically computes nominal inches, kg/m mass, MT/m mass, total tonnage, pipe piece counts, and single pipe weights.
   - Automatically calculates circumferential coating factors ($m^2/m$) and total surface areas ($m^2$).
2. **Cost Layer Build-Up:**
   - Automatically calculates Yield Loss (Delivered Steel $\times$ 2%).
   - Automatically solves Finance Cost based on the chosen payment term rate.
   - Automatically converts $m^2$ coating rates to ₹/MT using mass and coating factors.
   - Automatically sums Bare Pipe Price, Ex-Mill Price, and Delivered Subtotal.
   - Automatically calculates Freight per MT based on pipe count, trailer capacity, and trip rate.
   - Automatically adds GST @ 18% and Grand Total.
3. **Per-Metre Pricing Breakdown:**
   - Automatically rounds Bare/m, Ext/m, Int/m, and Freight/m to whole integers using `ROUNDUP(..., 0)`.
   - Automatically sums components into the Total Calculated Unit Price per metre.
   - Automatically applies the override (if present) to determine the Effective Price.
   - Automatically computes Line Total, Total Project GST, and Grand Project Value.
4. **Offer & Document Generation:**
   - Automatically formats and exports the official Commercial Offer matching Sheet1 of Unipro 2.

---

## 5. Location Part & Migration Plan for `Mega_Pipes_Quotation_DEV`

### Status of Location Functionality:
- The user correctly observed: *"and also the location part is not updated in the Mega_Pipes_Quotation_DEV cause it is not updated"*.
- In `DEV/Mega_Pipes_Quotation_DEV.xlsm`, the workbook was frozen under strict SHA-256 release lockdown (`74B5D738FB...`).
- In `DEV/Mega_Pipes_Quotation_PILOT_COPY.xlsm`, we successfully integrated and validated:
  1. The **Location & Territory Summary UI** on both the `Quote` sheet and `Compare` sheet.
  2. Automatic PIN code to distance lookup (using `data/pin_lat_lon.json` and `data/mppl_plants.json`).
  3. Real-time plant distance comparison between **Jabalpur (782 km)** and **Khopoli (1,821 km)**.
  4. Competitor Freight Parity table (comparing Ratnamani, Surya Roshni, Welspun, Jindal Saw, AMNS, etc.).

### Safe Migration Steps to Update DEV:
When authorized by the user, the location functionality will be migrated into DEV using the following safe procedure:
1. **Pre-Migration DEV Backup:** Take timestamped backup `DEV/Mega_Pipes_Quotation_DEV_PRE_LOCATION_UPDATE.xlsm`.
2. **VBA & Sheet Synchronisation:** Import the updated `frmLocationPicker`, `modLocation`, and sheet formulas validated in `PILOT_COPY`.
3. **Automated Verification:** Run the 153-test calculation and COM test suite to verify 100% pass rate.
4. **Baseline Lock:** Record new approved SHA-256 checksum for DEV.

---

## 6. Mathematical Parity Verification Table (Unipro 4-Line Order)

| Metric | Line 1 (1243x8.4) | Line 2 (1317x8.4) | Line 3 (1367x8.4) | Line 4 (1417x8.4) | Project Total |
|---|---|---|---|---|---|
| **Quantity (MTRS)** | 2,000 | 350 | 500 | 4,400 | **7,250 m** |
| **Mass (MT/m)** | 0.255756 | 0.271085 | 0.281443 | 0.291801 | — |
| **Quantity (MT)** | 511.511 | 94.880 | 140.721 | 1,283.923 | **2,031.036 MT** |
| **Total Pipes (Nos)**| 167 | 30 | 42 | 367 | **606 Nos** |
| **Base Steel (₹/MT)** | ₹62,000 | ₹62,000 | ₹62,000 | ₹62,000 | ₹62,000 |
| **Yield Loss (₹/MT)**| ₹1,240 | ₹1,240 | ₹1,240 | ₹1,240 | ₹1,240 |
| **Conversion (₹/MT)**| ₹4,000 | ₹4,000 | ₹4,000 | ₹4,000 | ₹4,000 |
| **Finance Cost (₹/MT)**| ₹1,024 | ₹1,014 | ₹1,011 | ₹1,008 | ₹1,012.52 |
| **Bare Price (₹/MT)**| ₹68,264 | ₹68,254 | ₹68,251 | ₹68,248 | ₹68,252.52 |
| **Ext Coating (₹/MT)**| ₹7,630.37 | ₹7,627.46 | ₹7,625.67 | ₹7,624.00 | ₹7,625.88 |
| **Int Coating (₹/MT)**| ₹3,815.19 | ₹3,813.73 | ₹3,812.83 | ₹3,812.00 | ₹3,812.94 |
| **Ex-Mill Price (₹/MT)**| ₹79,709.56 | ₹79,695.18 | ₹79,689.50 | ₹79,684.00 | ₹79,691.34 |
| **Freight (₹/MT)** | ₹7,676 | ₹7,248 | ₹6,985 | ₹6,740 | ₹7,016.44 |
| **Subtotal (₹/MT)** | ₹87,385.56 | ₹86,943.18 | ₹86,674.50 | ₹86,424.00 | ₹86,707.78 |
| **Bare Price / metre**| **₹17,459** | **₹18,503** | **₹19,209** | **₹19,915** | — |
| **Ext Coating / metre**| **₹1,952** | **₹2,068** | **₹2,147** | **₹2,225** | — |
| **Int Coating / metre**| **₹976** | **₹1,034** | **₹1,074** | **₹1,113** | — |
| **Freight / metre** | **₹1,964** | **₹1,965** | **₹1,966** | **₹1,967** | — |
| **Total Price / metre**| **₹22,351** | **₹23,570** | **₹24,396** | **₹25,220** | — |
| **Line Value (Excl GST)**| **₹4,47,02,000** | **₹82,49,500** | **₹1,21,98,000** | **₹11,09,68,000** | **₹17,61,17,500** |

---

## 7. Recommended Next Steps for Implementation

1. **Retain Calculation Parity:** Ensure the Python engine and VBA modules adhere to this exact rounding and formula sequence.
2. **Present Intuitive UI:** Provide clean input cells for the manual fields (Base Steel, Conversion, Freight Rate, Override Price) while locking all calculated formula cells to prevent human errors.
3. **Roll Out Location UI to DEV:** Once the commercial team reviews this report, proceed with migrating the location and territory UI from `PILOT_COPY` into `DEV`.
