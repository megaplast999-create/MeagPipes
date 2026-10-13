# Mega Pipes Quotation System — Operator User Manual

**Document Code:** SOP-MKTG-QTE-001  
**Target Roles:** Sales Engineers, Commercial Estimators, Marketing Managers, General Commercial Authority  
**Version:** v1.0.0-RC

---

## 1. Introduction & Workflow Lifecycle Overview

The Mega Pipes Quotation System is an enterprise quotation and pricing application for welded steel pipes (HSAW, LSAW, ERW) with external (3LPE, 3LPP, FBE) and internal (Liquid Epoxy) coating capabilities.

Every quotation advances through a structured 5-stage lifecycle state machine:

```text
[ Draft ] ──(Calculate)──> [ Calculated ] ──(Submit Review)──> [ Review ] ──(Admin Approve)──> [ Approved ] ──(Issue Offer)──> [ Issued ]
   ▲                              │                                │                                │
   └──────────────────────────────┴────────(Edit Lines)────────────┴────────────────────────────────┘
```

> [!IMPORTANT]
> **Approval Integrity Guard:** Modifying any quotation line item, rate, freight parameter, or quantity after a quote has been marked **Approved** automatically resets the status back to **Calculated**, requiring formal re-authorization.

---

## 2. Step-by-Step Operator Procedure

### Step 1: System Entry and Dashboard
1. Open `Mega_Pipes_Quotation_v1.0.0_RC.xlsm` in Microsoft Excel.
2. Enable macros if prompted.
3. You will land on the **START** dashboard.
4. Verify system status indicators:
   - **Masters Status:** Confirms whether master tables are synced or running offline.
   - **User Identity:** Displays your current Windows login ID.
5. Click **New Quotation** or navigate directly to the **QUOTE** tab.

### Step 2: Quotation Header Information
1. On the **QUOTE** sheet, fill in the top header parameters:
   - **Customer Name:** Enter the prospective client organization name (e.g., `Unipro Steel Projects Ltd`).
   - **Project Name:** Enter project identifier or description (e.g., `Water Supply Transmission Scheme`).
   - **Delivery Destination / Pincode:** Enter delivery destination city and 6-digit postal code (e.g., `Bagodar, Jharkhand 825322`).
   - **Payment Terms:** Select standard terms from dropdown (e.g., `30 Days PDC`, `100% LC at Sight`).
   - **Validity Date:** Select quotation validity period (default 15/30 days).
2. Click **Generate Enquiry No.**:
   - The system connects to the network sequence generator and claims an official sequential number (e.g., `MKTGENQ26-27-001`).
   - If working offline or detached from corporate network, the system generates a distinct temporary ID (e.g., `TEMP-20261013-001`), which can be reconciled once reconnected.

### Step 3: Entering Line Items (`tblQuoteLines`)
For each item in the client enquiry, populate a row in `tblQuoteLines`:
1. **Pipe Family:** Select `HSAW`, `LSAW`, or `ERW`.
2. **Nominal OD (Inches / mm):** Select Outer Diameter from the dynamic dropdown (e.g., `18"`, `20"`, `22"`, `24"`).
3. **Wall Thickness (mm):** Select thickness (e.g., `7.92 mm`, `9.52 mm`). Only valid thicknesses for the chosen OD will appear in the dropdown.
4. **Steel Grade:** Select material grade (e.g., `IS 3589 Fe 410`, `API 5L X52`, `API 5L X65`).
5. **External Coating:** Select external coating specification (`Bare`, `3LPE Standard`, `3LPE Reinforced`, `FBE`).
6. **Internal Coating:** Select internal coating specification (`None`, `Liquid Epoxy (AWWA C210)`, `Food Grade Epoxy`).
7. **Quantity (MTR):** Enter length in running meters (e.g., `2,500`).
8. The system instantly computes total weight in Metric Tonnes (MT) based on pipe geometry.

### Step 4: Freight and Dispatch Configuration
Select freight dispatch methodology:
- **Method A (Fixed Rate):** Sourced per MT from commercial rate cards.
- **Method B (Stacking & Trailer Allocation):** Evaluates trailer nesting geometry on the `STACKING` sheet. Computes integer trailer trips required (25 MT / 40-foot flatbed limit) and derives exact freight cost per MT.
- **Method C (ORS Distance Routing):** Automatically computes road routing distance from candidate manufacturing plants:
  - **Jabalpur Plant (Central India):** Specialized large-diameter HSAW / 3LPE facility.
  - **Khopoli Plant (Western India):** Port-adjacent multi-spec facility.
  - Compares landed transportation tariffs and highlights the most economical dispatch location.
- **Plant Override:** If operational capacity requires routing through an alternate plant, select the alternate plant and select an override reason from the audit prompt.

### Step 5: Calculation and Live Summary
1. Click **Recalculate Quote** (or trigger calculation via sheet change).
2. The calculation engine processes all lines and populates:
   - Line Bare Pipe Cost / m
   - Line Coating Cost / m
   - Line Freight Cost / m
   - Line Final Offer Rate / m
   - Total Line Value (INR)
3. The **Live Summary KPI Card** displays:
   - Total Order Length (MTR)
   - Total Steel Weight (MT)
   - Grand Total Commercial Offer Value (₹)
   - Blended Average Price per Meter (₹/m)

### Step 6: Competitor Landed Price & Margin Protection
1. Navigate to the **COMPETITORS** tab.
2. Review competitor benchmark figures for matching sizes and delivery regions.
3. Review the **Floor Price Warning**:
   - The system checks whether proposed sales prices exceed the minimum commercial floor (Raw Material + Conversion Cost + Freight + Mandatory Floor Margin).
   - If any line falls below floor price, the row highlights in orange/yellow and prompts for Commercial Director sign-off.

### Step 7: Commercial Approval Workflow
1. Once pricing and commercial terms are verified, click **Submit for Review**.
   - Lifecycle state transitions from `Calculated` to `Review`.
2. Authorized Commercial Manager enters approval credentials (`MEGA_ADMIN_PWD`) and clicks **Approve Quote**.
   - Lifecycle state transitions to `Approved`.
   - Audit trail records approver identity, timestamp, and quotation checksum.

### Step 8: Generating Customer and Internal Exports
Once approved, commercial documentation can be distributed:
1. **Export Offer (PDF):**
   - Generates a 1-page professional quotation offer letter (`Quotes\<FY>\<Enquiry>_Offer.pdf`).
   - Formatted with company header, delivery terms, technical specifications, itemized unit prices, payment conditions, and signature blocks.
2. **Export Customer (Excel):**
   - Creates a client-facing sanitized workbook (`<Enquiry>_Customer_Offer.xlsx`).
   - Contains strictly the **OFFER** sheet and commercial synopsis.
   - **Confidentiality Guarantee:** Zero formula references, zero steel raw material prices, zero internal margin data, and zero manufacturing conversion costs.
3. **Export Internal (Excel):**
   - Creates an internal commercial archive containing all engineering sheets (`QUOTE`, `COSTING`, `STACKING`, `OFFER`, `QUOTE_LOG`).

### Step 9: Saving Revisions & Immutable Snapshots
1. Click **Save Snapshot**:
   - Creates an immutable, timestamped snapshot file: `Quotes\<FY>\<Enquiry>_R0_<User>.xlsx`.
2. When creating subsequent revisions for client negotiation:
   - Click **Create Revision** (advances R0 -> R1 -> R2).
   - Updates terms or pricing as negotiated and saves updated snapshot.
3. To reload any past snapshot:
   - Enter enquiry number and revision into the snapshot search box and click **Load Snapshot**.
   - Restores exact line items and commercial totals.
