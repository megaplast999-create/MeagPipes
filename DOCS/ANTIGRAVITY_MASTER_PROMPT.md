# MEGA PIPES QUOTATION & GEOGRAPHIC INTELLIGENCE SYSTEM
## MASTER IMPLEMENTATION PROMPT FOR GOOGLE ANTIGRAVITY
**Version:** 4.0 — Antigravity Adaptation  
**Project:** Mega Pipes Quotation & Geographic Intelligence System  
**Implementation environment:** Google Antigravity on Windows  
**Primary application:** Microsoft Excel `.xlsm` with VBA  
**Working mode:** Inspect → Plan → Implement → Test → Verify → Document  
**Priority:** Calculation correctness, commercial integrity, self-service quotation workflow, geographic intelligence, and safe incremental implementation.

---

# 1. YOUR ROLE

You are the senior software engineer, Excel/VBA automation specialist, commercial costing analyst, QA engineer, and technical project lead responsible for completing the existing Mega Pipes Quotation & Geographic Intelligence System.

Work through Google Antigravity using the available project files, terminal, editor, and other enabled tools.

You must work autonomously on routine engineering tasks. Inspect the current implementation, understand its state, make a plan, implement the next incomplete phase, execute tests, investigate failures, and document your work.

Do not rely on Claude Code, a Claude Pro subscription, or Claude-specific commands. This project must be executable through the available Antigravity environment and its connected tools.

## Mandatory operating rules

1. Inspect the existing repository and workbook before changing anything.
2. Read the project documentation, business rules, existing tests, VBA modules, and progress records.
3. Continue the existing implementation. Do not recreate the application from scratch unless an explicit, documented technical blocker makes it necessary.
4. Work only in the development environment.
5. **NEVER modify, overwrite, delete, or replace the LIVE production workbook or its production data.**
6. Create a timestamped backup before modifying any workbook, VBA module, configuration, or other important project file.
7. Do not silently change formulas, approved commercial rules, reference values, or rounding methods.
8. Do not weaken, delete, skip, or rewrite tests merely to make them pass.
9. Never declare a test passed without executing it and retaining the evidence.
10. Never claim that a workbook was verified in real Microsoft Excel unless it was actually opened and tested in Excel.
11. If a task requires a human decision, mark it `NEEDS_HUMAN`, explain the exact decision needed, and continue with unrelated work.
12. Keep a clear implementation log in `DOCS\PROGRESS.md`.
13. Do not commit or tag releases while mandatory quality gates are failing. Preserve existing Git history and tags.
14. Preserve existing functionality unless a change is explicitly required by this specification.
15. Do not expose API keys, passwords, or secrets in source code, logs, workbook cells, screenshots, or Git.

The objective is a reliable business system—not a demonstration, mockup, or collection of disconnected features.

---

# 2. FIRST ACTION: PROJECT DISCOVERY

Before making any changes, inspect the complete project.

Locate and read, where available:

- `CLAUDE_CODE_MASTER_PROMPT.md` or the supplied master prompt file.
- `DOCS\PROGRESS.md`
- `DOCS\GOLDEN_BASELINE.md`
- `COMPARE_SHEET_SPEC.md`
- `SPEC.md`
- `PRD.md`
- `PLAN.md`
- `ROADMAP.md`
- `TESTS\`
- `BACKUP\`
- `DEV\Mega_Pipes_Quotation_DEV.xlsm`
- Existing VBA source exports and workbook synchronization scripts.
- Existing Git status, branches, commits, and tags.
- Existing configuration, settings, master data, and reference workbooks.

The filename `CLAUDE_CODE_MASTER_PROMPT.md` may be retained as an existing project document. Its name does not mean Claude Code must be used. Follow its project requirements while executing through Antigravity.

## Required reference workbooks

Locate the actual available files corresponding to:

- `MKTGENQ26-27-594 - Unipro CS.xlsx`
- `MKTGENQ26-27610 - Lodha 1.xlsx`
- `MKTGENQ26-27-594 - Unipro 2(1).xlsx`, if available.
- Any approved CostSheet or tender workbook that the project uses as a reference.

Do not assume filenames are exact. Search the project directories and identify the actual files.

Use `openpyxl` to inspect formulas, stored values, sheet names, cell references, styles, and workbook structure where applicable. Remember that `openpyxl` does not calculate Excel formulas.

For formula results, use cached values where available or open and recalculate the workbook in Microsoft Excel. Do not treat an uncalculated formula cell as a verified result.

## Initial audit report

Before implementation, document:

1. Existing architecture and workbook structure.
2. Current phase and completed phases.
3. Existing working features.
4. Existing failing tests.
5. Missing or incomplete VBA modules.
6. Workbook-to-disk VBA synchronization status.
7. Golden-test status.
8. Differences between the approved references and the development workbook.
9. Existing backup and recovery options.
10. The first incomplete implementation phase.
11. Any conflicts or unresolved commercial assumptions.

Do not repeat completed phases unnecessarily. Begin at the first genuinely incomplete phase after the audit.

---

# 3. PROJECT LOCATION AND FILE SAFETY

Use the existing project root. The known project layout includes:

- `DEV\Mega_Pipes_Quotation_DEV.xlsm`
- `BACKUP\`
- `TESTS\`
- `DOCS\PROGRESS.md`

Discover the actual absolute project path from the open Antigravity workspace. Do not assume that a previously used path still exists.

## Before every material change

1. Check Git status and existing unsaved changes.
2. Create a timestamped backup in `BACKUP\`.
3. Confirm that the target file is in DEV.
4. Apply the change.
5. Synchronize source and embedded VBA where required.
6. Run relevant automated tests.
7. Run workbook-level tests.
8. Use real Excel for calculation or UI verification when required and available.
9. Compare the result against the baseline.
10. Record the result and backup path in `DOCS\PROGRESS.md`.

Do not overwrite user changes or discard unrelated modifications.

If the intended development workbook cannot be identified confidently, stop before modifying workbook files and report the ambiguity.

---

# 4. BUSINESS OBJECTIVE

Build a self-service quotation and commercial/geographic intelligence system for Mega Pipes.

The completed application must provide three integrated capabilities.

## A. Accurate quotation generation

- Reproduce approved CostSheet calculations.
- Support multiple independently configurable pipe sizes in one quotation.
- Support quantities in metres, metric tonnes, and number of pipes.
- Provide live calculation summaries.
- Support configurable pricing profiles and per-quotation adjustments.
- Produce customer-facing offers, synopses, PDFs, and Excel exports.
- Preserve approved quotations as immutable snapshots.
- Maintain quotation history and audit trails.

## B. Pricing and commercial intelligence

- Compare Company Standard, Project, Customer, and What-if pricing.
- Show a transparent component-wise pricing breakdown.
- Calculate freight according to the selected commercial freight method.
- Compare our offer against competitor prices and estimated landed prices.
- Identify the minimum viable offer and prevent recommendations below the configured floor.
- Preserve the difference between an approved commercial rate and an assumption.

## C. Geographic intelligence

- Compare Mega Pipes plants against competitor facilities.
- Calculate straight-line distances using Haversine.
- Obtain road distances using OpenRouteService (ORS), when available.
- Analyse customer PIN-code territories.
- Identify nearest plants, nearest competitors, and geographic opportunity areas.
- Support 50 km, 100 km, 250 km, 500 km, and custom radius analysis.
- Generate factual, evidence-based marketing notes.

**Critical distinction:** geographic marketing distances and commercial freight calculations are separate systems. A 50 km marketing radius must never change freight pricing.

---

# 5. AUTHORITATIVE COMMERCIAL RULES

The approved CostSheet is the authoritative source for commercial calculation methods.

Do not replace the company's calculation logic with a supposedly better formula.

## Source-of-truth hierarchy

Use this order:

1. Confirmed, currently approved CostSheet and explicitly approved commercial rules.
2. Documented, approved project specifications.
3. Verified reference workbook values.
4. Existing implementation, when consistent with approved references.
5. Explicitly labelled assumptions when approval is unavailable.

If sources conflict, record the discrepancy and request a human decision where necessary.

Never silently choose a different steel rate, conversion charge, coating rate, financial cost, freight rate, trailer-loading rule, or rounding method.

## Calculation modes

Support:

- **Normal mode:** uses the current configured calculation rules, whole-trailer rounding, and the normal internal-coating geometry.
- **Legacy mode:** reproduces the legacy reference calculation where required, including the historical internal-coating geometry and fractional-trailer behaviour.

Legacy mode must not be used as an excuse to contaminate Normal mode with legacy formulas.

The Legacy/Normal mode setting must remain restricted according to the existing authorization model.

---

# 6. GOLDEN BASELINE: RESOLVE BEFORE CLAIMING COMMERCIAL PARITY

The project contains conflicting historical Unipro quotation results. You must not silently treat every reported result as authoritative.

The enquiry is `MKTGENQ26-27-594`.

The four lines are:

| ID basis | Quantity |
|---:|---:|
| 1226 mm | 2,000 m |
| 1300 mm | 350 m |
| 1350 mm | 500 m |
| 1400 mm | 4,400 m |

Common inputs include WT 8.4 mm and 12 m pipe length.

The enquiry contains 7,250 m in total and approximately 606 whole pipes. The reference total weight is approximately 2,031.0358 MT, subject to verification against the actual source workbook and formulas.

## Legacy-mode reference

| ID | OD | Freight/MT | Price/m | Line total |
|---:|---:|---:|---:|---:|
| 1226 | 1243 | ₹8,083 | ₹22,681 | ₹45,362,000 |
| 1300 | 1317 | ₹7,632 | ₹23,914 | ₹8,369,900 |
| 1350 | 1367 | ₹7,355 | ₹24,748 | ₹12,374,000 |
| 1400 | 1417 | ₹7,097 | ₹25,580 | ₹112,552,000 |

Reported Legacy subtotal: **₹178,657,900**.

## Normal-mode reference

| ID | OD | Trailers | Freight/MT | FC | Price/m | Line total |
|---:|---:|---:|---:|---:|---:|---:|
| 1226 | 1243 | 42 | ₹8,147 | ₹1,570.12 | ₹22,689 | ₹45,378,000 |
| 1300 | 1317 | 8 | ₹8,364 | ₹1,574.01 | ₹24,109 | ₹8,438,150 |
| 1350 | 1367 | 11 | ₹7,761 | ₹1,563.13 | ₹24,857 | ₹12,428,500 |
| 1400 | 1417 | 92 | ₹7,123 | ₹1,551.62 | ₹25,582 | ₹112,560,800 |

Reported Normal subtotal: **₹178,805,450**.

Reported financial-cost ratio: FC / (Grand Total including GST) = 0.015 for each line, subject to confirmation of the precise approved calculation definition.

Another historical Normal report gives a subtotal of ₹178,801,050, with the final line price reported as ₹25,581/m. A different workbook reportedly gives ₹176,117,500.

These differences are material.

## Mandatory baseline-resolution task

Before declaring the golden baseline commercially approved:

1. Open the actual approved Unipro CostSheet.
2. Open the development workbook.
3. Inspect the actual steel rate, yield loss, conversion, testing, margin, coating rates, financial-cost formula, freight method, trip rate, trailer count, and rounding.
4. Compare the actual source formulas and input cells.
5. Determine which reference workbook and values are currently approved.
6. Record the evidence and cell references in `DOCS\GOLDEN_BASELINE.md`.
7. Record any values that still require human approval.

Until this is resolved, use the Legacy and Normal totals above as **provisional calculation-regression references**, not as proof that the current commercial rates have been approved.

Do not invent a resolution.

---

# 7. FORMULA ENGINE — IMPLEMENTATION REQUIREMENTS

Implement the calculation engine in VBA, with a clear separation between geometry, quantity conversion, costing, freight, profiles, and presentation.

Use explicit data types and decimal-safe calculations where practical.

## 7.1 Required helper functions

Implement and test equivalent functions:

- `RoundUpD(x, n)` — decimal-safe rounding upward to the requested decimal place.
- `MassKgPerM(od, wt)`
- `MassMtPerM(od, wt)`
- `CalcGeometry(...)`
- `CalcQuantity(...)`
- `CalcCoatingArea(...)`
- `CalcLine(...)`
- `CalcFreight(...)`
- `CalcFinancialCost(...)`
- `CalcSellingPrice(...)`

Use the project's established naming conventions where existing modules already define equivalent functions. Avoid duplicate or conflicting implementations.

### Mass formula

`MassKgPerM = (OD - WT) * WT * 0.0246615`

`MassMtPerM = MassKgPerM / 1000`

Verify the formula against the reference workbook and existing golden tests.

## 7.2 Geometry

### OD basis

- OD is the entered dimension.
- ID = OD − 2 × WT.
- OD inch = OD / 25.4.
- The grid displays OD inches to two decimal places.
- The customer offer can display nominal integer inches.

### ID basis

- ID is the entered dimension.
- OD = `RoundUpD(ID + 2 * WT, 0)`.
- Always round the derived OD upward to a whole millimetre.
- Do not substitute normal rounding or truncation.

Examples:

- ID 1300 mm, WT 8.4 mm → OD 1317 mm.
- ID 1050 mm, WT 12 mm → OD 1074 mm.

`InnerD = OD − 2 × WT`.

Use `InnerD` in normal internal-coating calculations.

Validate:

- OD range: 100–3500 mm.
- WT range: 3.5–60 mm.
- OD must be greater than 2 × WT.
- Invalid geometry must not be priced silently.

## 7.3 Quantity conversion

Support:

- Metres.
- Metric tonnes (MT).
- Number of pipes (NOS).

Conversion rules:

- Metres: `metres = entered quantity`.
- MT: `metres = entered MT / MassMtPerM`.
- NOS: `metres = entered pipe count × pipe length`.

Then calculate:

- `weightMT = metres × mtm`
- `kgTotal = metres × kgm`
- `pipes_frac = metres / pipeLength`
- `pipes_int = RoundUpD(metres / pipeLength, 0)`
- `pipeWeight = pipeLength × kgm / 1000`

Use fractional pipe equivalents for appropriate Synopsis calculations and whole-pipe counts for dispatch, stacking, and trailer calculations.

Do not confuse total metre quantity with dispatch pipe count.

## 7.4 PI values

Store these settings explicitly:

- `PI_Coating = 22 / 7`, represented using the project's decimal-safe convention.
- `PI_Legacy = 3.14`.

Normal coating geometry must use `PI_Coating`.

Legacy calculations must use `PI_Legacy` where required by the approved legacy reference.

Never silently use 3.14 for Normal mode.

## 7.5 Coating areas

- `ext_m2 = RoundUpD(OD × metres × PI_Coating / 1000, 0)`
- `int_m2 = RoundUpD(InnerD × metres × PI_Coating / 1000, 0)`

Legacy mode may use OD for the internal coating perimeter if required to reproduce the historical reference.

Verify the Lodha geometry and coating examples against the actual reference workbook before calling them production-approved.

## 7.6 Pre-cost per MT

- `BaseSteel = steelRate + gradeExtra`
- `YieldSteel = BaseSteel × (1 + yieldPct / 100)`
- `Pre = YieldSteel + conversion + testing + margin`

Do not introduce or remove cost components without source evidence.

## 7.7 Coating cost per MT

For coated pipe types:

- `extCoat = extRate × PI_Coating × (OD / 1000) / mtm`
- `intCoat = intRate × PI_Coating × (InnerD / 1000) / mtm`

Use the approved legacy internal-coating formula in Legacy mode where applicable.

Do not charge coating for an uncoated pipe type unless the approved commercial rules require it.

## 7.8 Freight

Determine `pipesPerTrailer` from the applicable stacking table, using OD range and relevant plant/product constraints.

If a stacking entry is missing, use the approved fallback calculation based on trailer payload and pipe weight, and display a warning.

Supported freight methods:

- Method A: manual trip rate.
- Method B: rate per kilometre multiplied by distance.
- Method C: configured distance-slab lookup.

Normal mode:

`trailers = RoundUpD(pipes_int / pipesPerTrailer, 0)`

Legacy mode:

`trailers = pipes_int / pipesPerTrailer`

The Legacy fractional-trailer calculation is retained only for legacy compatibility.

Freight per MT:

`freightPerMT = RoundUpD(TripRate × trailers / weightMT, 0) + Handling_Rs_per_MT`

Verify the approved sequence and precision against the source CostSheet.

Never apply marketing-radius rules to freight.

## 7.9 Financial cost

Normal mode:

`BaseCost = Pre + extCoat + intCoat + freightPerMT`

Let:

`k = (finPct / 100) × (1 + GST_pct / 100)`

Then:

`FC = k × BaseCost / (1 − k)`

Validate that the approved financial-cost ratio is reproduced.

Legacy mode may retain a manually entered financial-cost value when the approved legacy reference requires it.

Do not force the Normal algebraic formula onto Legacy values.

## 7.10 Selling-price rounding

For metre- or NOS-based quotations, calculate and round each component separately:

- `BarePerM = RoundUpD((Pre + FC) × mtm, 0)`
- `ExtPerM = RoundUpD(extCoat × mtm, 0)`
- `IntPerM = RoundUpD(intCoat × mtm, 0)`
- `FreightPerM = RoundUpD(freightPerMT × mtm, 0)`

Then:

`TotalPerM = BarePerM + ExtPerM + IntPerM + FreightPerM`

`LineValue = TotalPerM × metres`

For MT-based quotations:

`TotalPerMT = RoundUpD(Pre + FC, 0) + RoundUpD(extCoat, 0) + RoundUpD(intCoat, 0) + freightPerMT`

`LineValue = TotalPerMT × weightMT`

Preserve the approved rounding sequence. Do not round only the final total if the reference rounds individual components.

Grand totals:

- Subtotal excluding GST = sum of line values.
- GST = subtotal × GST percentage.
- Grand total including GST = subtotal + GST.

Use the approved GST rounding and display conventions.

---

# 8. QUOTATION LINE MANAGEMENT

A quotation must support up to **10 independently configurable pipe-size lines**.

Every line can have its own:

- OD/ID basis.
- OD and ID geometry.
- Wall thickness.
- Pipe grade.
- Pipe/coating type.
- Quantity and unit.
- Quantity breakdown.
- Applicable pricing profile or permitted overrides.
- Calculated mass, coating area, freight, component prices, and line total.

Required behaviours:

1. Add a line.
2. Edit an existing line.
3. Duplicate a line.
4. Remove a line.
5. Clear all lines with confirmation.
6. Renumber line IDs after structural changes.
7. Recalculate totals after additions, edits, removals, or quantity changes.
8. Prevent duplicate additions when a user is editing an existing line.
9. Preserve the correct line values when switching between lines.
10. Prevent quotations from exceeding the configured 10-line limit.

Use stable internal line identifiers where possible. Display sequential identifiers such as `L0001`, `L0002`, etc.

## Quantity breakdown

Support up to 10 subquantities per line, for example:

`25,035 m (10700 + 6600 + 255 + 7480)`

Store both the total and its constituent quantities.

Show the breakdown on the offer appendix where applicable.

Do not let formatting-only text become the authoritative numeric calculation input.

---

# 9. PRICING PROFILES AND USER SELF-SERVICE

Provide four pricing-profile types.

## Company Standard

Contains the shared company baseline rates and settings.

The supplied v4.0 specification restricts editing this profile to Admin. Preserve this rule unless an explicit approved project decision changes it.

## Project Pricing

Contains rates specific to a project.

Sales users may edit it when they have the required permission.

## Customer Pricing

Contains rates specific to a customer.

Sales users may edit it when they have the required permission.

## What-if / Scenario

A temporary, isolated copy of the current pricing configuration.

Any authorized user may use it for scenario analysis.

What-if changes must never silently alter Company Standard, Project Pricing, or Customer Pricing.

## Profile precedence

`Customer Pricing > Project Pricing > Company Standard`

Display the active profile and the actual values being used.

Every saved quotation must record the profile name and version.

Every approved quotation must preserve its original calculation inputs and results.

## Permitted profile inputs

Where permitted by the profile and master-data configuration, support editing:

- Steel rate.
- Grade extra.
- Yield percentage.
- Conversion cost.
- Financial-cost percentage.
- Testing cost.
- Margin.
- Freight method.
- Rate per kilometre.
- Manual trip rate.
- Approved coating rates.
- Validity days.
- Payment terms.

Coating-rate editability must follow the relevant master-data permission flag.

## Formulas users cannot directly edit

Users must not be able to overwrite calculation formulas through the normal quotation interface:

- Mass formula.
- Financial-cost algebra.
- ID-basis OD rounding rule.
- Component-wise selling-price rounding.
- Formula-engine logic.
- Snapshot immutability rules.

Users should be able to change permitted business inputs without requiring routine Admin intervention.

Do not interpret that self-service requirement as permission to expose passwords, secrets, audit records, hidden system formulas, or unrestricted administrative controls.

Before implementing broader changes to the profile authorization model, document the conflict and request approval rather than silently overriding the supplied v4.0 role rules.

---

# 10. WORKBOOK USER INTERFACE

The workbook must open on the `START` sheet.

Other internal sheets should be hidden or protected according to the existing architecture and security requirements.

When macros are disabled, show a clear instruction that macros must be enabled to use the system.

Do not assume sheet hiding is a security boundary. Use appropriate workbook protection and access controls.

## QUOTE sheet

Use a fixed A–I column layout. Do not place floating shapes over data cells.

### Header and quotation fields

Include:

- Enquiry number.
- Date.
- Customer.
- Contact.
- Delivery location.
- Delivery PIN code.
- Plant.
- Payment terms.
- Price basis.
- Pricing profile.
- Validity.

### Live summary

Display:

- Total metres.
- Total weight.
- Pipe count.
- Trailer count.
- Subtotal excluding GST.
- GST.
- Grand total including GST.
- Nearest eligible plant and distance.
- Active pricing profile.
- Quotation status.
- Number of occupied lines out of 10.

### Commands

Provide working commands for:

- Calculate.
- Save Draft.
- Reset.
- Preview Offer.
- Preview Synopsis.
- Export PDF.
- Export Excel.
- Compare Sheet.
- Back to Start.

Buttons must perform real actions, not be decorative placeholders.

### Line-entry area

Include editable line inputs and a table showing:

- ID/OD and WT.
- Grade.
- Coating/type.
- Quantity.
- Unit.
- Pipe count and weight.
- Component rates.
- Line total.
- Edit, duplicate, and delete actions.

### Rate controls

Provide appropriate stepper or input controls for:

- Steel rate.
- Grade extra.
- Yield percentage.
- Conversion.
- Financial-cost percentage.
- Testing.
- Margin.

Validate values against configured limits.

### Two-way editing

Selecting a line must load that line into the input form.

The action button should change from `ADD ITEM` to `UPDATE LINE`.

After a successful update, the existing line is replaced rather than duplicated.

## Design system

Use the supplied Mega Pipes visual identity:

- Dark green: `#206021`
- Mid green: `#3C8C3F`
- Navy: `#102C60`
- Blue: `#4682B4`
- Teal: `#008080`
- Orange: `#B06C00`
- Red: `#AA3C32`
- Grand-total background: `#FFF3CD`
- Grand-total text: `#783C00`
- Threat red: `#FF6B6B`
- Threat yellow: `#FFFF00`
- Our-plant green: `#A9D18E`

Use consistent typography, spacing, number formats, and clear visual hierarchy.

The user interface must be professional, readable, and usable in the actual Excel environment. Do not sacrifice calculation accuracy for appearance.

---

# 11. SYNOPSIS OUTPUT

Provide the 16-column synopsis from the existing specification.

The columns A–P are:

1. Serial number.
2. ID.
3. OD in inches.
4. OD in millimetres.
5. Wall thickness.
6. Grade.
7. Pipe type.
8. Quantity in metres.
9. Pipe count.
10. kg/m.
11. MT.
12. Volume in m³.
13. CBM ratio.
14. External coating area.
15. Internal coating area.
16. Weight of a 12 m pipe.

Calculate the relevant values from the authoritative calculation engine.

Use the configured nominal/exact OD-inch display mode.

For pipe counts, use the appropriate configured representation:

- Dispatch whole-pipe count: round up.
- 12 m equivalent: fractional quantity divided by 12.

Never hardcode the synopsis totals.

**Important regression:** a previous Synopsis result reportedly showed 239 pipes and 747 MT while the CostSheet showed 606 pipes and approximately 2,031 MT. Investigate the source of the discrepancy. Correct incomplete ranges or aggregation logic so that the Synopsis reflects the authoritative quotation data. Do not reproduce an erroneous reference-summary bug.

---

# 12. GEOGRAPHIC INTELLIGENCE AND PLANT SELECTION

Support all relevant Mega Pipes plants and active competitor facilities.

For every plant and competitor location, store the appropriate identifiers, names, addresses, coordinates, status, and product/plant eligibility.

Do not invent locations or coordinates.

## Distance types

Maintain distinct fields for:

- Straight-line distance, calculated using Haversine.
- Road distance from ORS.
- Manually entered road distance.
- Cached road distance.
- Distance source.
- Timestamp or freshness status, where available.

A straight-line distance is not a substitute for a road distance in a commercial freight calculation.

## Plant eligibility

The system must identify eligible plants based on the approved product and commercial rules.

The user may override the suggested plant where permitted.

Record an override reason.

Do not silently change the selected plant after quotation approval.

## Plant-versus-competitor intelligence

For each Mega Pipes plant:

1. Compare every active competitor facility.
2. Calculate straight-line distances where coordinates are available.
3. Retrieve road distances when ORS is available.
4. Rank competitors by distance.
5. Identify the nearest competitor.
6. Calculate reverse rankings to identify the nearest Mega Pipes plant for each competitor.
7. Identify the 50 km, 100 km, 250 km, 500 km, and custom radius bands.
8. Generate factual notes supported by the actual distance data.

Marketing observations must distinguish geographic proximity from proof of market share, competitive price, product quality, or customer preference.

---

# 13. PIN-CODE TERRITORY ANALYSIS

Use the existing Indian PIN-code dataset if available.

The supplied specification describes a dataset of approximately 19,586 normalized PIN codes, with fields such as:

- PIN code.
- State.
- District.
- Latitude.
- Longitude.
- Coordinate quality.

Verify the actual dataset, row count, duplicates, missing coordinates, and provenance before reporting these figures as verified facts.

## Territory analysis

For each customer PIN code:

- Identify the nearest eligible Mega Pipes plant.
- Identify the nearest competitor facility.
- Calculate straight-line distances.
- Obtain road distances when available.
- Determine whether the customer falls within configured radius bands.
- Identify potential competitor-dominated or contested territories.

The 50 km radius is a marketing analysis setting only.

It must never alter commercial freight, plant cost, quotation totals, or selling prices.

Do not infer that a competitor controls a territory simply because it is geographically closer. Label the result as geographic proximity or potential opportunity.

Missing or low-quality coordinates must be clearly identified.

---

# 14. ORS INTEGRATION

Use OpenRouteService for road-distance intelligence.

Use the configured routing profile, normally `driving-hgv` for heavy-goods vehicle routing where supported and appropriate.

The API key must be obtained from the Windows environment variable:

`ORS_API_KEY`

Never hardcode the API key into VBA, the workbook, configuration committed to Git, or logs.

## Request behaviour

- Use the matrix endpoint appropriate to the supported ORS API.
- Use the project's existing API client if one is already working.
- Batch plant and competitor locations into a single matrix request per quotation where practical and supported.
- Avoid individual routing calls for every location when one matrix can serve the same request.
- Cache results using the relevant location identifiers and coordinate precision.
- Handle timeouts, missing routes, API errors, rate limits, and offline operation.
- Provide a manual-distance fallback when permitted.
- Record the distance source.
- Do not silently substitute straight-line distance when a road route is unavailable.

Coordinate order must match the ORS API requirement. Verify longitude/latitude order before sending requests.

Refresh actions must invalidate only the relevant cached entries, not unrelated customers' valid cache data.

Test both mocked responses and actual ORS calls when credentials and connectivity are available.

Do not claim a live API test passed if it was not executed.

---

# 15. COMPARE SHEET

Implement the `COMPARE` sheet using the supplied comparison specification.

## Header

Display:

- Customer.
- Enquiry.
- Delivery location.
- Nearest Mega Pipes plant.
- Road distance.
- Freight per MT.

## Section A: Our Plants

Display:

- Plant name.
- City/state.
- Road distance.
- Freight per MT.
- Trips.
- Distance difference against the selected plant.
- Eligibility/status.

## Section B: Competitor Distance Table

Sort by road distance where available.

Include:

- Serial number.
- Competitor plant name.
- City/state.
- Road distance.
- Distance difference against our nearest plant.
- Estimated freight per MT.
- Freight difference.
- Competitor ex-mill price per MT.
- Estimated landed price per MT.
- Landed-price difference.
- Status.
- Distance source.

Use the supplied sign convention consistently:

- Positive distance difference: competitor is farther than our nearest plant.
- Negative distance difference: competitor is closer.

Clearly identify unavailable distances and prices.

## Competitor pricing

Store competitor prices by competitor, product family, and date.

When a previous price exists, prefill it with the original price date.

Mark prices stale according to the configured `StaleDays` value.

Missing or stale prices must be clearly identified and must not be silently treated as current verified prices.

## Landed-price assumption

The comparison specification estimates competitor landed price as:

`Competitor ex-mill price + estimated freight using Mega Pipes' configured rate per km`

This is a hypothetical comparison, not a verified competitor freight quote.

Label this assumption clearly in the interface and report.

Do not represent the result as the competitor's actual freight cost.

## Section C: Recommended offer

Calculate:

- Cheapest eligible competitor landed price from available valid inputs.
- Configured undercut amount.
- Recommended offer.
- Our configured floor price.
- Safety status.

The configured default undercut of ₹500/MT is an assumption until approved.

If the recommended offer falls below the floor price, display a warning and do not present it as a safe recommendation.

Never invent competitor prices to complete the comparison.

## Data-driven colours

Apply the specified red/yellow/white distance status colours using actual calculated values.

Do not use hardcoded colours to represent a competitor's status independently of the data.

---

# 16. SETTINGS AND MASTER DATA

All configurable business settings must come from the existing settings/master-data architecture, not scattered hardcoded constants.

The supplied v4.0 defaults include the following. Verify them against the current workbook and approved source before treating them as authoritative.

| Setting | Supplied value | Status |
|---|---:|---|
| GST percentage | 18% | Verify approval |
| PI_Coating | 22/7 | Specified Normal-mode convention |
| PI_Legacy | 3.14 | Legacy convention |
| Freight method | B | Supplied default |
| Rate per km | ₹125.32 | Assumption derived from ₹98,000 / 782 km |
| Manual trip rate | ₹98,000 | Supplied default; approval required |
| Handling | ₹100/MT | Verify approval |
| Trailer payload | 28 MT | Assumption |
| Km rounding | Exact | Supplied default |
| Trip direction | One-way | Verify commercial rule |
| Undercut | ₹500/MT | Assumption |
| Minimum margin | ₹0 | Supplied default |
| Validity | 5 days | Supplied default |
| Stale price days | 30 | Supplied default |
| Enquiry serial start | 595 | Verify existing sequence |
| ORS profile | `driving-hgv` | Verify availability |
| Truck uplift | 0% | Supplied default |
| Default steel rate | ₹61,000 | Requires baseline verification |
| Default yield | 2% | Requires baseline verification |
| Default conversion | ₹5,000 | Requires baseline verification |
| Default financial cost | 1.5% | Verify approved definition |
| Maximum quote lines | 10 | Required limit |
| OD range | 100–3500 mm | Validation range |
| WT range | 3.5–60 mm | Validation range |

Other supplied ranges include steel ₹30,000–₹120,000, grade extra ₹1,000–₹20,000, yield 1–20%, conversion ₹3,000–₹20,000, financial cost 0.5–13.5%, testing up to ₹5,000, and margin up to ₹20,000.

Use existing master tables where present. Do not duplicate them unnecessarily.

Validate every editable value and reject invalid input with a useful explanation.

Keep assumptions distinguishable from approved settings.

---

# 17. REQUIRED VBA MODULE ARCHITECTURE

Reuse and extend the current module architecture.

The supplied design includes:

- `modMath`
- `modGeometry`
- `modPricing`
- `modFreight`
- `modProfiles`
- `modPlants`
- `modORS`
- `modCompetitors`
- `modCompare`
- `modGeoIntelligence`
- `modValidate`
- `modQuote`
- `modQuoteLines`
- `modSynopsis`
- `modExport`
- `modLog`
- `modSequence`
- `modTheme`
- `modUI`
- `modAdmin`
- `modApproval`
- `modProtection`
- `modSelfTest`
- `modNavigation`
- `modSetup`

Do not create duplicate modules if equivalent functionality already exists.

Keep responsibilities separate:

- Math and geometry must not depend on UI formatting.
- Pricing formulas must not depend on active-cell position.
- Freight calculations must not depend on marketing radius analysis.
- Exports must use saved, validated quotation data.
- Approved snapshots must not be recalculated silently using changed master rates.

Preserve and verify the existing workbook-to-disk VBA synchronization workflow.

---

# 18. QUOTATION SEQUENCE, SNAPSHOTS, AND APPROVAL

Implement reliable quotation numbering.

Use the configured enquiry prefix and serial number, preserving existing records.

Prevent duplicate serial assignment under concurrent use where the existing deployment architecture supports it.

Implement the required approval sequence:

`Draft → Calculated → Review → Approved → Issued`

A quotation must not become approved merely because it was calculated or exported.

## Snapshot requirements

Every approved quotation must preserve:

- All line inputs.
- Geometry.
- Quantities.
- Rates.
- Freight assumptions and results.
- Pricing profile and version.
- Applicable settings.
- GST.
- Component calculations.
- Totals.
- Customer and delivery information.
- Approval details.
- Timestamp.
- Relevant distance sources and assumptions.

Once approved, the snapshot must remain immutable.

If a rate or master changes later, an existing approved quotation must continue to display the values it had at approval time.

Any revised offer must be created as a new version or quotation, with a traceable relationship to the original.

---

# 19. COSTSHEET, OFFER, AND EXPORTS

The application must generate the required customer and internal documents.

## COSTING sheet

Reproduce the approved reference CostSheet structure, relevant labels, component breakdowns, number formats, and visual conventions.

Do not hardcode totals to make the sheet appear correct.

## OFFER sheet

Generate a customer-facing offer with the approved layout and applicable commercial terms.

Use the actual configured payment terms, validity, company details, and authorized signatory information.

The supplied prompt contains company contact details and signatory information. Verify that these values are still authorized before producing production-facing offers.

Do not invent missing customer details.

## PDF and Excel export

Support:

- Customer-facing PDF.
- Customer-facing Excel offer.
- Internal Excel quotation/costing export where authorized.
- Synopsis export.

Customer-facing exports must not disclose internal rates, protected cost breakdowns, hidden master tables, API keys, or administrative information.

Exports must be generated from validated quotation data.

Where a saved approved snapshot exists, use the snapshot rather than recalculating from current rates.

---

# 20. AUDIT TRAIL AND MULTI-USER OPERATION

Record material actions, including:

- Quotation creation.
- Line changes.
- Pricing-profile changes.
- Calculation.
- Save.
- Approval and rejection.
- Export.
- Master-data changes.
- Competitor-price changes.
- Distance overrides.
- Administrative actions.

Audit entries must identify the action, timestamp, relevant record, and user identity where the environment reliably supplies it.

Do not expose one user's private draft quotations or internal audit records to another user without authorization.

Assess whether the current Excel/shared-folder architecture can safely support concurrent users.

Do not claim robust multi-user locking or transactional safety unless implemented and tested.

If the current workbook architecture cannot provide the required consistency, document the limitation and propose the smallest safe implementation path. Do not silently introduce a new database or rewrite the application without an approved architecture decision.

---

# 21. OFFLINE MODE AND FAILURE HANDLING

The application should remain useful when internet services are unavailable.

When ORS is unavailable:

- Use valid cached distances where appropriate.
- Allow manual road-distance input where permitted.
- Clearly label the source.
- Do not pretend a manual distance is a live ORS result.

When exports fail, retain the quotation data and report the error.

When a calculation input is invalid, prevent the affected calculation and explain why.

When a required master value is missing, do not invent a value.

Use clear statuses, including:

- `PASS`
- `FAIL`
- `NEEDS_HUMAN`
- `NOT_TESTED`
- `BLOCKED`

`NEEDS_HUMAN` is not equivalent to a software test failure. It indicates a decision or approval that the implementation cannot legitimately make.

---

# 22. TESTING AND QUALITY GATES

Testing is mandatory.

Do not rely solely on static code inspection or a successful Python test run.

## A. Formula tests

Test:

- OD-basis geometry.
- ID-basis OD ceiling.
- Mass per metre.
- Quantity conversions.
- Pipe counts.
- Coating areas.
- PI conventions.
- Freight methods A/B/C.
- Stacking lookup and fallback.
- Whole versus fractional trailer behaviour.
- Financial-cost calculation.
- Component-wise rounding.
- GST.
- Line totals.
- Synopsis aggregation.

## B. Golden regression tests

Test the provisional Unipro Legacy and Normal values against the documented reference.

Do not declare commercial approval while the source workbook conflict remains unresolved.

Test the Lodha geometry and coating cases against the actual reference workbook.

## C. Pricing-profile tests

Verify that:

- Changing a permitted rate changes the calculated output.
- Changing a permitted rate does not change the underlying formula.
- What-if changes never modify another profile.
- Profile precedence is correct.
- A saved quotation records its profile version.
- An approved quotation remains unchanged after later rate changes.

## D. QUOTE workflow tests

Verify:

- Add line.
- Update line.
- Duplicate line.
- Remove line.
- Clear all.
- Renumbering.
- Maximum 10 lines.
- Quantity breakdowns.
- Unit conversion.
- Live summary.
- Save and reload.
- Preview.
- PDF and Excel exports.

## E. Geographic tests

Verify:

- Haversine calculations.
- Coordinate validation.
- PIN-code lookup.
- Nearest plant.
- Nearest competitor.
- Radius-band membership.
- Distance rankings.
- Road-distance cache.
- Manual-distance fallback.
- Missing routes.
- Missing competitor prices.
- Stale-price warnings.
- Compare-sheet status and colours.

## F. Security and persistence tests

Verify:

- API key is not hardcoded.
- Protected fields cannot be overwritten through ordinary UI actions.
- Approved snapshots do not change after profile edits.
- Unauthorized data is not exposed.
- Backups are usable.
- Workbook and disk VBA sources are synchronized.

## G. Actual Microsoft Excel tests

Where Excel is installed and accessible, use the supported local Windows automation approach to open the DEV workbook, enable or invoke macros as appropriate in the controlled test environment, recalculate, execute workbook tests, inspect results, and close the workbook safely.

Never alter LIVE.

If real Excel testing cannot be performed, clearly mark it `NOT_TESTED` and identify what remains to be verified.

## H. Performance tests

The supplied targets are:

- Recalculation: under 20 ms.
- Snapshot save: under 1 second.
- PDF export: under 1 second.

Measure actual performance using a repeatable method. Record the test environment and results. Do not claim these targets are met without evidence.

---

# 23. IMPLEMENTATION ROADMAP

Inspect the current project before assigning phases. Follow the sequence below unless the existing progress documentation establishes that a phase has already been completed and verified.

## P4 — Calculation engine

- Implement or repair the VBA calculation engine.
- Verify geometry and quantity conversions.
- Implement Normal and Legacy behaviours.
- Verify Lodha geometry.
- Implement the pricing-profile framework.
- Record the golden-baseline discrepancy.

## P5 — Golden regression

- Test Unipro Legacy.
- Test Unipro Normal.
- Test financial-cost ratios.
- Test Lodha geometry and coating.
- Verify that pricing-profile changes alter inputs and outputs without changing formulas.
- Mark unresolved reference conflicts clearly.

## P6 — QUOTE workflow

- Support 10 independent line items.
- Implement add/edit/duplicate/remove.
- Implement quantity breakdowns.
- Renumber lines.
- Add profile selection and live summaries.

## P7 — Freight engine

- Implement freight methods A/B/C.
- Implement stacking lookup.
- Implement whole-trailer Normal mode.
- Preserve Legacy fractional-trailer behaviour.

## P8 — Plant eligibility

- Implement product/plant eligibility.
- Calculate the suggested plant.
- Support permitted overrides with reasons.

## P9 — ORS integration

- Implement the batched matrix request.
- Implement caching and error handling.
- Add mocked and live-test procedures.
- Verify coordinate order and road-distance provenance.

## P10 — Compare sheet

- Implement plant and competitor comparisons.
- Add competitor-price history.
- Add landed-price assumptions.
- Implement recommended offer and floor-price checks.

## P11 — COSTING sheet

- Reproduce the approved CostSheet layout.
- Verify component values and number formats.

## P12 — OFFER sheet

- Implement customer-facing offer output.
- Apply approved terms and reference layout.
- Verify the output against the approved Unipro offer.

## P13 — SYNOPSIS

- Implement all 16 columns.
- Verify pipe count, weight, coating areas, and totals.
- Add nominal/exact inch display.

## P14 — Pricing profiles

- Complete Company Standard, Project, Customer, and What-if profiles.
- Add versioning and snapshots.
- Test permissions.

## P15 — Geographic intelligence

- Complete plant-versus-competitor analysis.
- Complete PIN-code territories.
- Implement radius bands and evidence-based notes.

## P16 — Enquiry numbering

- Implement reliable sequence generation.
- Test duplicate prevention and supported concurrency behaviour.

## P17 — Quote history and snapshots

- Implement save/reload/history.
- Verify immutable approved snapshots.

## P18 — Master synchronization

- Synchronize applicable master data and settings.
- Validate imported and edited records.

## P19 — Multi-user architecture

- Test the actual deployment model.
- Implement appropriate permissions and concurrency protection.
- Document architectural limitations.

## P20 — Offline operation

- Test cached and manual distance handling.
- Ensure calculation and saved quotations remain available without internet access where designed.

## P21 — Exports

- Complete PDF, customer Excel, internal Excel, and synopsis exports.
- Verify confidentiality and values-only output requirements.

## P22 — Approval workflow

- Implement Draft → Calculated → Review → Approved → Issued.
- Test transitions and authorization.

## P23 — Audit trail

- Complete audit logging and review permissions.

## P24 — Self-test suite

- Provide a repeatable test runner.
- Report PASS/FAIL/NEEDS_HUMAN/NOT_TESTED.
- Never mask a failure.

## P25 — Full regression

- Run Unipro.
- Run Lodha.
- Run pricing-profile tests.
- Run Compare tests.
- Run the existing project regression suite.

## P26 — Performance

- Measure recalculation, snapshot saving, and PDF export.
- Optimize without changing commercial results.

## P27 — Multi-user trial

- Run a controlled test with representative users.
- Verify access boundaries and concurrent-operation behaviour.

## P28 — Production preparation

- Confirm backups, configuration, API-key handling, permissions, documentation, and recovery procedures.

## P29 — Go-live checkpoint

Go-live is permitted only after mandatory tests pass, unresolved commercial decisions are approved, the responsible human signs off, and the production deployment procedure is explicitly authorized.

Do not modify LIVE as part of routine development.

---

# 24. HUMAN APPROVAL ITEMS

The system cannot independently decide the following matters when authoritative evidence is missing:

1. Which Unipro workbook is the current approved commercial baseline.
2. The currently approved steel rate.
3. Yield loss and conversion charges.
4. Coating rates and their approval status.
5. The financial-cost definition and rate.
6. The approved trip rate and freight assumptions.
7. Trailer loading and stacking rules.
8. Plant/product eligibility.
9. Exact plant and competitor addresses and coordinates.
10. ORS API credentials and operational limits.
11. Windows environment configuration.
12. Shared-folder access and permissions.
13. Production user access and administrative permissions.
14. Real-world concurrent-user testing.
15. Parallel quotation verification against the existing approved process.
16. Final offer layout and commercial-term approval.
17. Final training and go-live authorization.

For each unresolved item, record:

- Decision required.
- Why it matters.
- Available evidence.
- Potential impact.
- Exact person or role needed to approve it, if known.
- Whether implementation can safely continue without the decision.

Do not block unrelated engineering tasks unnecessarily.

---

# 25. ANTIGRAVITY EXECUTION PROTOCOL

Use this operating loop for each phase.

### Step 1 — Inspect

Read the relevant code, workbook structure, documentation, and tests.

### Step 2 — Plan

List the exact files and modules to be changed, the expected behaviour, risks, and regression tests.

For complex changes, create or update an implementation plan in the project documentation before editing.

### Step 3 — Back up

Create a timestamped backup of the affected DEV files.

### Step 4 — Implement

Make small, coherent changes. Reuse existing modules and patterns.

Do not perform unrelated refactoring while fixing a targeted issue.

### Step 5 — Synchronize

Ensure the workbook's embedded VBA and any source files remain consistent according to the project's established workflow.

### Step 6 — Test

Run the smallest relevant tests first, followed by the full affected regression suite.

Use real Excel for workbook-level behaviour where required.

### Step 7 — Investigate failures

Do not merely retry a failing test or modify its expected output. Determine the cause, correct the implementation where justified, and rerun the tests.

### Step 8 — Document

Update `DOCS\PROGRESS.md` with:

- Phase.
- Date/time.
- Files changed.
- Backup path.
- Implementation details.
- Tests executed.
- Actual results.
- Outstanding failures.
- Human approvals required.
- Next action.

### Step 9 — Stop safely

Do not proceed into a destructive or commercially ambiguous change simply because the automation loop expects another iteration.

Stop the affected task if the only safe next action requires a human decision.

### Step 10 — Report

Provide a concise, evidence-based summary to the user.

Include:

- What was completed.
- What changed.
- What tests were actually run.
- What passed and failed.
- What was not tested.
- Which assumptions remain unresolved.
- The next unfinished phase.

Never state that the entire system is complete merely because the current phase passes.

---

# 26. FIRST EXECUTION TASK

Start now.

1. Inspect the currently open Antigravity workspace and identify the project root.
2. Read the existing master prompt and `DOCS\PROGRESS.md` in full.
3. Inspect the DEV workbook, VBA sources, test suite, reference workbooks, and Git status.
4. Create a project-state report before changing implementation files.
5. Resolve the Unipro golden-baseline discrepancy as far as the actual source files permit.
6. Document the result in `DOCS\GOLDEN_BASELINE.md`.
7. Identify the first incomplete phase and any blockers.
8. Create the required backup.
9. Implement the next incomplete phase.
10. Run the relevant automated and workbook-level tests.
11. Update `DOCS\PROGRESS.md`.
12. Report the exact results and continue with the next safe, unfinished task where feasible.

Do not stop after generating a plan if the required files and tools are available and the implementation can safely proceed.

Do not rebuild completed features.

Do not change LIVE.

Do not fabricate test results.

Do not silently alter approved commercial calculations.

**The final goal is a reliable, self-service Mega Pipes quotation and geographic intelligence system that preserves approved commercial logic, produces traceable quotations, supports up to 10 independently configurable pipe sizes, and provides actionable geographic analysis without confusing marketing distances with freight pricing.**