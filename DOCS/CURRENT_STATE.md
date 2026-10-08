# Current State Audit

Audit date: 2026-10-07

## Current workbook status

- Development workbook: `DEV/Mega_Pipes_Quotation_DEV.xlsm`
- Package size at audit: 56,966 bytes
- The package contains a real `xl/vbaProject.bin`.
- OOXML workbook sheets are present in the required order:
  `START`, `QUOTE`, `QUOTE_LINES`, `COSTING`, `OFFER`, `MASTERS`, `SETTINGS`,
  `CUSTOMERS`, `PLANTS`, `DISTANCE_CACHE`, and `QUOTE_LOG`.
- The workbook includes a START Form Control drawing and an embedded
  `Workbook_Open` signature.

## P3 checkpoint

The reusable P3 source modules `modMath`, `modPricing`, and `modSelfTest` are
present in the DEV source and embedded project. The explicit CURRENT/LEGACY
calculation golden gate is passed. The calculation engine must not advance to
quote-line orchestration or later phases until the remaining regression gates
are completed.
- A temporary Excel lock file (`DEV/~$Mega_Pipes_Quotation_DEV.xlsm`) was
  present during the audit; it is not a project deliverable and indicates that
  Excel may have had the workbook open.
- No worksheet formulas or worksheet data-validation records were found in the
  package audit.

### P3 calculation mode decision

The shared calculation engine now explicitly supports CURRENT and LEGACY
contracts through `RateInput.legacyMode`. CURRENT remains authoritative for
new quotations. LEGACY reproduces the original reference workbook's MT/m
coating basis, 3.14 circumference constant, common OD coating circumference,
and fractional freight quantity.

The mode-specific golden run passed in real Excel:

- LEGACY: `178657900`
- CURRENT contract: `154499250`

Evidence is in `TESTS\vba_p3_golden_20261008.txt`. Full quote-line
orchestration is now wired through `AddQuoteLine` and verified for the
current ID/Metres, active OD/Metres, MT, Nos, and legacy fixture cases.
Calculated OD `1243` remains covered by the ID-basis case. Invalid master
inputs are rejected by P2, while direct invalid geometry is rejected by
`CalcLine`.

## Existing VBA modules

Source files under `VBA/Modules`:

- `modCalculation.bas`
- `modGeometry.bas`
- `modMain.bas`
- `modNavigation.bas`
- `modValidation.bas`
- `modMasterData.bas`
- `modSequence.bas`
- `modQuoteLines.bas`
- `modSetup.bas`
- `modQuoteHeader.bas`
- `modDependentDropdowns.bas`
- `VBA/Classes/ThisWorkbook.cls` (document-module export)

Embedded VBA signatures confirmed in `xl/vbaProject.bin`:

- `modCalculation`
- `modGeometry`
- `modMain`
- `modNavigation`
- `modValidation`
- `SetupCustomerDropdown`
- `Workbook_Open`
- `EnterSystem`

The previously missing embedded `modSetup` and `ThisWorkbook` event are now
preserved as source exports. The complete source-to-embedded inventory is
documented in [`VBA_RECONCILIATION.md`](./VBA_RECONCILIATION.md). New source
modules were imported into DEV and compile/runtime smoke-tested in Excel on
2026-10-08. The evidence is in `TESTS/excel_runtime_p1_20261008.txt`.

## Existing Python scripts and tests

- `TESTS/build_dev_workbook.py`
- `TESTS/calculation_engine.py`
- `TESTS/phase1_architecture_check.py`
- `TESTS/test_calculation_engine.py`

The Python calculation suite currently covers geometry, quantity conversion,
coating cost, financial cost, selling-price rounding, and invalid geometry.
The structure test targets `DEV/Mega_Pipes_Quotation_DEV.xlsx`, not the current
macro-enabled workbook.

## Existing master data

Seeded data is embedded in the workbook:

- Products: four documented HSAW/HFW bare/coated products.
- Sizes: documented HSAW ID/OD and HFW OD seeds.
- Thickness: documented HSAW and HFW seeds with bounds.
- Grades: documented HSAW and HFW grade lists.
- Coatings: documented external and internal placeholders, including confirmed
  PU 500 micron and Food Grade Epoxy rates.
- Customer: `C001`, Unipro Techno Infrastructure Pvt Ltd, Rakesh Gupta.
- Plants: P01 and P02 OURS plants.
- Payment terms: three documented terms.
- Settings: documented rates, defaults, validation bounds, and ORS profile.

The `MASTERS` folder itself contains no source files.

## Existing tests

- Phase 1 OOXML structure validation exists but currently validates the XLSX
  build artifact.
- Python calculation tests pass in the configured environment.
- No Excel/VBA runtime test has been run in this environment.

## Missing functionality

- Excel import/compile reconciliation for the new header and line modules.
- Contact-person and delivery lookup.
- Plant, payment-term, validity, and price-basis dropdowns.
- Safe quote numbering.
- Quote-line add/remove workflow.
- Dependent product/family/basis/size/thickness/grade/coating validation.
- Excel/VBA calculation orchestration and costing population.
- Freight, plant-distance, manual-distance, ORS, and competitor architecture.
- COSTING and OFFER generation.
- Snapshot, customer/internal export, and PDF export.
- Centralized audit logging, admin workflow, protection, offline mode, and
  self-test procedures.

The new source-only stage adds master-driven header validation setup,
customer-field lookup, a local DEV enquiry sequence, validated quote-line
insertion, and helper-range dependent-dropdown architecture. These modules are
not yet embedded in the workbook.

Phase 5 source work adds one repeatable public `SetupQuoteHeader` procedure
that validates required master tables and rebuilds named-range validation for
B8, B14, B15, and B17. It is not embedded or Excel-runtime tested.

The workbook package now contains the reconciled source modules, and embedded
`modMasterData` exposes `SetupQuoteHeader`. The header integration checkpoint
is verified locally in Excel; shared-folder concurrency and later phases remain
unfinished.

## Known discrepancies

- The embedded project contains `SetupCustomerDropdown` and `Workbook_Open`
  signatures represented by the preserved source exports.
- The prompt reports that the customer dropdown found two rows; the current
  workbook customer table must be checked for blank or unintended rows before
  production use.
- The reference quotation has a documented synopsis weight discrepancy.
- Excel runtime behavior has not been claimed or marked as passed.
- The initial Excel integration stage was blocked by Trust Center settings;
  after approved AccessVBOM and trusted-location changes, the DEV P1
  integration smoke test passed. See [`VBA_RUNTIME_TEST_REPORT.md`](./VBA_RUNTIME_TEST_REPORT.md).

## Risks

- Updating the workbook without first exporting/reconciling embedded VBA can
  lose manually added procedures.
- The workbook is currently usable only after Excel has successfully opened the
  manually converted XLSM; package inspection alone is not runtime validation.
- Without worksheet data validation and centralized VBA orchestration, users can
  enter inconsistent master IDs and quote values.
- No ORS key or external API integration may be added to source or workbook.
