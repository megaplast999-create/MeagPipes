# VBA Runtime Test Report

## STATUS

**P1 DEV CHECKPOINT: PASS**

**P2 DEV CHECKPOINT: PASS**

**P3 CHECKPOINT: PASS (VERIFIED 2026-10-12)**

### 2026-10-12 Recovery and Verification Audit

On 2026-10-12, a comprehensive diagnosis and recovery was conducted on `DEV\Mega_Pipes_Quotation_DEV.xlsm`:
1. **Compilation Diagnosis**:
   - `Debug -> Compile VBAProject` was initially greyed out because Excel loaded pre-compiled p-code streams.
   - Upon project dirtying, `Compile VBAProject` correctly became enabled.
   - In the prior embedded workbook state, `modQuoteHeader` had accidentally been overwritten with a duplicate copy of `modMasterData` procedures (`SetupQuoteHeader`, `SetupPlantDropdown`, etc.), causing procedure name ambiguity and removing `InitializeNewQuote` and `GetConfiguredValidity`.
   - `modSelfTest` in the embedded workbook was missing `RunP3InvalidMasterGuardTest`.
2. **Repairs Applied**:
   - Replaced embedded `modQuoteHeader` with authoritative source from `VBA\Modules\modQuoteHeader.bas` (restoring `InitializeNewQuote`, `HandleCustomerSelection`, `GetConfiguredValidity` and removing duplicates).
   - Replaced embedded `modSelfTest` with authoritative source from `VBA\Modules\modSelfTest.bas` (adding `RunP3InvalidMasterGuardTest`).
   - Clean VBE compilation executed successfully; `Compile VBAProject` disabled after compilation.
3. **Runtime Test Verification**:
   - `RunP3GoldenTests`: PASS (LEGACY 178,657,900, CURRENT 178,801,050, diff = 0).
   - `RunP3InvalidGeometryTest`: PASS.
   - `RunP3InvalidMasterGuardTest`: PASS ("Size is not available for the selected family and basis.").
   - `InitializeNewQuote`: PASS.
   - `SetupMasterDrivenHeader`: PASS.
   - `Sheet2.Worksheet_Change` -> `HandleCustomerSelection`: PASS (customer lookup verified).
   - Python test suite: 9 passed in pytest.

4. **Normal Mode Coating Unit & Golden Test Fixture Recovery (16:39)**:
   - Corrected normal-mode coating denominator in `modPricing.bas` (`massKgPerM` -> `massMtPerM`).
   - Aligned normal-mode test expectation to audited Option 1 (`Line 4 = 25581/m`, `Total = 178801050`).
   - Clean VBE compilation verified (Compile VBAProject disabled).
   - All tests PASS in real Excel 16.0 with zero difference (`TESTS\vba_p3_golden_20261008.txt`).

### Latest P3 mode verification

On 2026-10-08, `modMath`, `modPricing`, and `modSelfTest` were backed up,
re-imported, and compiled in the DEV workbook. The shared engine now
explicitly supports CURRENT and LEGACY contracts. `RunP3GoldenTests` produced
passing evidence:

- LEGACY total: `178657900`
- CURRENT contract total: `154499250`
- Tolerance: `<1`
- Python tests: `9 passed`

This does not claim full quote-line orchestration or P1/P2 regression. No
commit or phase tag has been created.

### Quote-line orchestration

`AddQuoteLine` now builds `QuoteLineInput`, loads `RateInput`, calls the
shared `CalcLine`, and writes available calculated fields back to
`tblQuoteLines`. The integrated runtime evidence covers CURRENT ID/Metres,
CURRENT MT, CURRENT Nos, CURRENT active OD, LEGACY line 1, and the full
four-line LEGACY fixture. The table has no columns for trailers or individual
coating/financial components; those remain in `LineResult` and are not
silently mapped to unrelated columns.

The active OD runtime case uses OD `1219`. Calculated OD `1243` remains
covered by the ID-basis case. The P2 invalid-master guard and direct
`CalcLine` invalid-geometry guard both pass in test-only wrappers.

P3 source modules were imported into the DEV workbook and the VBE compile
command was executed. A qualified call to `modSelfTest.RunP3GoldenTests`
does not return a result; Excel enters VBE break mode and no
`TESTS\vba_p3_golden_20261008.txt` evidence file is produced. P3 is not
claimed as passed and no phase tag or commit was created.

The initial Excel/VBA checkpoint was blocked by Trust Center settings. After
approved AccessVBOM and trusted-location changes, Excel 16.0 was opened
through pywin32, the project reported 23 components, six drifted source
modules were imported into DEV, and the VBE compile command was executed.

The workbook package later changed from 56,966 bytes to 94,838 bytes and now
contains the five pending module streams. This indicates that an import/save
occurred outside the available automation inspection path. It is not treated
as a compile or runtime pass.

Phase 5 source preparation added one public `SetupQuoteHeader` procedure to
`modMasterData`. It checks required master tables and rebuilds named-range
validation for B8, B14, B15, and B17. This source procedure is not embedded or
runtime tested.

## BACKUP

Created before the attempted import:

- `BACKUP\pre_vba_import_20261007_175511.xlsm`

## IMPORTED MODULES

The package now contains module streams for:

- `modMasterData.bas`
- `modSequence.bas`
- `modQuoteLines.bas`
- `modQuoteHeader.bas`
- `modDependentDropdowns.bas`

The required public `SetupQuoteHeader` procedure is present in the source
`modMasterData.bas` but was not found in the extracted embedded module;
`SetupMasterDrivenHeader` remains the embedded procedure name.

The existing embedded modules were not blindly replaced:

- `ThisWorkbook`
- `modSetup`
- `modCalculation`
- `modGeometry`
- `modMain`
- `modNavigation`
- `modValidation`

## COMPILE RESULT

**EXECUTED**

Excel's VBE compile command was executed after import. The subsequent local
header and quote-line smoke test completed successfully.

## RUNTIME TEST RESULT

**PASS (P1 DEV SMOKE TEST)**

The verified tests are recorded in
`TESTS/excel_runtime_p1_20261008.txt`. This is not a claim that the full
ten-test production runtime suite has passed.

## FAILED TESTS

No P1 smoke-test failures.

The package-level audit did verify that the workbook remains a genuine XLSM
with `xl/vbaProject.bin`, the expected sheets/tables, and the original
embedded VBA signatures. It did not verify Excel execution.

## SOURCE/EMBEDDED RECONCILIATION

- Embedded baseline remains authoritative.
- `modSetup.SetupCustomerDropdown` and `ThisWorkbook.Workbook_Open` were
  preserved in source before this stage.
- No new source module was imported into the embedded project.
- B14 Plant, B15 Payment Terms, and B17 Price Basis remain unclaimed as
  runtime-tested Excel dropdowns.
- Embedded module streams now include the five pending modules, but compile
  status and the exact embedded source are not verified.
- No duplicate source procedures were found.
- No unintended embedded import occurred.

## WORKBOOK PATH

`DEV\Mega_Pipes_Quotation_DEV.xlsm`

## VBA PROJECT PRESENT

Yes. Package inspection confirms `xl\vbaProject.bin`.

## LIVE CHANGES

None.

## Required next boundary action

The pending modules must be imported and compiled from within desktop Excel's
VBE after the project is safely accessible. Compilation and runtime results
must then be recorded by a subsequent audit; this report intentionally does
not claim them.
