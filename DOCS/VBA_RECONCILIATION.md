# Embedded VBA Reconciliation

Audit date: 2026-10-07

## Method

The embedded project was read from
`DEV/Mega_Pipes_Quotation_DEV.xlsm` with `oletools.olevba`. The source tree was
searched with a procedure-name inventory over every `.bas` and `.cls` file.
This is package/source inspection only; Microsoft Excel compilation and runtime
execution were not performed.

## Authority decision

The embedded VBA project is the authoritative baseline before further
implementation. It is the only version known to have been opened and embedded
by Microsoft Excel, and it contains procedures not present in the old source
exports. No embedded code was deleted or overwritten during this audit.

After the exported baseline is reconciled, newly added source modules are the
candidate implementation for the next import/review stage. They are not
considered active workbook functionality until imported and compiled in Excel.

## Procedure inventory

| Module | Procedure | Source exists? | Embedded exists? | Duplicate? | Conflict? | Authoritative version |
|---|---|---:|---:|---:|---:|---|
| `ThisWorkbook` | `Workbook_Open` | Yes, `VBA/Classes/ThisWorkbook.cls` | Yes | No | No | Embedded baseline until Excel import |
| `modCalculation` | `ConvertQuantityToMetres` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modCalculation` | `CalculatePipeCount` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modCalculation` | `CalculateWeightMT` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modCalculation` | `CalculateCoatingCostPerMT` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modCalculation` | `CalculateFinancialCost` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modGeometry` | `CalculateOD` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modGeometry` | `CalculateID` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modGeometry` | `CalculateMassKgPerM` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modMain` | `ShowStartScreen` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modNavigation` | `GoToStart` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modNavigation` | `EnterSystem` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modNavigation` | `GoToQuote` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modNavigation` | `GoToCosting` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modNavigation` | `GoToOffer` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modValidation` | `IsValidPipeLength` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modValidation` | `IsValidQuantity` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modValidation` | `IsValidGeometry` | Yes | Yes | No | No | Embedded/source identical baseline |
| `modSetup` | `SetupCustomerDropdown` | Yes, newly exported | Yes | No | No | Embedded baseline; source export preserves it |
| `modMasterData` | `SetupPlantDropdown` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `SetupPaymentTermsDropdown` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `SetupPriceBasisDropdown` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `SetupMasterDrivenHeader` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `RefreshCustomerFields` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `ApplyListValidation` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `TableColumnAddress` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `FindTableRow` | Yes, new | No | No | No | Source candidate pending import |
| `modMasterData` | `TableColumnIndex` | Yes, new | No | No | No | Source candidate pending import |
| `modSequence` | `AllocateEnquiryNumber` | Yes, new | No | No | No | Source candidate pending import |
| `modSequence` | `FindSettingRow` | Yes, new | No | No | No | Source candidate pending import |
| `modSequence` | `GetSettingValue` | Yes, new | No | No | No | Source candidate pending import |
| `modQuoteLines` | `AddQuoteLine` | Yes, new | No | No | No | Source candidate pending import |
| `modQuoteHeader` | `InitializeNewQuote` | Yes, new | No | No | No | Source candidate pending import |
| `modQuoteHeader` | `HandleCustomerSelection` | Yes, new | No | No | No | Source candidate pending import |
| `modQuoteHeader` | `GetConfiguredValidity` | Yes, new | No | No | No | Source candidate pending import |
| `modDependentDropdowns` | `SetupDependentDropdownArchitecture` | Yes, new | No | No | No | Source candidate pending import |
| `modDependentDropdowns` | `BuildHelperList` | Yes, new | No | No | No | Source candidate pending import |
| `modDependentDropdowns` | `ApplyNamedListValidation` | Yes, new | No | No | No | Source candidate pending import |
| `modDependentDropdowns` | `FindTable` | Yes, new | No | No | No | Source candidate pending import |

The embedded project also contains eleven empty worksheet class streams
(`Sheet1` through `Sheet11`). They have no procedures and require no source
export for this stage.

## Conflicts and duplicate risk

- No duplicate procedure names exist in the source tree.
- No source procedure conflicts with an embedded procedure.
- `SetupCustomerDropdown` is now exported as `modSetup.bas`; it was not
  recreated under another module name.
- `Workbook_Open` is stored in the special `ThisWorkbook` document module, not
  as a standard module.
- The new modules are absent from the embedded project and must remain marked
  pending until Excel import and compile verification.

## Required future reconciliation action

When Excel is available for the import stage:

1. Preserve the current workbook backup.
2. Import only the source components marked “Source candidate pending import”.
3. Do not import `modSetup` or the `ThisWorkbook` class a second time; they are
   already represented in the embedded baseline unless replacing them after a
   deliberate source review.
4. Compile the VBA project in Excel.
5. Re-export all standard modules and the `ThisWorkbook` event to `VBA`.
6. Re-run the source/embedded inventory and package audit.

Until those steps occur, Excel runtime status remains **PENDING**.
