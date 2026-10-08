# VBA Architecture

This document describes the intended VBA source structure. Importable source
for the initial calculation and navigation layer is now present, but it is not
embedded in the workbook yet.

## Planned standard modules

- `modQuote`
- `modQuoteLines`
- `modCalculation`
- `modPricing`
- `modFreight`
- `modORS`
- `modOffer`
- `modExport`
- `modSnapshot`
- `modSequence`
- `modMasterSync`
- `modValidation`
- `modSecurity`
- `modUtilities`

Implemented source modules:

- `modMain`
- `modNavigation`
- `modGeometry`
- `modCalculation`
- `modValidation`
- `modMasterData`
- `modSequence`
- `modQuoteLines`

## Planned classes

- `clsQuote`
- `clsQuoteLine`
- `clsCalculationResult`
- `clsPlant`
- `clsCustomer`
- `clsMasterSync`

The `VBA\Modules`, `VBA\Classes`, and `VBA\Forms` folders are reserved for
source exports. UserForms, classes, full quote orchestration, exports, and
runtime event wiring remain deferred. See
[`VBA_IMPORT.md`](./VBA_IMPORT.md) for the blocked embedding procedure.
