# P1/P2 VBA Dependency Audit

Audit date: 2026-10-08  
Checkpoint audited: `382f2d1` (`p2-dropdowns`)

| Procedure | Module | Callers | Purpose | Used by P1/P2 | Duplicate/overlap risk | Recommended action |
|---|---|---|---|---:|---|---|
| `SetupMasterDrivenHeader` | `modMasterData` | `InitializeNewQuote` | Applies customer, plant, payment-term, and price-basis validation | Yes, through New Quote | Low; distinct from `SetupQuoteHeader` because it uses direct table addresses | Preserve and reuse for New Quote |
| `SetupQuoteHeader` | `modMasterData` | Direct Excel runtime test; setup entry point | Validates required masters, creates named ranges, and applies header validation | Yes, directly verified | Low; complements rather than duplicates `SetupMasterDrivenHeader` | Preserve as the robust named-range setup entry point |
| `RefreshCustomerFields` | `modMasterData` | `HandleCustomerSelection`; direct Excel runtime test | Looks up the selected customer and writes contact/address | Yes | Low | Preserve; call after customer selection |
| `InitializeNewQuote` | `modQuoteHeader` | Direct Excel runtime test; intended New Quote action | Allocates enquiry number, date, validity, default unit, and header setup | Yes | Low | Preserve; extend only through explicit orchestration |
| `HandleCustomerSelection` | `modQuoteHeader` | Intended customer-change action | Delegates customer lookup to `RefreshCustomerFields` | Yes, indirectly verified | Low | Preserve |
| `GetConfiguredValidity` | `modQuoteHeader` | `InitializeNewQuote` | Reads `ValidityDays` from `tblSettings` | Yes | Low | Preserve |
| `SetupDependentDropdownArchitecture` | `modDependentDropdowns` | Direct P2 runtime test; intended setup action | Builds product/family/basis/coating helper lists and applies validation | Yes | Low | Preserve |
| `SetupQuoteLineDropdowns` | `modDependentDropdowns` | `SetupDependentDropdownArchitecture`; `RequestNewSize` | Builds family/basis-dependent size, WT, and grade lists | Yes | Low | Preserve and call after relevant input changes |
| `RequestNewSize` | `modDependentDropdowns` | Direct P2 runtime test; intended Add Size action | Validates range, adds pending size, writes request CSV, refreshes size lists | Yes | Low | Preserve; later connect to a UI |
| `AddQuoteLine` | `modQuoteLines` | Direct P1/P2 runtime tests; intended Add Line action | Validates master-backed line inputs and writes a DRAFT row | Yes | Low | Preserve; P3 calculation should consume its stored values |
| `CalculateOD` | `modGeometry` | No P1/P2 caller | Calculates OD from OD or ID basis | No | Medium with future `modPricing` geometry logic | Reuse from P3 calculation orchestration; do not duplicate |
| `CalculateID` | `modGeometry` | No P1/P2 caller | Calculates ID from OD or ID basis | No | Medium with future pricing logic | Reuse |
| `CalculateMassKgPerM` | `modGeometry` | No P1/P2 caller | Calculates kg/m from OD and WT | No | Medium with future mass logic | Reuse |
| `ConvertQuantityToMetres` | `modCalculation` | No P1/P2 caller | Converts Metres, MT, or Nos to metres | No | High with future pricing quantity logic | Preserve compatibility wrapper; route new code through it or a shared core |
| `CalculatePipeCount` | `modCalculation` | No P1/P2 caller | Rounds pipe count upward | No | Medium | Reuse |
| `CalculateWeightMT` | `modCalculation` | No P1/P2 caller | Converts metres and kg/m to MT | No | Medium | Reuse |
| `CalculateCoatingCostPerMT` | `modCalculation` | No P1/P2 caller | Calculates coating cost per MT using a diameter | No | High because internal coating needs ID in normal mode | Preserve; add explicit internal/external P3 logic elsewhere without changing this contract |
| `CalculateFinancialCost` | `modCalculation` | No P1/P2 caller | Solves financial cost from base cost and GST-adjusted percentage | No | Low | Reuse |
| `IsValidPipeLength` | `modValidation` | `AddQuoteLine` | Checks 6–13 m pipe length | Yes | Low | Preserve |
| `IsValidQuantity` | `modValidation` | `AddQuoteLine` | Checks quantity > 0 | Yes | Low | Preserve |
| `IsValidGeometry` | `modValidation` | No P1/P2 caller | Checks positive and ordered geometry | No | Low | Reuse in P3 |

## Conclusion

P1/P2 dependencies are clean. There are no procedures that should be
deleted, renamed, or replaced. Existing geometry and calculation procedures
are currently unused by the verified header/dropdown/line-entry workflow and
are available for P3 reuse. P3 should add reusable pricing types and
orchestration while preserving the existing public function contracts.
