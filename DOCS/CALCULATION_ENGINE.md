# Calculation Engine Status

The workbook currently contains structural tables and seeded reference data.
The independently testable reference implementation is
`TESTS/calculation_engine.py`; it does not write workbook formulas or pretend
to be embedded VBA.

Implemented reference rules:

- OD basis uses `OD = Size` and `ID = OD - 2 * WT`.
- ID basis uses `ID = Size` and rounds `OD = ID + 2 * WT` upward to a whole
  millimetre.
- Mass uses `((OD - WT) * WT * 0.0246615)` kg/m.
- `massKgPerM` is kg/m and `massMtPerM = massKgPerM / 1000` is MT/m.
- Metres, MT, and Nos quantities preserve the exact commercial quantity while
  calculating a rounded pipe count for logistics.
- `pre`, financial cost, coating cost, and freight are expressed in Rs/MT.
- Selling-price-per-metre components convert Rs/MT to Rs/m using MT/m:
  `BarePerM`, `ExtPerM`, `IntPerM`, and `FreightPerM`.
- `TotalPerMT` remains the rounded Rs/MT total, while `LineValue` uses the
  exact commercial metres or MT quantity.
- Coating cost supports rate x circumference / mass.
- Financial cost uses the documented GST-adjusted percentage equation.
- Selling-price components use positive whole-unit rounding.

The exact reference quotation totals are implemented as the explicit LEGACY
mode fixture. CURRENT remains the authoritative contract for new quotations.

P3 VBA source modules now exist in `VBA/Modules/modMath.bas` and
`VBA/Modules/modPricing.bas`, with `modSelfTest.bas` providing a golden-test
entry point. The kg/m-to-MT/m unit mismatch in the four per-metre components
was corrected without changing geometry, freight/MT, financial cost,
`TotalPerMT`, or golden values. The golden test now reaches numerical
comparison, but the unchanged legacy and normal totals still fail; full
Excel-runtime validation remains blocked.

## Explicit calculation modes

The shared VBA engine supports two explicit contracts through
`RateInput.legacyMode`:

- **CURRENT** is authoritative for new quotations. Coating uses kg/m for
  Rs/MT coating cost, freight uses rounded pipes and trailers, and the
  current financial and selling-price rules remain unchanged.
- **LEGACY** reproduces the original reference workbook. Coating uses MT/m
  and the workbook's common OD circumference factor for both coating rows,
  freight uses the fractional pipe/trailer quantity before rounding, and the
  legacy financial fixture remains explicit.

The public `LineResult.MassPerM` remains kg/m in both modes. `massMtPerM` is
an internal conversion and is used by the legacy coating and selling-price
rules where required. The two modes share geometry, mass, quantity, and
result contracts; only the confirmed historical rules branch explicitly.

## Golden fixture forensic result

The diagnostic golden run agrees with the reference workbook:

- LEGACY per-metre values: `22681`, `23914`, `24748`, `25580`
- LEGACY total: `178657900`
- CURRENT contract total: `154499250`
- Python reference tests: `9 passed`
- VBA tolerance: `<1`

The historical fixture is therefore valid as an explicit LEGACY contract and
is not used to redefine CURRENT production pricing.
