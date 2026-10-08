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
- Metres, MT, and Nos quantities preserve the exact commercial quantity while
  calculating a rounded pipe count for logistics.
- Coating cost supports rate x circumference / mass.
- Financial cost uses the documented GST-adjusted percentage equation.
- Selling-price components use positive whole-unit rounding.

The exact reference quotation totals require the complete cost, distance, and
freight inputs. They are therefore not forced into the implementation. Any
future mismatch must be reported against the responsible business rule.

VBA integration, Excel formulas, dependent dropdowns, ORS calls, and runtime
Excel validation remain pending.
