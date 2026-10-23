# Cost-sheet automation: formulas, manual vs automatic, fixes

Source of truth: `CS Stacking 10-09-26` (Unipro 594, 23/09/2026). Verified in
`TESTS/test_costsheet_automation.py` (stacking = 176,117,500; nesting = 174,274,350).

## Formulas (per line = one column of the sheet)
| Item | Formula |
|---|---|
| OD (ID basis) | ROUNDUP(ID + 2*WT, 0); ID (OD basis) = OD - 2*WT |
| Mass MT/m | (OD-WT)*WT*0.0246615/1000 |
| Qty MT / pipes | metres*mass ; ROUNDUP(metres/12) |
| Pre-cost | (Steel+GradeExtra)*(1+Yield%) + Conversion + Testing + Margin |
| Ext / Int coating Rs/MT | Rate * 3.14 * (OD/1000) / mass  (sheet uses OD for BOTH; setting IntCoatingBasis=ID uses ID) |
| Freight Rs/MT | ROUNDUP(TripRate * trips / lineMT) + 100 ; trips = (metres/12)/pipesPerTrailer (fractional) |
| Financial cost | typed value, else k*base/(1-k), k = Fin% * (1+GST), base = pre+ext+int+freight |
| Per-metre prices | ROUNDUP(value Rs/MT * mass) for bare(pre+FC), ext, int, freight; summed |
| Line value | price/m * metres; GST 18% on the sum |
| "total" column | SUMPRODUCT(row, lineMT)/total MT |

## Manual (typed) / automatic
Manual per quote or line: size, WT, basis, qty, grade, coating choice, trip rate (transporter quote),
optional typed FC, steel rate/conversion overrides, optional selling-price override, optional loading-plan trips.
Automatic: everything in the table above, pipes per trailer (tblStacking, else payload/pipe weight),
coating rate (from the coating master), COSTING sheet, totals, GST.

## SETTINGS keys (auto-created by `EnsureCostPolicySettings`; defaults = reference sheet)
`FreightTrailerRounding` FRACTIONAL|WHOLE, `IntCoatingBasis` OD|ID, `FCMethod` ALGEBRAIC|SIMPLE,
`DefaultExtCoatingRate` 500, `DefaultIntCoatingRate` 250.

## Defects found and fixed
1. Rates/coatings/trip rate/pipes-per-trailer rebuilt in 4 places with hard-coded 500/250/3.14/4; now one loader (`modPricing.LoadLineInputsFromRow`).
2. Editing a line dropped its typed FC and overrides; recalculation ignored the coating chosen at add-time (product-ID hack). Fixed by the loader.
3. Add-line charged BOTH coatings when only one was chosen. Now per side.
4. Suggested FC used a shortcut (1,019) vs the sheet's own fixed point (1,031). Now algebraic.
5. Python engine switched method when steel==62000 or conversion==4000. Now explicit policy (`REFERENCE_POLICY`).
6. COSTING sheet: ROUND instead of ROUNDUP per metre (rows did not sum, e.g. 22,350 vs 22,351), 3.14159265 coating factor, "total" column faked with `*0.08` / FC% constants, trailers shown as 153 whole while priced on 151.04 trips. Now taken straight from CalcLine.
7. Pipes-per-trailer table `tblStacking` was never used.
8. Reference sheet itself: Synopsis totals sum rows 11:14 and miss line 4; FC typed 1024 vs its own check cell 1031 (stale).

## Not changed / open
- Trip rate stays manual (no destination PIN on QUOTE; COMPARE uses fixed 782/1821 km).
- Pooled "nesting" trips: Python `calc_costsheet(pooled_trips=...)`; VBA supports per-line optional column `ManualTrips` only.
- VBA could not be compiled here: import the .bas files, run `RunCostSheetParityTests`.
- Old VBA self-tests (modSelfTest) use the old defaults and are untouched.
