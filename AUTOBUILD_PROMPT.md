# MEGA PIPES QUOTATION WORKBOOK — AUTOBUILD PROMPT
# Version 3.1 — Includes modCompare.bas and Full COMPARE Sheet Intelligence

## VBA MODULE LIST
- `modMath.bas`: Precision decimal arithmetic, RoundUpD, MassMtPerM, HexToLong
- `modPricing.bas`: LineInput, RateInput, LineResult, CalcLine pricing engine
- `modFreight.bas`: Stacking rules, trailer allocation, trip rate models A/B/C
- `modPlants.bas`: Eligibility, nearest plant selection, plant override audit
- `modORS.bas`: OpenRouteService HGV matrix, caching, coordinates formatting
- `modCompetitors.bas`: Competitor pricing model, floor margin protection, stale badge
- `modCompare.bas`: Competitive intelligence dashboard (Sections A, B, C), live ex-mill price entry, dynamic row threat highlighting, values-only guarantee
- `modQuote.bas`: Quotation workflow, header reading, live calculation synchronization
- `modExport.bas`: COSTING, OFFER, SYNOPSIS, Customer/Internal Excel sanitization, PDF export
- `modSequence.bas`: Sequential enquiry allocation, file locking, offline TEMP fallback
- `modLog.bas`: Audit logging, snapshot archival and restoration
- `modAdmin.bas`: Master data sync, password authentication
- `modUI.bas`: Event handling, form navigation, live feedback banners
- `modProtection.bas`: Cell locking, sheet protection management
- `modSelfTest.bas`: Unified self-test suite and golden regression validation

## COMPARE SHEET SPECIFICATION
`modCompare.bas` builds the COMPARE sheet:
1. Reads our selected plant and its km from the current quote.
2. Reads all competitor distances from tblDistanceCache.
3. Reads ex-mill prices from tblCompPrices (matching product family).
4. Sorts by km ascending.
5. Applies the row colour rules at runtime (never hardcoded row numbers).
6. Calculates the recommended offer.
7. Writes all values only (no formulas in the COMPARE sheet).

The COMPARE sheet is refreshed automatically when:
- The distance lookup is completed (Find Distance button).
- A competitor ex-mill price is entered or changed.
- The selected plant is changed.
