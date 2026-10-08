# Workbook Architecture

## Purpose

The development workbook is the structural foundation for the Mega Pipes
quotation system. Phase 1 establishes sheet responsibilities, table names, and
reserved areas without implementing quotation calculations or production
workflows.

## Sheet responsibilities

- `START`: development landing page and implementation status.
- `QUOTE`: reserved quotation header/input area.
- `QUOTE_LINES`: transactional line-item table.
- `COSTING`: reserved costing sections.
- `OFFER`: reserved offer sections.
- `MASTERS`: master-data registry and reserved master tables.
- `SETTINGS`: structured system settings table.
- `CUSTOMERS`: customer master table.
- `PLANTS`: plant master table.
- `DISTANCE_CACHE`: distance-cache storage table.
- `QUOTE_LOG`: audit-log storage table.

## Master and transactional data

Master data is separated from transactional data. Master tables provide
reference data for products, dimensions, grades, coatings, plants, customers,
payment terms, and settings. `tblQuoteLines` stores quote transaction data and
uses IDs rather than embedding master records.

Lookup logic and calculation formulas are not implemented in this phase.

## Naming conventions

Structured tables use the `tbl` prefix and descriptive PascalCase names, such
as `tblQuoteLines`, `tblCustomers`, and `tblSettings`.

Identifiers use descriptive names such as `QuoteID`, `LineID`, `ProductID`,
`CustomerID`, `PlantID`, and `LogID`. Row numbers are not permanent IDs.

## Future architecture

The intended flow is:

`MASTER DATA -> QUOTE INPUT -> QUOTE LINE DATA -> CALCULATION ENGINE -> COSTING -> OFFER`

Future exports may produce customer-facing output, internal output, PDFs, and
quote snapshots. These features are not implemented in Phase 1.

Future VBA responsibilities are documented separately in
[`VBA_ARCHITECTURE.md`](./VBA_ARCHITECTURE.md).

## DEV/LIVE separation

All Phase 1 work is confined to `DEV`. The `LIVE` workbook is not modified.
Promotion to `LIVE` requires separate testing and explicit approval.
