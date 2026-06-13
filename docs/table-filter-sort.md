# Tables, Filters, And Sort

Structured table operations use native Office.js table APIs instead of range-only mutations.

## Implemented Flow

- Read table metadata, headers, values, formulas, text, number formats, filters, and sort state.
- Project table reads by selected columns, row offset/limit, and requested facets to keep large-table payloads small.
- Create, resize, reorder columns, append rows, update rows, clear constants while keeping formulas, set total row, set style, and copy table structure.
- Apply, clear, preserve, and validate table filters through Office.js filter criteria.
- Apply, clear, and preserve table sort state through Office.js sort fields.

## Safety

Mutating table/filter/sort tools create a backup before execution. The backend backs up the current table range when mutating an existing table, or the target range when creating/copying a structure.
Column reorder keeps the existing Excel table object and uses native copy operations through a temporary worksheet, avoiding full-table value transfer through MCP.

## Template Behavior

Table template validation currently compares live table metadata with registered template table fingerprints. Deeper table-structure repair will build on the same registered template payload used by style and formula repair.
