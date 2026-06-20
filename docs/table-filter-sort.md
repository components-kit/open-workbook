# Tables, Filters, And Sort

Structured table operations use native Office.js table APIs instead of range-only mutations.

## Implemented Flow

- Read table metadata, headers, values, formulas, text, number formats, filters, and sort state.
- Project table reads by selected columns, row offset/limit, and requested facets to keep large-table payloads small.
- Create, resize, reorder columns, append rows, update rows, clear constants while keeping formulas, set total row, set style, and copy table structure.
- Apply, clear, and validate table filters through Office.js filter criteria.
- Apply and clear table sort state through Office.js sort fields.
- Apply combined table views with filters, sort fields, and optional clear-filter/clear-sort flags in one Office.js transaction.

For `excel.agent.run`, table filter previews use `values.filters`. The canonical filter shape is:

```json
{ "column": "Status", "criteria": { "filterOn": "Values", "values": ["Open"] } }
```

The agent surface also normalizes common shorthand shapes such as `{ "column": "Status", "value": "Open" }`, `{ "column": "Status", "criterion": "Open" }`, and `{ "column": "Status", "filterType": "text", "value": "Open" }` before calling the Office.js table filter API.

Structured table sort previews accept `values.sortBy` or `values.column` plus `values.direction` or `values.ascending`; text-only amount sort inference remains a fallback.

When one request combines table filters and sorting, use `intent.action: "apply_table_view"` with `values.filters` and `values.sort.fields`. The backend stores one preview, applies one direct table transaction, and returns one proof instead of splitting filter and sort into separate operations.

## Safety

Mutating table/filter/sort tools create a backup before execution. The backend backs up the current table range when mutating an existing table, or the target range when creating/copying a structure.
Column reorder keeps the existing Excel table object and uses native copy operations through a temporary worksheet, avoiding full-table value transfer through MCP.

## Template Behavior

Table template validation currently compares live table metadata with registered template table fingerprints. Deeper table-structure repair will build on the same registered template payload used by style and formula repair.
