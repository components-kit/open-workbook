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

The agent surface also normalizes common shorthand shapes such as `{ "column": "Status", "value": "Open" }`, `{ "column": "Status", "criterion": "Open" }`, `{ "column": "Status", "filterType": "text", "value": "Open" }`, `values.filter`, and top-level `values.column` plus `values.value`, `values.criterion`, or `values.criteria` before calling the Office.js table filter API.

Structured table sort previews accept `values.sortBy` or `values.column` plus `values.direction` or `values.ascending`; text-only amount sort inference remains a fallback. Direction phrases such as "lowest to highest", "highest to lowest", "A to Z", and "Z to A" are normalized.

To show all data again, use `intent.action: "clear_table_filters"` for a structured table. Clear/remove/reset filter phrasing sent through `filter_range` also clears table filters when the resolved target is a table. For ordinary worksheet ranges, clear/remove/reset filter phrasing emits `range.clear_autofilter`.

To add filter controls to all columns in an ordinary range, use `intent.action: "filter_range"` with a resolved `target.sheetName` and `target.range`; this enables the Excel autofilter UI across the target range. Criteria-based filtering remains table-scoped in this pass.

When one request combines table filters and sorting, use `intent.action: "apply_table_view"` with `values.filters` and `values.sort.fields`. Include `values.clearFilters` or `values.clearSort` when the new view should reset existing state first. The backend stores one preview, applies one direct table transaction, and returns one proof instead of splitting filter and sort into separate operations.

## Safety

Mutating table/filter/sort tools create a backup before execution. The backend backs up the current table range when mutating an existing table, or the target range when creating/copying a structure.
Column reorder keeps the existing Excel table object and uses native copy operations through a temporary worksheet, avoiding full-table value transfer through MCP.

## Template Behavior

Table template validation currently compares live table metadata with registered template table fingerprints. Deeper table-structure repair will build on the same registered template payload used by style and formula repair.
