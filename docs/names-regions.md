# Names and Regions

Names and regions give agents stable workbook targets that survive normal sheet edits better than ad hoc range addresses.

## Named Items

`excel.names.*` uses native Office.js named-item APIs.

Supported tools:

- `excel.names.list`: lists workbook and worksheet scoped names.
- `excel.names.get`: returns one named item, including formula/value metadata and range address when the name resolves to a range.
- `excel.names.create`: creates a workbook or worksheet scoped name from a formula or range reference.
- `excel.names.update`: updates formula/reference, comment, or visibility.
- `excel.names.delete`: deletes a named item.

When `sheetName` is provided, the operation uses that worksheet's scoped name collection. Without `sheetName`, it uses workbook scope.

## Regions

`excel.region.*` is an Open Workbook registry over workbook areas that agents can address by intent, not only by A1 notation.

Supported tools:

- `excel.region.detect`: returns registered regions plus candidates from Excel named ranges and sheet used ranges.
- `excel.region.register`: registers a reusable region and can optionally create a matching Excel named range.
- `excel.region.list`: lists runtime-registered regions.
- `excel.region.get`: resolves a registered region or an Excel named range.
- `excel.region.clear_values`: clears values while preserving formatting.
- `excel.region.write_values`: writes values while preserving formatting.
- `excel.region.fill`: optionally clears first, then writes values.

Region mutations compile to normal `range.*` batch operations, so they use the same backup, telemetry, and rollback path as direct range writes.

## Current Limits

The region registry is runtime-local for now. Persisting regions to workbook custom XML or local project storage should be a future milestone before relying on regions across MCP restarts.
