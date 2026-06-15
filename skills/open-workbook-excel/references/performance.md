# Performance

Fast Excel automation is part of correctness. Slow workflows push agents toward unsafe shortcuts.

## Defaults

- Prefer bulk Office.js operations over per-cell actions.
- Use 2D matrices for values, formulas, and number formats.
- Use table APIs for table-shaped work.
- Group writes by workbook, sheet, and contiguous range.
- Read only requested properties.
- Use `excel.runtime.get_capabilities` before trying expensive or host-limited features.

## Reads

- For user-visible reporting, read display text.
- For computation, read raw values.
- For formula review, read formulas or formula patterns.
- For formatting review, read number formats or styles only for the relevant range.
- For table analysis, read the table instead of a sheet-sized range.

Avoid workbook-wide reads unless the task is search, validation, audit, or discovery.

## Writes

- Use one batch or plan for related edits.
- Preflight large generated batches before applying them.
- Keep matrix shapes exact: rows and columns must match the target range.
- Let Open Workbook chunk large values/formulas/number formats through safe row-based chunk plans.
- Avoid alternating read/write/read/write loops. Read once, compute, apply once, validate once.
- If work is queued or applying, report the progress message to the user and wait/poll the job or transaction rather than starting parallel mutations.

For very large writes, consider a plan preview first so the agent can expose scope, chunk count, and rollback coverage before applying.

Automatic timeout retry is limited to style-only batches because repeating the same style is safe. Values, formulas, and number formats should be chunked before execution when preflight recommends it. For tables, pivots, charts, and structure, inspect transaction status before retrying.

## Formulas

- Prefer formula pattern tools for repeated formula layouts.
- Trace precedents/dependents before editing source ranges used by reports, charts, pivots, or formulas.
- Recalculate with `excel.formula.recalculate` or `excel.workbook.calculate` when the workflow requires fresh computed values.

## Telemetry

Read tool results for duration, sync count, payload bytes, cells read/written, range count, chunk count, engine, and warnings. Use that telemetry to explain unusually slow or partial work.
