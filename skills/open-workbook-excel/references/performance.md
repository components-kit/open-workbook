# Performance

Fast Excel automation is part of correctness. Slow workflows push agents toward unsafe shortcuts.

## Defaults

- On the public surface, use `excel.agent.run`; the backend owns compact reads, summaries, proof, resource handles, and response budgeting.
- Prefer bulk Office.js operations over per-cell actions.
- Use 2D matrices for values, formulas, and number formats.
- Use table APIs for table-shaped work.
- Group writes by workbook, sheet, and contiguous range.
- Read only requested properties.
- Backend routing should start unknown targets with lookup capabilities, then inspect one candidate.
- Backend routing should start known large scopes with compact summaries and schemas.
- Use `excel.runtime.get_capabilities` before trying expensive or host-limited features.

## Reads

- For user-visible reporting, read display text.
- For computation, read raw values.
- For formula review, read formulas or formula patterns.
- For formatting review, read number formats or styles only for the relevant range.
- For table analysis, read the table instead of a sheet-sized range.
- On the public surface, use `excel.agent.run` `mode: "find"` or `mode: "answer"` for unknown targets and exploratory reads.
- Internally, use lookup before reading cells and use compact range/table reads before full reads.

Avoid workbook-wide reads unless the task is search, validation, audit, or discovery.

Compact reads return payload/token telemetry, truncation status, and continuation metadata. Page only when more rows or columns are needed.
When a compact response includes `resourceUri`, fetch it only if the full detail is required.

## Writes

- Use one batch or plan for related edits.
- On the public agent surface, group related range value edits with `values.patches` in one `preview_update`, then call `apply_update` once for the returned operation.
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
For compact reads, also inspect `estimatedTokens`, `truncated`, and `nextPage` before requesting more workbook data.
For mutation results, prefer the returned `compactProof` summary over reading changed ranges back into the model.
