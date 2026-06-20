# Performance Contract

Open Workbook treats speed as a correctness requirement because slow workbook automation encourages unsafe shortcuts.

## Rules

- All writes route through the batch compiler.
- Individual write tools are syntactic sugar over batch operations.
- Do not call `context.sync()` inside per-cell loops.
- Use 2D array assignments for values and formulas.
- Load only requested Office.js properties.
- Start unknown target discovery through `excel.agent.run` `mode: "find"`; the backend can route internally to lookup capabilities such as `excel.lookup.search_workbook` or `excel.lookup.resolve_range`.
- On the default agent surface, let `excel.agent.run` choose lightweight structure metadata, targeted reads, and compact proof.
- Internally, start known-scope large workbook reads with compact context: `excel.workbook.get_summary`, `excel.workbook.get_used_range_summary`, `excel.sheet.get_summary`, `excel.table.get_schema`, or `excel.range.get_summary`.
- Internally, prefer `excel.table.read_compact` and `excel.range.read_compact` for exploratory reads, broad used ranges, and large tables; page or project only the rows/columns needed for the task.
- Use `excel.agent.run` `mode: "validate"` when validation counts and examples are enough; the backend can route internally to `excel.validate.compact` and return resource links for full details.
- Reserve full reads for explicit user requests, audits, exports, or exact-data tasks that genuinely need every cell or facet.
- Group operations by workbook, sheet, and contiguous range.
- Chunk large payloads before Office.js limits.
- On the public agent surface, describe large generated batches through `excel.agent.run`; the backend preflights them and decides whether to apply directly, submit to the queue, or chunk safely.
- Describe grouped range work through one `excel.agent.run` preview instead of many parallel single-range writes; internally, related values, number formats, styles, clears, and autofit route to `excel.range.*_many` operations and queued parent jobs when needed.
- For extracted booking/client-image tables, use one `replace_range_with_styled_table` preview/apply so values, clearing, autofit, and style copies are compiled together.
- Surface queued or long-running mutations through transaction progress instead of hiding them behind parallel write calls.
- For chunked work, surface returned job progress so the user sees one update with chunk progress.
- Suspend calculation and screen updating only around large operations.

## Large Range Execution

Large matrix writes are split into row chunks before assignment to Office.js ranges:

- values
- formulas
- number formats

The chunker preserves whole rows and uses the configured cell limit to avoid sending very large matrices through a single Office.js range assignment. Structural operations such as row insertion, table resize, chart creation, and PivotTable creation remain native Office.js object calls because splitting them would change semantics.

Style-only batches can be retried adaptively after an add-in timeout because reapplying the same range style is idempotent. Value, formula, and number-format matrices are split before execution when preflight can safely chunk by rows. Structural mutations do not auto-retry after a timeout because Excel may have applied part or all of the original request.

## Telemetry

Every read/write result records:

- duration in milliseconds
- sync count
- payload bytes
- cells read or written
- range count
- chunk count
- engine name and version
- warning count

Compact summary and read tools also return `payloadBytes`, rough `estimatedTokens`, `truncated`, and optional `nextPage` metadata so agents can stop before sending oversized workbook context to the model.
Compact tools accept `responseMode` values of `brief`, `standard`, or `verbose`; compact-profile defaults to `brief`, storing full details locally and returning `contextId`/`resourceUri` handles plus proof metadata. Compact-profile responses are guarded by shared output limits for summaries, warnings, examples, issue lists, schema/sample fields, and final result size. Budget-limited compact responses return `resourceUri` handles for full details plus `truncated`, `omittedCounts`, `fullResult`, and telemetry. Mutation results include `compactProof` metadata so agents can report changed cells/ranges, warnings, backup IDs, and rollback availability without reading changed sheets back.

For MCP clients that resend all tool schemas with each generation, workbook payload savings can be hidden by tool-schema overhead. `owb mcp` exposes the public `excel.agent.run` surface so normal agents avoid primitive compact-capability chains. Raw range facet reads such as `excel.range.read_values` are not part of the public agent surface; backend routing should use compact reads so sparse ranges can be bounded, trimmed, and paged. `excel.runtime.get_capabilities` reports the public tool list separately from the internal backend capability catalog.

Snapshot creation and mutation results can contain full before/after workbook payloads for rollback. Public MCP responses return metadata, IDs, counts, rollback proof, `nextActionRecommendation`, `reasoningHints`, confidence metadata, and a `contextId`/`resourceUri` for the full detail instead of embedding snapshot, backup, or diff bodies inline. Snapshot compact tools accept nested `budget` controls for tighter response caps. Stored detail is exposed as MCP resource links returned by `excel.agent.run` rather than separate compact-resource tools.

For retry-prone mutation paths, pass a stable `idempotencyKey`. Replays with the same key and payload return `idempotentReplay: true` without applying the edit again; the same key with a different payload is rejected. Successful compact mutations also invalidate stale workbook-local compact resources and report `invalidatedContextIds`, preventing old read or validation context from being reused after an edit.

Lookup capabilities are cheaper than workbook scans when the backend does not know where data lives. They return ranked sheet/table/column/header/entity/range candidates and encoded `matchId` values; normal agents should ask through `excel.agent.run` `mode: "find"` or retry with a returned candidate instead of calling lookup primitives directly.

Initial releases collect telemetry first, then set hard SLOs from real workbook tests.

Use `corepack pnpm diagnose:session -- <log-file>` on OpenCode or MCP logs when tool-call count looks wrong. The report highlights repeated `excel.agent.run` calls, missed batching opportunities, preview/apply imbalance, and when `operation_status` should replace repeated apply attempts.

## Token Budget Examples

Compact workflows should usually cut large workbook context by 80-99%:

- Let `excel.agent.run` answer table questions from schema plus a compact page instead of sending a 10,000-row table.
- Let `excel.agent.run` answer used-range questions from summaries plus projected compact facets instead of full used-range matrices.
- Let internal compact range reads use sparse output for mostly empty worksheets instead of large rectangular matrices.
- Use `mode: "find"` instead of reading every sheet's first 100 rows to find a target column.
- Use validation or snapshot comparison through `excel.agent.run`; fetch returned resource links only for full detail review.

## Model Routing

Compact discovery, schema inspection, deterministic table/range mutations, validation proof, and rollback reporting are good candidates for cheaper models because Open Workbook handles the Excel mechanics locally. Escalate to larger models when the task requires ambiguous business reasoning, complex transformation design, or user-facing narrative synthesis from multiple compact evidence sources.

Use `excel.agent.run` for deterministic workbook facts such as totals, counts, missing values, duplicates, column profiling, or basic numeric summaries. The backend can route internally to analysis workflows and return compact proof plus `contextId` instead of sending raw rows to the model.

Use `excel.agent.run` `mode: "rollback"` for recovery tasks that should roll back or restore, recalculate, validate, and return compact proof in one call.

## Synthetic Core Benchmark

Run:

```bash
corepack pnpm benchmark:synthetic
```

This builds the repo, then runs `tests/benchmarks/synthetic.mjs` against compiled core modules. The benchmark emits JSON for:

- large value matrix chunking
- large formula matrix chunking
- batch compilation across many range writes
- compact large range read response size and estimated token savings
- compact validation failure response size and estimated token savings
- compact mutate/validate/diff response size and estimated token savings

These numbers are useful for detecting obvious core regressions in chunking, compilation overhead, and compact response-size caps. They are not real Excel SLOs because they do not include Office.js, workbook calculation, add-in websocket latency, Excel rendering, or platform-specific host behavior.

Real workbook SLOs should be calibrated from telemetry captured on macOS and Windows with representative workbook sizes before hard latency budgets are enforced in CI.
