# Performance

Fast Excel automation is part of correctness. Slow workflows push agents toward unsafe shortcuts.

## Defaults

- On the public surface, use `excel.agent.run`; the backend owns compact reads, summaries, proof, resource handles, and response budgeting.
- Optimize for cost per completed user task, not cost per individual tool call. A good read-only task usually finishes from one `excel.agent.run` result. A small exact value mutation should usually finish in one `auto` call after workbook write access is allowed for the session; broader or risky mutations should be preview plus apply, with validation only when needed.
- Prefer bulk Office.js operations over per-cell actions.
- Use 2D matrices for values, formulas, and number formats.
- Use table APIs for table-shaped work.
- Group writes by workbook, sheet, and contiguous range.
- Read only requested properties.
- Express unknown targets with `mode: "find"` or `targetHints`; the backend starts with lookup internally.
- Express known large scopes with `mode: "answer"` and a narrow target; the backend starts with compact summaries and schemas internally.
- Let `agent.run` report host capability warnings before expensive or host-limited features.

## Reads

- For user-visible reporting, read display text.
- For computation, read raw values.
- For formula review, use `read_formulas` for exact formula/status proof or `read_formula_patterns` for repeated layouts. Never infer formula existence from displayed values or numbers alone.
- For formatting review, read number formats or styles only for the relevant range.
- For table analysis, read the table instead of a sheet-sized range.
- On the public surface, use `excel.agent.run` `mode: "find"` or `mode: "answer"` for unknown targets and exploratory reads.
- The backend should use lookup before reading cells and compact range/table reads before full reads.

Avoid workbook-wide reads unless the task is search, validation, audit, or discovery.

Compact reads return payload/token telemetry, truncation status, result handles, and continuation metadata. Page only when more rows or columns are needed.
When a compact response includes inline `values`, `encodedValues`, `rows`, or `sparseRows`, use them directly before asking for more data. If `valueEncoding.kind` is `domain_dictionary_by_column`, decode `encodedValues` by replacing integer codes with `valueEncoding.columns[position].domain[code]`; the full raw matrix remains behind `fullResultUri`.
Wide table/range results may include `inlineColumnProjection`. Treat the selected inline columns as high-signal columns chosen from selection, requested fields, cached column roles, validation/format hints, and cardinality. Ask for full detail only when omitted columns are required for the task.
When only a preview is inline and the next step requires exact rows, raw values, transformation input, validation evidence, or audit detail, call `excel.agent.run` once with `continuation.fullResultUri` or paste the returned `resultUri`/`fullResultUri` in `request`. `excel://...` handles are not HTTP URLs; never use Webfetch, browser fetch, curl, or HTTP tooling for them.

## Task Completion And Freshness

- If the result has `taskOutcome: "final_answer"` and `maxRecommendedFollowupCalls: 0`, answer the user now. Do not call again to “double check” unless the result says data is stale, unavailable, or incomplete.
- If the result has `taskOutcome: "apply_complete"` and `maxRecommendedFollowupCalls: 0`, report the returned proof and stop.
- If the result has `taskOutcome: "preview_ready"`, the next call should normally be only `apply_update` with the returned `operationId` and `confirmationToken` after user approval.
- Reuse `continuation.workbookContextId`, `resultUri`, and `fullResultUri`. If `continuation.freshness` is present and the workbook content/structure version has not changed, prefer cached context/result handles over rediscovery reads.
- Treat freshness as a first-pass safety check: unchanged fingerprints or no overlapping journal entries mean the agent can skip rereading latest context. If the backend reports stale context, changed target fingerprints, or overlapping changes, refresh the relevant target only.

## Writes

- Use one batch or plan for related edits.
- On the public agent surface, group related small exact range value edits with `values.patches` in one `auto` call. Use `preview_update` plus one `apply_update` when the grouped edit is large, broad, ambiguous, formula/style/template/table-related, or user-reviewable.
- For multiple explicit value edits in the same user instruction, group all known target/value pairs in one `values.patches` call, even when rows describe different business topics. Do not split independent exact edits into separate calls unless the grouped call fails with actionable details.
- For dropdown option inspection, treat `data_validation_summary` as complete inline proof. Do not spend follow-up calls on `fullResultUri` unless the user requested raw audit metadata. Missing option additions should update one source-list cell/range when available, or use one dropdown-rule preview/apply for inline validation lists.
- For repeated values, number formats, styles, clears, or autofit across related ranges, use one grouped preview; the backend compiles it to internal `*_many` range operations so stale checks, backups, telemetry, and rollback stay single-operation.
- For broad deterministic value transforms, use `transform_values` so Open Workbook scans the target internally and returns bounded examples instead of full-column payloads.
- For row-aware derivations, use `derive_values` so Open Workbook reads source/target columns internally, verifies row alignment, and compiles only changed target cells. Use `formula_like` for calculations such as Payment Variance = Actual Amount - Cash Amount instead of reading full source/target columns into the model.
- For full-range formula repair from one repeated A1 pattern, use `write_formulas` with one `values.formula` and the full target range. Let Open Workbook expand relative references; do not generate large formula matrices or dummy value matrices in model context.
- For transaction settlement bundles, use `settle_reconciliation` so Payment Variance, Reconciliation Note, and Detail Notes are previewed as one grouped plan instead of separate formula/note calls.
- For batch sheet renames from one deterministic rule, use `transform_sheets` so Open Workbook previews one bounded workbook-structure plan instead of looping through individual sheet calls.
- Ask for preview before applying large generated changes.
- Keep matrix shapes exact: rows and columns must match the target range.
- Let Open Workbook chunk large values/formulas/number formats through safe row-based chunk plans.
- Avoid alternating read/write/read/write loops. Read once, compute, apply once, validate once.
- If work is queued or applying, report the progress message to the user and wait for the existing operation rather than starting parallel mutations.

For very large writes, consider a plan preview first so the agent can expose scope, chunk count, and rollback coverage before applying.

Automatic timeout retry is limited to style-only batches because repeating the same style is safe. Values, formulas, and number formats should be chunked before execution when preflight recommends it. For tables, pivots, charts, and structure, inspect transaction status before retrying.

## Formulas

- Ask `agent.run` for `read_formulas` on exact formula questions and formula-pattern repair when repeated formula layouts are involved.
- Ask for dependency-aware preview before editing source ranges used by reports, charts, pivots, or formulas.
- Keep formula writes, formula repairs, and broad formula-like derivations preview-first with validation.
- Use `intent.action: "calculate"` when the workflow requires fresh computed values.

## Telemetry

Read tool results for duration, sync count, payload bytes, cells read/written, range count, chunk count, engine, and warnings. Use that telemetry to explain unusually slow or partial work.
For compact reads, also inspect `estimatedTokens`, `truncated`, and `nextPage` before requesting more workbook data.
For mutation results, prefer the returned `compactProof` summary over reading changed ranges back into the model.
