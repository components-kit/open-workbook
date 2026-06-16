---
name: open-workbook-excel
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP, including inspecting workbooks, reading or writing ranges, updating tables, preserving templates, repairing formulas/styles, creating pivots/charts, validating reports, saving/exporting files, coordinating multiple agents, or choosing the fastest reliable Excel MCP tool instead of slow manual spreadsheet automation."
---

# Open Workbook Excel

Use Open Workbook MCP for live desktop Excel work. Prefer it over manual UI automation, per-cell scripts, or offline `.xlsx` rewrites when the workbook is open in Excel and the user expects formatting, formulas, filters, tables, pivots, charts, backups, and rollback safety to survive.

## First Calls

On the default MCP surface, call only:

```text
excel.agent.run
```

Use `mode: "prepare"` to cache workbook metadata, `mode: "find"` to locate sheets/tables/headers/named ranges/regions, `mode: "answer"` for targeted compact reads and deterministic summaries, `mode: "auto"` for normal user requests, `mode: "preview_update"` when a manual review step is wanted before a write, `mode: "apply_update"` only with the returned `confirmationToken`, `mode: "rollback"` for recovery, and `mode: "validate"` after risky changes. Treat `resourceLinks`, `nextAction`, `proof`, and `telemetry` as the primary contract.

`auto` may apply clearly scoped low-risk value edits after preview checks. Respect `nextAction` when `auto` returns `PREVIEW_READY`, `NEEDS_INPUT`, `AMBIGUOUS_TARGET`, `VALIDATION_FAILED`, or `manual_review`; do not force formula, structural, broad, sparse, stale, or ambiguous edits through a value write.

Pass natural-language targets directly when the user speaks casually, such as "June financial sheet" or "amount column in transactions". The backend resolves them against cached workbook metadata and returns `AMBIGUOUS_TARGET` with candidates when the request is too broad.

The tools below are the advanced compatibility/debug surface. Use them only when `OPEN_WORKBOOK_MCP_SURFACE=advanced` exposes them or when the user explicitly asks for primitive tool control.

Start every workbook session with:

```text
excel.workflow.prepare_session
```

If `excel.workflow.prepare_session` is unavailable, call `excel.runtime.get_status`, `excel.runtime.get_active_context`, `excel.runtime.get_capabilities`, `excel.workbook.get_workbook_map`, and `excel.collab.get_status`.

If the add-in is disconnected, ask the user to start their agent UI so it launches the configured Open Workbook MCP command, then open Excel and load the Open Workbook add-in. For manual troubleshooting, run `npx -y @components-kit/open-workbook@latest mcp` and retry. Do not fake workbook state from stale assumptions.

## Tool Selection

- When the target is unknown, use lookup first: `excel.lookup.search_workbook`, `excel.lookup.find_headers`, `excel.lookup.find_tables_by_columns`, `excel.lookup.find_entity`, or `excel.lookup.resolve_range`, then inspect one candidate with `excel.lookup.inspect_match`.
- Start with compact context tools for known large scopes: `excel.workbook.get_summary`, `excel.workbook.get_used_range_summary`, `excel.sheet.get_summary`, `excel.table.get_schema`, and `excel.range.get_summary`.
- Prefer `excel.range.read_compact` and `excel.table.read_compact` for exploratory data reads. Treat brief proof plus `contextId`/`resourceUri` as enough unless exact cell bodies are required.
- Use `excel.validate.compact` for validation proof when counts/examples are enough. Fetch `excel://compact/{resource_id}` details only if the user or task needs the full report.
- If a tool returns `nextActionRecommendation: "answer_now"`, answer from the compact proof instead of continuing to inspect or validate unless the user asked for exhaustive audit.
- Treat high-confidence compact proofs with `reasoningHints` such as "Agent can answer now" as a stop signal. Fetch `contextId` details only for failures, warnings, low confidence, truncation, or explicit audit requests.
- For long sessions or token debugging, call `excel.compact.context_stats` instead of listing or fetching full resources.
- Use `excel.range.read_*` for scoped cell data and metadata.
- Use `excel.table.*` for structured table rows, filters, sorts, totals, and table resizing.
- Use `excel.template.*`, `excel.style.*`, and `excel.formula.*` when preserving or repairing templates matters.
- Use `excel.plan.*` for previewable multi-step changes that need rollback and stale-target checks.
- Use `excel.batch.preflight` before large generated writes, then choose direct apply, queued submit, or `excel.batch.submit_chunked` with job progress based on the recommendation.
- Use `excel.batch.*` for compact, direct range mutations that still need backups, fingerprints, permissions, and transaction logging. Provide `idempotencyKey` on mutating tools when retrying or when the agent may be interrupted.
- Use `excel.workflow.create_formula_sheet` for standard sheet creation with values, formulas, number formats, and formula validation in one response.
- Use `excel.workflow.repair_formula_errors` for formula error repair when you have an error range plus a source formula range or exact formulas to write.
- Use `excel.workflow.preview_risky_edit` for scoped risky edits that need before/after snapshots, a diff, and rollback preview in one response. Provide at least one minimal scoped operation and leave `apply` enabled unless the user asked for preview only.
- Use `excel.workflow.create_template_report` for standard template report creation that needs region clear/fill, style comparison/repair, and validation in one response.
- Use `excel.workflow.create_pivot_chart_summary` for standard PivotTable plus chart summary tasks that need create, refresh, and validation in one response.
- Use `excel.workflow.inspect_analyze` for deterministic analysis such as missing values, duplicates, type profiling, and basic numeric summaries instead of asking the model to calculate from raw rows.
- Use `excel.workflow.rollback_validate` when a recovery task should rollback or restore, recalculate, validate, and return proof in one call.
Combined mutating workflows include a read-only preflight payload before mutation; still prefer `excel.workflow.prepare_session` first when possible.
- After a mutation returns `invalidatedContextIds`, do not reuse those old compact resources; read fresh compact context only if the next step truly needs it.
- Use `excel.validate.*` before and after risky changes.
- Use `excel.backup.*`, `excel.snapshot.*`, `excel.transaction.*`, and `excel.job.*` for recovery, audit, long-running progress, rollback previews, and rollback chains.
- Use `excel.task.*`, `excel.lock.*`, `excel.collab.*`, and `excel.conflict.*` for multi-agent workbook work.

For detailed routing, read `references/tool-selection.md`.

## Reliability Rules

- Never bypass Open Workbook's safety lifecycle for mutations: permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, and rollback metadata.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as 2D matrices over contiguous ranges.
- Never pad a broad range write with `null` or blanks when only a smaller range is intended. Write the smallest changed rectangle or use explicit clear tools.
- Read only the workbook properties needed for the task. Avoid broad workbook scans unless the task is audit, validation, search, or repair.
- Do not use full range/table reads when compact summaries, schemas, samples, or projected pages are enough. Full reads are for explicit user requests, audits, exports, or exact-data tasks.
- Treat `CAPABILITY_UNAVAILABLE`, partial capability warnings, and Office.js host limits as real results. Explain them and choose a supported path.
- Preserve existing template conventions over generic formatting rules.
- After mutation, validate the affected area and surface backups, transaction IDs, warnings, diffs, and rollback options.
- Treat `compactProof` on mutation results as the default reportable proof. Do not read back whole changed sheets unless exact cell data is required.

## Critical Recipes

- Sheet/formula, template-report, pivot/chart, risky-edit, and formula-repair tasks: prefer the matching `excel.workflow.*` tool, then validate compactly.
- Large table work: inspect schema first, read compact pages only when row data is needed, mutate with table-native tools, then validate tables/filters.
- Unknown target work: resolve the sheet/table/column/range with lookup tools before reading data.
- Snapshot/diff/rollback proof: prefer compact diff or risky-edit workflow results; fetch full compact resources only for detailed review.
- Multi-agent work: use collaboration/task/lock tools before mutation.

## Workflow References

- Read `references/tool-selection.md` to choose the most efficient MCP interface.
- Read `references/workflows.md` for common Excel task recipes.
- Read `references/reliability.md` for validation, rollback, stale-plan, and failure handling.
- Read `references/performance.md` before large reads/writes or latency-sensitive tasks.
- Read `references/multi-agent.md` when more than one agent or task may touch the same workbook.
