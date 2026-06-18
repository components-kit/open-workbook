---
name: open-workbook-excel
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP with the public excel.agent.run workflow, including inspecting workbooks, reading or writing ranges, updating tables, preserving templates, repairing formulas/styles, creating pivots/charts, validating reports, saving/exporting files, coordinating multiple agents, or normalizing multilingual user requests into structured agent intent instead of using offline spreadsheet automation."
---

# Open Workbook Excel

Use Open Workbook MCP for live desktop Excel work. It is the required first path when the workbook is open in Excel and the user expects current cell values, unsaved edits, formatting, formulas, filters, tables, pivots, charts, backups, and rollback safety to survive.

Do not inspect or modify an open workbook with shell scripts, Python, openpyxl, pandas, manual UI automation, or offline `.xlsx` parsing unless the user explicitly asks for offline file analysis or the MCP path is unavailable and you first ask permission to use a non-live fallback. If Open Workbook MCP is connected but cannot return the needed data, report that MCP limitation instead of silently bypassing it.

## First Calls

On the default MCP surface, call only:

```text
excel.agent.run
```

Use `mode: "prepare"` to cache lightweight workbook structure metadata, `mode: "find"` to locate sheets/tables/headers/named ranges/regions, `mode: "answer"` for targeted compact reads and deterministic summaries, `mode: "preview_update"` when a manual review step is wanted before a write or table append, `mode: "apply_update"` only with the returned `confirmationToken`, `mode: "rollback"` for recovery, and `mode: "validate"` after risky changes. Omitted mode or `mode: "auto"` remains compatible for casual prompts, but explicit modes are more predictable for agent UIs. Treat `resourceLinks`, `nextAction`, `proof`, and `telemetry` as the primary contract.

`auto` may answer vague `.xlsx` review, workbook overview, current sheet, and table-list questions from lightweight cached metadata, compare two explicitly named sheets in one call, and apply clearly scoped low-risk value edits after preview checks. Targeted row/value/formula/comparison requests upgrade the context with sheet samples only when needed. Respect `nextAction` when `auto` returns `PREVIEW_READY`, `NEEDS_INPUT`, `AMBIGUOUS_TARGET`, `VALIDATION_FAILED`, or `manual_review`; do not force formula, style, template, structural, broad, sparse, stale, or ambiguous edits through a value write.

Pass natural-language targets directly when the user speaks casually, such as "June financial sheet" or "amount column in transactions". The backend resolves them against cached workbook metadata and returns `AMBIGUOUS_TARGET` with candidates when the request is too broad.

For non-English or mixed-language requests, keep `request` in the user's original language for audit and final response context, but normalize machine-routing fields when you can infer them: use canonical English `intent.action` values such as `read_values`, `read_schema`, `write_values`, `write_formulas`, `format_range`, `clear_values`, `append_table_rows`, `sort_table`, `filter_range`, `autofit`, `copy_template_sheet`, `calculate`, or `save`; include `intent.targetHints` with workbook-native labels and useful translated aliases; pass explicit `target` and `values` when known. Do not ask the backend to translate the whole prompt. Answer the user in their language unless they ask otherwise.

If `excel.agent.run` returns `AMBIGUOUS_TARGET`, choose one returned candidate and retry with the same `workbookContextId` plus `target.candidateId`. For exact value reads, call `mode: "answer"` with `target.sheetName` and `target.range` when available, or put the sheet name and A1 range clearly in the request. Requests for rows, samples, actual values, raw monthly sheets, or explicit A1 ranges should return live read proof; schema/columns/headers-only questions can return cached metadata. For raw month sheets without Excel Tables, ask for the exact sheet/range or the transaction/invoice section by month, such as `Apr 2026 invoice rows`. For small direct value edits with explicit range and values, use `auto` and let the backend apply safely in one call. For related edits across multiple ranges, send one `mode: "preview_update"` call with `values.patches`, where each patch has `target.sheetName`, `target.range`, and a 2D `values` matrix; then call `mode: "apply_update"` once with the returned `operationId` and `confirmationToken`. To add rows to an Excel table on the default surface, call `mode: "preview_update"` with `target.candidateId` or `target.tableName` and `values.rows`, then call `mode: "apply_update"` with the returned `operationId` and `confirmationToken`.

Open Workbook's primitive operation catalog, compact reads, batch/plan/workflow tools, validation, backup, snapshot, transaction, job, lock, and collaboration capabilities are backend-owned for normal agents. Do not ask for or assume a separate primitive MCP surface in user workflows.

If the add-in is disconnected, ask the user to start their agent UI so it launches the configured Open Workbook MCP command, then open Excel and load the Open Workbook add-in. For manual troubleshooting, run `npx -y @components-kit/open-workbook@latest mcp` and retry. Do not fake workbook state from stale assumptions.

## Agent Run Inputs

On the default surface, keep the interaction in `excel.agent.run` instead of composing primitive reads, compact resources, validation, and mutation tools yourself:

- Use `mode: "prepare"` for workbook identity and lightweight structure context.
- Use `mode: "find"` when the target is unknown.
- Use `mode: "answer"` for targeted reads, deterministic summaries, schemas, raw monthly sections, and comparisons.
- Use `mode: "preview_update"` and then `mode: "apply_update"` for user-reviewable changes.
- Use `mode: "validate"` or `mode: "rollback"` for post-change checks and recovery.

The backend owns compacting, cached metadata, target resolution, live reads, proof, validation, rollback metadata, and resource links for normal agents. Treat returned `resourceLinks`, `compactProof`, `invalidatedContextIds`, `nextAction`, and telemetry as evidence, not as mandatory follow-up calls.

The primitive capabilities below are internal/backend routing concepts, not normal agent calls:

- Lookup, compact summaries, `excel.range.read_compact`, `excel.table.read_compact`, `excel.validate.compact`, and `excel.compact.get_resource` are composed by the backend to keep `agent.run` responses bounded.
- Runtime and workbook diagnostics such as `excel.runtime.get_status`, `excel.runtime.get_capabilities`, and `excel.workbook.get_workbook_map` are backend capability concepts; normal agents should request status, prepare, find, or answer through `excel.agent.run`.
- Range/table/template/style/formula/plan/batch/workflow/validation/backup/snapshot/transaction/job/task/lock/collaboration capabilities remain part of the backend operation catalog, including `excel.plan.*` and `excel.batch.*`.
- If `agent.run` returns compact proof or `nextAction: "answer_now"`, answer from that proof instead of fetching stored detail unless the user requested an audit.

For detailed routing, read `references/tool-selection.md`.

## Reliability Rules

- Never bypass Open Workbook's safety lifecycle for mutations: permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, and rollback metadata.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as 2D matrices over contiguous ranges. For repeated zone/range edits, use `values.patches` in one grouped preview instead of issuing one tool call per range.
- Never pad a broad range write with `null` or blanks when only a smaller range is intended. Write the smallest changed rectangle or use explicit clear tools.
- Read only the workbook properties needed for the task. Avoid broad workbook scans unless the task is audit, validation, search, or repair.
- Do not use full range/table reads when compact summaries, schemas, samples, or projected pages are enough. Full reads are for explicit user requests, audits, exports, or exact-data tasks.
- Do not use shell commands, Python, openpyxl, pandas, or offline `.xlsx` parsing when the workbook is already open and connected through Open Workbook MCP. If the default MCP answer path cannot return the needed live values, report the MCP limitation and ask before using a non-live fallback.
- Treat `CAPABILITY_UNAVAILABLE`, partial capability warnings, and Office.js host limits as real results. Explain them and choose a supported path.
- Preserve existing template conventions over generic formatting rules.
- After mutation, validate the affected area and surface backups, transaction IDs, warnings, diffs, and rollback options.
- Treat `compactProof` on mutation results as the default reportable proof. Do not read back whole changed sheets unless exact cell data is required.

## Critical Recipes

- Sheet/formula, template-report, pivot/chart, risky-edit, and formula-repair tasks: use `excel.agent.run` modes and let the backend select the matching workflow.
- Large table work: use `prepare`, `find`, `answer`, `preview_update`, and `apply_update`; the backend should inspect schema, read compact pages only when row data is needed, mutate with table-native capabilities, then validate.
- Multi-range value updates: use one grouped `preview_update` with `values.patches`, then one `apply_update`. Do not split related patches unless apply returns a hard failure with actionable issue details.
- Unknown target work: use `mode: "find"` or retry with a returned `target.candidateId`.
- Snapshot/diff/rollback proof: prefer compact diff or risky-edit workflow results; fetch full compact resources only for detailed review.
- Multi-agent work: describe the coordination need through `excel.agent.run`; the backend owns collaboration, task, lock, transaction, and job capabilities.

## Workflow References

- Read `references/tool-selection.md` to choose the most efficient MCP interface.
- Read `references/workflows.md` for common Excel task recipes.
- Read `references/reliability.md` for validation, rollback, stale-plan, and failure handling.
- Read `references/performance.md` before large reads/writes or latency-sensitive tasks.
- Read `references/multi-agent.md` when more than one agent or task may touch the same workbook.
