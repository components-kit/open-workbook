---
name: open-workbook-skills
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP with the public excel.agent.run workflow: inspect workbooks, read/write ranges, update tables, preserve templates, repair formulas/styles, create pivots/charts, validate reports, save/export files, coordinate multiple agents, or normalize multilingual workbook requests into structured intent."
---

# Open Workbook Skills

Use Open Workbook MCP for live desktop Excel work. It is the required first path when the workbook is open in Excel and the user expects current cell values, unsaved edits, formatting, formulas, filters, tables, pivots, charts, snapshots, backups, transaction history, and rollback safety to survive.

Do not inspect or modify a connected live workbook with shell scripts, Python, openpyxl, pandas, manual UI automation, or offline `.xlsx` parsing unless the user explicitly asks for offline file analysis or the MCP path is unavailable and the user approves the fallback.

## Public Surface

Normal agents call one tool:

```text
excel.agent.run
```

Use modes deliberately:

- `status`: check Excel/add-in readiness and shared-agent coordination.
- `prepare`: cache workbook identity, sheet/table/name/region metadata, and collaboration state.
- `find`: locate candidate sheets, tables, headers, named ranges, regions, formulas, and summary blocks.
- `answer`: targeted live reads, schemas, summaries, comparisons, and deterministic inspection.
- `preview_update`: preview mutations and get an `operationId` plus `confirmationToken`.
- `apply_update`: apply only a previewed operation with the returned token.
- `operation_status`: check a returned operation without reissuing preview/apply.
- `cancel_operation`: cancel a pending preview before apply starts.
- `validate`: validate workbook, sheet, table, formula, style, filter, print-layout, or unintended-change scope.
- `rollback`: inspect or apply recovery through returned transaction, snapshot, or backup guidance.
- `auto`: compatibility route for casual prompts; explicit modes are preferred when the workflow step is known.

The backend owns primitive Excel capabilities, compact reads, batches, plans, validation, snapshots, backups, locks, jobs, transactions, resources, and collaboration records. Do not ask for or assume a separate primitive MCP surface in user workflows.

## Structured Intent

When the caller LLM can infer routing, pass structured fields alongside the natural request:

- `intent.action`: canonical English action such as `read_values`, `append_table_rows`, `write_values`, `validate_workbook`, `create_backup`, `restore_file_backup`, `rollback_validate`, `save`, or another supported backend action.
- `intent.targetHints`: workbook-native labels plus useful translated aliases.
- `target`: explicit `sheetName`, `range`, `tableName`, named item, region, or returned `candidateId`.
- `values`: 2D values, formulas, styles, table rows, patches, template data, backup IDs, or confirmation data.

For mutations, do not put the write payload only in `request`. Send structured data: `write_values` needs `target.sheetName`, `target.range`, and `values.values` as a 2D matrix; table appends use `values.rows`; multi-range edits use `values.patches`.

Use `detailLevel` conservatively: `workbook_summary` for metadata-only workbook context, `semantic_index` for role-aware workbook targets/candidates, `sheet_summary` for “look at/check/how is this sheet” overview requests without live cell reads, `table_sample` for a bounded live table sample, and `full_table` only when the user explicitly asks for all rows, every value, or full table contents.

For multilingual requests, preserve the original language in `request`, normalize routing fields to canonical English when clear, keep target hints in the workbook/user language, and answer the user in their language unless asked otherwise. Do not translate the whole task into English and discard the original wording.

Read `references/agent-run.md` for mode contracts, preview/apply behavior, multilingual routing, and result handling. Read `references/capability-map.md` when you need exact internal action groups.

## Workflow Semantics

- Raw duplicate: use `copy_sheet` only when the user explicitly asks for an exact copy including old values.
- Create from template: use template workflows to preserve structure, formulas, styles, validation, and layout while clearing old data regions for fresh entry.
- Apply style from template: use `copy_style_from_template`; the source/template sheet is a style source only and must not be duplicated or mutated.
- Replace styled table: use `replace_range_with_styled_table` to clear stale layout, write headers/rows, copy header/body style samples, and autofit in one preview/apply workflow.
- Inspect current styling: call `excel.agent.run` with `mode: "answer"`, `intent.action: "read_style_summary"`, and an exact range or current selection. Use `read_style_fingerprint` for template comparison, not normal user-facing style inspection.
- Formatting errors: use `format_diagnostics` before mutating. It returns raw value, displayed text, formulas, number formats, style summary, likely issues, and suggested fix actions.
- Filters, borders, dropdowns, and conditional formatting: send canonical intent early. Use `intent.action: "filter_range"` for autofilters, `clear_table_filters` for table filter removal, `format_range` for borders/fills/fonts/alignment, `write_data_validation` for dropdown/list cells, and `write_conditional_formatting` for formula-based formatting. Do not retry as sheet creation or value writes when the target sheet already exists.
- Column changes: use `insert_columns`/`delete_columns` for sheet structure and `reorder_range_columns` for plain range swaps/reorders; use `reorder_table_columns` only for real Excel tables.
- Date text cleanup: diagnose first, then use `parse_dates` only on exact date columns/ranges. If conversion needs explicit known dates, use one grouped `write_values` preview with per-range `numberFormat`.

## Safety Rules

- Never bypass permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, or rollback metadata for mutations.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as narrow 2D matrices or grouped patches.
- Do not fetch the full sheet before every preview. Reuse `workbookContextId`; preview may perform narrow metadata/fingerprint checks for safety, but agents should only read values first when target resolution is ambiguous, the user refers to current selection, or the change depends on existing values/styles.
- For “what is this workbook”, “look at this workbook/sheet”, “where is the invoice/customer/receipt/template area”, or similar context requests, use `workbook_summary`, `semantic_index`, or `sheet_summary`. Do not fetch `table_sample` or `full_table` unless the user asks for actual rows/values.
- For border-only, fill-only, font-only, alignment-only, or number-format-only clearing, use `clear_style_dimensions` with `values.dimensions` such as `["borders"]`. Use `clear_formats` only when the user asks to remove all formatting.
- For date-display verification, use answer-mode reads or `format_diagnostics`; wording that includes “format” is not by itself a mutation request.
- For OCR output, screenshots, forms, invoices, shipment documents, booking images, or other field/value data that must become a styled horizontal table, use one `preview_update` with `intent.action: "replace_range_with_styled_table"` and one `apply_update`; do not split clear/write/autofit/style copy into separate calls. If old borders/fills must disappear, do not clear leftovers by writing blanks.
- If a response returns `status: "NEEDS_WORKFLOW_REDIRECT"` during a real mutation, stop the fragmented plan and call the suggested grouped `preview_update` workflow instead. If you were only reading existing styles, retry as `mode: "answer"` with `intent.action: "read_style_summary"` and do not apply the suggested style-write workflow.
- Never pad broad ranges with blanks or `null` when only a smaller rectangle should change.
- Treat `CAPABILITY_UNAVAILABLE`, unsupported, partial, disconnected, and Office.js host-limit warnings as real results.
- After risky mutations, validate affected scopes and report important proof: transaction IDs, backup IDs, warnings, diffs, rollback options, compact proof, and telemetry.
- Do not invent agent identities. Multiple MCP sessions get trusted identity from the daemon/MCP adapter.

## Reference Router

- Read `references/agent-run.md` for `excel.agent.run` modes, structured intent, multilingual inputs, and proof/result handling.
- Read `references/capability-map.md` for internal backend action groups and which mode usually routes to them.
- Read `references/tool-selection.md` for choosing the efficient public workflow without exposing primitive calls.
- Read `references/workflows.md` for common inspect, read/analyze, write, template, formula, table, pivot/chart, backup, and export recipes.
- Read `references/reliability.md` for validation, rollback, stale previews, conflicts, locks, disconnected add-ins, and capability limits.
- Read `references/performance.md` before large reads/writes, latency-sensitive tasks, or token-heavy workbook context.
- Read `references/multi-agent.md` when multiple agents, sessions, sheets, or tasks may touch the same workbook.
