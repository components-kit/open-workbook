---
name: open-workbook-excel
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP with the public excel.agent.run workflow: inspect workbooks, read/write ranges, update tables, preserve templates, repair formulas/styles, create pivots/charts, validate reports, save/export files, coordinate multiple agents, or normalize multilingual workbook requests into structured intent."
---

# Open Workbook Excel

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

For multilingual requests, preserve the original language in `request`, normalize routing fields to canonical English when clear, keep target hints in the workbook/user language, and answer the user in their language unless asked otherwise. Do not translate the whole task into English and discard the original wording.

Read `references/agent-run.md` for mode contracts, preview/apply behavior, multilingual routing, and result handling. Read `references/capability-map.md` when you need exact internal action groups.

## Safety Rules

- Never bypass permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, or rollback metadata for mutations.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as narrow 2D matrices or grouped patches.
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
