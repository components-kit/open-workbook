---
name: open-workbook-excel
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP, including inspecting workbooks, reading or writing ranges, updating tables, preserving templates, repairing formulas/styles, creating pivots/charts, validating reports, saving/exporting files, coordinating multiple agents, or choosing the fastest reliable Excel MCP tool instead of slow manual spreadsheet automation."
---

# Open Workbook Excel

Use Open Workbook MCP for live desktop Excel work. Prefer it over manual UI automation, per-cell scripts, or offline `.xlsx` rewrites when the workbook is open in Excel and the user expects formatting, formulas, filters, tables, pivots, charts, backups, and rollback safety to survive.

## First Calls

Start every workbook session with:

```text
excel.workflow.prepare_session
```

If `excel.workflow.prepare_session` is unavailable, call `excel.runtime.get_status`, `excel.runtime.get_active_context`, `excel.runtime.get_capabilities`, `excel.workbook.get_workbook_map`, and `excel.collab.get_status`.

If the add-in is disconnected, ask the user to start their agent UI so it launches the configured Open Workbook MCP command, then open Excel and load the Open Workbook add-in. For manual troubleshooting, run `npx -y @components-kit/open-workbook@latest mcp` and retry. Do not fake workbook state from stale assumptions.

## Tool Selection

- Use `excel.range.read_*` for scoped cell data and metadata.
- Use `excel.table.*` for structured table rows, filters, sorts, totals, and table resizing.
- Use `excel.template.*`, `excel.style.*`, and `excel.formula.*` when preserving or repairing templates matters.
- Use `excel.plan.*` for previewable multi-step changes that need rollback and stale-target checks.
- Use `excel.batch.preflight` before large generated writes, then choose direct apply, queued submit, or `excel.batch.submit_chunked` with job progress based on the recommendation.
- Use `excel.batch.*` for compact, direct range mutations that still need backups, fingerprints, permissions, and transaction logging.
- Use `excel.workflow.create_formula_sheet` for standard sheet creation with values, formulas, number formats, and formula validation in one response.
- Use `excel.workflow.repair_formula_errors` for formula error repair when you have an error range plus a source formula range or exact formulas to write.
- Use `excel.workflow.preview_risky_edit` for scoped risky edits that need before/after snapshots, a diff, and rollback preview in one response. Provide at least one minimal scoped operation and leave `apply` enabled unless the user asked for preview only.
- Use `excel.workflow.create_template_report` for standard template report creation that needs region clear/fill, style comparison/repair, and validation in one response.
- Use `excel.workflow.create_pivot_chart_summary` for standard PivotTable plus chart summary tasks that need create, refresh, and validation in one response.
Combined mutating workflows include a read-only preflight payload before mutation; still prefer `excel.workflow.prepare_session` first when possible.
- Use `excel.validate.*` before and after risky changes.
- Use `excel.backup.*`, `excel.snapshot.*`, `excel.transaction.*`, and `excel.job.*` for recovery, audit, long-running progress, rollback previews, and rollback chains.
- Use `excel.task.*`, `excel.lock.*`, `excel.collab.*`, and `excel.conflict.*` for multi-agent workbook work.

For detailed routing, read `references/tool-selection.md`.

## Reliability Rules

- Never bypass Open Workbook's safety lifecycle for mutations: permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, and rollback metadata.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as 2D matrices over contiguous ranges.
- Never pad a broad range write with `null` or blanks when only a smaller range is intended. Write the smallest changed rectangle or use explicit clear tools.
- Read only the workbook properties needed for the task. Avoid broad workbook scans unless the task is audit, validation, search, or repair.
- Treat `CAPABILITY_UNAVAILABLE`, partial capability warnings, and Office.js host limits as real results. Explain them and choose a supported path.
- Preserve existing template conventions over generic formatting rules.
- After mutation, validate the affected area and surface backups, transaction IDs, warnings, diffs, and rollback options.

## Critical Recipes

- Sheet with formulas: prefer `excel.workflow.create_formula_sheet`; if using separate tools, create the sheet, write labels/constants with `excel.range.write_values`, write formulas with `excel.range.write_formulas`, write number formats with `excel.range.write_number_formats`, then validate with `excel.formula.validate` or `excel.validate.no_formula_errors`. Do not put formula strings into `write_values`.
- Large table reorder/filter/sort/append/update: always inspect the table first with `excel.table.get_info` or `excel.workbook.get_workbook_map`. Use bounded `excel.table.read` only when row data is needed, then use `excel.table.reorder_columns`, `excel.table.apply_filters`, `excel.table.sort`, `excel.table.append_rows`, or `excel.table.update_rows`. For table filters, prefer `excel.table.apply_filters` over generic filter tools. After mutation, validate with `excel.validate.tables` for table changes and `excel.validate.filters` for filter/sort changes. Never clear, recreate, or rewrite the whole table for these tasks.
- Formula repair: prefer `excel.workflow.repair_formula_errors` when the error range and source formula range are known. Otherwise find errors, read nearby formula patterns, inspect dependencies, repair with `excel.formula.repair_patterns`, `excel.formula.fill_down`, `excel.formula.fill_right`, or scoped `excel.range.write_formulas`, recalculate if needed, then validate. Never convert formulas to values unless explicitly requested.
- Template report: prefer `excel.workflow.create_template_report`; if using separate tools, create the sheet from template, clear/fill declared regions, compare styles, repair style drift, and validate against the template.
- Snapshot/diff/rollback preview: after discovery, prefer `excel.workflow.preview_risky_edit` for a scoped risky edit with a non-empty minimal operation list. If using separate tools, create a before snapshot, make the scoped change, create an after snapshot, call `excel.diff.summarize` or `excel.snapshot.compare` with both snapshot IDs, then use rollback preview tools without actually rolling back unless asked.
- Pivot/chart summary: prefer `excel.workflow.create_pivot_chart_summary`; if using separate tools, check capability, create the PivotTable, refresh it, create/update the chart, refresh/update the chart source, and validate the PivotTable source before reporting success.
- Multi-agent work: start with `excel.collab.get_status`, create or inspect the task, acquire the narrowest lock, follow conflict guidance, perform or plan the scoped operation, then release locks.

## Workflow References

- Read `references/tool-selection.md` to choose the most efficient MCP interface.
- Read `references/workflows.md` for common Excel task recipes.
- Read `references/reliability.md` for validation, rollback, stale-plan, and failure handling.
- Read `references/performance.md` before large reads/writes or latency-sensitive tasks.
- Read `references/multi-agent.md` when more than one agent or task may touch the same workbook.
