---
name: open-workbook-excel
description: Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP: inspect workbooks, read or write ranges, update tables, preserve templates, repair formulas/styles, create pivots/charts, validate reports, save/export files, coordinate multiple agents, or choose the fastest reliable Excel MCP tool instead of slow manual spreadsheet automation.
---

# Open Workbook Excel

Use Open Workbook MCP for live desktop Excel work. Prefer it over manual UI automation, per-cell scripts, or offline `.xlsx` rewrites when the workbook is open in Excel and the user expects formatting, formulas, filters, tables, pivots, charts, backups, and rollback safety to survive.

## First Calls

Start every workbook session with:

```text
excel.runtime.get_status
excel.runtime.get_active_context
excel.runtime.get_capabilities
excel.workbook.get_workbook_map
excel.collab.get_status
```

If the add-in is disconnected, ask the user to start their agent UI so it launches the configured Open Workbook MCP command, then open Excel and load the Open Workbook add-in. For manual troubleshooting, run `npx -y @components-kit/open-workbook@latest mcp` and retry. Do not fake workbook state from stale assumptions.

## Tool Selection

- Use `excel.range.read_*` for scoped cell data and metadata.
- Use `excel.table.*` for structured table rows, filters, sorts, totals, and table resizing.
- Use `excel.template.*`, `excel.style.*`, and `excel.formula.*` when preserving or repairing templates matters.
- Use `excel.plan.*` for previewable multi-step changes that need rollback and stale-target checks.
- Use `excel.batch.*` for compact, direct range mutations that still need backups, fingerprints, permissions, and transaction logging.
- Use `excel.validate.*` before and after risky changes.
- Use `excel.backup.*`, `excel.snapshot.*`, and `excel.transaction.*` for recovery, audit, rollback previews, and rollback chains.
- Use `excel.task.*`, `excel.lock.*`, `excel.collab.*`, and `excel.conflict.*` for multi-agent workbook work.

For detailed routing, read `references/tool-selection.md`.

## Reliability Rules

- Never bypass Open Workbook's safety lifecycle for mutations: permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, and rollback metadata.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as 2D matrices over contiguous ranges.
- Read only the workbook properties needed for the task. Avoid broad workbook scans unless the task is audit, validation, search, or repair.
- Treat `CAPABILITY_UNAVAILABLE`, partial capability warnings, and Office.js host limits as real results. Explain them and choose a supported path.
- Preserve existing template conventions over generic formatting rules.
- After mutation, validate the affected area and surface backups, transaction IDs, warnings, diffs, and rollback options.

## Workflow References

- Read `references/tool-selection.md` to choose the most efficient MCP interface.
- Read `references/workflows.md` for common Excel task recipes.
- Read `references/reliability.md` for validation, rollback, stale-plan, and failure handling.
- Read `references/performance.md` before large reads/writes or latency-sensitive tasks.
- Read `references/multi-agent.md` when more than one agent or task may touch the same workbook.
