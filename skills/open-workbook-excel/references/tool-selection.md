# Tool Selection

Choose the narrowest Open Workbook MCP interface that preserves workbook intent.

## Session And Discovery

- Runtime health: `excel.runtime.get_status`
- Full workflow setup: `excel.workflow.prepare_session`
- Active workbook and selection context: `excel.runtime.get_active_context`
- Host/tool capability matrix: `excel.runtime.get_capabilities`
- Workbook structure: `excel.workbook.get_workbook_map`
- Open workbooks: `excel.workbook.list_open_workbooks`
- Sheets: `excel.sheet.list`, `excel.sheet.get_info`, `excel.sheet.get_used_range`

Prefer `excel.workflow.prepare_session` as the first call. If capabilities are unknown because the add-in is disconnected, stop and ask for runtime setup instead of guessing.

## Reading Data

- Values only: `excel.range.read_values`
- Formulas only: `excel.range.read_formulas`
- Display text for user-visible output: `excel.range.read_display_text`
- Formatting: `excel.range.read_number_formats`, `excel.range.read_styles`
- Full cell payload: `excel.range.read_full`
- Search and diagnostics: `excel.range.search`, `excel.range.find_blank_cells`, `excel.range.find_errors`
- Table-shaped data: `excel.table.read`; pass `columns`, `rowOffset`, `rowLimit`, and include flags for large tables

Use explicit sheet/address ranges whenever possible. Use used-range or workbook-wide scans only for audits, validation, search, or discovery. For simple reads, use the facet-specific range tools instead of `excel.range.read_full`.

## Writing Data

- Simple range values: `excel.range.write_values`
- Simple formulas: `excel.range.write_formulas`
- Number formats: `excel.range.write_number_formats`
- Styles: `excel.range.write_styles`
- Combined formula sheet creation: `excel.workflow.create_formula_sheet`
- Combined formula error repair: `excel.workflow.repair_formula_errors`
- Multiple related range edits: `excel.batch.validate`, `excel.batch.dry_run`, `excel.batch.apply`
- Scoped risky edit with diff and rollback preview: `excel.workflow.preview_risky_edit`
- Previewable or user-reviewable edits: `excel.plan.create`, `excel.plan.preview`, `excel.plan.apply`
- Stale plan handling: `excel.plan.refresh_preview`, `excel.plan.rebase`

Prefer `excel.workflow.create_formula_sheet` for standard new sheets that need values, formulas, number formats, and validation. Prefer `excel.workflow.repair_formula_errors` for formula error repair when an error range and source formula range are known. Prefer `excel.workflow.preview_risky_edit` when a risky scoped edit should apply and return before/after snapshots, a diff, transaction id, and rollback preview in one response. Provide at least one minimal scoped operation and leave `apply` enabled unless the user asked for preview only. Prefer `excel.plan.*` when a user should review a diff before applying, when formulas/templates are at risk, or when rollback clarity matters. Prefer `excel.batch.*` for compact, well-scoped range mutations.
Combined mutating workflows include an internal read-only preflight payload before mutation; still prefer `excel.workflow.prepare_session` first when possible.
Values, formulas, formats, and styles are separate workbook facets. Use `excel.range.write_values` for constants only, `excel.range.write_formulas` for formulas beginning with `=`, and `excel.range.write_number_formats` for display formats. Do not merge formulas into a values write just because the matrix is convenient.
Do not pad a large range write with `null`, blank strings, or unchanged cells when only one cell or a smaller rectangle should change. Use the smallest changed target range, or use `excel.range.clear_values_keep_format` for intentional clearing.

When a noninteractive agent run cannot apply a mutation, create an `excel.plan.create` draft with operations that use actual MCP tool names and nested `args`. Example operations should look like `{ "tool": "excel.range.write_formulas", "args": { ... } }`, not vague prose or broad workbook rewrites.

## Tables, Filters, And Sorts

- Inspect tables: `excel.table.list`, `excel.table.get_info`, `excel.table.read`
- Append or update rows: `excel.table.append_rows`, `excel.table.update_rows`
- Reorder columns: `excel.table.reorder_columns`
- Resize or structure changes: `excel.table.resize`, `excel.table.copy_structure`
- Filters: `excel.filter.get_filters`, `excel.filter.apply`, `excel.filter.clear`, `excel.filter.validate`
- Sorts: `excel.sort.apply`, `excel.sort.clear`

Use table tools instead of range tools when the target is an Excel table. This preserves headers, totals rows, filters, structured references, and table styles.
Do not clear and recreate a table to reorder columns; use `excel.table.reorder_columns`, or stop for confirmation before any destructive rebuild.
For large tables, avoid full `excel.table.read` unless the user asks for all rows. Pass `rowLimit`, `rowOffset`, and `columns` for targeted discovery, then mutate with table-native tools.

## Templates, Styles, And Formulas

- Template discovery and registration: `excel.template.detect_templates`, `excel.template.register`, `excel.template.list`
- Combined template report creation: `excel.workflow.create_template_report`
- New period/report sheet: `excel.template.create_sheet_from_template`
- Template data fills: `excel.template.clear_data_regions`, `excel.template.fill_regions`
- Template validation/repair: `excel.template.validate_sheet_against_template`, `excel.template.repair_sheet_from_template`
- Style fingerprints and copy: `excel.style.get_fingerprint`, `excel.style.compare_fingerprint`, `excel.style.copy_from_template`
- Formula patterns and repair: `excel.workflow.repair_formula_errors`, `excel.formula.read_patterns`, `excel.formula.copy_patterns`, `excel.formula.validate`, `excel.formula.repair_patterns`
- Formula dependency safety: `excel.formula.get_dependency_graph`, `excel.formula.trace_precedents`, `excel.formula.trace_dependents`

Use `excel.workflow.create_template_report` for standard period/report creation from a registered template. Use these tools when the workbook has established formatting, formula patterns, reports, print layout, or period templates.
Formula repair requires both diagnosis and a repair action: use `excel.workflow.repair_formula_errors` when possible, or find formula errors, read formula patterns, inspect dependencies when references may shift, repair with formula-aware tools, then validate. Do not use `excel.formula.convert_to_values` as a repair.

## Pivots And Charts

- Pivot inspect/refresh: `excel.pivot.list`, `excel.pivot.get_info`, `excel.pivot.refresh`, `excel.pivot.refresh_all`
- Combined pivot/chart summary: `excel.workflow.create_pivot_chart_summary`
- Pivot create/rebuild: `excel.pivot.create`, `excel.pivot.rebuild_with_source`
- Pivot template work: `excel.pivot.get_fingerprint`, `excel.pivot.compare_fingerprint`, `excel.pivot.copy_from_template`, `excel.pivot.repair_from_template`
- Chart inspect/create/update: `excel.chart.list`, `excel.chart.get_info`, `excel.chart.create`, `excel.chart.update_data_source`, `excel.chart.copy_from_template`

Use `excel.workflow.create_pivot_chart_summary` for standard summary requests that need a PivotTable plus chart and post-create validation. For host-limited pivot/chart dimensions, report capability warnings and avoid claiming deterministic replay.

## Validation And Recovery

- Workbook/sheet health: `excel.validate.workbook`, `excel.validate.sheet`
- Formula errors: `excel.validate.no_formula_errors`, `excel.formula.find_errors`
- Broken references: `excel.validate.no_broken_references`
- Template/style/table/filter checks: `excel.validate.template_consistency`, `excel.validate.styles`, `excel.validate.tables`, `excel.validate.filters`
- Unintended changes: `excel.validate.no_unintended_changes`
- Combined risky edit workflow: `excel.workflow.preview_risky_edit`
- Snapshots and diffs: `excel.snapshot.create`, `excel.snapshot.compare`, `excel.diff.summarize`
- Rollback: `excel.plan.rollback`, `excel.transaction.preview_rollback`, `excel.transaction.rollback`, `excel.transaction.preview_rollback_chain`, `excel.transaction.rollback_chain`
- File backups: `excel.backup.create_file`, `excel.backup.verify`, `excel.backup.restore_file`

Always present validation issues with severity and target, not vague success/failure language.
Diff tools require two snapshots. For risky scoped edits, use `excel.workflow.preview_risky_edit` after discovery when available, with a non-empty minimal operation list. If using separate tools, create a before snapshot, make the scoped edit, create an after snapshot, call `excel.diff.summarize` or `excel.snapshot.compare`, then preview rollback with `excel.transaction.preview_rollback` or `excel.transaction.preview_rollback_chain`.
