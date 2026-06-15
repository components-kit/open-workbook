# Tool Selection

Choose the narrowest Open Workbook MCP interface that preserves workbook intent.

## Session And Discovery

- Runtime health: `excel.runtime.get_status`
- Full workflow setup: `excel.workflow.prepare_session`
- Active workbook and selection context: `excel.runtime.get_active_context`
- Host/tool capability matrix: `excel.runtime.get_capabilities`
- Workbook structure: `excel.workbook.get_workbook_map`
- Compact workbook structure: `excel.workbook.get_summary`, `excel.workbook.get_used_range_summary`
- Workbook lookup: `excel.lookup.search_workbook`, `excel.lookup.find_headers`, `excel.lookup.find_tables_by_columns`, `excel.lookup.find_entity`, `excel.lookup.resolve_range`, `excel.lookup.inspect_match`
- Open workbooks: `excel.workbook.list_open_workbooks`
- Sheets: `excel.sheet.list`, `excel.sheet.get_info`, `excel.sheet.get_summary`, `excel.sheet.get_used_range`

Prefer `excel.workflow.prepare_session` as the first call. If capabilities are unknown because the add-in is disconnected, stop and ask for runtime setup instead of guessing.

## Reading Data

- Compact range/table discovery: `excel.range.get_summary`, `excel.table.get_schema`
- Unknown target lookup: `excel.lookup.search_workbook`, `excel.lookup.find_headers`, `excel.lookup.find_tables_by_columns`, `excel.lookup.find_entity`, `excel.lookup.resolve_range`
- Focused candidate preview: `excel.lookup.inspect_match`
- Compact bounded reads: `excel.range.read_compact`, `excel.table.read_compact`
- Stored compact details: `excel.compact.get_resource`, `excel.compact.list_resources`
- Values, formulas, display text, number formats, and styles: use `excel.range.read_compact` with the matching include flags
- Full stored detail: use `excel.compact.get_resource` with the `resourceUri` returned by compact reads
- Search and diagnostics: `excel.range.search`, `excel.range.find_blank_cells`, `excel.range.find_errors`
- Table data: use `excel.table.read_compact` with `columns`, `rowOffset`, `maxRows`, and include flags

Use lookup tools before reading when the target sheet, table, header, entity, or range is unknown. Use compact summaries or schemas before cell bodies once the target scope is known. Use explicit sheet/address ranges whenever possible. Use used-range or workbook-wide scans only for audits, validation, search, or discovery. For simple reads, use one projected `excel.range.read_compact` or `excel.table.read_compact` call instead of separate raw reads.

## Writing Data

- Simple range values: `excel.range.write_values`
- Simple formulas: `excel.range.write_formulas`
- Number formats: `excel.range.write_number_formats`
- Styles: `excel.range.write_styles`
- Grouped styles: `excel.range.write_styles_many`
- Combined formula sheet creation: `excel.workflow.create_formula_sheet`
- Combined formula error repair: `excel.workflow.repair_formula_errors`
- Multiple related range edits: `excel.batch.validate`, `excel.batch.preflight`, `excel.batch.dry_run`, `excel.batch.apply`
- Long-running related range edits: `excel.batch.submit`
- Preflighted chunked range edits: `excel.batch.submit_chunked`
- Chunked long-running jobs: `excel.job.get`, `excel.job.wait`, `excel.job.cancel`
- Scoped risky edit with diff and rollback preview: `excel.workflow.preview_risky_edit`
- Previewable or user-reviewable edits: `excel.plan.create`, `excel.plan.preview`, `excel.plan.apply`
- Stale plan handling: `excel.plan.refresh_preview`, `excel.plan.rebase`
- Queued mutation status: `excel.transaction.get`, `excel.transaction.wait`, `excel.transaction.cancel`

Prefer `excel.workflow.create_formula_sheet` for standard new sheets that need values, formulas, number formats, and validation. Prefer `excel.workflow.repair_formula_errors` for formula error repair when an error range and source formula range are known. Prefer `excel.workflow.preview_risky_edit` when a risky scoped edit should apply and return before/after snapshots, a diff, transaction id, and rollback preview in one response. Provide at least one minimal scoped operation and leave `apply` enabled unless the user asked for preview only. Prefer `excel.plan.*` when a user should review a diff before applying, when formulas/templates are at risk, or when rollback clarity matters. Prefer `excel.batch.*` for compact, well-scoped range mutations. Preflight large generated batches before applying them.
Combined mutating workflows include an internal read-only preflight payload before mutation; still prefer `excel.workflow.prepare_session` first when possible.
Values, formulas, formats, and styles are separate workbook facets. Use `excel.range.write_values` for constants only, `excel.range.write_formulas` for formulas beginning with `=`, and `excel.range.write_number_formats` for display formats. Do not merge formulas into a values write just because the matrix is convenient.
Do not pad a large range write with `null`, blank strings, or unchanged cells when only one cell or a smaller rectangle should change. Use the smallest changed target range, or use `excel.range.clear_values_keep_format` for intentional clearing.
Use `excel.range.write_styles_many` for report styling that touches many ranges, such as title bands, headers, zebra rows, status colors, or type colors. Use single `excel.range.write_styles` only for one contiguous range. Large grouped style updates may return queued parent jobs; wait or poll their job IDs.
If a mutation returns queued or applying transaction or job state, tell the user the workbook update is still running, then call `excel.job.wait`, `excel.job.get`, `excel.transaction.wait`, or `excel.transaction.get` instead of launching parallel writes. Use cancel tools only for queued work that should not start.

When a noninteractive agent run cannot apply a mutation, create an `excel.plan.create` draft with operations that use actual MCP tool names and nested `args`. Example operations should look like `{ "tool": "excel.range.write_formulas", "args": { ... } }`, not vague prose or broad workbook rewrites.

## Tables, Filters, And Sorts

- Inspect tables: `excel.table.list`, `excel.table.get_info`, `excel.table.get_schema`, `excel.table.read_compact`
- Append or update rows: `excel.table.append_rows`, `excel.table.update_rows`
- Reorder columns: `excel.table.reorder_columns`
- Resize or structure changes: `excel.table.resize`, `excel.table.copy_structure`
- Filters: inspect with `excel.table.get_info`; mutate with `excel.table.apply_filters`, `excel.table.clear_filters`, or `excel.table.preserve_filters`
- Sorts: `excel.table.sort`

Use table tools instead of range tools when the target is an Excel table. This preserves headers, totals rows, filters, structured references, and table styles.
Before table reorder, filter, sort, append, or update operations, call `excel.table.get_info` or `excel.workbook.get_workbook_map`. For table filters, use `excel.table.apply_filters` rather than generic filter tools. After table or filter/sort changes, use `excel.validate.compact` for proof unless exact validation details are required.
Do not clear and recreate a table to reorder columns; use `excel.table.reorder_columns`, or stop for confirmation before any destructive rebuild.
For large tables, use `excel.table.get_schema` first, then `excel.table.read_compact` with `maxRows`, `rowOffset`, and `columns` for targeted discovery before mutating with table-native tools.

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
- Compact validation proof: `excel.validate.compact`
- Formula errors: `excel.validate.no_formula_errors`, `excel.formula.find_errors`
- Broken references: `excel.validate.no_broken_references`
- Template/style/table/filter checks: `excel.validate.template_consistency`, `excel.validate.styles`, `excel.validate.tables`, `excel.validate.filters`
- Unintended changes: `excel.validate.no_unintended_changes`

Prefer `excel.validate.compact` when issue counts, severity counts, and a few examples are enough. It returns a compact resource URI for the full report.
- Combined risky edit workflow: `excel.workflow.preview_risky_edit`
- Snapshots and diffs: `excel.snapshot.create`, `excel.snapshot.get_compact`, `excel.snapshot.compare_compact`, `excel.diff.summarize`, `excel.diff.get_compact`
- Rollback: `excel.plan.rollback`, `excel.transaction.preview_rollback`, `excel.transaction.rollback`, `excel.transaction.preview_rollback_chain`, `excel.transaction.rollback_chain`
- File backups: `excel.backup.create_file`, `excel.backup.verify`, `excel.backup.restore_file`

Always present validation issues with severity and target, not vague success/failure language.
Diff tools require two snapshots. For risky scoped edits, use `excel.workflow.preview_risky_edit` after discovery when available, with a non-empty minimal operation list. If using separate tools, create a before snapshot, make the scoped edit, create an after snapshot, call `excel.diff.summarize`, `excel.diff.get_compact`, or `excel.snapshot.compare_compact`, then preview rollback with `excel.transaction.preview_rollback` or `excel.transaction.preview_rollback_chain`.
