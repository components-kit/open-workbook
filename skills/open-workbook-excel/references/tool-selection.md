# Tool Selection

Choose the narrowest Open Workbook MCP interface that preserves workbook intent.

## Session And Discovery

- Runtime health: `excel.runtime.get_status`
- Active workbook and selection context: `excel.runtime.get_active_context`
- Host/tool capability matrix: `excel.runtime.get_capabilities`
- Workbook structure: `excel.workbook.get_workbook_map`
- Open workbooks: `excel.workbook.list_open_workbooks`
- Sheets: `excel.sheet.list`, `excel.sheet.get_info`, `excel.sheet.get_used_range`

If capabilities are unknown because the add-in is disconnected, stop and ask for runtime setup instead of guessing.

## Reading Data

- Values only: `excel.range.read_values`
- Formulas only: `excel.range.read_formulas`
- Display text for user-visible output: `excel.range.read_display_text`
- Formatting: `excel.range.read_number_formats`, `excel.range.read_styles`
- Full cell payload: `excel.range.read_full`
- Search and diagnostics: `excel.range.search`, `excel.range.find_blank_cells`, `excel.range.find_errors`
- Table-shaped data: `excel.table.read`

Use explicit sheet/address ranges whenever possible. Use used-range or workbook-wide scans only for audits, validation, search, or discovery.

## Writing Data

- Simple range values: `excel.range.write_values`
- Simple formulas: `excel.range.write_formulas`
- Number formats: `excel.range.write_number_formats`
- Styles: `excel.range.write_styles`
- Multiple related range edits: `excel.batch.validate`, `excel.batch.dry_run`, `excel.batch.apply`
- Previewable or user-reviewable edits: `excel.plan.create`, `excel.plan.preview`, `excel.plan.apply`
- Stale plan handling: `excel.plan.refresh_preview`, `excel.plan.rebase`

Prefer `excel.plan.*` when a user should review a diff, when formulas/templates are at risk, or when rollback clarity matters. Prefer `excel.batch.*` for compact, well-scoped range mutations.

## Tables, Filters, And Sorts

- Inspect tables: `excel.table.list`, `excel.table.get_info`, `excel.table.read`
- Append or update rows: `excel.table.append_rows`, `excel.table.update_rows`
- Resize or structure changes: `excel.table.resize`, `excel.table.copy_structure`
- Filters: `excel.filter.get_filters`, `excel.filter.apply`, `excel.filter.clear`, `excel.filter.validate`
- Sorts: `excel.sort.apply`, `excel.sort.clear`

Use table tools instead of range tools when the target is an Excel table. This preserves headers, totals rows, filters, structured references, and table styles.

## Templates, Styles, And Formulas

- Template discovery and registration: `excel.template.detect_templates`, `excel.template.register`, `excel.template.list`
- New period/report sheet: `excel.template.create_sheet_from_template`
- Template data fills: `excel.template.clear_data_regions`, `excel.template.fill_regions`
- Template validation/repair: `excel.template.validate_sheet_against_template`, `excel.template.repair_sheet_from_template`
- Style fingerprints and copy: `excel.style.get_fingerprint`, `excel.style.compare_fingerprint`, `excel.style.copy_from_template`
- Formula patterns and repair: `excel.formula.read_patterns`, `excel.formula.copy_patterns`, `excel.formula.validate`, `excel.formula.repair_patterns`
- Formula dependency safety: `excel.formula.get_dependency_graph`, `excel.formula.trace_precedents`, `excel.formula.trace_dependents`

Use these tools when the workbook has established formatting, formula patterns, reports, print layout, or period templates.

## Pivots And Charts

- Pivot inspect/refresh: `excel.pivot.list`, `excel.pivot.get_info`, `excel.pivot.refresh`, `excel.pivot.refresh_all`
- Pivot create/rebuild: `excel.pivot.create`, `excel.pivot.rebuild_with_source`
- Pivot template work: `excel.pivot.get_fingerprint`, `excel.pivot.compare_fingerprint`, `excel.pivot.copy_from_template`, `excel.pivot.repair_from_template`
- Chart inspect/create/update: `excel.chart.list`, `excel.chart.get_info`, `excel.chart.create`, `excel.chart.update_data_source`, `excel.chart.copy_from_template`

For host-limited pivot/chart dimensions, report capability warnings and avoid claiming deterministic replay.

## Validation And Recovery

- Workbook/sheet health: `excel.validate.workbook`, `excel.validate.sheet`
- Formula errors: `excel.validate.no_formula_errors`, `excel.formula.find_errors`
- Broken references: `excel.validate.no_broken_references`
- Template/style/table/filter checks: `excel.validate.template_consistency`, `excel.validate.styles`, `excel.validate.tables`, `excel.validate.filters`
- Unintended changes: `excel.validate.no_unintended_changes`
- Snapshots and diffs: `excel.snapshot.create`, `excel.snapshot.compare`, `excel.diff.summarize`
- Rollback: `excel.plan.rollback`, `excel.transaction.preview_rollback`, `excel.transaction.rollback`, `excel.transaction.preview_rollback_chain`, `excel.transaction.rollback_chain`
- File backups: `excel.backup.create_file`, `excel.backup.verify`, `excel.backup.restore_file`

Always present validation issues with severity and target, not vague success/failure language.
