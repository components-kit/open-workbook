# Tool Surface

The full namespace is represented in the protocol catalog, but MCP only exposes capability-gated tools.

## Status Model

- `stable`: exposed by default.
- `preview`: exposed only with `OPEN_WORKBOOK_PREVIEW_TOOLS=1`.
- `planned`: listed by `excel.runtime.get_capabilities`, not callable yet.
- `unsupported`: listed by capabilities as intentionally unavailable. OCR is unsupported for now because agents can perform OCR externally and call range/table/template tools with extracted data.

`excel.runtime.get_capabilities` is the source of truth for the complete catalog, resources, prompts, and runtime connection status.

## Stable Tool Groups

- Runtime: `excel.runtime.get_status`, `excel.runtime.get_active_context`
- Workbook: `excel.workbook.list_open_workbooks`, `excel.workbook.get_workbook_info`, `excel.workbook.get_workbook_map`, `excel.workbook.snapshot`, `excel.workbook.refresh_snapshot`, `excel.workbook.get_snapshot`, `excel.workbook.detect_external_changes`, `excel.workbook.calculate`, `excel.workbook.save`, `excel.workbook.save_as`, `excel.workbook.create_backup`, `excel.workbook.restore_backup`, `excel.workbook.export_copy`, `excel.workbook.close`
- Sheet: `excel.sheet.list`, `excel.sheet.get_info`, `excel.sheet.create`, `excel.sheet.copy`, `excel.sheet.rename`, `excel.sheet.delete`, `excel.sheet.hide`, `excel.sheet.unhide`, `excel.sheet.protect`, `excel.sheet.unprotect`, `excel.sheet.clear`, `excel.sheet.get_used_range`, `excel.sheet.set_tab_color`
- Range: `excel.range.read_values`, `excel.range.read_formulas`, `excel.range.read_number_formats`, `excel.range.read_display_text`, `excel.range.read_styles`, `excel.range.read_full`, `excel.range.write_values`, `excel.range.write_formulas`, `excel.range.write_number_formats`, `excel.range.write_styles`, `excel.range.clear`, `excel.range.clear_values`, `excel.range.clear_formats`, `excel.range.clear_values_keep_format`, `excel.range.copy`, `excel.range.move`, `excel.range.insert_rows`, `excel.range.delete_rows`, `excel.range.insert_columns`, `excel.range.delete_columns`, `excel.range.autofit_columns`, `excel.range.autofit_rows`, `excel.range.merge`, `excel.range.unmerge`
- Advanced range reads: `excel.range.read_hyperlinks`, `excel.range.read_comments`, `excel.range.read_notes`, `excel.range.read_merged_cells`, `excel.range.read_data_validation`, `excel.range.read_conditional_formatting`, `excel.range.search`, `excel.range.find_blank_cells`, `excel.range.find_errors`
- Batch and plan: `excel.batch.validate`, `excel.batch.dry_run`, `excel.batch.apply`, `excel.plan.create`, `excel.plan.preview`, `excel.plan.apply`, `excel.plan.rollback`
- Snapshot and diff: `excel.snapshot.create`, `excel.snapshot.refresh`, `excel.snapshot.get`, `excel.snapshot.compare`, `excel.snapshot.invalidate`, `excel.snapshot.list`, `excel.snapshot.delete`, `excel.diff.create`, `excel.diff.summarize`, `excel.diff.get_details`, `excel.diff.export_json`, `excel.diff.export_html`
- Events: `excel.events.subscribe`, `excel.events.unsubscribe`, `excel.events.get_recent`, `excel.events.clear`, `excel.events.set_debounce`
- Templates: `excel.template.detect_templates`, `excel.template.register`, `excel.template.unregister`, `excel.template.get`, `excel.template.list`, `excel.template.infer_regions`, `excel.template.create_sheet_from_template`, `excel.template.clear_data_regions`, `excel.template.fill_regions`, `excel.template.validate_sheet_against_template`, `excel.template.repair_sheet_from_template`
- Style and formula preservation: `excel.style.get_fingerprint`, `excel.style.compare_fingerprint`, `excel.style.copy_from_template`, `excel.style.apply_style`, `excel.style.validate_consistency`, `excel.style.repair_consistency`, `excel.style.get_theme`, `excel.style.apply_theme`, `excel.style.copy_column_widths`, `excel.style.copy_row_heights`, `excel.style.copy_borders`, `excel.style.copy_fills`, `excel.style.copy_fonts`, `excel.style.copy_alignment`, `excel.style.copy_number_formats`, `excel.style.copy_conditional_formatting`, `excel.style.copy_data_validation`, `excel.style.copy_freeze_panes`, `excel.style.copy_print_settings`, `excel.style.copy_page_layout`, `excel.style.copy_hidden_rows_columns`, `excel.formula.read_patterns`, `excel.formula.copy_patterns`, `excel.formula.fill_down`, `excel.formula.fill_right`, `excel.formula.validate`, `excel.formula.validate_against_template`, `excel.formula.repair_patterns`, `excel.formula.find_errors`, `excel.formula.find_circular_references`, `excel.formula.trace_precedents`, `excel.formula.trace_dependents`, `excel.formula.convert_to_values`, `excel.formula.recalculate`, `excel.formula.explain`
- Tables: `excel.table.list`, `excel.table.get_info`, `excel.table.read`, `excel.table.create`, `excel.table.resize`, `excel.table.append_rows`, `excel.table.update_rows`, `excel.table.clear_data_keep_formulas`, `excel.table.clear_filters`, `excel.table.apply_filters`, `excel.table.preserve_filters`, `excel.table.sort`, `excel.table.set_total_row`, `excel.table.set_style`, `excel.table.copy_structure`, `excel.table.validate_against_template`
- Filters and sort: `excel.filter.get_filters`, `excel.filter.apply`, `excel.filter.clear`, `excel.filter.preserve_from_template`, `excel.filter.validate`, `excel.sort.apply`, `excel.sort.clear`, `excel.sort.preserve_from_template`
- PivotTables and charts: `excel.pivot.list`, `excel.pivot.get_info`, `excel.pivot.create`, `excel.pivot.refresh`, `excel.pivot.refresh_all`, `excel.pivot.update_source`, `excel.pivot.copy_from_template`, `excel.pivot.validate_source`, `excel.chart.list`, `excel.chart.get_info`, `excel.chart.create`, `excel.chart.update_data_source`, `excel.chart.copy_from_template`, `excel.chart.refresh`, `excel.chart.delete`, `excel.chart.validate_against_template`
- Names and regions: `excel.names.list`, `excel.names.get`, `excel.names.create`, `excel.names.update`, `excel.names.delete`, `excel.region.detect`, `excel.region.register`, `excel.region.list`, `excel.region.get`, `excel.region.clear_values`, `excel.region.write_values`, `excel.region.fill`
- Validation: `excel.validate.workbook`, `excel.validate.sheet`, `excel.validate.template_consistency`, `excel.validate.formulas`, `excel.validate.styles`, `excel.validate.tables`, `excel.validate.filters`, `excel.validate.print_layout`, `excel.validate.no_broken_references`, `excel.validate.no_formula_errors`, `excel.validate.no_unintended_changes`
- Repair: `excel.repair.style_from_template`, `excel.repair.formulas_from_template`, `excel.repair.filters_from_template`, `excel.repair.table_structure`, `excel.repair.print_layout`, `excel.repair.named_ranges`, `excel.repair.formula_errors`, `excel.repair.merged_cells`
- Cleaning: `excel.clean.detect_header_row`, `excel.clean.normalize_headers`, `excel.clean.trim_whitespace`, `excel.clean.remove_duplicates`, `excel.clean.parse_dates`, `excel.clean.parse_numbers`, `excel.clean.standardize_currency`, `excel.clean.fill_missing_values`, `excel.clean.split_column`, `excel.clean.merge_columns`, `excel.clean.detect_outliers`, `excel.clean.fuzzy_match`
- Permissions: `excel.permissions.get`, `excel.permissions.set`, `excel.permissions.require_confirmation`, `excel.permissions.set_scope`, `excel.permissions.allow_destructive_actions`, `excel.permissions.allow_macro_execution`, `excel.permissions.lock_regions`, `excel.permissions.unlock_regions`

## Resources

- `excel://runtime/status`
- `excel://workbooks`
- `excel://workbooks/{workbook_id}/map`
- `excel://workbooks/{workbook_id}/sheets`
- `excel://workbooks/{workbook_id}/templates`
- `excel://workbooks/{workbook_id}/snapshots/{snapshot_id}`
- `excel://workbooks/{workbook_id}/plans/{plan_id}/diff`

## Rule

No mutating tool should bypass the backend safety lifecycle. It must validate permissions and create a backup before Excel receives writes, whether execution goes through the batch engine or a native Office.js object API.

## Implementation Notes

`excel.batch.apply`, `excel.plan.apply`, and all stable mutating sheet/range tools route through backend snapshots, backup records, target-region conflict checks, and add-in Office.js execution.

`excel.plan.rollback` and `excel.workbook.restore_backup` restore captured range snapshots. Full file-copy restore is tracked separately because it requires file-level user or OS involvement.

Workbook file lifecycle is explicit about host limitations. `excel.workbook.save` and `excel.workbook.close` use Office.js. `excel.workbook.save_as` and true `.xlsx` `excel.workbook.export_copy` require a future native host bridge; today they return capability-unavailable results, and `export_copy` creates a persistent snapshot backup as a fallback.

Template repair creates a backup before mutating the target sheet, then copies template styles/formulas/layout through Office.js and validates the target sheet against the registered template fingerprint.

Style fidelity tools capture granular style fingerprints for a sheet or address, compare dimensions independently, and create backups before copying style dimensions. Native Office.js format copy is used for fills, fonts, borders, alignment, number formats, conditional formatting, and data validation so Excel preserves more formatting detail than the protocol models manually. Column widths and row heights are copied explicitly. Workbook theme, freeze pane, print setting, page layout, and hidden row/column replay currently return capability-status warnings where Office.js does not expose a deterministic cross-platform replay path.

Formula intelligence tools capture R1C1 formula patterns, compare pattern matrices, copy formulas from templates, fill formulas down or right, convert formulas to values with backup, recalculate workbooks, and return lightweight formula explanations. Formula error validation scans used ranges through Office.js special-cell APIs. Circular-reference and formula dependency graph tracing currently return explicit capability-status results until those paths can be normalized across Excel hosts.

Table, filter, and sort mutations use native Office.js table APIs. The backend captures a backup over the affected table range or target structure range before mutating.

PivotTable and chart tools use native Office.js APIs. Pivot creation supports source ranges or structured tables and destination cells. Chart creation/update supports source range, chart type, series orientation, title, position, and style. Pivot source reassignment and deep chart/pivot template copy currently return capability-status metadata where Office.js does not expose a safe deterministic operation.

Advanced range reads use native Office.js metadata APIs where available. Comments and legacy notes currently return explicit unsupported warnings because reliable address mapping is not implemented yet.

Named-item tools use Office.js workbook and worksheet scoped `NamedItem` collections. Region tools maintain a runtime registry of reusable sheet/address targets and can resolve existing Excel named ranges as regions. Region writes and fills route through the standard batch lifecycle.

Permission tools manage runtime policy for writes, destructive actions, workbook actions, confirmation requirements, sheet/region scope, and locked regions. `excel.batch.apply` and region/range cleaning writes are checked before Excel receives the request; direct table mutations also check scope and locked regions.

Cleaning tools read range values, transform them in the backend, and write results through the same backup-aware batch path. Detection tools such as header detection, outlier detection, and fuzzy match are read-only.

Validation tools return structured reports with `ok`, issue severity counts, categorized issues, and optional supporting data. Formula and broken-reference validators inspect workbook used ranges through compact Office.js range-area summaries. Template consistency validators reuse registered template fingerprints.

Repair tools return structured repair reports. Style and formula repairs use registered templates and create backups before mutation. Table-structure repair uses the existing table copy path. Repair categories that Office.js cannot safely execute yet return `CAPABILITY_UNAVAILABLE` with a specific reason code.

## Preview Runtime Tools

Set `OPEN_WORKBOOK_PREVIEW_TOOLS=1` before starting `owb mcp` to expose:

- `excel.runtime.connect_addin`
- `excel.runtime.disconnect_addin`
- `excel.runtime.ping_addin`
- `excel.runtime.get_capabilities`
- `excel.runtime.get_selection`
- `excel.runtime.set_active_workbook`
- `excel.runtime.set_active_sheet`

These are marked preview because they establish runtime-control contracts that can affect session routing.
