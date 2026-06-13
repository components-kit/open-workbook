# Workflows

These are default Open Workbook MCP workflows. Adjust scope and validation to the workbook risk.

## Inspect A Workbook

1. Call `excel.runtime.get_status`.
2. Call `excel.runtime.get_active_context`.
3. Call `excel.runtime.get_capabilities`.
4. Call `excel.workbook.get_workbook_map`.
5. For a specific sheet, call `excel.sheet.get_used_range` and scoped `excel.range.read_*` tools.

Use display text when reporting what a user sees. Use values/formulas/number formats when making calculations or edits.

## Read And Analyze Data

1. Identify the smallest sheet, table, region, or range that answers the question.
2. Use `excel.table.read` for Excel tables; otherwise use `excel.range.read_values` or `excel.range.read_full`.
3. For formula-sensitive analysis, also call `excel.range.read_formulas` or `excel.formula.read_patterns`.
4. For data-quality work, use `excel.clean.detect_header_row`, `excel.range.find_blank_cells`, `excel.range.find_errors`, and relevant validators.

Do not read the whole workbook when a named table, used range, or explicit range is enough.

## Write Values Safely

1. Resolve workbook, sheet, and target address.
2. Use `excel.batch.validate` or `excel.plan.preview` for non-trivial changes.
3. Apply through `excel.batch.apply` or `excel.plan.apply`.
4. Validate the target with `excel.validate.no_unintended_changes`, `excel.validate.no_formula_errors`, or a scoped validator.
5. Return transaction IDs, backup IDs, warnings, and rollback options.

For a one-range value update, direct `excel.range.write_values` is acceptable because it routes through the backend safety lifecycle.

## Create A New Period Sheet From Template

1. Call `excel.template.list` or `excel.template.detect_templates`.
2. Register the template if needed with `excel.template.register`.
3. Create the sheet with `excel.template.create_sheet_from_template`.
4. Clear declared input/data regions with `excel.template.clear_data_regions`.
5. Fill declared regions with `excel.template.fill_regions`.
6. Validate with `excel.template.validate_sheet_against_template`.

If no registered template exists, warn before mutation. Preserve the template's formatting, formulas, filters, print layout, tables, freeze panes, and named ranges.

## Repair Formulas Or Styles

1. Identify the template and target sheet/range.
2. Capture current state with `excel.snapshot.create` when the repair may affect user work.
3. Compare fingerprints with `excel.style.compare_fingerprint` or formula patterns with `excel.formula.validate_against_template`.
4. Repair with `excel.repair.style_from_template`, `excel.repair.formulas_from_template`, `excel.style.copy_from_template`, or `excel.formula.repair_patterns`.
5. Validate again and report any Office.js capability-status warnings.

Do not convert formulas to values unless the user explicitly asks or the workflow requires a static export.

## Clean A Sheet Or Range

1. Detect headers with `excel.clean.detect_header_row`.
2. Prefer read-only detectors first: outliers, fuzzy matches, blanks, errors.
3. For mutations, apply a scoped cleaning tool such as `excel.clean.trim_whitespace`, `excel.clean.normalize_headers`, `excel.clean.parse_dates`, `excel.clean.parse_numbers`, `excel.clean.remove_duplicates`, or `excel.clean.fill_missing_values`.
4. Validate formulas, tables, and unintended changes after cleaning.

Cleaning writes must stay within the requested sheet, range, table, or registered region.

## Update Tables, Filters, Or Sorts

1. Inspect with `excel.table.get_info` and `excel.filter.get_filters`.
2. Use `excel.table.append_rows` or `excel.table.update_rows` for data.
3. Use `excel.table.resize` only when structure must change.
4. Preserve or reapply filters with `excel.filter.apply`, `excel.filter.preserve_from_template`, or `excel.table.preserve_filters`.
5. Validate with `excel.validate.tables` and `excel.validate.filters`.

Avoid raw range writes inside table bodies when table tools can express the intent.

## Create Or Update Pivots And Charts

1. Confirm active host support with `excel.runtime.get_capabilities`.
2. Inspect source tables/ranges and existing objects.
3. For pivots, use `excel.pivot.validate_source`, `excel.pivot.create`, or `excel.pivot.rebuild_with_source`.
4. For charts, use `excel.chart.create` or `excel.chart.update_data_source`.
5. For template parity, use pivot/chart fingerprint or copy-from-template tools.
6. Refresh and validate before reporting success.

When Office.js cannot expose deterministic pivot/chart dimensions, return the capability warning instead of inventing proof.

## Save, Export, And Back Up

1. Use `excel.workbook.save` for normal save.
2. Use `excel.workbook.export_copy` for a true `.xlsx` copy when supported.
3. Use `excel.backup.create_file` for durable file backup records.
4. Use `excel.backup.verify` before trusting a backup path.
5. Use `excel.workbook.save_as` only when the native file bridge is configured.

Full file replacement of an open workbook is host-limited. Prefer safe open-as-new restore unless the user confirms replacement.
