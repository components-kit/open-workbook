# Workflows

These are default Open Workbook MCP workflows. Adjust scope and validation to the workbook risk.

## Inspect A Workbook

On the public surface, call `excel.agent.run` with `mode: "prepare"` first. Then use `mode: "find"` for discovery and `mode: "answer"` for targeted reads.

The backend can combine runtime status, active context, capabilities, workbook map, and collaboration state before a workflow mutates. It should use scoped compact range/table reads only after the target is known.

Use display text when reporting what a user sees. Use values/formulas/number formats when making calculations or edits.

## Read And Analyze Data

1. Identify the smallest sheet, table, region, or range that answers the question.
2. On the default surface, use `excel.agent.run` with `mode: "answer"` and a clear target or `target.candidateId`; include the sheet/range, raw monthly transaction/invoice section, or row/sample/value wording when actual cell data is needed.
3. Let the backend use table-native compact reads for Excel tables and range compact reads for normal ranges.
4. For formula-sensitive analysis, include formulas in the compact read or use a matched formula workflow.
5. For data-quality work, ask through `agent.run`; the backend can use header detection, blank/error scans, and relevant validators.

Do not read the whole workbook when a named table, used range, or explicit range is enough.

## Write Values Safely

On the public surface, use `excel.agent.run` with `mode: "preview_update"` and then `mode: "apply_update"` for scoped value edits. The primitive steps below describe backend behavior.

1. Resolve workbook, sheet, and target address.
2. Use `excel.batch.validate` or `excel.plan.preview` for non-trivial changes.
3. Apply through `excel.batch.apply` or `excel.plan.apply`.
4. Validate the target with `excel.validate.no_unintended_changes`, `excel.validate.no_formula_errors`, or a scoped validator.
5. Return transaction IDs, backup IDs, warnings, and rollback options.

For a one-range value update, direct `excel.range.write_values` is acceptable because it routes through the backend safety lifecycle.

For new sheets with formulas and number formats, prefer `excel.workflow.create_formula_sheet`.

Combined mutating workflows return an internal `preflight` payload with runtime status, active context, capabilities, workbook map, and collaboration state before they mutate. On the public surface, use `excel.agent.run` preview/apply modes and let the backend select the matched workflow.

## Create A New Period Sheet From Template

Prefer `excel.workflow.create_template_report` for standard template report creation.

1. Call `excel.template.list` or `excel.template.detect_templates`.
2. Register the template if needed with `excel.template.register`.
3. Create the sheet with `excel.template.create_sheet_from_template`.
4. Clear declared input/data regions with `excel.template.clear_data_regions`.
5. Fill declared regions with `excel.template.fill_regions`.
6. Validate with `excel.template.validate_sheet_against_template`.

If no registered template exists, warn before mutation. Preserve the template's formatting, formulas, filters, print layout, tables, freeze panes, and named ranges.
Do not replace this workflow with `excel.sheet.copy` unless the user asks for a raw sheet duplicate or the template tool is unavailable.

## Repair Formulas Or Styles

Prefer `excel.workflow.repair_formula_errors` for ordinary formula error repairs when you can identify the error range plus a source formula or explicit formula matrix.

1. Find formula errors with `excel.formula.find_errors`, `excel.range.find_errors`, or `excel.validate.no_formula_errors`.
2. Read neighboring patterns with `excel.formula.read_patterns` and dependencies with `excel.formula.get_dependency_graph` or trace tools.
3. Capture current state with `excel.snapshot.create` when the repair may affect user work.
4. Repair with `excel.formula.repair_patterns`, `excel.formula.fill_down`, `excel.formula.fill_right`, or scoped `excel.range.write_formulas`; for styles use `excel.style.copy_from_template` or `excel.style.repair_consistency`.
5. Recalculate when formulas changed, validate with `excel.formula.validate` or `excel.validate.no_formula_errors`, and report warnings.

Do not convert formulas to values unless the user explicitly asks or the workflow requires a static export.

## Clean A Sheet Or Range

1. Detect headers with `excel.clean.detect_header_row`.
2. Prefer read-only detectors first: outliers, fuzzy matches, blanks, errors.
3. For mutations, apply a scoped cleaning tool such as `excel.clean.trim_whitespace`, `excel.clean.normalize_headers`, `excel.clean.parse_dates`, `excel.clean.parse_numbers`, `excel.clean.remove_duplicates`, or `excel.clean.fill_missing_values`.
4. Validate formulas, tables, and unintended changes after cleaning.

Cleaning writes must stay within the requested sheet, range, table, or registered region.

## Update Tables, Filters, Or Sorts

On the default surface, append table rows with `excel.agent.run` `mode: "preview_update"` plus `target.candidateId` or `target.tableName` and `values.rows`, then apply with the returned `confirmationToken`.

Backend table workflow:

1. Inspect with `excel.table.get_info`.
2. Use projected `excel.table.read_compact` options when only some columns or rows are needed.
3. Use `excel.table.reorder_columns` for column order changes.
4. Use `excel.table.append_rows` or `excel.table.update_rows` for data.
5. Use `excel.table.resize` only when structure must change.
6. Preserve or reapply filters with `excel.table.apply_filters` or `excel.table.preserve_filters`.
7. Validate with `excel.validate.compact`.

Avoid raw range writes inside table bodies when table tools can express the intent.
Avoid full-table rewrites for layout changes such as column reorder; they are slow on large tables and can break table identity or dependent objects.

## Create Or Update Pivots And Charts

For a standard summary PivotTable plus chart, prefer `excel.workflow.create_pivot_chart_summary`.

1. Confirm active host support with `excel.runtime.get_capabilities`.
2. Inspect source tables/ranges and existing objects.
3. For pivots, validate first with `excel.pivot.validate_source`, then use `excel.pivot.create` or `excel.pivot.rebuild_with_source`.
4. For charts, use `excel.chart.create` or `excel.chart.update_data_source`.
5. For template parity, use pivot/chart fingerprint or copy-from-template tools.
6. Refresh pivots and charts with `excel.pivot.refresh`/`excel.pivot.refresh_all` and `excel.chart.refresh`/`excel.chart.update_data_source`, then validate before reporting success.

When Office.js cannot expose deterministic pivot/chart dimensions, return the capability warning instead of inventing proof.

## Snapshot, Diff, And Rollback Preview

Use `excel.workflow.preview_risky_edit` after discovery when the requested edit is scoped and the user expects proof, diff, and rollback preview. Pass a non-empty minimal operation list and leave `apply` enabled unless the user asked for preview only. It creates the before snapshot, plan preview, scoped apply, after snapshot, diff, and rollback preview in one response.

If the combined workflow is unavailable:

1. Create a before snapshot with `excel.snapshot.create` or `excel.workbook.snapshot`.
2. Apply only the scoped edit requested by the user.
3. Create an after snapshot.
4. Compare with `excel.diff.summarize`, `excel.diff.create`, or `excel.snapshot.compare`.
5. Preview rollback with `excel.transaction.preview_rollback` or `excel.transaction.preview_rollback_chain`.

Do not actually roll back unless the user explicitly asks for rollback apply.
Do not stop after creating a plan or making the edit; a snapshot/diff/rollback-preview workflow is incomplete until the diff and rollback preview tools have both run.

## Save, Export, And Back Up

1. Use `excel.workbook.save` for normal save.
2. Use `excel.workbook.export_copy` for a true `.xlsx` copy when supported.
3. Use `excel.backup.create_file` for durable file backup records.
4. Use `excel.backup.verify` before trusting a backup path.
5. Use `excel.workbook.save_as` only when the native file bridge is configured.

Full file replacement of an open workbook is host-limited. Prefer safe open-as-new restore unless the user confirms replacement.
