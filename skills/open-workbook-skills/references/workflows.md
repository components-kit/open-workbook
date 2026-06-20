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

On the public surface, use `excel.agent.run` with `mode: "preview_update"` and then `mode: "apply_update"` for scoped value edits. Pass `intent.action`, `target`, and `values` when known; the backend owns the primitive safety lifecycle.

1. Resolve workbook, sheet, and target address.
2. Preview non-trivial changes instead of applying blindly.
3. Apply only with the returned confirmation token.
4. Validate the target through `excel.agent.run` `mode: "validate"` for unintended changes, formula errors, or a scoped validator.
5. Return transaction IDs, backup IDs, warnings, and rollback options.

For a one-range value update, keep the call on `excel.agent.run`; the backend may route internally to range write capabilities.

For new sheets with formulas and number formats, describe a formula-sheet request through `excel.agent.run`; the backend may route to its formula-sheet workflow.

Combined mutating workflows return an internal `preflight` payload with runtime status, active context, capabilities, workbook map, and collaboration state before they mutate. On the public surface, use `excel.agent.run` preview/apply modes and let the backend select the matched workflow.

## Create A New Period Sheet From Template

For standard template report creation, describe the report through `excel.agent.run`; the backend may route to its template-report workflow.

1. Use `prepare`/`find` to identify the template and target period.
2. Use `preview_update` with a clear template-sheet request.
3. Apply only with the returned confirmation token.
4. Validate the result.

If no registered template exists, warn before mutation. Preserve the template's formatting, formulas, filters, print layout, tables, freeze panes, and named ranges.
Do not replace this workflow with `excel.sheet.copy` unless the user asks for a raw sheet duplicate or the template tool is unavailable.

## Repair Formulas Or Styles

For ordinary formula error repairs, use `excel.agent.run` with the error range plus a source formula or explicit formula matrix when known.

1. Use `answer`/`find` to identify the error range or style target.
2. Use `preview_update` with `intent.action: "write_formulas"` or `format_range` when the repair is clear.
3. Apply only after preview confirmation.
4. Recalculate/validate through `agent.run` and report warnings.

Do not convert formulas to values unless the user explicitly asks or the workflow requires a static export.

## Clean A Sheet Or Range

1. Use `find`/`answer` to locate the range and inspect headers.
2. Prefer read-only analysis before mutation.
3. For mutations, use `preview_update` with a scoped target and structured values/patches.
4. Validate formulas, tables, and unintended changes after cleaning.

Cleaning writes must stay within the requested sheet, range, table, or registered region.

## Update Tables, Filters, Or Sorts

On the default surface, append table rows with `excel.agent.run` `mode: "preview_update"` plus `target.candidateId` or `target.tableName` and `values.rows`, then apply with the returned `confirmationToken`.

Backend table workflow, expressed through `excel.agent.run`:

1. Use `prepare`/`find` to identify the table.
2. Use `answer` for schema or targeted row reads.
3. Use `preview_update` for appends, filters, sorts, or updates.
4. Apply once with the returned confirmation token.
5. Validate through `agent.run`.

Avoid raw range writes inside table bodies when table tools can express the intent.
Avoid full-table rewrites for layout changes such as column reorder; they are slow on large tables and can break table identity or dependent objects.

## Create Or Update Pivots And Charts

For a standard summary PivotTable plus chart, describe the source, pivot fields, and chart goal through `excel.agent.run`; the backend may route to its pivot/chart summary workflow.

1. Use `prepare`/`find` to identify source tables/ranges and existing objects.
2. Use `preview_update` for the requested pivot/chart summary workflow.
3. Apply only after confirmation.
4. Validate before reporting success.

When Office.js cannot expose deterministic pivot/chart dimensions, return the capability warning instead of inventing proof.

## Snapshot, Diff, And Rollback Preview

Use `excel.agent.run` `mode: "preview_update"` after discovery when the requested edit is scoped and the user expects proof, diff, and rollback preview. The backend may route to its risky-edit workflow to create the before snapshot, plan preview, scoped apply, after snapshot, diff, and rollback preview.

If the combined workflow is unavailable:

1. Ask `agent.run` for a preview/update path that returns rollback proof.
2. Apply only the scoped edit requested by the user.
3. Ask `agent.run` for validation or rollback proof.

Do not actually roll back unless the user explicitly asks for rollback apply.
Do not stop after creating a plan or making the edit; a snapshot/diff/rollback-preview workflow is incomplete until the diff and rollback preview tools have both run.

## Save, Export, And Back Up

1. Use `excel.agent.run` with `intent.action: "save"` for normal save.
2. Ask for an export/copy/backup through `agent.run`; report host capability warnings.
3. Verify and report backup or transaction metadata returned by the backend.

Full file replacement of an open workbook is host-limited. Prefer safe open-as-new restore unless the user confirms replacement.
