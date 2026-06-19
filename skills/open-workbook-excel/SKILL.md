---
name: open-workbook-excel
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP with the public excel.agent.run workflow, including inspecting workbooks, reading or writing ranges, updating tables, preserving templates, repairing formulas/styles, creating pivots/charts, validating reports, saving/exporting files, coordinating multiple agents, or normalizing multilingual user requests into structured agent intent instead of using offline spreadsheet automation."
---

# Open Workbook Excel

Use Open Workbook MCP for live desktop Excel work. It is the required first path when the workbook is open in Excel and the user expects current cell values, unsaved edits, formatting, formulas, filters, tables, pivots, charts, backups, and rollback safety to survive.

Do not inspect or modify an open workbook with shell scripts, Python, openpyxl, pandas, manual UI automation, or offline `.xlsx` parsing unless the user explicitly asks for offline file analysis or the MCP path is unavailable and you first ask permission to use a non-live fallback. If Open Workbook MCP is connected but cannot return the needed data, report that MCP limitation instead of silently bypassing it.

## First Calls

On the default MCP surface, call only:

```text
excel.agent.run
```

Use `mode: "status"` to check connection and shared-agent coordination, `mode: "prepare"` to cache lightweight workbook structure metadata, `mode: "find"` to locate sheets/tables/headers/named ranges/regions, `mode: "answer"` for targeted compact reads and deterministic summaries, `mode: "preview_update"` when a manual review step is wanted before a write or table append, `mode: "apply_update"` only with the returned `confirmationToken`, `mode: "rollback"` for recovery, and `mode: "validate"` after risky changes. Omitted mode or `mode: "auto"` remains compatible for casual prompts, but explicit modes are more predictable for agent UIs. Treat `resourceLinks`, `nextAction`, `proof`, `telemetry`, and returned `collaboration` summaries as the primary contract.

`auto` may answer vague `.xlsx` review, workbook overview, current sheet, and table-list questions from lightweight cached metadata, compare two explicitly named sheets in one call, and apply clearly scoped low-risk value edits after preview checks. Targeted row/value/formula/comparison requests upgrade the context with sheet samples only when needed. Respect `nextAction` when `auto` returns `PREVIEW_READY`, `NEEDS_INPUT`, `AMBIGUOUS_TARGET`, `VALIDATION_FAILED`, or `manual_review`; do not force formula, style, template, structural, broad, sparse, stale, or ambiguous edits through a value write.

Pass natural-language targets directly when the user speaks casually, such as "June financial sheet" or "amount column in transactions". The backend resolves them against cached workbook metadata and returns `AMBIGUOUS_TARGET` with candidates when the request is too broad.

For non-English or mixed-language requests, keep `request` in the user's original language for audit and final response context, but normalize machine-routing fields when you can infer them: use canonical English `intent.action` values such as `read_values`, `read_schema`, `list_open_workbooks`, `get_workbook_info`, `refresh_workbook_snapshot`, `get_workbook_snapshot`, `detect_external_changes`, `restore_workbook_backup`, `export_local_config`, `import_local_config`, `embed_local_config`, `read_embedded_local_config`, `import_embedded_local_config`, `close_workbook`, `prepare_session`, `create_formula_sheet`, `create_template_report`, `create_pivot_chart_summary`, `preview_risky_edit`, `inspect_analyze`, `rollback_validate`, `read_formula_patterns`, `get_formula_dependency_graph`, `trace_formula_precedents`, `trace_formula_dependents`, `validate_formula_range`, `validate_formula_against_template`, `find_formula_errors`, `explain_formula`, `copy_formula_patterns`, `fill_formula_down`, `fill_formula_right`, `repair_formula_patterns`, `convert_formulas_to_values`, `recalculate_formulas`, `read_named_item`, `create_name`, `update_name`, `delete_name`, `read_region`, `register_region`, `clear_region_values`, `write_region_values`, `fill_region`, `read_range_compact`, `get_range_summary`, `read_hyperlinks`, `read_comments`, `read_notes`, `read_merged_cells`, `read_data_validation`, `read_conditional_formatting`, `search_range`, `find_blank_cells`, `find_range_errors`, `write_values`, `write_formulas`, `write_number_formats`, `format_range`, `read_style_fingerprint`, `compare_style_fingerprint`, `get_theme`, `apply_theme`, `copy_style_from_template`, `repair_style_consistency`, `repair_style_from_template`, `repair_formulas_from_template`, `repair_filters_from_template`, `repair_table_structure`, `repair_print_layout`, `repair_named_ranges`, `repair_formula_errors`, `repair_merged_cells`, `detect_header_row`, `normalize_headers`, `trim_whitespace`, `remove_duplicates`, `parse_dates`, `parse_numbers`, `standardize_currency`, `fill_missing_values`, `split_column`, `merge_columns`, `detect_outliers`, `fuzzy_match`, `clear_range`, `clear_values`, `clear_values_raw`, `clear_formats`, `copy_range`, `move_range`, `write_styles_many`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, `merge_range`, `unmerge_range`, `append_table_rows`, `update_table_rows`, `create_table`, `resize_table`, `reorder_table_columns`, `clear_table_data`, `clear_table_filters`, `sort_table`, `filter_range`, `set_table_total_row`, `set_table_style`, `copy_table_structure`, `validate_table_against_template`, `create_sheet`, `copy_sheet`, `rename_sheet`, `delete_sheet`, `hide_sheet`, `unhide_sheet`, `protect_sheet`, `unprotect_sheet`, `clear_sheet`, `set_sheet_tab_color`, `autofit`, `autofit_rows`, `copy_template_sheet`, `detect_templates`, `register_template`, `unregister_template`, `read_template`, `list_templates`, `infer_template_regions`, `clear_template_data_regions`, `fill_template_regions`, `validate_sheet_against_template`, `repair_sheet_from_template`, `create_snapshot`, `create_backup`, `list_snapshots`, `read_snapshot`, `compare_snapshots`, `refresh_snapshot`, `invalidate_snapshot`, `delete_snapshot`, `list_backups`, `read_backup`, `verify_backup`, `pin_backup`, `unpin_backup`, `delete_backup`, `create_file_backup`, `restore_file_backup`, `prune_backups`, `validate_compact`, `validate_workbook`, `validate_sheet`, `validate_template_consistency`, `validate_formulas`, `validate_styles`, `validate_tables`, `validate_filters`, `validate_print_layout`, `validate_no_broken_references`, `validate_no_formula_errors`, `validate_no_unintended_changes`, `calculate`, or `save`; include `intent.targetHints` with workbook-native labels and useful translated aliases; pass explicit `target` and `values` when known. Do not ask the backend to translate the whole prompt. Answer the user in their language unless they ask otherwise.

If `excel.agent.run` returns `AMBIGUOUS_TARGET`, choose one returned candidate and retry with the same `workbookContextId` plus `target.candidateId`. For exact value reads, call `mode: "answer"` with `target.sheetName` and `target.range` when available, or put the sheet name and A1 range clearly in the request. Requests for rows, samples, actual values, raw monthly sheets, or explicit A1 ranges should return live read proof; schema/columns/headers-only questions can return cached metadata. Range metadata reads use answer-mode actions such as `read_hyperlinks`, `read_comments`, `read_merged_cells`, `read_data_validation`, `read_conditional_formatting`, `search_range`, `find_blank_cells`, and `find_range_errors`; range row/column insert/delete and merge/unmerge requests use preview/apply. For raw month sheets without Excel Tables, ask for the exact sheet/range or the transaction/invoice section by month, such as `Apr 2026 invoice rows`. Template registry reads use `mode: "answer"` with `list_templates`, `read_template`, `detect_templates`, or `infer_template_regions`; template registration, unregister, repair, data-region clear, and explicit data-region fill use `mode: "preview_update"` then `apply_update`. For `fill_template_regions`, pass explicit `values.regions` or `values.regionValues`; do not ask it to infer new data. Style fingerprint reads and comparisons use `read_style_fingerprint` or `compare_style_fingerprint` with `mode: "answer"`; style copy and repair use `copy_style_from_template`, `repair_style_consistency`, or `repair_style_from_template` with `mode: "preview_update"` then `apply_update`. Formula and table repair use `repair_formulas_from_template` or `repair_table_structure` with preview/apply; filter, print-layout, named-range, formula-error, and merged-cell repair intents return structured capability reports when the host path is not safely executable. Cleaning inspections use `mode: "answer"`; cleaning transforms use `mode: "preview_update"` then `apply_update` with a concrete range. For small direct value edits with explicit range and values, use `auto` and let the backend apply safely in one call. For related edits across multiple ranges, send one `mode: "preview_update"` call with `values.patches`, where each patch has `target.sheetName`, `target.range`, and a 2D `values` matrix; then call `mode: "apply_update"` once with the returned `operationId` and `confirmationToken`. To add rows to an Excel table on the default surface, call `mode: "preview_update"` with `target.candidateId` or `target.tableName` and `values.rows`, then call `mode: "apply_update"` with the returned `operationId` and `confirmationToken`.

Open Workbook's primitive operation catalog, compact reads, batch/plan/workflow tools, validation, backup, snapshot, transaction, job, task, lock, and collaboration capabilities are backend-owned for normal agents. Do not ask for or assume a separate primitive MCP surface in user workflows. Multiple MCP sessions get trusted agent identities from the daemon/MCP adapter; do not put user-provided agent IDs in the prompt as a substitute for runtime identity.

If the add-in is disconnected, ask the user to start their agent UI so it launches the configured Open Workbook MCP command, then open Excel and load the Open Workbook add-in. For manual troubleshooting, run `npx -y @components-kit/open-workbook@latest mcp` and retry. Do not fake workbook state from stale assumptions.

## Agent Run Inputs

On the default surface, keep the interaction in `excel.agent.run` instead of composing primitive reads, compact resources, validation, and mutation tools yourself:

- Use `mode: "prepare"` for workbook identity and lightweight structure context.
- Use `mode: "status"` or `prepare` before shared workbook work to see compact collaboration state.
- Use `mode: "find"` when the target is unknown.
- Use `mode: "answer"` for targeted reads, deterministic summaries, schemas, raw monthly sections, and comparisons.
- Use `mode: "preview_update"` and then `mode: "apply_update"` for user-reviewable changes.
- Use `mode: "validate"` or `mode: "rollback"` for post-change checks and recovery.

The backend owns compacting, cached metadata, target resolution, live reads, proof, validation, rollback metadata, and resource links for normal agents. Treat returned `resourceLinks`, `compactProof`, `invalidatedContextIds`, `nextAction`, and telemetry as evidence, not as mandatory follow-up calls.

The primitive capabilities below are internal/backend routing concepts, not normal agent calls:

- Lookup, compact summaries, `excel.range.read_compact`, `excel.table.read_compact`, and `excel.validate.compact` are composed by the backend to keep `agent.run` responses bounded; full details come back as MCP resource links.
- Runtime and workbook diagnostics such as `excel.runtime.get_status`, `excel.runtime.get_capabilities`, and `excel.workbook.get_workbook_map` are backend capability concepts; normal agents should request status, prepare, find, or answer through `excel.agent.run`.
- Range/table/template/style/formula/plan/batch/workflow/validation/backup/snapshot/transaction/job/task/lock/collaboration capabilities remain part of the backend operation catalog, including `excel.plan.*` and `excel.batch.*`.
- If `agent.run` returns compact proof or `nextAction: "answer_now"`, answer from that proof instead of fetching stored detail unless the user requested an audit.

For detailed routing, read `references/tool-selection.md`.

## Reliability Rules

- Never bypass Open Workbook's safety lifecycle for mutations: permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, and rollback metadata.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as 2D matrices over contiguous ranges. For repeated zone/range edits, use `values.patches` in one grouped preview instead of issuing one tool call per range.
- Never pad a broad range write with `null` or blanks when only a smaller range is intended. Write the smallest changed rectangle or use explicit clear tools.
- Read only the workbook properties needed for the task. Avoid broad workbook scans unless the task is audit, validation, search, or repair.
- Do not use full range/table reads when compact summaries, schemas, samples, or projected pages are enough. Full reads are for explicit user requests, audits, exports, or exact-data tasks.
- Do not use shell commands, Python, openpyxl, pandas, or offline `.xlsx` parsing when the workbook is already open and connected through Open Workbook MCP. If the default MCP answer path cannot return the needed live values, report the MCP limitation and ask before using a non-live fallback.
- Treat `CAPABILITY_UNAVAILABLE`, partial capability warnings, and Office.js host limits as real results. Explain them and choose a supported path.
- Preserve existing template conventions over generic formatting rules.
- After mutation, validate the affected area and surface backups, transaction IDs, warnings, diffs, and rollback options.
- Treat `compactProof` on mutation results as the default reportable proof. Do not read back whole changed sheets unless exact cell data is required.

## Critical Recipes

- Sheet create/copy/rename/hide/unhide/tab-color, formula, template-report, pivot/chart, risky-edit, and formula-repair tasks: use `excel.agent.run` modes and let the backend select the matching workflow.
- Large table work: use `prepare`, `find`, `answer`, `preview_update`, and `apply_update`; the backend should inspect schema, read compact pages only when row data is needed, mutate with table-native capabilities, then validate.
- Multi-range value updates: use one grouped `preview_update` with `values.patches`, then one `apply_update`. Do not split related patches unless apply returns a hard failure with actionable issue details.
- Unknown target work: use `mode: "find"` or retry with a returned `target.candidateId`.
- Snapshot/diff/rollback proof: prefer compact diff or risky-edit workflow results; fetch full compact resources only for detailed review.
- Multi-agent work: call `status` or `prepare`, inspect returned `collaboration` counts and samples, then describe the coordination need through `excel.agent.run`; the backend owns collaboration, task, lock, transaction, conflict, and job capabilities.

## Workflow References

- Read `references/tool-selection.md` to choose the most efficient MCP interface.
- Read `references/workflows.md` for common Excel task recipes.
- Read `references/reliability.md` for validation, rollback, stale-plan, and failure handling.
- Read `references/performance.md` before large reads/writes or latency-sensitive tasks.
- Read `references/multi-agent.md` when more than one agent or task may touch the same workbook.
