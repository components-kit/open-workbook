# Tool Selection

Use the public Open Workbook surface:

```text
excel.agent.run
```

Do not call backend primitive capabilities directly in normal MCP clients. Lookup, compact reads, summaries, batches, plans, validation, snapshots, backups, transactions, jobs, locks, pivots, charts, and cleaning capabilities are internal backend concepts. Mention them only when explaining what the backend did or why a host capability is limited.

## Agent Run Modes

- `status`: check connection readiness and compact multi-agent collaboration state.
- `prepare`: cache workbook identity and lightweight structure for sheets, tables, named ranges, regions, and sheet kind; shared workbook sessions also return compact collaboration state.
- `find`: search cached metadata for candidate sheets, tables, headers, named ranges, regions, summary blocks, and formula regions.
- `answer`: perform targeted live reads and deterministic summaries; cached metadata is enough for schema-only answers.
- `auto`: compatibility route for casual prompts; explicit modes are preferred when the workflow step is known.
- `preview_update`: resolve targets, preview safe writes/table appends, block unsafe edits, and return an operation plus confirmation token.
- `apply_update`: apply only a previewed update with the returned confirmation token.
- `rollback`: map transaction or backup identifiers to recovery.
- `validate`: validate after risky changes.

If `excel.agent.run` returns `AMBIGUOUS_TARGET`, select one candidate and retry with `target.candidateId` plus the same `workbookContextId`. Do not switch to shell, Python, openpyxl, pandas, or offline `.xlsx` parsing for a connected live workbook.

If `excel.agent.run` returns `NEEDS_WORKFLOW_REDIRECT`, stop the fragmented operation sequence and call the suggested grouped `preview_update` workflow. Do not continue chunking adjacent style, clear, autofit, table, or formula operations.

## Structured Intent

When the user's LLM/client can infer intent, pass structured fields instead of relying only on English keyword matching:

- `intent.action`: one of `read_values`, `read_schema`, `list_open_workbooks`, `get_workbook_info`, `refresh_workbook_snapshot`, `get_workbook_snapshot`, `detect_external_changes`, `restore_workbook_backup`, `export_local_config`, `import_local_config`, `embed_local_config`, `read_embedded_local_config`, `import_embedded_local_config`, `close_workbook`, `prepare_session`, `create_formula_sheet`, `create_template_report`, `create_pivot_chart_summary`, `preview_risky_edit`, `inspect_analyze`, `rollback_validate`, `read_formula_patterns`, `get_formula_dependency_graph`, `trace_formula_precedents`, `trace_formula_dependents`, `validate_formula_range`, `validate_formula_against_template`, `find_formula_errors`, `explain_formula`, `copy_formula_patterns`, `fill_formula_down`, `fill_formula_right`, `repair_formula_patterns`, `convert_formulas_to_values`, `recalculate_formulas`, `read_named_item`, `create_name`, `update_name`, `delete_name`, `read_region`, `register_region`, `clear_region_values`, `write_region_values`, `fill_region`, `find_target`, `read_range_compact`, `get_range_summary`, `read_hyperlinks`, `read_comments`, `read_notes`, `read_merged_cells`, `read_data_validation`, `read_conditional_formatting`, `search_range`, `find_blank_cells`, `find_range_errors`, `write_values`, `write_formulas`, `write_number_formats`, `format_range`, `write_data_validation`, `write_conditional_formatting`, `clear_style_dimensions`, `read_style_summary`, `format_diagnostics`, `read_style_fingerprint`, `compare_style_fingerprint`, `get_theme`, `apply_theme`, `copy_style_from_template`, `repair_style_consistency`, `repair_style_from_template`, `repair_formulas_from_template`, `repair_filters_from_template`, `repair_table_structure`, `repair_print_layout`, `repair_named_ranges`, `repair_formula_errors`, `repair_merged_cells`, `detect_header_row`, `normalize_headers`, `trim_whitespace`, `remove_duplicates`, `parse_dates`, `parse_numbers`, `standardize_currency`, `fill_missing_values`, `split_column`, `merge_columns`, `detect_outliers`, `fuzzy_match`, `clear_range`, `clear_values`, `clear_values_raw`, `clear_formats`, `copy_range`, `move_range`, `write_styles_many`, `reorder_range_columns`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, `merge_range`, `unmerge_range`, `append_table_rows`, `update_table_rows`, `create_table`, `resize_table`, `reorder_table_columns`, `clear_table_data`, `clear_table_filters`, `sort_table`, `filter_range`, `set_table_total_row`, `set_table_style`, `copy_table_structure`, `validate_table_against_template`, `create_sheet`, `copy_sheet`, `rename_sheet`, `delete_sheet`, `hide_sheet`, `unhide_sheet`, `protect_sheet`, `unprotect_sheet`, `clear_sheet`, `set_sheet_tab_color`, `autofit`, `autofit_rows`, `copy_template_sheet`, `detect_templates`, `register_template`, `unregister_template`, `read_template`, `list_templates`, `infer_template_regions`, `clear_template_data_regions`, `fill_template_regions`, `validate_sheet_against_template`, `repair_sheet_from_template`, `create_snapshot`, `create_backup`, `list_snapshots`, `read_snapshot`, `compare_snapshots`, `refresh_snapshot`, `invalidate_snapshot`, `delete_snapshot`, `list_backups`, `read_backup`, `verify_backup`, `pin_backup`, `unpin_backup`, `delete_backup`, `create_file_backup`, `restore_file_backup`, `prune_backups`, `validate_compact`, `validate_workbook`, `validate_sheet`, `validate_template_consistency`, `validate_formulas`, `validate_styles`, `validate_tables`, `validate_filters`, `validate_print_layout`, `validate_no_broken_references`, `validate_no_formula_errors`, `validate_no_unintended_changes`, `calculate`, or `save`.
- `intent.targetHints`: short target clues from the workbook, including original labels and useful translated aliases.
- `target`: explicit `sheetName`, `range`, `tableName`, or returned `candidateId` when known.
- `values`: 2D matrices, `values.patches`, formulas, styles, or table rows for updates.

Structured intent is only a routing hint. The backend still performs target resolution, ambiguity checks, preview/apply confirmation, stale checks, permissions, locks, backups, validation, and rollback bookkeeping.

## Multi-Agent Coordination

For shared workbooks, start with `mode: "status"` or `mode: "prepare"` and inspect the returned `collaboration` summary. It includes counts and bounded samples for active agents, open tasks, active locks, queued/applying/blocked transactions, conflicts, and recent events.

Use the shared daemon when multiple MCP sessions may touch the same workbook. Multiple workbooks are isolated by workbook id; multiple sheets in the same workbook can proceed in parallel when scopes do not overlap; same-sheet work can proceed when ranges/objects do not conflict and otherwise queues, waits, or returns conflict guidance. Do not invent or pass agent identities in prompt text; the daemon/MCP adapter supplies trusted runtime identity.

## Multilingual Requests

For non-English prompts:

1. Preserve the user's original wording in `request`.
2. Normalize action intent into canonical English `intent.action` when clear.
3. Include `intent.targetHints` in the workbook's language plus translated aliases when helpful.
4. Pass explicit `target` and `values` when the user gave them or when the agent can safely infer them.
5. Reply in the user's language unless they ask for another language.

Do not translate the entire workbook task into English and discard the original request. Do not assume backend deterministic keyword parsing fully understands every language; the caller LLM should provide structured intent when language may matter.

## Reading Data

- Use `mode: "prepare"` once to get `workbookContextId` for follow-up calls.
- Use `mode: "find"` when the target is unknown.
- Use `list_open_workbooks`, `get_workbook_info`, `export_local_config`, or `read_embedded_local_config` with `mode: "answer"` for workbook-level metadata and local config inspection.
- Use `mode: "answer"` with `target.candidateId`, explicit `target.sheetName`/`target.range`, or a clear natural-language target.
- Ask for rows, samples, actual values, formulas, raw monthly sections, or explicit A1 ranges when live cell data is required.
- Ask for schema, columns, or headers when cached metadata is enough.
- For formula pattern, dependency, trace, error, explanation, or template comparison inspection, use `read_formula_patterns`, `get_formula_dependency_graph`, `trace_formula_precedents`, `trace_formula_dependents`, `find_formula_errors`, `explain_formula`, or `validate_formula_against_template` with an explicit formula `target.sheetName` and `target.range` when applicable.
- For named items and registered regions, use `read_named_item` or `read_region` with `target.candidateId` from `find` or `target.entity` when the name is exact.
- For registered-region mutations, use `register_region`, `clear_region_values`, `write_region_values`, or `fill_region` with `values.regionName`; include explicit `values.values` or `values.rows` for writes/fills.

Avoid prompts that require workbook-wide scans unless the user asked for audit, validation, search, or repair.

## Writing Data

- Use `mode: "preview_update"` for user-reviewable edits.
- Use `mode: "apply_update"` once with the returned top-level `operationId` and top-level `confirmationToken`; do not put the token inside `values`.
- For one small explicit value edit, `mode: "auto"` may apply safely after backend checks.
- For `write_values`, send `target.sheetName`, `target.range`, and `values.values` as a 2D matrix; row data embedded only in `request` is not used for safe writes.
- For related range edits, send one `values.patches` preview and one apply call.
- For range inspection, use `read_range_compact`, `get_range_summary`, `read_hyperlinks`, `read_comments`, `read_notes`, `read_merged_cells`, `read_data_validation`, `read_conditional_formatting`, `search_range`, `find_blank_cells`, or `find_range_errors` in answer mode with a concrete range. For historical labeling/classification examples across sheets, use `find_similar_rows` with the current row/range/table instead of reading whole prior sheets.
- For formulas, use `intent.action: "write_formulas"` and formula matrices beginning with `=`. For formula-based cell coloring or row highlighting, use `intent.action: "write_conditional_formatting"` with `values.formula` or `values.rule.formula` plus `values.style`; do not route conditional-format rules through `write_formulas`.
- For formula pattern repair from a registered template, use `repair_formula_patterns` with `values.templateId` and `target.sheetName`, then apply the returned preview once.
- For current styling, use `mode: "answer"` with `intent.action: "read_style_summary"` and an exact `target` or current selection. For formatting errors, use `intent.action: "format_diagnostics"` before mutating. If a read-style request ever returns a write-style preview or `NEEDS_WORKFLOW_REDIRECT`, do not apply it; retry the read with the explicit answer-mode style intent. For direct style writes, use `intent.action: "format_range"` with a narrow `target`; include border payloads when adding borders. For dropdown/list cells, use `read_data_validation` first when the user asks what is allowed, and use `intent.action: "write_data_validation"` with `values.validation.source`, `values.validation.formula1`, or `values.options` only when changing the rule. For autofilters, use `intent.action: "filter_range"` with the header/data range; use `clear_table_filters` for table filter removal. Do not retry these requests as sheet creation or generic value writes. For border-only/fill-only/font-only/alignment-only cleanup, use `intent.action: "clear_style_dimensions"` with `values.dimensions`, for example `["borders"]`. Use `clear_formats` only when the user asks to remove all formatting. For style audit/comparison, use `read_style_fingerprint` or `compare_style_fingerprint` with explicit source and target sheets/ranges. For workbook theme requests, use `get_theme` or `apply_theme` in answer mode and report the returned capability status; do not synthesize theme changes when the host says unavailable. For template style copy or repair, use `copy_style_from_template` or `repair_style_consistency` with explicit source/destination or `values.templateId`, then apply once.
- For style-only requests such as “make Booking look like Employees,” use `copy_style_from_template`; never duplicate the source/template sheet or change source values.
- For OCR, screenshots, forms, invoices, shipment documents, booking images, or other field/value data that must become a horizontal styled table, use `replace_range_with_styled_table` with `values.headers`, `values.row` or `values.rows`, optional `values.clearRange`, and optional `values.headerStyleSource` / `values.bodyStyleSource`. This workflow clears stale layout, writes the table, styles header/body, and autofits in one preview/apply path.
- For explicit repair intents, use `repair_style_from_template`, `repair_formulas_from_template`, or `repair_table_structure` with `mode: "preview_update"` and then apply once. Use `repair_filters_from_template`, `repair_print_layout`, `repair_named_ranges`, `repair_formula_errors`, or `repair_merged_cells` with `mode: "answer"` to get a structured repair report when the host cannot safely execute the category.
- For date text such as `26/6/26`, diagnose with `format_diagnostics` or read the exact date columns first. Prefer `parse_dates` on exact date ranges with `values.patches[*].target.range` and `numberFormat: "dd/mm/yyyy"`; if you already know the desired dates, use one grouped `write_values` preview with per-range `numberFormat`.
- For number formats, use `intent.action: "write_number_formats"` with `values.numberFormat` or `values.numberFormats`.
- For cleaning inspections, use `detect_header_row`, `detect_outliers`, or `fuzzy_match` with `mode: "answer"`. For cleaning transforms, use `normalize_headers`, `trim_whitespace`, `remove_duplicates`, `parse_dates`, `parse_numbers`, `standardize_currency`, `fill_missing_values`, `split_column`, or `merge_columns` with a concrete range, then apply once.
- For clearing values, use `intent.action: "clear_values"` to preserve formatting or `clear_values_raw` for the raw range clear operation, and state the intended scope.
- For copy/move, use `intent.action: "copy_range"` or `move_range` with explicit `values.source` and `values.destination` targets.
- For structural range edits, use `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, `merge_range`, or `unmerge_range` with an explicit sheet/range, then apply the returned preview once. For plain range column swaps/reorders, use `reorder_range_columns` with `values.columnOrder`; reserve `reorder_table_columns` for Excel tables.
- For table row changes, use `append_table_rows` or `update_table_rows` with explicit `target.tableName` and `values.rows`.
- For table structure changes, use `create_table`, `resize_table`, `reorder_table_columns`, or `copy_table_structure` with explicit table/sheet/address fields, then apply the returned preview once.
- For table presentation, use `clear_table_filters`, `set_table_total_row`, or `set_table_style` with explicit `target.tableName`.
- For sheet structure and presentation, use `create_sheet`, `copy_sheet`, `rename_sheet`, `hide_sheet`, `unhide_sheet`, or `set_sheet_tab_color`; pass `target.sheetName` for existing sheets, `values.newSheetName` for copy/rename, `values.sheetName` for create, and `values.color` for tab color. Use raw `copy_sheet` only for an exact duplicate including old values.
- For template registry reads, use `list_templates`, `read_template`, `detect_templates`, or `infer_template_regions` with `mode: "answer"` and `values.templateId` when needed.
- For template lifecycle and repair, use `register_template`, `unregister_template`, `clear_template_data_regions`, `fill_template_regions`, or `repair_sheet_from_template` with `mode: "preview_update"` and then apply once. Pass `values.templateId`, `target.sheetName` or `values.sourceSheetName`, `values.targetSheetName`, and optional `values.dataRegions` or repair dimensions when known. For `fill_template_regions`, pass explicit `values.regions` entries or a `values.regionValues` map keyed by declared data-region address.
- For “create from template,” prefer a clean template copy: preserve headers, formulas, styles, validation, and layout, then clear old data regions so fresh rows follow the same pattern. Do not leave previous business data in the new sheet unless the user explicitly asks for a raw duplicate.
- For local safety captures, use `create_snapshot` or `create_backup` with `mode: "preview_update"`; pass `target.sheetName` and `target.range` for scoped captures when possible, then apply once with the returned confirmation token.
- For safety artifact inspection, use `list_snapshots`, `read_snapshot`, `list_backups`, or `read_backup` with `mode: "answer"`; pass `values.snapshotId` or `values.backupId` for reads. Snapshot reads return compact metadata and payload counts, not the full stored payload.
- For safety artifact lifecycle work, use `compare_snapshots` and `verify_backup` with `mode: "answer"`; use `create_file_backup`, `restore_file_backup`, `prune_backups`, `refresh_snapshot`, `invalidate_snapshot`, `delete_snapshot`, `pin_backup`, `unpin_backup`, or `delete_backup` with `mode: "preview_update"` and then apply with the returned confirmation token.
- For workbook lifecycle and local config mutations, use `restore_workbook_backup`, `import_local_config`, `embed_local_config`, `import_embedded_local_config`, or `close_workbook` with `mode: "preview_update"` and then apply once. Use `refresh_workbook_snapshot`, `get_workbook_snapshot`, and `detect_external_changes` with `mode: "answer"` when you need workbook-scoped snapshot proof.
- For validation-specific checks, use the matching `validate_*` action with `mode: "validate"`; pass `target.sheetName`, `target.tableName`, `target.range`, or snapshot IDs in `values` when the validation scope needs them.

Never pad broad ranges with blanks or nulls when only a smaller rectangle should change. Never split related patches into many separate calls unless apply returns a hard failure with actionable issue details.

## Tables, Templates, Formulas, Pivots, Charts

Describe the user's goal through `excel.agent.run` and let the backend choose internal table, template, formula, pivot, chart, validation, snapshot, diff, backup, and transaction capabilities.

- For table appends, pass `target.candidateId` or `target.tableName` with `values.rows`, preview first, then apply.
- For table sort/filter, pass `intent.action: "sort_table"` or `filter_range` plus a table/range target. If one user request combines table filters and sorting, use `intent.action: "apply_table_view"` with `values.filters` and `values.sort.fields` so the backend previews and applies one table mutation.
- For template, formula repair, pivot/chart, and risky-edit tasks, use `prepare`, `find`, `answer`, `preview_update`, and `validate` as appropriate; template metadata reads can stay in `answer`, while template mutations and repairs must preview/apply; report host capability warnings honestly.
- For save/recalculate, use `intent.action: "save"` or `calculate`; the backend maps these to workbook-level operations.

## Result Handling

Treat returned `resourceLinks`, `compactProof`, `proof`, `continuation`, `invalidatedContextIds`, `invalidatedResourceUris`, `nextAction`, warnings, backup IDs, transaction IDs, and telemetry as evidence. Reuse `workbookContextId` or the returned `continuation` on follow-up `excel.agent.run` calls. `excel://...` handles are MCP/Open Workbook handles, not web URLs; never use Webfetch, browser fetch, curl, or HTTP tooling for them. Use inline `values`, `rows`, or `sparseRows` when present. If exact rows, raw values, validation evidence, transformation input, or audit detail is needed and only a preview is inline, call `excel.agent.run` once with `continuation.fullResultUri` or paste the returned `resultUri`/`fullResultUri` in `request`. If `nextAction` says `answer_now`, answer from the proof and inline data unless the task truly needs more detail.
