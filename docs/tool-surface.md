# Tool Surface

The full namespace is represented in the protocol catalog, but MCP only exposes capability-gated tools.

## Status Model

- `stable`: exposed by default.
- `preview`: exposed only with `OPEN_WORKBOOK_PREVIEW_TOOLS=1`.
- `planned`: intentionally omitted from the production MCP surface until the contract and implementation are ready.
- `unsupported`: listed by capabilities only when a host or adapter cannot provide a deterministic implementation.

`excel.runtime.get_capabilities` is the source of truth for the complete catalog, resources, prompts, runtime connection status, and active Excel host capability status. When an add-in is connected, it includes the host platform, Office version when available, Office API-set support such as `ExcelApi` versions, and derived feature statuses for ranges, tables, pivots, charts, and known host-limited paths. Without a connected add-in, it returns a daemon fallback that marks Excel host capabilities as `unknown`.

## Stable Tool Groups

- Runtime: `excel.runtime.get_status`, `excel.runtime.get_capabilities`, `excel.runtime.get_active_context`
- Workbook: `excel.workbook.list_open_workbooks`, `excel.workbook.get_workbook_info`, `excel.workbook.get_workbook_map`, `excel.workbook.snapshot`, `excel.workbook.refresh_snapshot`, `excel.workbook.get_snapshot`, `excel.workbook.detect_external_changes`, `excel.workbook.calculate`, `excel.workbook.save`, `excel.workbook.save_as`, `excel.workbook.create_backup`, `excel.workbook.restore_backup`, `excel.workbook.export_copy`, `excel.workbook.export_local_config`, `excel.workbook.import_local_config`, `excel.workbook.embed_local_config`, `excel.workbook.read_embedded_local_config`, `excel.workbook.import_embedded_local_config`, `excel.workbook.close`
- File backups: `excel.backup.create_file`, `excel.backup.list`, `excel.backup.get`, `excel.backup.verify`, `excel.backup.restore_file`, `excel.backup.delete`, `excel.backup.prune`, `excel.backup.pin`, `excel.backup.unpin`
- Sheet: `excel.sheet.list`, `excel.sheet.get_info`, `excel.sheet.create`, `excel.sheet.copy`, `excel.sheet.rename`, `excel.sheet.delete`, `excel.sheet.hide`, `excel.sheet.unhide`, `excel.sheet.protect`, `excel.sheet.unprotect`, `excel.sheet.clear`, `excel.sheet.get_used_range`, `excel.sheet.set_tab_color`
- Range: `excel.range.read_values`, `excel.range.read_formulas`, `excel.range.read_number_formats`, `excel.range.read_display_text`, `excel.range.read_styles`, `excel.range.read_full`, `excel.range.write_values`, `excel.range.write_formulas`, `excel.range.write_number_formats`, `excel.range.write_styles`, `excel.range.clear`, `excel.range.clear_values`, `excel.range.clear_formats`, `excel.range.clear_values_keep_format`, `excel.range.copy`, `excel.range.move`, `excel.range.insert_rows`, `excel.range.delete_rows`, `excel.range.insert_columns`, `excel.range.delete_columns`, `excel.range.autofit_columns`, `excel.range.autofit_rows`, `excel.range.merge`, `excel.range.unmerge`
- Advanced range reads: `excel.range.read_hyperlinks`, `excel.range.read_comments`, `excel.range.read_notes`, `excel.range.read_merged_cells`, `excel.range.read_data_validation`, `excel.range.read_conditional_formatting`, `excel.range.search`, `excel.range.find_blank_cells`, `excel.range.find_errors`
- Batch and plan: `excel.batch.validate`, `excel.batch.dry_run`, `excel.batch.apply`, `excel.plan.create`, `excel.plan.preview`, `excel.plan.refresh_preview`, `excel.plan.rebase`, `excel.plan.apply`, `excel.plan.rollback`
- Multi-agent coordination: `excel.task.create`, `excel.task.claim`, `excel.task.update`, `excel.task.set_progress`, `excel.task.add_blocker`, `excel.task.resolve_blocker`, `excel.task.evaluate_schedule`, `excel.task.resume_ready`, `excel.task.complete`, `excel.task.fail`, `excel.task.cancel`, `excel.task.list`, `excel.task.get`, `excel.collab.get_status`, `excel.collab.list_agents`, `excel.collab.list_tasks`, `excel.collab.list_locks`, `excel.collab.list_transactions`, `excel.collab.get_conflicts`, `excel.collab.get_recent_events`, `excel.lock.get_policy`, `excel.lock.set_policy`, `excel.lock.acquire`, `excel.lock.renew`, `excel.lock.release`, `excel.conflict.get_guidance`, `excel.conflict.explain`, `excel.conflict.get_telemetry`, `excel.conflict.clear_telemetry`, `excel.transaction.list`, `excel.transaction.get`, `excel.transaction.preview_rollback`, `excel.transaction.rollback`, `excel.transaction.preview_rollback_chain`, `excel.transaction.rollback_chain`
- Snapshot and diff: `excel.snapshot.create`, `excel.snapshot.refresh`, `excel.snapshot.get`, `excel.snapshot.compare`, `excel.snapshot.invalidate`, `excel.snapshot.list`, `excel.snapshot.delete`, `excel.diff.create`, `excel.diff.summarize`, `excel.diff.get_details`, `excel.diff.export_json`, `excel.diff.export_html`
- Events: `excel.events.subscribe`, `excel.events.unsubscribe`, `excel.events.get_recent`, `excel.events.clear`, `excel.events.set_debounce`
- Templates: `excel.template.detect_templates`, `excel.template.register`, `excel.template.unregister`, `excel.template.get`, `excel.template.list`, `excel.template.infer_regions`, `excel.template.create_sheet_from_template`, `excel.template.clear_data_regions`, `excel.template.fill_regions`, `excel.template.validate_sheet_against_template`, `excel.template.repair_sheet_from_template`
- Style and formula preservation: `excel.style.get_fingerprint`, `excel.style.compare_fingerprint`, `excel.style.copy_from_template`, `excel.style.apply_style`, `excel.style.validate_consistency`, `excel.style.repair_consistency`, `excel.style.get_theme`, `excel.style.apply_theme`, `excel.style.copy_column_widths`, `excel.style.copy_row_heights`, `excel.style.copy_borders`, `excel.style.copy_fills`, `excel.style.copy_fonts`, `excel.style.copy_alignment`, `excel.style.copy_number_formats`, `excel.style.copy_conditional_formatting`, `excel.style.copy_data_validation`, `excel.style.copy_freeze_panes`, `excel.style.copy_print_settings`, `excel.style.copy_page_layout`, `excel.style.copy_hidden_rows_columns`, `excel.formula.read_patterns`, `excel.formula.copy_patterns`, `excel.formula.fill_down`, `excel.formula.fill_right`, `excel.formula.validate`, `excel.formula.validate_against_template`, `excel.formula.repair_patterns`, `excel.formula.find_errors`, `excel.formula.find_circular_references`, `excel.formula.get_dependency_graph`, `excel.formula.trace_precedents`, `excel.formula.trace_dependents`, `excel.formula.convert_to_values`, `excel.formula.recalculate`, `excel.formula.explain`
- Tables: `excel.table.list`, `excel.table.get_info`, `excel.table.read`, `excel.table.create`, `excel.table.resize`, `excel.table.append_rows`, `excel.table.update_rows`, `excel.table.clear_data_keep_formulas`, `excel.table.clear_filters`, `excel.table.apply_filters`, `excel.table.preserve_filters`, `excel.table.sort`, `excel.table.set_total_row`, `excel.table.set_style`, `excel.table.copy_structure`, `excel.table.validate_against_template`
- Filters and sort: `excel.filter.get_filters`, `excel.filter.apply`, `excel.filter.clear`, `excel.filter.preserve_from_template`, `excel.filter.validate`, `excel.sort.apply`, `excel.sort.clear`, `excel.sort.preserve_from_template`
- PivotTables and charts: `excel.pivot.list`, `excel.pivot.get_info`, `excel.pivot.create`, `excel.pivot.refresh`, `excel.pivot.refresh_all`, `excel.pivot.update_source`, `excel.pivot.copy_from_template`, `excel.pivot.delete`, `excel.pivot.validate_source`, `excel.pivot.get_capability_matrix`, `excel.pivot.get_fingerprint`, `excel.pivot.compare_fingerprint`, `excel.pivot.diff`, `excel.pivot.repair_from_template`, `excel.pivot.rebuild_with_source`, `excel.chart.list`, `excel.chart.get_info`, `excel.chart.create`, `excel.chart.update_data_source`, `excel.chart.copy_from_template`, `excel.chart.refresh`, `excel.chart.delete`, `excel.chart.validate_against_template`
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
- `excel://workbooks/{workbook_id}/sheets/{sheet_name}/used-range`
- `excel://workbooks/{workbook_id}/sheets/{sheet_name}/style-fingerprint`
- `excel://workbooks/{workbook_id}/sheets/{sheet_name}/formula-patterns`
- `excel://workbooks/{workbook_id}/tables`
- `excel://workbooks/{workbook_id}/templates`
- `excel://workbooks/{workbook_id}/snapshots/{snapshot_id}`
- `excel://workbooks/{workbook_id}/plans/{plan_id}/diff`

Resources return JSON and are registered through the MCP server. Workbook/sheet resources that need live workbook state return explicit add-in disconnected or range unavailable errors when Excel is not connected instead of fabricating stale data.

## Prompts

Generic workbook prompts are registered with the MCP server and return safe tool-use workflows:

- `excel.prompts.create_next_month_sheet`
- `excel.prompts.clean_current_sheet`
- `excel.prompts.fix_formula_errors`
- `excel.prompts.format_like_template`
- `excel.prompts.validate_report_before_saving`
- `excel.prompts.create_summary_report`

## Rule

No mutating tool should bypass the backend safety lifecycle. It must validate permissions and create a backup before Excel receives writes, whether execution goes through the batch engine or a native Office.js object API.

## Implementation Notes

`excel.batch.apply`, `excel.plan.apply`, and all stable mutating sheet/range tools route through backend snapshots, backup records, target-region conflict checks, and add-in Office.js execution.

`excel.plan.rollback` and `excel.workbook.restore_backup` restore captured range snapshots. Full file-copy restore is tracked separately because it requires file-level user or OS involvement.

Multi-agent coordination runs through task, lock, and transaction records. The shared daemon serializes `excel.batch.apply` and `excel.plan.apply` through a transaction queue, attaches agent/task/transaction metadata to results, and exposes collaboration status for multiple MCP adapters.

Plan refresh is conservative. `excel.plan.refresh_preview` and `excel.plan.rebase` re-snapshot planned target ranges and only refresh stored fingerprints when those target ranges are unchanged since the previous preview. If a user or another agent changed a target range, refresh/rebase returns `TARGET_REGION_CHANGED` and leaves the plan blocked until the agent creates a new plan.

Transaction rollback is conservative. `excel.transaction.preview_rollback` checks for later applied transactions that overlap the target scopes. `excel.transaction.rollback` only delegates to plan rollback when the preview is clean and plan metadata exists.

Rollback chains are explicit. `excel.transaction.preview_rollback_chain` finds later related transactions that must be rolled back newest-first with the target transaction. `excel.transaction.rollback_chain` requires the confirmation token returned by preview when more than one transaction is affected, and it blocks the chain if any included transaction lacks plan rollback metadata.

Task progress is first-class collaboration state. `excel.task.set_progress` updates percentage and current step, while `excel.task.add_blocker` and `excel.task.resolve_blocker` let agents show waiting/conflict states such as "waiting for table lock" without hiding the task in generic failure status.

Task scheduling uses dependencies, open blockers, and active lock conflicts. `excel.task.evaluate_schedule` returns ready/waiting/blocked decisions, and `excel.task.resume_ready` applies cleared wait states so agents can continue when dependencies and locks are resolved.

Lock leases are explicit. `excel.lock.acquire`, `excel.lock.renew`, and `excel.lock.release` let agents reserve scopes during planning or guarded user workflows. `excel.lock.get_policy` and `excel.lock.set_policy` control default TTL, transaction TTL, maximum TTL, and whether manual locks are allowed. Lock conflicts include the active lock id and expiry timestamp where available, and scheduler decisions expose `nextRetryAt` for lock waits.

Conflict payloads use specific codes where possible: `STRUCTURE_CONFLICT`, `TABLE_CONFLICT`, `FORMULA_DEPENDENCY_CONFLICT`, `DERIVED_OBJECT_CONFLICT`, `NAMED_RANGE_CONFLICT`, and `LOCK_CONFLICT`. `excel.collab.get_conflicts` includes guidance, while `excel.conflict.get_guidance` and `excel.conflict.explain` return structured resolution steps such as wait/retry, split scope, hand off to a task owner, refresh a plan, preview a rollback chain, repair from backup, or ask for manual review. `excel.conflict.get_telemetry` summarizes repeated contention by code, action, scope, task, and agent, including open versus cleared waits; `excel.conflict.clear_telemetry` resets that local telemetry.

Workbook file lifecycle is explicit about host limitations. `excel.workbook.save` and `excel.workbook.close` use Office.js. `excel.workbook.save_as` requires a native file bridge configured with `OPEN_WORKBOOK_FILE_BRIDGE_URL`; without it the tool returns capability status. `excel.workbook.export_copy` creates a persistent snapshot backup, then writes a true `.xlsx` from Office.js compressed file slices on Excel desktop hosts, using the native bridge first when configured. `excel.backup.create_file` creates a durable file-backup manifest with path, size, checksum, source snapshot backup id when available, pin state, verification status, and restore status. `excel.backup.verify` checks file existence and checksum, `excel.backup.pin` and `excel.backup.unpin` control retention, and `excel.backup.prune` deletes only unpinned backups selected by age or per-workbook count. `excel.backup.restore_file` defaults to safe `open-as-new` recovery and returns the verified file path; `replace-open-workbook` requires confirmation, creates a pinned emergency backup, and uses the native bridge on macOS/Windows to close, replace, and reopen the workbook. `restore-into-open-workbook` remains unsupported because full workbook replacement inside an already-open workbook is not deterministic. `excel.runtime.get_status` accepts `probeFileBridge` for a live native bridge health check when agents need to confirm reachability and supported host operations.
File-backup lifecycle operations also emit collaboration audit events so multi-agent clients can show backup creation, verification, restore, retention, and pin/unpin activity through `excel.collab.get_status`.

Workbook local config is portable daemon metadata. `excel.workbook.export_local_config` returns versioned JSON for registered templates, regions, and workbook-scoped permission metadata. `excel.workbook.import_local_config` loads that JSON into the local registry with merge or overwrite behavior. `excel.workbook.embed_local_config` writes the same JSON into a namespaced workbook custom XML part when the connected Excel host exposes that Office.js API. `excel.workbook.read_embedded_local_config` and `excel.workbook.import_embedded_local_config` read it back from the workbook. Local config export/import does not save or mutate workbook cells; embedded config mutates workbook metadata and therefore runs through the transaction queue and workbook-level permission checks.

Template repair creates a backup before mutating the target sheet, then copies template styles/formulas/layout through Office.js and validates the target sheet against the registered template fingerprint.

Style fidelity tools capture granular style fingerprints for a sheet or address, compare dimensions independently, and create backups before copying style dimensions. Native Office.js format copy is used for fills, fonts, borders, alignment, number formats, conditional formatting, and data validation so Excel preserves more formatting detail than the protocol models manually. Column widths and row heights are copied explicitly. Workbook theme, freeze pane, print setting, page layout, and hidden row/column replay currently return capability-status warnings where Office.js does not expose a deterministic cross-platform replay path.

Formula intelligence tools capture R1C1 formula patterns, compare pattern matrices, copy formulas from templates, fill formulas down or right, convert formulas to values with backup, recalculate workbooks, parse formula dependency edges, trace precedents/dependents, and return lightweight formula explanations. Dependency graph nodes distinguish local ranges, structured table references, and external workbook references; structured references resolve to precise header/data/totals/all table ranges when table metadata is available, and dynamic spill references expand when spill metadata is available. Formula writes include parsed local range/table dependencies in pre-commit lock checks. Formula error validation scans used ranges through Office.js special-cell APIs. Circular-reference enumeration currently returns explicit capability-status results until it can be normalized across Excel hosts.

Table, filter, and sort mutations use native Office.js table APIs. The backend captures a backup over the affected table range or target structure range before mutating.

PivotTable and chart tools use native Office.js APIs. Pivot creation supports source ranges or structured tables, destination cells, row/column/filter/data field layout, aggregation, number formats, and layout flags. Pivot reads include layout and hierarchy metadata where Office.js exposes it. Pivot capability matrix reports supported, partial, unsupported, and unknown dimensions for the active host. Pivot validation can check expected source fields, axis placement, data-field aggregation, number formats, and layout settings before an agent mutates a report. Pivot fingerprints hash deterministic source, layout, data-field, and output-shape metadata with warnings for Office.js-limited dimensions; compare and diff tools use those fingerprints for review. Pivot template copy replays deterministic options, layout, hierarchy order, data aggregation/number formats, and basic field settings through a transaction-backed path when the target pivot has matching source fields; agents can restrict replay with `dimensions`. Pivot copy/repair responses include `capabilityStatus` plus warnings for deep template dimensions that Office.js cannot safely replay, such as in-place source reassignment, PivotChart-specific settings, slicers/timelines, item-level manual filters/sorts, grouping details, calculated fields/items, and host-specific settings. Pivot repair runs template copy and returns before/after fingerprint comparison. Pivot rebuild creates a new PivotTable from a desired source and can replay a template; `replaceExisting` performs an explicit delete/create flow with separate backups and transactions. Chart creation/update supports source range, chart type, series orientation, title, position, and style. Chart template copy replays deterministic chart metadata such as type, style, title, and geometry through a transaction-backed path. Pivot source reassignment still returns capability-status metadata because Office.js does not expose a safe deterministic operation.

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
- `excel.runtime.get_selection`
- `excel.runtime.set_active_workbook`
- `excel.runtime.set_active_sheet`

These are marked preview because they establish runtime-control contracts that can affect session routing.

`excel.runtime.get_selection` returns the active workbook plus the current selected range. The selection includes the range `address`, `rowCount`, `columnCount`, `cellCount`, `isSingleCell`, and `startCell`/`endCell` metadata with A1 addresses, one-based `row`/`column` values, and zero-based `rowIndex`/`columnIndex` values. For multi-cell selections, `startCell` is the top-left cell of the selected range, not necessarily Excel's internal active cell inside that selection.
