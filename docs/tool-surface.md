# Tool Surface

The full namespace is represented in the internal capability catalog, but MCP exposes only the public agent workflow tool.

## Status Model

- `stable`: implemented catalog capability; only `excel.agent.run` is public in MCP `tools/list`.
- `preview`: reserved for future internal capability metadata; preview capabilities are not exposed as MCP tools.
- `planned`: not shipped in the current catalog; unfinished names stay out of the tool surface until their contract and implementation are ready.
- `unsupported`: listed by capabilities only when a host or adapter cannot provide a deterministic implementation.

`owb mcp` exposes the agent workflow surface. Normal agents see one public tool, `excel.agent.run`, while Open Workbook keeps the compact discovery, range, table, batch, plan, workflow, validation, diff, snapshot, job, and rollback tools as internal backend capabilities. This keeps `tools/list` small and lets the backend own workbook discovery, metadata reuse, target resolution, preview/apply, validation, rollback, and compact proof generation.

The full protocol catalog remains available to backend/runtime tests and internal orchestration. It is not a normal MCP client surface; agents should not manually compose primitive compact reads or mutation tools.

## Agent Surface

Default public tool:

- `excel.agent.run`: send workbook intent through `request` and optional `mode`, `intent`, `workbookContextId`, `target`, `values`, `operationId`, and `confirmationToken`. The backend performs deterministic orchestration over internal Excel capabilities and returns structured compact output. Natural-language targets are resolved against cached sheets, tables, headers, named ranges, registered regions, summary blocks, formula regions, and the derived semantic workbook index; close matches return `AMBIGUOUS_TARGET` with candidates instead of guessing. Candidates include retry hints when budget allows; retry with `target.candidateId` from a returned candidate to force that target in a follow-up call. Caller LLMs may provide canonical English `intent.action` hints from the protocol schema, including workbook lifecycle, range read/metadata/mutation, formula, names, regions, style, repair, cleaning, table, sheet, template, snapshot, backup, validation, calculate, and save actions. These hints improve routing but never bypass target resolution, preview/apply confirmation, risk policy, stale checks, or validation. Caller LLMs may also provide `intent.targetHints`; those hints are a bounded scoring signal for candidate ranking, not an exact target override, so explicit targets and ambiguity checks still win. Workbook overview, semantic-index, style-overview, sheet-count, table-list, named-range, and vague `.xlsx` review questions answer from lightweight structure metadata before deeper reads. Schema/header-only requests can answer from cached metadata; style-overview requests combine cached metadata with bounded style fingerprints to return current style context, column groups, grouped-header suggestions, and workflow hints without full data reads; formula pattern/dependency/trace/error/explanation requests use formula-native runtime reads; range metadata requests can inspect hyperlinks, comments, notes, merged cells, data validation, conditional formatting, blanks, search hits, and errors; row/value/sample/A1-range requests perform live targeted reads. Explicit two-sheet comparison requests read both targets in one agent call. Raw monthly sheets without Excel Tables can resolve exact sheet/range references and generic semantic header blocks. Related range value edits can be grouped with `values.patches`, where each patch has `target.sheetName`, `target.range`, and a 2D `values` matrix; the backend returns one preview operation that should be applied once.

For multilingual prompts, the caller should keep `request` in the user's original language and normalize machine-routing fields into the canonical schema. Use English `intent.action` enum values, original and translated `intent.targetHints`, explicit `target` when known, and structured `values` for edits. The backend does not translate the full prompt; deterministic orchestration treats structured intent as a routing hint while preserving all safety checks.

Supported modes:

- `auto`: compatibility route for casual prompts; answer/find/prepare, and apply only clearly scoped low-risk value edits after the same preview checks. Risky edits return `PREVIEW_READY`, `NEEDS_INPUT`, or `AMBIGUOUS_TARGET`. Common style, formula, and template-copy intents route to dedicated previews instead of generic value writes.
- `status`: return runtime/add-in status.
- `prepare`: build or refresh the workbook metadata cache and return `workbookContextId`.
- `find`: locate sheets, tables, columns, ranges, regions, and summary blocks from cached metadata.
- `answer`: answer when deterministic cached metadata and targeted reads are enough; otherwise return candidates and `NEEDS_INPUT`. Schema/header/column-only requests return cached table or range metadata when possible; requests for actual rows, values, samples, raw monthly sections, or explicit A1 ranges perform live targeted reads.
- `preview_update`: prepare a scoped workbook change and return `operationId` plus `confirmationToken`.
- `apply_update`: apply a previously previewed operation.
- `operation_status`: inspect a previewed, applying, applied, failed, expired, or cancelled operation by `operationId`.
- `cancel_operation`: cancel a pending preview before apply starts.
- `validate`: run compact validation.
- `rollback`: reserved for conservative rollback orchestration.

Mutation rule: `preview_update` never applies workbook edits. It can preview scoped range value writes, grouped multi-range value patches, raw value clears, row/column insert/delete operations, plain range column reorders, merge/unmerge operations, formula writes, style updates, data-validation dropdowns, formula-based conditional formatting, template style copy/repair, sheet create/copy/rename/hide/unhide/tab-color updates, local snapshot/backup creation, workbook backup/config/close actions, template-sheet cleanup, template registry mutations, template-backed sheet repair, explicit table-row appends, and composed booking/OCR table replacement through the default agent tool. `auto` may preview and apply a clearly authorized, scoped value edit when safety checks pass; structural, destructive, broad, sparse, stale, ambiguous, style, formula, template, sheet, backup, workbook, or table append edits stop before mutation and return `nextAction`. Manual applying requires a second `excel.agent.run` call with `mode: "apply_update"`, the previewed `operationId`, and the matching `confirmationToken`. Pending previews store workbook and target-specific fingerprints; apply returns `STALE_CONTEXT` when sheet, table, named-range, formula-region, selection, or target metadata has drifted after preview. If a grouped or composed preview succeeds, agents should call `apply_update` once and should not split these related patches unless apply returns a hard failure with actionable details.

Styling workflows are first-class agent actions. Use `style_overview` for one low-read styling inspection, `improve_visual_readability` for broad safe visual styling, and `grouped_header` for structural two-level header bands. `improve_visual_readability` keeps grouped headers suggest-only, preserves grouped/header bands by default, and applies safe body styling, widths, alignment, date/money formats, filters, and highlights through one preview/apply lifecycle; `layout`, `validation`, and `freeze_panes` remain explicit opt-in buckets. If a visual preview reports `operationCount: 0` or `nextAction: "answer_now"`, agents must not call `apply_update` or decompose the task into primitive style calls. `grouped_header` accepts group payloads shaped as `{ label, startColumn, endColumn }`, `{ label, columns: [...] }`, or `{ label, range: "A:B" }`, and agents must create a fresh grouped-header preview rather than reusing a visual-readability `operationId`.

For booking images, OCR output, or client screenshots that need headers rotated into a value table while preserving workbook style, use `intent.action: "replace_range_with_styled_table"` with explicit target, values, optional `values.clearRange`, and optional `values.headerStyleSource`/`values.bodyStyleSource`. This compiles clear/write/autofit/style-copy steps into one previewed operation and one apply. Do not issue separate clear, write, autofit, and style calls for this workflow.

Agent results include `structuredContent`, a text fallback, compact proof ranges, reusable candidate ids, resource links such as `excel://agent/contexts/{workbook_context_id}`, `excel://agent/contexts/{workbook_context_id}/semantic-index`, `excel://agent/operations/{operation_id}`, and `excel://agent/results/{result_id}`, plus telemetry for payload bytes, estimated tokens, elapsed time, cache reuse, route decision, workflow route, metadata/read policy, semantic index status, caller intent source/action/acceptance, target hint count/usage, matched action handler, operation risk, target fingerprint status, metadata cache status, auto-apply decisions, internal read count, full-read cell count, candidate count, resource-link count, and estimated token savings. `budget.maxPayloadBytes`, `budget.maxEstimatedTokens`, and `budget.maxExamples` bound inline agent output; larger context is summarized or moved behind resource links. Reuse `workbookContextId` or `continuation` on follow-up calls. `excel://` handles are MCP/Open Workbook handles, not HTTP URLs; retrieve result resources through MCP resources or another `excel.agent.run` call, only when the user explicitly needs full detail.

`excel.runtime.get_capabilities` is available internally through agent status/capability flows. Its public catalog view reports only `excel.agent.run`; backend capability metadata reports the internal Excel capability catalog separately. When an add-in is connected, runtime capability data includes the host platform, Office version when available, Office API-set support such as `ExcelApi` versions, and derived feature statuses for ranges, tables, pivots, charts, and known host-limited paths. Without a connected add-in, it returns a daemon fallback that marks Excel host capabilities as `unknown`.

## Stable Tool Groups

The groups below describe the internal capability catalog used by backend orchestration and tests. They are not exposed as the normal MCP client surface.
The backend also tracks whether a capability is the public agent entrypoint, currently backed by an agent action handler, or internal-only. Capability planning status separates covered routes, unit-contract work, future orchestration candidates, host-limited paths, and deferred capabilities. This classification is for future test planning; it does not add new agent orchestration routes.

- Agent: `excel.agent.run`
- Runtime: `excel.runtime.get_status`, `excel.runtime.connect_addin`, `excel.runtime.disconnect_addin`, `excel.runtime.ping_addin`, `excel.runtime.get_capabilities`, `excel.runtime.get_active_context`, `excel.runtime.get_selection`, `excel.runtime.set_active_workbook`, `excel.runtime.set_active_sheet`
- Workbook: `excel.workbook.list_open_workbooks`, `excel.workbook.get_workbook_info`, `excel.workbook.get_workbook_map`, `excel.workbook.get_summary`, `excel.workbook.get_used_range_summary`, `excel.workbook.snapshot`, `excel.workbook.refresh_snapshot`, `excel.workbook.get_snapshot`, `excel.workbook.detect_external_changes`, `excel.workbook.calculate`, `excel.workbook.save`, `excel.workbook.save_as`, `excel.workbook.create_backup`, `excel.workbook.restore_backup`, `excel.workbook.export_copy`, `excel.workbook.export_local_config`, `excel.workbook.import_local_config`, `excel.workbook.embed_local_config`, `excel.workbook.read_embedded_local_config`, `excel.workbook.import_embedded_local_config`, `excel.workbook.close`
- File backups: `excel.backup.create_file`, `excel.backup.list`, `excel.backup.get`, `excel.backup.verify`, `excel.backup.restore_file`, `excel.backup.delete`, `excel.backup.prune`, `excel.backup.pin`, `excel.backup.unpin`
- Sheet: `excel.sheet.list`, `excel.sheet.get_info`, `excel.sheet.get_summary`, `excel.sheet.create`, `excel.sheet.copy`, `excel.sheet.copy_clean_data_regions`, `excel.sheet.rename`, `excel.sheet.delete`, `excel.sheet.hide`, `excel.sheet.unhide`, `excel.sheet.protect`, `excel.sheet.unprotect`, `excel.sheet.clear`, `excel.sheet.get_used_range`, `excel.sheet.set_tab_color`, `excel.sheet.freeze_panes`
- Range: `excel.range.read_compact`, `excel.range.get_summary`, `excel.range.read_hyperlinks`, `excel.range.read_comments`, `excel.range.read_notes`, `excel.range.read_merged_cells`, `excel.range.read_data_validation`, `excel.range.read_conditional_formatting`, `excel.range.search`, `excel.range.find_blank_cells`, `excel.range.find_errors`, `excel.range.write_values`, `excel.range.write_values_many`, `excel.range.write_formulas`, `excel.range.write_number_formats`, `excel.range.write_number_formats_many`, `excel.range.write_styles`, `excel.range.write_styles_many`, `excel.range.write_data_validation`, `excel.range.write_conditional_formatting`, `excel.range.clear_style_dimensions`, `excel.range.clear_style_dimensions_many`, `excel.range.clear`, `excel.range.clear_many`, `excel.range.clear_values`, `excel.range.clear_formats`, `excel.range.clear_formats_many`, `excel.range.clear_values_keep_format`, `excel.range.copy`, `excel.range.move`, `excel.range.reorder_columns`, `excel.range.insert_rows`, `excel.range.delete_rows`, `excel.range.insert_columns`, `excel.range.delete_columns`, `excel.range.hide_columns`, `excel.range.unhide_columns`, `excel.range.autofit_columns`, `excel.range.autofit_rows`, `excel.range.autofit_many`, `excel.range.merge`, `excel.range.unmerge`
- Lookup: `excel.lookup.search_workbook`, `excel.lookup.resolve_range`, `excel.lookup.inspect_match`
- Batch: `excel.batch.apply`, `excel.batch.submit`, `excel.batch.submit_chunked`, `excel.batch.preflight`, `excel.batch.validate`, `excel.batch.dry_run`
- Workflow: `excel.workflow.prepare_session`, `excel.workflow.create_formula_sheet`, `excel.workflow.create_template_report`, `excel.workflow.create_pivot_chart_summary`, `excel.workflow.repair_formula_errors`, `excel.workflow.preview_risky_edit`, `excel.workflow.inspect_analyze`, `excel.workflow.rollback_validate`
- Plan: `excel.plan.create`, `excel.plan.preview`, `excel.plan.refresh_preview`, `excel.plan.rebase`, `excel.plan.apply`, `excel.plan.rollback`
- Jobs: `excel.job.list`, `excel.job.get`, `excel.job.wait`, `excel.job.cancel`
- Tasks: `excel.task.create`, `excel.task.claim`, `excel.task.update`, `excel.task.set_progress`, `excel.task.add_blocker`, `excel.task.resolve_blocker`, `excel.task.evaluate_schedule`, `excel.task.resume_ready`, `excel.task.complete`, `excel.task.fail`, `excel.task.cancel`, `excel.task.list`, `excel.task.get`
- Collaboration: `excel.collab.get_status`, `excel.collab.list_agents`, `excel.collab.list_tasks`, `excel.collab.list_locks`, `excel.collab.list_transactions`, `excel.collab.get_conflicts`, `excel.collab.get_recent_events`
- Locks: `excel.lock.get_policy`, `excel.lock.set_policy`, `excel.lock.acquire`, `excel.lock.renew`, `excel.lock.release`
- Conflicts: `excel.conflict.get_guidance`, `excel.conflict.explain`, `excel.conflict.get_telemetry`, `excel.conflict.clear_telemetry`
- Transactions: `excel.transaction.get`, `excel.transaction.list`, `excel.transaction.wait`, `excel.transaction.cancel`, `excel.transaction.preview_rollback`, `excel.transaction.rollback`, `excel.transaction.preview_rollback_chain`, `excel.transaction.rollback_chain`
- Snapshot: `excel.snapshot.create`, `excel.snapshot.refresh`, `excel.snapshot.get_compact`, `excel.snapshot.compare_compact`, `excel.snapshot.invalidate`, `excel.snapshot.list`, `excel.snapshot.delete`
- Events: `excel.events.subscribe`, `excel.events.unsubscribe`, `excel.events.get_recent`, `excel.events.clear`, `excel.events.set_debounce`
- Templates: `excel.template.detect_templates`, `excel.template.register`, `excel.template.unregister`, `excel.template.get`, `excel.template.list`, `excel.template.infer_regions`, `excel.template.create_sheet_from_template`, `excel.template.clear_data_regions`, `excel.template.fill_regions`, `excel.template.validate_sheet_against_template`, `excel.template.repair_sheet_from_template`
- Style: `excel.style.get_fingerprint`, `excel.style.compare_fingerprint`, `excel.style.copy_from_template`, `excel.style.apply_style`, `excel.style.validate_consistency`, `excel.style.repair_consistency`, `excel.style.get_theme`, `excel.style.apply_theme`, `excel.style.copy_column_widths`, `excel.style.copy_row_heights`, `excel.style.copy_borders`, `excel.style.copy_fills`, `excel.style.copy_fonts`, `excel.style.copy_alignment`, `excel.style.copy_number_formats`, `excel.style.copy_conditional_formatting`, `excel.style.copy_data_validation`
- Formula: `excel.formula.read_patterns`, `excel.formula.copy_patterns`, `excel.formula.fill_down`, `excel.formula.fill_right`, `excel.formula.validate`, `excel.formula.validate_against_template`, `excel.formula.repair_patterns`, `excel.formula.find_errors`, `excel.formula.get_dependency_graph`, `excel.formula.trace_precedents`, `excel.formula.trace_dependents`, `excel.formula.convert_to_values`, `excel.formula.recalculate`, `excel.formula.explain`
- Tables: `excel.table.list`, `excel.table.get_info`, `excel.table.get_schema`, `excel.table.read_compact`, `excel.table.create`, `excel.table.resize`, `excel.table.reorder_columns`, `excel.table.append_rows`, `excel.table.update_rows`, `excel.table.clear_data_keep_formulas`, `excel.table.clear_filters`, `excel.table.apply_filters`, `excel.table.sort`, `excel.table.apply_view`, `excel.table.set_total_row`, `excel.table.set_style`, `excel.table.copy_structure`, `excel.table.validate_against_template`
- PivotTables: `excel.pivot.list`, `excel.pivot.get_info`, `excel.pivot.create`, `excel.pivot.refresh`, `excel.pivot.refresh_all`, `excel.pivot.update_source`, `excel.pivot.copy_from_template`, `excel.pivot.delete`, `excel.pivot.validate_source`, `excel.pivot.get_capability_matrix`, `excel.pivot.get_fingerprint`, `excel.pivot.compare_fingerprint`, `excel.pivot.diff`, `excel.pivot.repair_from_template`, `excel.pivot.rebuild_with_source`
- Charts: `excel.chart.list`, `excel.chart.get_info`, `excel.chart.create`, `excel.chart.update_data_source`, `excel.chart.copy_from_template`, `excel.chart.refresh`, `excel.chart.delete`, `excel.chart.validate_against_template`
- Names: `excel.names.list`, `excel.names.get`, `excel.names.create`, `excel.names.update`, `excel.names.delete`
- Regions: `excel.region.detect`, `excel.region.register`, `excel.region.list`, `excel.region.get`, `excel.region.clear_values`, `excel.region.write_values`, `excel.region.fill`
- Validation: `excel.validate.workbook`, `excel.validate.compact`, `excel.validate.sheet`, `excel.validate.template_consistency`, `excel.validate.formulas`, `excel.validate.styles`, `excel.validate.tables`, `excel.validate.filters`, `excel.validate.print_layout`, `excel.validate.no_broken_references`, `excel.validate.no_formula_errors`, `excel.validate.no_unintended_changes`
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
- `excel://compact/{resource_id}`
- `excel://agent/contexts/{workbook_context_id}`
- `excel://agent/contexts/{workbook_context_id}/semantic-index`
- `excel://agent/operations/{operation_id}`
- `excel://agent/results/{result_id}`

Resources return JSON and are registered through the MCP server. Workbook/sheet resources that need live workbook state return explicit add-in disconnected or range unavailable errors when Excel is not connected instead of fabricating stale data.

## Prompts

Generic workbook prompts are registered with the MCP server and return safe tool-use workflows:

- `excel.prompts.create_next_month_sheet`
- `excel.prompts.clean_current_sheet`
- `excel.prompts.fix_formula_errors`
- `excel.prompts.format_like_template`
- `excel.prompts.field_value_image_to_styled_table`
- `excel.prompts.booking_image_to_styled_table`
- `excel.prompts.validate_report_before_saving`
- `excel.prompts.create_summary_report`

## Rule

No mutating tool should bypass the backend safety lifecycle. It must validate permissions and create a backup before Excel receives writes, whether execution goes through the batch engine or a native Office.js object API.

## Implementation Notes

Internal mutation routes such as batch, workflow preview, plan apply, and stable sheet/range operations route through backend snapshots, backup records, target-region conflict checks, and add-in Office.js execution. They are invoked by backend orchestration behind `excel.agent.run`, not exposed as separate MCP tools.

On the public agent surface, `excel.agent.run` with `mode: "prepare"` is the preferred first call. Prepare builds lightweight structure metadata for fast workbook identity, sheet, table, and named-range context; later targeted answers upgrade that context with sheet samples only when actual rows, values, formulas, comparisons, or raw header sections are needed. Internally, workflow preflight combines runtime status, active context, capabilities, workbook map, and collaboration state before mutations.

Runtime readiness is stricter than socket connection. Status payloads include `connectionState` (`disconnected`, `connected_no_workbook`, `ready`, or `stale`), `activeAddinReachable`, and `activeWorkbookAvailable`. `excel.agent.run` probes the add-in quickly before workbook work; stale or unresponsive sessions return `NEEDS_INPUT` with taskpane reload guidance instead of waiting on a long Office.js RPC timeout.

`excel.agent.run` `mode: "prepare"` returns a server-side `workbookContextId`. Ambiguous public-surface calls return candidates with ids; pass the selected id back as `target.candidateId` with the same `workbookContextId` to continue without switching surfaces. Internally, workbook summaries, sheet summaries, compact range/table reads, and result resources can reuse that ID to avoid repeated workbook discovery; missing or ambiguous sheet/table targets return structured candidates.

Use `detailLevel` to control how much workbook data the backend fetches before response compaction:

- `workbook_summary`: metadata-only workbook summary.
- `semantic_index`: role-aware workbook target index built from cached metadata.
- `sheet_summary`: metadata-only summary for one sheet.
- `table_sample`: targeted live table sample, compact inline response.
- `full_table`: targeted live table read with full details stored behind result resources.

Lookup and compact context tools reduce token use by exposing target candidates and workbook structure before cell data. Use `excel.lookup.search_workbook` or `excel.lookup.resolve_range` when the target sheet, table, column, entity, or range is unknown. Use `excel.lookup.inspect_match` to read one bounded candidate preview instead of probing whole sheets.

`excel.workbook.get_summary`, `excel.workbook.get_used_range_summary`, `excel.sheet.get_summary`, `excel.table.get_schema`, and `excel.range.get_summary` return dimensions, table/schema metadata, truncation status, `payloadBytes`, and rough `estimatedTokens` without full cell matrices. Use them before broad reads when the target scope is already known.

`excel.range.read_compact` and `excel.table.read_compact` are the internal token-saving read paths. They default to `responseMode: "brief"`, storing full page/sample details locally while returning schema/window proof, `contextId`, `resourceUri`, `truncated`, `nextPage`, `payloadBytes`, and `estimatedTokens`. Range compact reads also accept `cellOutput: "auto" | "matrix" | "sparse"`; auto mode keeps dense small matrices readable and returns coordinate-aware `sparseRows` plus `emptySummary` for mostly empty ranges. `agent.run` uses these patterns to return bounded proof and resource links without asking agents to fetch full cell bodies unless the user requests detail.

Large or budget-limited agent results are summarized inline and point to backend-owned context resources when full detail is needed. Compact-profile outputs include `truncated`, `budgetExceeded`, `omittedCounts`, proof references, and token telemetry when data is bounded. Agents reach reusable workbook context and pending operation detail through MCP resource links returned by `excel.agent.run`; compact resource management remains backend-owned rather than advertised as separate callable Excel capabilities. Snapshot and diff detail follows this pattern through internal capability routes such as `excel.snapshot.get_compact` and `excel.snapshot.compare_compact`.

`excel.validate.compact` returns validation issue counts, severity counts, categories, and a few examples while storing the full validation report behind a compact resource URI. Use it before reporting validation proof unless the user needs the full issue list inline.

Workbook-mutating MCP tool results include an additive `compactProof` object with changed cells/ranges, backup/rollback metadata, warning counts, validation counts when present, and rough response token telemetry.
Compact-profile results may include `nextActionRecommendation`, `reasoningHints`, `confidence`, and `confidenceReasons`; agents should answer after `answer_now` unless the user requested exhaustive audit or full payload inspection.

After successful mutations, stale workbook context and stored result resources for the same workbook context are invalidated conservatively. Mutation proofs include `invalidatedContextIds` and `invalidatedResourceUris` when old handles were removed, so agents should not reuse those handles after an edit.

Combined mutating workflows also include an internal read-only preflight before mutation and return that `preflight` payload. On the public surface, use `excel.agent.run` preview/apply modes.

`excel.workflow.create_formula_sheet` is a combined formula-sheet workflow for common input-sheet tasks. It creates the sheet, writes constants and formulas through the batch lifecycle, applies number formats, and validates the formula range before returning.

`excel.workflow.create_template_report` is a combined template-report workflow for period/report sheet tasks. It creates the target sheet from a registered template, clears and fills declared data regions, compares style fingerprints, repairs styles, and validates the result against the template.

`excel.workflow.create_pivot_chart_summary` is a combined report-object workflow for PivotTable plus chart tasks. It checks PivotTable capability status, creates the PivotTable, refreshes it, creates the chart, updates/refreshes the chart source, and validates PivotTable source metadata in one response.

`excel.workflow.repair_formula_errors` is a combined formula repair workflow. It validates/finds formula errors, reads formula patterns, builds the dependency graph, repairs formulas with a scoped pattern fill or explicit formula matrix write, then validates the repaired target range.

`excel.workflow.preview_risky_edit` is a combined safe-edit interface for agents that need one call to perform a scoped risky edit and return proof. It requires at least one scoped operation, captures a before snapshot, creates and previews a plan, applies the scoped operations by default, captures an after snapshot, compares snapshots, and returns rollback preview metadata for the transaction. Set `apply: false` only when the user asked to stop after snapshot and plan preview. The workflow blocks sparse/null-padded value writes by default because those often indicate a tiny intended edit expanded into a broad overwrite; use a smaller target range, an explicit clear operation, or `allowSparseOverwrite` only for intentional broad overwrites.

`excel.workflow.inspect_analyze` is a read-only local analysis workflow for table or range profiling. It reads the selected data inside MCP, computes shape, missing values, duplicate rows, formula-error counts, type inference, and numeric min/max/average summaries, then stores the full analysis behind a compact context resource.

`excel.workflow.rollback_validate` combines transaction rollback or backup restore with workbook recalculation and compact workbook validation. Use it instead of separate rollback, calculate, and validate calls when a recovery action should end with proof that the workbook is usable.

On the public agent surface, workflow `intent.action` values return deterministic `workflow_plan` reports through `excel.agent.run`. Mutating workflow plans describe the composed backend steps and required internal capabilities, then stop for manual review or an existing preview/apply route instead of applying a multi-step workbook mutation implicitly.

Compact table reads accept optional column, row-window, and include flags so agents can inspect large tables without transferring the entire body.

`excel.batch.preflight` estimates operation count, touched cells, payload bytes, destructive level, and whether a batch should run as synchronous apply, queued submit, or chunked submit. Agents should preflight large or generated batches before execution instead of discovering size problems through add-in timeouts.

`excel.batch.submit` queues a batch mutation and returns transaction progress immediately. Use it when a mutation may take longer than an MCP client timeout, then follow with `excel.transaction.wait` or `excel.transaction.get`. `excel.batch.submit_chunked` preflights a large batch, splits safely chunkable style/value/formula/number-format work, queues child transactions, and returns one parent job. `excel.batch.apply` still waits when the writer is idle, but if another mutation is active or queued it returns queued transaction progress instead of blocking behind the queue.

Grouped range operations (`excel.range.write_values_many`, `excel.range.write_number_formats_many`, `excel.range.write_styles_many`, `excel.range.clear_many`, `excel.range.clear_formats_many`, and `excel.range.autofit_many`) apply related range entries through one reversible batch path. Agents should use one grouped preview for related values, number formats, styles, stale layout cleanup, and autofit work instead of launching many parallel single-range calls. Large style entry lists are split into queued parent jobs; set `OPEN_WORKBOOK_STYLE_BATCH_CHUNK_SIZE` to override the default chunk size.

`excel.job.list`, `excel.job.get`, `excel.job.wait`, and `excel.job.cancel` track parent jobs that split one user-visible update into multiple queued transactions. Jobs expose aggregate chunk progress and cancellation for queued child transactions. Applying chunks are not interrupted mid-Office.js call.

`excel.plan.rollback` and `excel.workbook.restore_backup` restore captured range snapshots. Full file-copy restore is tracked separately because it requires file-level user or OS involvement.

Multi-agent coordination runs through backend task, lock, transaction, conflict, collaboration, and job records. The shared daemon serializes internal batch and plan applies through a transaction queue, attaches trusted MCP agent/task/transaction metadata to results, and exposes compact collaboration status for multiple MCP adapters through `excel.agent.run` `status` and `prepare` answers. Pending previews remember the authoring agent identity; `apply_update` runs under that stored identity even when a later call comes from another MCP adapter. Internal transaction records include queue position and progress messages for queued work, and internal job records aggregate chunk progress so users see one parent progress object instead of many transaction IDs.

Plan refresh is conservative. `excel.plan.refresh_preview` and `excel.plan.rebase` re-snapshot planned target ranges and only refresh stored fingerprints when those target ranges are unchanged since the previous preview. If a user or another agent changed a target range, refresh/rebase returns `TARGET_REGION_CHANGED` and leaves the plan blocked until the agent creates a new plan.

Transaction rollback is conservative. `excel.transaction.preview_rollback` checks for later applied transactions that overlap the target scopes. `excel.transaction.rollback` only delegates to plan rollback when the preview is clean and plan metadata exists.

Rollback chains are explicit. `excel.transaction.preview_rollback_chain` finds later related transactions that must be rolled back newest-first with the target transaction. `excel.transaction.rollback_chain` requires the confirmation token returned by preview when more than one transaction is affected, and it blocks the chain if any included transaction lacks plan rollback metadata.

Task progress is first-class collaboration state. Internal task records track percentage, current step, and blockers so agents can show waiting/conflict states such as "waiting for table lock" without hiding the task in generic failure status. Normal MCP clients observe this through the compact `collaboration` summary returned by `excel.agent.run`.

Task scheduling uses dependencies, open blockers, and active lock conflicts. `excel.task.evaluate_schedule` returns ready/waiting/blocked decisions, and `excel.task.resume_ready` applies cleared wait states so agents can continue when dependencies and locks are resolved.

Lock leases are explicit. `excel.lock.acquire`, `excel.lock.renew`, and `excel.lock.release` let agents reserve scopes during planning or guarded user workflows. `excel.lock.get_policy` and `excel.lock.set_policy` control default TTL, transaction TTL, maximum TTL, and whether manual locks are allowed. Lock conflicts include the active lock id and expiry timestamp where available, and scheduler decisions expose `nextRetryAt` for lock waits.

Conflict payloads use specific codes where possible: `STRUCTURE_CONFLICT`, `TABLE_CONFLICT`, `FORMULA_DEPENDENCY_CONFLICT`, `DERIVED_OBJECT_CONFLICT`, `NAMED_RANGE_CONFLICT`, and `LOCK_CONFLICT`. `excel.collab.get_conflicts` includes guidance, while `excel.conflict.get_guidance` and `excel.conflict.explain` return structured resolution steps such as wait/retry, split scope, hand off to a task owner, refresh a plan, preview a rollback chain, repair from backup, or ask for manual review. `excel.conflict.get_telemetry` summarizes repeated contention by code, action, scope, task, and agent, including open versus cleared waits; `excel.conflict.clear_telemetry` resets that local telemetry.

Workbook file lifecycle is explicit about host limitations. `excel.workbook.save` and `excel.workbook.close` use Office.js. `excel.workbook.save_as` requires a native file bridge configured with `OPEN_WORKBOOK_FILE_BRIDGE_URL`; without it the tool returns capability status. `excel.workbook.export_copy` creates a persistent snapshot backup, then writes a true `.xlsx` from Office.js compressed file slices on Excel desktop hosts, using the native bridge first when configured. `excel.backup.create_file` creates a durable file-backup manifest with path, size, checksum, source snapshot backup id when available, pin state, verification status, and restore status. `excel.backup.verify` checks file existence and checksum, `excel.backup.pin` and `excel.backup.unpin` control retention, and `excel.backup.prune` deletes only unpinned backups selected by age or per-workbook count. `excel.backup.restore_file` defaults to safe `open-as-new` recovery and returns the verified file path; `replace-open-workbook` requires confirmation, creates a pinned emergency backup, and uses the native bridge on macOS/Windows to close, replace, and reopen the workbook. `restore-into-open-workbook` remains unsupported because full workbook replacement inside an already-open workbook is not deterministic. `excel.runtime.get_status` accepts `probeFileBridge` for a live native bridge health check when agents need to confirm reachability and supported host operations.
File-backup lifecycle operations also emit collaboration audit events so multi-agent clients can show backup creation, verification, restore, retention, and pin/unpin activity through `excel.agent.run` status or prepare summaries.

Workbook local config is portable daemon metadata. `excel.workbook.export_local_config` returns versioned JSON for registered templates, regions, and workbook-scoped permission metadata. `excel.workbook.import_local_config` loads that JSON into the local registry with merge or overwrite behavior. `excel.workbook.embed_local_config` writes the same JSON into a namespaced workbook custom XML part when the connected Excel host exposes that Office.js API. `excel.workbook.read_embedded_local_config` and `excel.workbook.import_embedded_local_config` read it back from the workbook. Local config export/import does not save or mutate workbook cells; embedded config mutates workbook metadata and therefore runs through the transaction queue and workbook-level permission checks.

Template repair creates a backup before mutating the target sheet, then copies template styles/formulas/layout through Office.js and validates the target sheet against the registered template fingerprint.

Style fidelity tools capture granular style fingerprints for a sheet or address, compare dimensions independently, and create backups before copying style dimensions. Native Office.js format copy is used for fills, fonts, borders, alignment, number formats, conditional formatting, and data validation so Excel preserves more formatting detail than the protocol models manually. Column widths and row heights are copied explicitly. Workbook theme, freeze pane, print setting, page layout, and hidden row/column replay currently return capability-status warnings where Office.js does not expose a deterministic cross-platform replay path.

Formula intelligence tools capture R1C1 formula patterns, compare pattern matrices, copy formulas from templates, fill formulas down or right, convert formulas to values with backup, recalculate workbooks, parse formula dependency edges, trace precedents/dependents, and return lightweight formula explanations. Dependency graph nodes distinguish local ranges, structured table references, and external workbook references; structured references resolve to precise header/data/totals/all table ranges when table metadata is available, and dynamic spill references expand when spill metadata is available. Formula writes include parsed local range/table dependencies in pre-commit lock checks. Formula error validation scans used ranges through Office.js special-cell APIs. Circular-reference enumeration currently returns explicit capability-status results until it can be normalized across Excel hosts.

Table, filter, and sort mutations use native Office.js table APIs. Table column reorder preserves the existing table object and uses native range copy through a temporary worksheet rather than clearing/recreating the table or transferring the full table body through MCP. The backend captures a backup over the affected table range or target structure range before mutating.

PivotTable and chart tools use native Office.js APIs. Pivot creation supports source ranges or structured tables, destination cells, row/column/filter/data field layout, aggregation, number formats, and layout flags. Pivot reads include layout and hierarchy metadata where Office.js exposes it. Pivot capability matrix reports supported, partial, unsupported, and unknown dimensions for the active host. Pivot validation can check expected source fields, axis placement, data-field aggregation, number formats, and layout settings before an agent mutates a report. Pivot fingerprints hash deterministic source, layout, data-field, and output-shape metadata with warnings for Office.js-limited dimensions; compare and diff tools use those fingerprints for review. Pivot template copy replays deterministic options, layout, hierarchy order, data aggregation/number formats, and basic field settings through a transaction-backed path when the target pivot has matching source fields; agents can restrict replay with `dimensions`. Pivot copy/repair responses include `capabilityStatus` plus warnings for deep template dimensions that Office.js cannot safely replay, such as in-place source reassignment, PivotChart-specific settings, slicers/timelines, item-level manual filters/sorts, grouping details, calculated fields/items, and host-specific settings. Pivot repair runs template copy and returns before/after fingerprint comparison. Pivot rebuild creates a new PivotTable from a desired source and can replay a template; `replaceExisting` performs an explicit delete/create flow with separate backups and transactions. Chart creation/update supports source range, chart type, series orientation, title, position, and style. Chart template copy replays deterministic chart metadata such as type, style, title, and geometry through a transaction-backed path. Pivot source reassignment still returns capability-status metadata because Office.js does not expose a safe deterministic operation.

Advanced range reads use native Office.js metadata APIs where available. Comments and legacy notes currently return explicit unsupported warnings because reliable address mapping is not implemented yet.

Named-item tools use Office.js workbook and worksheet scoped `NamedItem` collections. Region tools maintain a runtime registry of reusable sheet/address targets and can resolve existing Excel named ranges as regions. Region writes and fills route through the standard batch lifecycle.

Permission tools manage runtime policy for writes, destructive actions, workbook actions, confirmation requirements, sheet/region scope, and locked regions. `excel.batch.apply` and region/range cleaning writes are checked before Excel receives the request; direct table mutations also check scope and locked regions.

Cleaning tools read range values, transform them in the backend, and write results through the same backup-aware batch path. Detection tools such as header detection, outlier detection, and fuzzy match are read-only.

Validation tools return structured reports with `ok`, issue severity counts, categorized issues, and optional supporting data. Formula and broken-reference validators inspect workbook used ranges through compact Office.js range-area summaries. Template consistency validators reuse registered template fingerprints.

Repair tools return structured repair reports. Style and formula repairs use registered templates and create backups before mutation. Table-structure repair uses the existing table copy path. Repair categories that Office.js cannot safely execute yet return `CAPABILITY_UNAVAILABLE` with a specific reason code.

## Runtime Selection

`excel.runtime.get_selection` returns the active workbook plus the current selected range. The selection includes the range `address`, `rowCount`, `columnCount`, `cellCount`, `isSingleCell`, and `startCell`/`endCell` metadata with A1 addresses, one-based `row`/`column` values, and zero-based `rowIndex`/`columnIndex` values. For multi-cell selections, `startCell` is the top-left cell of the selected range, not necessarily Excel's internal active cell inside that selection.
