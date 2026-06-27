---
name: open-workbook-skills
description: "Use when an agent needs to automate live Microsoft Excel workbooks through Open Workbook MCP with the public excel.agent.run workflow: inspect workbooks, read/write ranges, update tables, preserve templates, repair formulas/styles, create pivots/charts, validate reports, save/export files, coordinate multiple agents, or normalize multilingual workbook requests into structured intent."
---

# Open Workbook Skills

Use Open Workbook MCP for live desktop Excel work. It is the required first path when the workbook is open in Excel and the user expects current cell values, unsaved edits, formatting, formulas, filters, tables, pivots, charts, snapshots, backups, transaction history, and rollback safety to survive.

Do not inspect or modify a connected live workbook with shell scripts, Python, openpyxl, pandas, manual UI automation, or offline `.xlsx` parsing unless the user explicitly asks for offline file analysis or the MCP path is unavailable and the user approves the fallback. For workbook/worksheet "look at", "inspect", "review", or "what is this" requests, do not run `ls`, `stat`, Python, or offline file checks as a preflight; call `excel.agent.run` first. If Open Workbook is connected but returns an empty live-read diagnostic or `cannot_complete`, report that Open Workbook live-read failure instead of switching to saved-file parsing.

## Public Surface

Normal agents call one tool:

```text
excel.agent.run
```

Use modes deliberately:

- `status`: check Excel/add-in readiness and shared-agent coordination.
- `prepare`: cache workbook identity, sheet/table/name/region metadata, and collaboration state.
- `find`: locate candidate sheets, tables, headers, named ranges, regions, formulas, and summary blocks.
- `answer`: targeted live reads, schemas, summaries, comparisons, and deterministic inspection.
- `preview_update`: preview mutations and get an `operationId` plus `confirmationToken` when review is required.
- `apply_update`: apply only a previewed operation with the returned token.
- `operation_status`: check a returned operation without reissuing preview/apply.
- `cancel_operation`: cancel a pending preview before apply starts.
- `validate`: validate workbook, sheet, table, formula, style, filter, print-layout, or unintended-change scope.
- `rollback`: inspect or apply recovery through returned transaction, snapshot, or backup guidance.
- `auto`: default route for casual prompts and small explicit value edits; safe narrow edits may apply in one call after workbook write access is allowed for the session unless `autoApply: false`.

The backend owns primitive Excel capabilities, compact reads, batches, plans, validation, snapshots, backups, locks, jobs, transactions, resources, and collaboration records. Do not ask for or assume a separate primitive MCP surface in user workflows.

`excel://...` result handles returned by Open Workbook are internal MCP/Open Workbook handles, not web URLs. Never use Webfetch, browser fetch, curl, or HTTP tooling for `resultUri`, `fullResultUri`, or `resourceLinks`. If hidden detail is genuinely required, call `excel.agent.run` once with `continuation.fullResultUri` or paste the `excel://...` handle in the `request`.

For "what is this workbook/file?", "summarize this workbook", "look into this Excel file", and similar overview requests, make one `excel.agent.run` call with `mode: "answer"` and `detailLevel: "workbook_summary"` or `detailLevel: "sheet_summary"`. When the response says `nextAction: "answer_now"` or `maxRecommendedFollowupCalls: 0`, answer the user immediately. Do not fetch `fullResultUri`, chunk-read worksheets, list MCP resources, or call low-level resource reads unless the user explicitly asks for all raw rows, every value, or exact cell contents.

## Structured Intent

When the caller LLM can infer routing, pass structured fields alongside the natural request:

- `intent.action`: canonical English action such as `read_values`, `read_formulas`, `append_table_rows`, `write_values`, `validate_workbook`, `create_backup`, `restore_file_backup`, `rollback_validate`, `save`, or another supported backend action.
- `intent.targetHints`: workbook-native labels plus useful translated aliases.
- `target`: explicit `sheetName`, `range`, `tableName`, named item, region, or returned `candidateId`.
- `values`: 2D values, formulas, styles, table rows, patches, template data, backup IDs, or confirmation data.

For mutations, do not put the write payload only in `request`. Send structured data: `write_values` needs `target.sheetName`, `target.range`, and `values.values` as a 2D matrix; table appends use `values.rows`; multi-range edits use `values.patches`.

### Multiple Updates

When one user instruction contains multiple explicit value edits, send one `excel.agent.run` call with `mode: "auto"` and `values.patches`. A different topic or separate row does not mean different tool call; same user instruction plus explicit ranges means one grouped patch. Do not issue parallel or sequential update calls unless the grouped call fails with actionable details.

Example:

```json
{
  "mode": "auto",
  "intent": { "action": "write_values" },
  "target": { "sheetName": "May 2026" },
  "values": {
    "patches": [
      { "target": { "sheetName": "May 2026", "range": "C19:E19" }, "values": [["700-5229", "maintenance note", "maintenance"]] },
      { "target": { "sheetName": "May 2026", "range": "D20:E20" }, "values": [["Owner cash top-up", "owner_cash_topup"]] },
      { "target": { "sheetName": "May 2026", "range": "H20:J20" }, "values": [[10000, 0, "Owner fund top-up"]] }
    ]
  }
}
```

When a sheet summary or semantic index identifies a section with header/data anchors, prefer anchor writes over raw coordinates. For “set/update the Klongtoey row Vendor Propose value” style tasks, send `values.semanticPatches` with `sectionId` or `sectionLabel`, `rowMatch` (`column` plus `value`), `columnMatch`, and `value`. This lets Open Workbook resolve the exact cell from the row label and column header, avoids reading whole sections, and keeps safe small edits eligible for one-call `mode: "auto"` after session write access exists.

Use `detailLevel` conservatively: `workbook_summary` for metadata-only workbook context, `semantic_index` for role-aware workbook targets/candidates, `sheet_summary` for “look at/check/how is this sheet” overview requests without live cell reads, `table_sample` for a bounded live table sample, and `full_table` only when the task requires all rows, every value, or full table contents.

For broad deterministic value changes, do not read whole columns into model context and then generate a large write matrix. Use `intent.action: "transform_values"` for one-column transforms such as add prefix/suffix, replace text, normalize whitespace/case, fill blanks, or map values. Use `intent.action: "derive_values"` for row-aware updates such as “fill Category from Suggested Category,” “copy X from Y if blank,” “extract ID from Description,” `formula_like` calculations such as Payment Variance = Actual Amount - Cash Amount, or conditional/lookup-style mappings. Use `intent.action: "settle_reconciliation"` for transaction settlement bundles that need Payment Variance plus Reconciliation Note and Detail Notes kept consistent. Use `intent.action: "transform_sheets"` for workbook structure batches such as adding a prefix/suffix to many sheet names. These workflows let Open Workbook scan rows or sheet metadata internally, preview bounded source/before/after examples, and apply only changed cells or sheet renames with rollback.

For multilingual requests, preserve the original language in `request`, normalize routing fields to canonical English when clear, keep target hints in the workbook/user language, and answer the user in their language unless asked otherwise. Do not translate the whole task into English and discard the original wording.

Read `references/agent-run.md` for mode contracts, preview/apply behavior, multilingual routing, and result handling. Read `references/capability-map.md` when you need exact internal action groups.

## Selection-First Targeting

When the user request points at the current place in Excel, prefer the current selection as the first target. This includes one cell, whole rows, whole columns, and rectangular ranges. A selected cell is normal while someone works in Excel; do not treat it as the target for broad workbook/worksheet overview requests unless the wording says "this", "here", "selected", "current cell/range/row/column", or asks for values/rows from the selected area.

Never tell the user you cannot detect the selected row/cell/range before calling `excel.agent.run`. The model cannot see Excel selection by itself, but Open Workbook can read it through the connected add-in. If the user says "what do you think about this?", "check this", "does this look right?", or similar vague wording, call `excel.agent.run` with `mode: "answer"` and no explicit `target`; let the backend resolve the live selection. Ask the user to select a cell/range or reload the taskpane only after `excel.agent.run` returns selection unavailable or stale.

Target priority:

1. Explicit user target in the prompt or structured `target`.
2. Current Excel selection when the user says "this", "here", "selected", "current cell/range/row/column", "look at this", "fix this", or asks for values/rows from the selected area.
3. Sheet, table, or workbook discovery only when no usable selection exists or the prompt clearly asks for broader context.

For reads, start with a bounded `answer` call on the selected range. For mutations, preview only the selected scope unless the user explicitly asks to expand to a table, sheet, or workbook. If a single selected cell is inside a table or header-shaped region and the request is vague ("this", "check this", "what do you think"), let the backend include the active record row as context while preserving proof of the exact selected cell.

## Workflow Semantics

- Raw duplicate: use `copy_sheet` only when the user explicitly asks for an exact copy including old values.
- Create from template: use template workflows to preserve structure, formulas, styles, validation, and layout while clearing old data regions for fresh entry.
- Apply style from template: use `copy_style_from_template`; the source/template sheet is a style source only and must not be duplicated or mutated.
- Replace styled table: use `replace_range_with_styled_table` to clear stale layout, write headers/rows, copy header/body style samples, and autofit in one preview/apply workflow.
- Inspect current styling: call `excel.agent.run` with `mode: "answer"`, `intent.action: "read_style_summary"`, and an exact range or current selection. For styling recommendations or best-practice review, call `intent.action: "style_overview"` or `detailLevel: "style_overview"` once; it returns current style context, column groups, grouped-header suggestions, and safe next workflow hints without reading full data rows. Use `read_style_fingerprint` for template comparison, not normal user-facing style inspection. When the user asks for a style/template reference from another sheet or month, call `intent.action: "find_style_references"` so Open Workbook returns bounded source candidates instead of reading whole sheets.
- Grouped-header summary: when the user asks to look at, examine, or summarize grouped headers such as row 1 grouped headers, call `excel.agent.run` once with `mode: "answer"` and a target sheet/range if known. The backend returns `answer.kind: "grouped_header_summary"` with spans, labels, merged count, and unmerged labels. Do not call `workbook_design_overview`, `semantic_index`, full result resources, or broad row value reads for the same summary.
- Workbook design review: when the user asks for a column-by-column recommendation such as whether each column should be free text, date, money/number, ID/text code, dropdown, or lookup/reference from another sheet, call `excel.agent.run` once with `mode: "answer"` and `intent.action: "workbook_design_overview"`. It returns column recommendations, related-sheet hints, format/dropdown/lookup suggestions, and next workflows from cached metadata. Do not manually read Customer, Bookings, Drivers, or empty data rows first; use targeted `read_data_validation`, `write_data_validation`, or lookup/formula previews only after the user chooses a recommendation.
- Visual readability: when the user asks to make a sheet cleaner, easier to read, office-ready, visually grouped, highlighted, or formatted broadly, call `excel.agent.run` with `mode: "preview_update"` and `intent.action: "improve_visual_readability"` instead of issuing many low-level style operations. Default to `values.visualReadability.styleDepth: "standard"`, preserve formulas, and apply only after the returned `operationId` and `confirmationToken` when `nextAction` is `call_apply_update` and metrics show `operationCount > 0`. If the preview reports `operationCount: 0` or `nextAction: "answer_now"`, do not call `apply_update` and do not decompose the styling into primitive `format_range` calls; explain the skipped reasons and ask for the supported next workflow. Style preservation defaults to `stylePreservationMode: "protected_regions"` so summary/template areas and grouped header bands stay guarded while ordinary table body styling, widths, alignment, and date/money formats can still be improved. Use `"strict"` only when the user asks to preserve every existing style, and `"none"` for an explicit redesign. Use `styleDepth: "basic"` for fast low-risk layout/format cleanup, `styleDepth: "comprehensive"` for deeper suggestions, and `values.visualReadability.applySuggestionBuckets` for explicit actionable suggestions: `["layout"]` for wrap/row-height styling, `["validation"]` for dropdown writes, and `["freeze_panes"]` for freeze rows/columns. Grouped headers are suggestions for wide tables by default; to apply inserted group rows, merged labels, and matching header colors, use `mode: "preview_update"` with `intent.action: "grouped_header"` and optional `values.groupedHeader.groups`. Prefer grouped-header groups as `{ "label": "...", "startColumn": "A", "endColumn": "B" }`; `{ "columns": ["A", "B"] }` and `{ "range": "A:B" }` are accepted. Never continue grouped_header with an old visual-readability `operationId`. For freeze column requests, use `freezePanes: { "columns": 1 }` or a clear request like "freeze first column". Keep formulas, summary blocks, reference-style layout cues, and print settings as separate confirmed workflows or preview-only when the host capability is unavailable.
- Structural styling permission: grouped headers insert rows and merge ranges, so they require structure/destructive permission. If `apply_update` returns `DESTRUCTIVE_ACTION_BLOCKED` or `PERMISSION_DENIED`, do not retry the stale preview. Ask the user for permission, then call `excel.agent.run` with `intent.action: "set_permissions"` and `values.permissions: { "allowWrites": true, "allowDestructiveActions": true, "scopeToWorkbook": true, "requireConfirmationFor": [] }`; after success, create a fresh grouped_header preview and apply that fresh operation.
- Multi-range merge and alignment: when the user asks to merge several grouped-header spans and center them, send one `preview_update` with `values.merges` or `values.entries` containing each `sheetName`/`range` and optional `style`. The preview must include `range.merge` operations before the `range.write_styles_many` operation. Do not claim a range was merged if the preview/apply only contains style updates.
- Visual styling safety: comprehensive validation/formula suggestions remain preview-only unless the user chooses an explicit bucket or separate workflow; do not apply dropdowns, formulas, inserted rows/columns, or summary blocks through the visual styling apply path.
- Formatting errors: use `format_diagnostics` before mutating. It returns raw value, displayed text, formulas, number formats, style summary, likely issues, and suggested fix actions.
- Historical labels and similar rows: when the user asks how something was labeled/classified before, or asks to look at another month/sheet for a data reference, call `intent.action: "find_similar_rows"` on the current row/range/table or requested prior sheet. Let Open Workbook search related sheets and return exact matched rows with proof; do not manually read broad prior-sheet ranges, fetch fullResultUri, or chunk rows looking for matches.
- Formulas: for exact checks such as “is this a formula?”, “raw formula”, “show formula”, or “formula in I165”, call `intent.action: "read_formulas"` with the exact target. Use `read_formula_patterns`, `validate_formula_against_template`, or formula dependency/trace actions for repeated layouts and reference comparisons. Never infer formula existence from displayed values or numbers alone. Formula writes, formula repairs, and broad formula-like derivations are preview/apply workflows with validation.
- Formula repairs over a full column/table should target the full data range, not sample rows. When one normal same-sheet A1 formula pattern should fill a full target range, send `intent.action: "write_formulas"` with `values.formula` as a single formula string and the full `target.range`; Open Workbook expands relative references such as `=H2-G2` to `=H3-G3` and beyond. Do not build a matching-sized formula array unless Open Workbook explicitly asks for `values.formulas`, and do not add dummy `values.values` to make formula writes route.
- Reconciliation conventions: for transaction sheets, inspect the reference month convention before changing Payment Variance, Reconciliation Note, or Detail Notes. Match columns by header/role, use `read_formulas` or `read_formula_patterns` for the variance convention, use `find_similar_rows` for note wording, or call `intent.action: "settle_reconciliation"` to preview one grouped formula/note repair that preserves Reconciliation Note and Detail Notes as separate columns.
- Dropdowns and allowed values: call `intent.action: "read_data_validation"` on the selected/current column or exact target before guessing from visible values. When it returns `answer.kind: "data_validation_summary"`, answer from the inline validation metadata/options and stop; do not fetch `fullResultUri`, chunk-read sheets, list MCP resources, or read raw rows unless the user explicitly asks for raw audit metadata. If the user asks to read values from a source-list sheet such as `Dropdown Lists`, read the actual cell values with `read_values` or a targeted range read; do not treat the sheet name itself as validation intent. When dropdown options are wrong, read validation/source-list proof first, then update exact source-list cells with `mode: "auto"` if it is a bounded value correction. If the dropdown has only an inline comma-list source, use one `preview_update` with `intent.action: "write_data_validation"` and the existing options plus the new option; apply that preview once after confirmation. Use `write_data_validation` only when the user wants to change the dropdown rule itself.
- Filters, borders, dropdowns, and conditional formatting: send canonical intent early. Use `intent.action: "filter_range"` for autofilters, `clear_table_filters` for table filter removal, `format_range` for borders/fills/fonts/alignment, `write_data_validation` for dropdown/list cells, and `write_conditional_formatting` for formula-based formatting. Do not retry as sheet creation or value writes when the target sheet already exists.
- Column changes: use `insert_columns`/`delete_columns` for sheet structure and `reorder_range_columns` for plain range swaps/reorders; use `reorder_table_columns` only for real Excel tables.
- Date text cleanup: diagnose first, then use `parse_dates` only on exact date columns/ranges. If conversion needs explicit known dates, use one grouped `write_values` preview with per-range `numberFormat`.

## Safety Rules

- Never bypass permissions, scoped locks, snapshots, backups, fingerprints, Office.js execution, validation, transaction records, or rollback metadata for mutations.
- For small explicit value edits the user already asked you to make, use `mode: "auto"`, `intent.action: "write_values"`, an explicit `target`, and structured `values`; leave `autoApply` unset. Once workbook write access is allowed for the session, do not ask the user to confirm every small exact edit. If `auto` returns `taskOutcome: "apply_complete"` or `maxRecommendedFollowupCalls: 0`, report the proof and stop. Use `autoApply: false` or `preview_update` when the user asks to review first, the edit is broad/risky/ambiguous, or the backend says preview is required.
- Never write cell-by-cell loops. Batch values, formulas, number formats, and styles as narrow 2D matrices or grouped patches.
- For broad column transforms, row-aware derivations, or batch sheet renames, use `transform_values`, `derive_values`, or `transform_sheets`; do not fetch full source/target columns or issue sheet-by-sheet calls when the backend can compile one plan.
- For formula-related tasks, preserve formula proof in the response. Do not ask for `fullResultUri` for ordinary exact formula checks when `read_formulas` returns formula/status proof inline.
- Do not fetch the full sheet before every preview. Reuse `workbookContextId`; preview may perform narrow metadata/fingerprint checks for safety, but agents should only read values first when target resolution is ambiguous, the user refers to current selection, or the change depends on existing values/styles.
- For “what is this workbook”, “look at this workbook/sheet”, “where is the invoice/customer/receipt/template area”, or similar context requests, use `workbook_summary`, `semantic_index`, or `sheet_summary`. Do not fetch `table_sample` or `full_table` unless the user asks for actual rows/values.
- For border-only, fill-only, font-only, alignment-only, or number-format-only clearing, use `clear_style_dimensions` with `values.dimensions` such as `["borders"]`. Use `clear_formats` only when the user asks to remove all formatting.
- For date-display verification, use answer-mode reads or `format_diagnostics`; wording that includes “format” is not by itself a mutation request.
- For OCR output, screenshots, forms, invoices, shipment documents, booking images, or other field/value data that must become a styled horizontal table, use one `preview_update` with `intent.action: "replace_range_with_styled_table"` and one `apply_update`; do not split clear/write/autofit/style copy into separate calls. If old borders/fills must disappear, do not clear leftovers by writing blanks.
- If a response returns `status: "NEEDS_WORKFLOW_REDIRECT"` during a real mutation, stop the fragmented plan and call the suggested grouped `preview_update` workflow instead. If you were only reading existing styles, retry as `mode: "answer"` with `intent.action: "read_style_summary"` and do not apply the suggested style-write workflow.
- Never pad broad ranges with blanks or `null` when only a smaller rectangle should change.
- Treat `CAPABILITY_UNAVAILABLE`, unsupported, partial, disconnected, and Office.js host-limit warnings as real results.
- After risky mutations, validate affected scopes and report important proof: transaction IDs, backup IDs, warnings, diffs, rollback options, compact proof, and telemetry.
- Do not invent agent identities. Multiple MCP sessions get trusted identity from the daemon/MCP adapter.

## Reference Router

- Read `references/agent-run.md` for `excel.agent.run` modes, structured intent, multilingual inputs, and proof/result handling.
- Read `references/capability-map.md` for internal backend action groups and which mode usually routes to them.
- Read `references/tool-selection.md` for choosing the efficient public workflow without exposing primitive calls.
- Read `references/workflows.md` for common inspect, read/analyze, write, template, formula, table, pivot/chart, backup, and export recipes.
- Read `references/reliability.md` for validation, rollback, stale previews, conflicts, locks, disconnected add-ins, and capability limits.
- Read `references/performance.md` before large reads/writes, latency-sensitive tasks, or token-heavy workbook context.
- Read `references/multi-agent.md` when multiple agents, sessions, sheets, or tasks may touch the same workbook.
