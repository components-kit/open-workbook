# Agent Run

`excel.agent.run` is the public Open Workbook MCP contract. Normal agents should keep the conversation on this tool and let the backend compose internal workbook capabilities.

## Modes

- `status`: connection readiness, active workbook availability, file bridge status, and compact collaboration state.
- `prepare`: workbook identity plus lightweight structure metadata for follow-up calls.
- `find`: target discovery across sheets, tables, headers, named items, regions, summary blocks, formulas, semantic roles, and workbook labels.
- `answer`: live reads, schema questions, deterministic summaries, comparisons, metadata inspection, and validation-style facts that do not mutate.
- `preview_update`: target resolution, permission checks, lock checks, safety capture, and preview for writes or workbook actions when review is required.
- `apply_update`: apply a previously previewed operation with its exact `operationId` and `confirmationToken`.
- `validate`: scoped validation after risky changes or when the user asks for proof.
- `rollback`: recovery inspection or confirmed rollback/restore using returned transaction, snapshot, or backup guidance.
- `auto`: default for casual prompts and small explicit value edits; safe narrow edits may preview and apply in one call after workbook write access is allowed for the session unless `autoApply: false`.

## Intent Fields

Use structured fields when the task is clear:

- `request`: keep the user's original wording.
- `intent.action`: canonical English routing action.
- `intent.targetHints`: short workbook/user-language labels plus translated aliases when useful.
- `target`: explicit sheet, range, table, named item, region, workbook, or returned `candidateId`.
- `values`: matrices, rows, formulas, styles, patches, template data, IDs, or other action payload.
- `workbookContextId`: reuse the value from `prepare`, `find`, or a prior result when returned.

Structured intent is a hint, not a bypass. The backend still resolves ambiguity, checks permissions, blocks unsafe edits, records snapshots/backups, runs Office.js, validates, and returns rollback metadata.

Use `context` to request the acquisition policy and `detailLevel` to request the returned preset shape. `context.strategy` is why/how to gather context, `context.scope` is where to start, and `context.include` is the facet list. Use `detailLevel: "semantic_index"` when you need a compact role-aware workbook map before choosing a sheet, table, template, form region, formula region, or style target. Use `workbook_summary` and `sheet_summary` for overview questions. Do not request table samples or full tables for broad context questions.

For overview questions such as "what is this workbook/file?", "look into this Excel file", or "summarize this workbook", one `mode: "answer"` call with `detailLevel: "workbook_summary"` or `detailLevel: "sheet_summary"` is normally the whole workflow. If the response has `nextAction: "answer_now"` or `maxRecommendedFollowupCalls: 0`, answer immediately. Do not follow `resourceLinks`, fetch `fullResultUri`, chunk-read sheets, list MCP resources, or call low-level resource reads unless the user explicitly asks for all raw rows, every value, or exact cell contents.

For exact formula questions, pass `intent.action: "read_formulas"` with the exact target when known. This covers "is this a formula?", "raw formula", "show formula", and "formula in I165". Do not infer formula existence from displayed values or numbers alone; use returned formula/status proof.

For read-only lookup/filter-like questions such as "show rows where Status = Unpaid", "which invoices are overdue?", or "list transactions in June", pass `intent.action: "query_rows"` with `values.where`, optional `values.return`, `values.limit`, and `values.format`. Use `filter_range` only when the user explicitly wants the visible Excel filter/view changed.

## Multilingual Requests

For non-English or mixed-language prompts:

1. Preserve the original user request.
2. Normalize `intent.action` to canonical English when clear.
3. Put workbook labels and user-language aliases in `intent.targetHints`.
4. Pass exact `target` and `values` if available.
5. Reply in the user's language unless instructed otherwise.

Do not rely on deterministic English keyword parsing for every language. The caller LLM should provide structured intent when language may affect routing.

## Preview And Apply

Use `auto` for small explicit value edits that the user already asked you to make. Include `intent.action: "write_values"` and `values.patches`; even a single-cell edit is one patch with its own `target`. Leave `autoApply` unset unless the user asks for preview-only behavior. If the user says "ask before editing", "show preview first", or similar, set `autoApply: false`. Once workbook write access is allowed for the session, do not ask the user to confirm every small exact edit. If `auto` returns `taskOutcome: "apply_complete"` or `maxRecommendedFollowupCalls: 0`, report the returned proof and stop.

Use `preview_update` for non-trivial mutations, broad edits, table appends, template actions, workbook lifecycle operations, backup lifecycle changes, or anything the user may want to review.

Formula writes, formula repairs, and broad formula-like derivations are preview/apply workflows with validation. Use `derive_values` with `formula_like` for row-aware calculations such as Payment Variance = Actual Amount - Cash Amount so the backend scans source/target columns and returns bounded source/before/after examples. When a full range should receive the same relative A1 formula pattern, use `write_formulas` with the full target range and one `values.formula`, for example `=H2-G2` on `I2:I244`; do not create a row-per-formula array unless the backend asks for it.

For styling review or "what would make this easier to read?" requests, use `mode: "answer"` with `intent.action: "style_overview"` or `detailLevel: "style_overview"` first. It returns current style context, column-role groups, grouped-header suggestions, and safe next workflow hints without full data-row reads.

For workbook design review requests such as "for each column decide free text/date/money/ID/dropdown/lookup" or "look at other sheets and recommend how this table should behave", use one `mode: "answer"` call with `intent.action: "workbook_design_overview"`. It returns the target shape, column-by-column behavior recommendations, date/money/text formats, dropdown candidates, lookup/reference candidates, related-sheet hints, and next workflows from cached metadata. Do not read Customer, Bookings, Drivers, or empty table rows manually before this overview. After the user chooses a recommendation, use the targeted workflow it names, such as `write_data_validation`, `derive_values`, or `improve_visual_readability`.

For broad readability/styling requests such as "make this cleaner", "make this easier to read", "office-ready", "highlight important issues", or "format this table", use `mode: "preview_update"` with `intent.action: "improve_visual_readability"`. Put options under `values.visualReadability`; use `styleDepth: "standard"` by default, `basic` for fast low-risk cleanup, and `comprehensive` only when the user wants deeper suggestions. The backend compiles column-first layout, width, alignment, number-format, grouping, formula/error highlight, and optional validation/formula suggestions. Style preservation defaults to `stylePreservationMode: "protected_regions"` so summary/template areas and grouped header bands stay guarded while ordinary table body styling, widths, alignment, and date/money formats can still be improved; use `"strict"` only when the user asks to preserve every existing style, and `"none"` for explicit redesign. Use `applySuggestionBuckets` to opt into actionable suggestions: `["layout"]` for wrap/row-height style writes, `["validation"]` for dropdown validation writes, and `["freeze_panes"]` for freeze rows/columns. Grouped headers are suggested for wide tables by default; to apply inserted group rows, merged labels, and matching header colors, use `mode: "preview_update"` with `intent.action: "grouped_header"` and optional `values.groupedHeader.groups`. Prefer groups shaped as `{ "label": "...", "startColumn": "A", "endColumn": "B" }`; `{ "columns": ["A", "B"] }` and `{ "range": "A:B" }` are accepted. For freeze column requests, include `freezePanes: { "columns": 1 }` or use a clear request like "freeze first column". Use `referenceStyle` for "make this look like that sheet" previews and `presentationMode` for print/export suggestions. Apply only when `nextAction` is `call_apply_update`, `operationCount > 0`, and the returned `operationId`/`confirmationToken` match the preview. If `operationCount` is `0` or `nextAction` is `answer_now`, do not call `apply_update` and do not decompose the styling into primitive `format_range` calls; explain the skipped reasons and ask for the supported next workflow. Never continue a new grouped-header preview with an `operationId` from an older visual-readability preview. Formula helpers, structure changes, reference-style layout cues, and print settings remain separate confirmed workflows or preview-only until their host capability is available.

Grouped headers are structural. If apply returns `DESTRUCTIVE_ACTION_BLOCKED` or `PERMISSION_DENIED`, ask for user approval to allow structure changes, then call:

```json
{
  "intent": { "action": "set_permissions" },
  "values": {
    "permissions": {
      "allowWrites": true,
      "allowDestructiveActions": true,
      "scopeToWorkbook": true,
      "requireConfirmationFor": []
    }
  }
}
```

After the permission update succeeds, create a fresh grouped_header preview and apply that fresh operation; do not retry a stale failed preview.

Visual styling safety: comprehensive validation/formula suggestions remain preview-only unless the user chooses an explicit bucket or separate workflow; do not apply dropdowns, formulas, inserted rows/columns, or summary blocks through the visual styling apply path.

OpenCode prompt examples:

```text
Use open-workbook. Inspect the active sheet with a style overview first, without reading every data cell. Suggest visual readability improvements including grouped headers, one consistent palette, safe widths, alignment, filters, number formats, and highlights. Do not apply yet.
```

```text
Preview a grouped_header workflow for this sheet. Add a higher-level grouped header row above the existing column headers, merge group labels, and use matching group colors. Wait for approval before apply_update.
```

```json
{
  "mode": "preview_update",
  "intent": { "action": "grouped_header" },
  "target": { "sheetName": "Invoices", "tableName": "InvoicesTable" },
  "values": {
    "stylePreservationMode": "none",
    "groupedHeader": {
      "groups": [
        { "label": "Status", "startColumn": "A", "endColumn": "B" },
        { "label": "Job Details", "startColumn": "C", "endColumn": "E" }
      ]
    }
  }
}
```

```text
Apply the safe visual readability preview in one apply_update. Include opt-in buckets layout, validation, and freeze_panes only if they were present in the preview.
```

Only call `apply_update` with the returned `operationId` and `confirmationToken`. If the backend reports stale context, target drift, ambiguity, missing permission, an active lock, or validation failure, stop and create a fresh preview or ask the user for direction.

If `auto` returns `taskOutcome: "apply_complete"`, stop and report the applied change. If `auto` returns `taskOutcome: "preview_ready"` or `nextAction: "call_apply_update"`, ask the user once unless the user's configuration/instruction already permits applying previews, then call one `apply_update` with the returned `operationId` and `confirmationToken`.

Small explicit value edits may use `auto` when the backend can prove the target and risk are narrow. Multiple explicit value edits from the same user instruction should use one `auto` call with `values.patches`; different topic does not mean different tool call when the targets and values are known. Use one grouped preview/apply only when the grouped edit is broad, risky, ambiguous, formula/style/template/table-related, or user-reviewable.

For dropdown option questions, call `intent.action: "read_data_validation"` once on the selected/current column or exact target. If the answer kind is `data_validation_summary`, the validation metadata and dropdown options are complete for the requested range; answer from that result and do not fetch `fullResultUri` unless the user explicitly asks for raw audit metadata. If the user asks to read values from a source-list sheet such as `Dropdown Lists`, read the actual cell values with `read_values` or a targeted range read; do not treat the sheet name itself as validation intent. To add one missing dropdown option, prefer a bounded `mode: "auto"` write to the returned source-list range/cell. If the dropdown rule/source range itself must change, use one `preview_update` with `intent.action: "write_data_validation"` and one `apply_update`, even when the existing target validation is mixed or inconsistent. Do not retry formula string variants, fetch resources, or test-write cells; if a fresh apply fails, report the exact error and stop.

## Result Handling

Treat `resourceLinks`, `proof`, `compactProof`, `continuation`, `invalidatedContextIds`, `invalidatedResourceUris`, `nextAction`, warnings, telemetry, backup IDs, transaction IDs, and rollback options as the evidence contract. Reuse `workbookContextId` or `continuation` on follow-up calls. `excel://...` handles are MCP/Open Workbook handles, not web URLs: never pass them to Webfetch, browser fetch, curl, or HTTP tooling. Use inline `values`, `rows`, or `sparseRows` when returned. If the next task needs exact rows, raw values, transformation input, validation evidence, full details, or an audit and only a preview is inline, call `excel.agent.run` once with `continuation.fullResultUri` or paste the returned `resultUri`/`fullResultUri` in `request`.

If `taskOutcome` is `final_answer` and `maxRecommendedFollowupCalls` is `0`, stop calling tools and answer the user from `finalAnswer`, `summary`, proof, and inline data. If `taskOutcome` is `preview_ready`, only continue with the required apply call after user approval.

Compact answers may use `encodedValues` with `valueEncoding.kind: "domain_dictionary_by_column"` for repeated readable values such as statuses, labels, vendors, dates, and categories. Decode integer codes with the per-column domain dictionary before reasoning. Wide rows may include `inlineColumnProjection`; use those high-signal columns first and retrieve `fullResultUri` only if omitted columns are necessary.

When `continuation.freshness` is present, carry it forward with the continuation. It is proof for cached workbook context/result reuse: unchanged workbook content version, structure hash, or no overlapping change journal entries means you can avoid rediscovery and reread calls. If freshness is stale or overlapping changes are reported, refresh only the affected target.

If `nextAction` is `answer_now`, answer from the returned proof and inline data. Retrieve full detail through `excel.agent.run` when the task cannot be completed correctly from the proof and inline data.

If the result is `AMBIGUOUS_TARGET`, retry with one returned `target.candidateId` and the same `workbookContextId`. Do not switch to offline parsing for a connected workbook.
