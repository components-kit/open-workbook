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

Use `detailLevel: "semantic_index"` when you need a compact role-aware workbook map before choosing a sheet, table, template, form region, formula region, or style target. Use `workbook_summary` and `sheet_summary` for overview questions. Do not request table samples or full tables for broad context questions.

## Multilingual Requests

For non-English or mixed-language prompts:

1. Preserve the original user request.
2. Normalize `intent.action` to canonical English when clear.
3. Put workbook labels and user-language aliases in `intent.targetHints`.
4. Pass exact `target` and `values` if available.
5. Reply in the user's language unless instructed otherwise.

Do not rely on deterministic English keyword parsing for every language. The caller LLM should provide structured intent when language may affect routing.

## Preview And Apply

Use `auto` for small explicit value edits that the user already asked you to make. Leave `autoApply` unset unless the user asks for preview-only behavior; if the user says "ask before editing", "show preview first", or similar, set `autoApply: false`. Once workbook write access is allowed for the session, do not ask the user to confirm every small exact edit.

Use `preview_update` for non-trivial mutations, broad edits, table appends, template actions, workbook lifecycle operations, backup lifecycle changes, or anything the user may want to review.

Only call `apply_update` with the returned `operationId` and `confirmationToken`. If the backend reports stale context, target drift, ambiguity, missing permission, an active lock, or validation failure, stop and create a fresh preview or ask the user for direction.

If `auto` returns `taskOutcome: "apply_complete"`, stop and report the applied change. If `auto` returns `taskOutcome: "preview_ready"` or `nextAction: "call_apply_update"`, ask the user once unless the user's configuration/instruction already permits applying previews, then call one `apply_update` with the returned `operationId` and `confirmationToken`.

Small explicit value edits may use `auto` when the backend can prove the target and risk are narrow. Related changes should use one grouped preview with `values.patches`, then one apply call.

## Result Handling

Treat `resourceLinks`, `proof`, `compactProof`, `continuation`, `invalidatedContextIds`, `invalidatedResourceUris`, `nextAction`, warnings, telemetry, backup IDs, transaction IDs, and rollback options as the evidence contract. Reuse `workbookContextId` or `continuation` on follow-up calls. `excel://...` handles are MCP/Open Workbook handles, not web URLs: never pass them to Webfetch, browser fetch, curl, or HTTP tooling. Use inline `values`, `rows`, or `sparseRows` when returned. If the next task needs exact rows, raw values, transformation input, validation evidence, full details, or an audit and only a preview is inline, call `excel.agent.run` once with `continuation.fullResultUri` or paste the returned `resultUri`/`fullResultUri` in `request`.

If `taskOutcome` is `final_answer` and `maxRecommendedFollowupCalls` is `0`, stop calling tools and answer the user from `finalAnswer`, `summary`, proof, and inline data. If `taskOutcome` is `preview_ready`, only continue with the required apply call after user approval.

Compact answers may use `encodedValues` with `valueEncoding.kind: "domain_dictionary_by_column"` for repeated readable values such as statuses, labels, vendors, dates, and categories. Decode integer codes with the per-column domain dictionary before reasoning. Wide rows may include `inlineColumnProjection`; use those high-signal columns first and retrieve `fullResultUri` only if omitted columns are necessary.

When `continuation.freshness` is present, carry it forward with the continuation. It is proof for cached workbook context/result reuse: unchanged workbook content version, structure hash, or no overlapping change journal entries means you can avoid rediscovery and reread calls. If freshness is stale or overlapping changes are reported, refresh only the affected target.

If `nextAction` is `answer_now`, answer from the returned proof and inline data. Retrieve full detail through `excel.agent.run` when the task cannot be completed correctly from the proof and inline data.

If the result is `AMBIGUOUS_TARGET`, retry with one returned `target.candidateId` and the same `workbookContextId`. Do not switch to offline parsing for a connected workbook.
