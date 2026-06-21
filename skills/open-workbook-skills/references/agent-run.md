# Agent Run

`excel.agent.run` is the public Open Workbook MCP contract. Normal agents should keep the conversation on this tool and let the backend compose internal workbook capabilities.

## Modes

- `status`: connection readiness, active workbook availability, file bridge status, and compact collaboration state.
- `prepare`: workbook identity plus lightweight structure metadata for follow-up calls.
- `find`: target discovery across sheets, tables, headers, named items, regions, summary blocks, formulas, semantic roles, and workbook labels.
- `answer`: live reads, schema questions, deterministic summaries, comparisons, metadata inspection, and validation-style facts that do not mutate.
- `preview_update`: target resolution, permission checks, lock checks, safety capture, and preview for writes or workbook actions.
- `apply_update`: apply a previously previewed operation with its exact `operationId` and `confirmationToken`.
- `validate`: scoped validation after risky changes or when the user asks for proof.
- `rollback`: recovery inspection or confirmed rollback/restore using returned transaction, snapshot, or backup guidance.
- `auto`: compatibility for casual prompts; use explicit modes for predictable agent workflows.

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

Use `preview_update` for non-trivial mutations, broad edits, table appends, template actions, workbook lifecycle operations, backup lifecycle changes, or anything the user may want to review.

Only call `apply_update` with the returned `operationId` and `confirmationToken`. If the backend reports stale context, target drift, ambiguity, missing permission, an active lock, or validation failure, stop and create a fresh preview or ask the user for direction.

Small explicit value edits may use `auto` when the backend can prove the target and risk are narrow. Related changes should use one grouped preview with `values.patches`, then one apply call.

## Result Handling

Treat `resourceLinks`, `proof`, `compactProof`, `continuation`, `invalidatedContextIds`, `invalidatedResourceUris`, `nextAction`, warnings, telemetry, backup IDs, transaction IDs, and rollback options as the evidence contract. Reuse `workbookContextId` or `continuation` on follow-up calls. `excel://...` handles are MCP/Open Workbook handles, not web URLs: never pass them to `webfetch`. When the user explicitly asks for full details, all rows, raw values, or an audit, call `excel.agent.run` again with the returned `resultUri` or `fullResultUri` in `request` or `continuation`.

If `nextAction` is `answer_now`, answer from the returned proof. Retrieve full detail through `excel.agent.run` only when the user asks for an audit or the proof says detail is required.

If the result is `AMBIGUOUS_TARGET`, retry with one returned `target.candidateId` and the same `workbookContextId`. Do not switch to offline parsing for a connected workbook.
