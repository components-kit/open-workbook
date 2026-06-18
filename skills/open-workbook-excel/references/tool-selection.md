# Tool Selection

Use the public Open Workbook surface:

```text
excel.agent.run
```

Do not call backend primitive capabilities directly in normal MCP clients. Lookup, compact reads, summaries, batches, plans, validation, snapshots, backups, transactions, jobs, locks, pivots, charts, and cleaning capabilities are internal backend concepts. Mention them only when explaining what the backend did or why a host capability is limited.

## Agent Run Modes

- `prepare`: cache workbook identity and lightweight structure for sheets, tables, named ranges, regions, and sheet kind.
- `find`: search cached metadata for candidate sheets, tables, headers, named ranges, regions, summary blocks, and formula regions.
- `answer`: perform targeted live reads and deterministic summaries; cached metadata is enough for schema-only answers.
- `auto`: compatibility route for casual prompts; explicit modes are preferred when the workflow step is known.
- `preview_update`: resolve targets, preview safe writes/table appends, block unsafe edits, and return an operation plus confirmation token.
- `apply_update`: apply only a previewed update with the returned confirmation token.
- `rollback`: map transaction or backup identifiers to recovery.
- `validate`: validate after risky changes.

If `excel.agent.run` returns `AMBIGUOUS_TARGET`, select one candidate and retry with `target.candidateId` plus the same `workbookContextId`. Do not switch to shell, Python, openpyxl, pandas, or offline `.xlsx` parsing for a connected live workbook.

## Structured Intent

When the user's LLM/client can infer intent, pass structured fields instead of relying only on English keyword matching:

- `intent.action`: one of `read_values`, `read_schema`, `find_target`, `write_values`, `write_formulas`, `format_range`, `clear_values`, `append_table_rows`, `sort_table`, `filter_range`, `autofit`, `copy_template_sheet`, `calculate`, or `save`.
- `intent.targetHints`: short target clues from the workbook, including original labels and useful translated aliases.
- `target`: explicit `sheetName`, `range`, `tableName`, or returned `candidateId` when known.
- `values`: 2D matrices, `values.patches`, formulas, styles, or table rows for updates.

Structured intent is only a routing hint. The backend still performs target resolution, ambiguity checks, preview/apply confirmation, stale checks, permissions, locks, backups, validation, and rollback bookkeeping.

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
- Use `mode: "answer"` with `target.candidateId`, explicit `target.sheetName`/`target.range`, or a clear natural-language target.
- Ask for rows, samples, actual values, formulas, raw monthly sections, or explicit A1 ranges when live cell data is required.
- Ask for schema, columns, or headers when cached metadata is enough.

Avoid prompts that require workbook-wide scans unless the user asked for audit, validation, search, or repair.

## Writing Data

- Use `mode: "preview_update"` for user-reviewable edits.
- Use `mode: "apply_update"` once with the returned `operationId` and `confirmationToken`.
- For one small explicit value edit, `mode: "auto"` may apply safely after backend checks.
- For related range edits, send one `values.patches` preview and one apply call.
- For formulas, use `intent.action: "write_formulas"` and formula matrices beginning with `=`.
- For styles, use `intent.action: "format_range"` with a narrow `target`.
- For clearing values, use `intent.action: "clear_values"` and state the intended scope.

Never pad broad ranges with blanks or nulls when only a smaller rectangle should change. Never split related patches into many separate calls unless apply returns a hard failure with actionable issue details.

## Tables, Templates, Formulas, Pivots, Charts

Describe the user's goal through `excel.agent.run` and let the backend choose internal table, template, formula, pivot, chart, validation, snapshot, diff, backup, and transaction capabilities.

- For table appends, pass `target.candidateId` or `target.tableName` with `values.rows`, preview first, then apply.
- For table sort/filter, pass `intent.action: "sort_table"` or `filter_range` plus a table/range target.
- For template, formula repair, pivot/chart, and risky-edit tasks, use `prepare`, `find`, `answer`, `preview_update`, and `validate` as appropriate; report host capability warnings honestly.
- For save/recalculate, use `intent.action: "save"` or `calculate`; the backend maps these to workbook-level operations.

## Result Handling

Treat returned `resourceLinks`, `compactProof`, `proof`, `invalidatedContextIds`, `nextAction`, warnings, backup IDs, transaction IDs, and telemetry as evidence. If `nextAction` says `answer_now`, answer from the proof instead of fetching more detail unless the user requested an audit.
