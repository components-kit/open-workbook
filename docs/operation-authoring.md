# Operation Authoring

Every workbook mutation must be explainable from the operation manifest before it ships. Run:

```bash
corepack pnpm operations:manifest
corepack pnpm operations:check
```

The manifest joins the protocol catalog, agent modes/actions, backend action handlers, risk policy, batch compiler, add-in executor, and host/backend method registry. If a new capability or operation is missing from the manifest, do not patch around the report; add the missing lifecycle mapping.

## Mutation Lifecycle

Use this contract for any new operation, direct mutation, or composed workflow:

- Define the protocol capability/action and keep `excel.agent.run` as the only normal MCP tool.
- Add an agent handler or explicitly mark the capability internal-only.
- Classify risk in the agent action policy.
- Route preview through `preview_update`; route apply through `apply_update` with `operationId` and `confirmationToken`.
- Validate target ambiguity, permissions, locked regions, and stale fingerprints before Excel writes.
- Use `applyBatch` for range/sheet operations; use `applyDirectTransaction` for native Office.js object APIs.
- Create the right backup before mutation: region backup for scoped range edits, sheet/workbook-copy backup for structural or workbook-level edits.
- Return the common mutation proof: `transactionId`, `warnings`, `backups`, `rollbackAvailable`, telemetry, and validation or step results when available.
- Pass stable idempotency keys for agent-applied direct mutations and composed workflows.
- Invalidate workbook metadata/context after structure or value mutations that can change used ranges, tables, formulas, or workbook shape.
- Document any Office.js host limit as capability metadata and warning output instead of pretending the operation is complete.

## Composite Workflows

Composite workflows should compile user intent into one previewed operation whenever the steps are logically one user action. Examples include template repair, style repair, grouped patches, and booking-image table replacement.

For related range work, prefer grouped operation kinds such as `range.write_values_many`, `range.write_number_formats_many`, `range.write_styles_many`, `range.clear_many`, `range.clear_formats_many`, and `range.autofit_many` over repeated primitive operations. The grouped operation must still carry per-range targets so fingerprints, lock scopes, backups, telemetry, and rollback evidence cover every affected range.

For duplicate-from-template work, use one operation. Default clean copies should use `sheet.copy_clean_data_regions` rather than `sheet.copy` followed by separate clear operations: copy the source sheet once, clear only declared data regions on the new sheet, preserve headers/formulas/styles/layout, and return one preview/apply proof. If the caller requests `copyMode: "with_data"` or otherwise explicitly asks to preserve existing data, use `sheet.copy` and report that data is retained.

For table requests that combine filtering and sorting, prefer `table.apply_view` over separate `table.apply_filters` and `table.sort` calls. The operation must preserve one preview/apply lifecycle, one table backup, and one direct Office.js table transaction.

For a booking or OCR-extracted table fill, prefer:

```text
excel.agent.run mode=preview_update intent.action=replace_range_with_styled_table
excel.agent.run mode=apply_update operationId=<returned> confirmationToken=<returned>
```

Do not split this into clear, write, autofit, and style-copy calls unless the composed workflow returns a hard failure that names the step to isolate.

## Operational Review

Before release, inspect:

- `corepack pnpm operations:manifest -- --json` for ownership and lifecycle gaps.
- `corepack pnpm capabilities:report` for coverage planning.
- `corepack pnpm diagnose:session -- <log-file>` for real agent traces with repeated calls or missed batching.

The most important failure mode is a mutating add-in call that bypasses preview, permission checks, backups, transaction locks, idempotency, or rollback metadata.
