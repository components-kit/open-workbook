# Multi-Agent Work

Open Workbook coordinates multiple agents through the shared `owb daemon`. Multiple agents may read and plan in parallel, but workbook mutations commit through a serialized transaction path.

## Start

1. Start with `excel.agent.run` `mode: "status"` or `mode: "prepare"` to establish workbook context and read the returned `collaboration` summary.
2. Describe substantial work and guarded scopes clearly in the request.
3. Use `auto` for small exact value edits in known non-overlapping ranges after session write access is allowed; use `preview_update` before broad, risky, table, formula, style, template, structural, or user-reviewable workbook changes.
4. Respect returned lock, task, transaction, or conflict guidance.

Use specific scopes: workbook, sheet, range, table, named range, chart, pivot, or template.

## Progress

- Report returned task, transaction, and progress metadata to the user.
- Surface blockers when waiting on locks, dependencies, user decisions, or manual review.
- Do not hide waits as generic failures.

Backend task capabilities are internal unless the public surface returns task metadata. Do not call or assume separate task, lock, transaction, or collaboration tools on the normal MCP surface.

## Locks

- Ask for preview before long analyses or edits over shared ranges.
- Use `write_values`, `write_formulas`, `write_styles`, `format_layout`, `table`, `chart`, `pivot`, `structure`, or `workbook` modes to match mutation risk.
- Let the backend handle lock acquisition, renewal, release, and conflict reporting.

If locks conflict, follow returned guidance: wait, split scope, or hand off.

Multiple sheets in one workbook can proceed in parallel when scopes do not overlap. Same-sheet work can proceed when ranges or objects do not conflict; otherwise the backend queues, blocks, or returns conflict guidance.

## Transactions

- Mutations are serialized by the daemon.
- Inspect returned transaction, job, and rollback metadata.
- Use `excel.agent.run` `mode: "rollback"` or returned rollback guidance before reverting.
- Use rollback-chain review when later dependent transactions overlap.

Report transaction IDs for auditability.

## Conflict Telemetry

Use returned conflict telemetry during long sessions to identify repeated contention by range, table, task, or agent. Use the result to split work or sequence tasks.
