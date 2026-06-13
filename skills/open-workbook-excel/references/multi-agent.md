# Multi-Agent Work

Open Workbook coordinates multiple agents through the shared `owb daemon`. Multiple agents may read and plan in parallel, but workbook mutations commit through a serialized transaction path.

## Start

1. Call `excel.collab.get_status`.
2. Create or claim a task with `excel.task.create` or `excel.task.claim` for substantial work.
3. Call `excel.task.evaluate_schedule` before starting scoped workbook changes.
4. Acquire a scoped lock with `excel.lock.acquire` for long planning work or guarded user workflows.

Use specific scopes: workbook, sheet, range, table, named range, chart, pivot, or template.

## Progress

- Use `excel.task.set_progress` for percentage and current step.
- Use `excel.task.add_blocker` when waiting on locks, dependencies, user decisions, or manual review.
- Use `excel.task.resolve_blocker` when a wait clears.
- Use `excel.task.complete`, `excel.task.fail`, or `excel.task.cancel` for terminal task state.

Do not hide waits as generic failures.

## Locks

- Use read locks for long analyses over shared ranges.
- Use `write_values`, `write_formulas`, `write_styles`, `format_layout`, `table`, `chart`, `pivot`, `structure`, or `workbook` modes to match mutation risk.
- Renew long-lived locks with `excel.lock.renew`.
- Release manual locks with `excel.lock.release` when the plan is applied or abandoned.

If locks conflict, call `excel.conflict.get_guidance` and wait, split scope, or hand off.

## Transactions

- `excel.batch.apply` and `excel.plan.apply` are serialized by the daemon.
- Inspect work with `excel.transaction.list` and `excel.transaction.get`.
- Use `excel.transaction.preview_rollback` before reverting.
- Use rollback-chain preview/apply when later dependent transactions overlap.

Report transaction IDs for auditability.

## Conflict Telemetry

Use `excel.conflict.get_telemetry` during long sessions to identify repeated contention by range, table, task, or agent. Use the result to split work or sequence tasks.
