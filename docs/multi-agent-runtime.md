# Multi-Agent Runtime

Open Workbook treats a live Excel workbook as a transaction target, not as a directly concurrent write surface.

## Core Rule

Multiple agents may read, analyze, create tasks, and preview plans in parallel. Workbook mutations commit through a serialized transaction path so Excel receives one Office.js write batch at a time.

```text
parallel reads and planning
        |
        v
task / plan / lock / transaction coordination
        |
        v
single serialized Office.js writer
```

## Implemented Foundation

- Agent records with `agentId`, optional name, client type, PID, and heartbeat timestamps.
- Task registry for multi-agent workbook work with goal, role, priority, progress, current step, blockers, scope, dependencies, plans, transactions, and rollback backup references.
- Lock manager for workbook, sheet, range, table, named range, chart, pivot, and template scopes.
- Lock modes for read, values, formulas, styles, layout, table, chart, pivot, structure, and workbook actions.
- Configurable lock lease policy for manual locks and transaction locks.
- Transaction records for queued/applying/applied/failed/blocked commits.
- Serialized writer queue around `excel.batch.apply` and `excel.plan.apply`.
- Conservative plan refresh/rebase checks for stale plans.
- Collaboration events for agent, task, transaction, lock, and conflict activity.
- MCP tools for task, collaboration, and transaction inspection.
- Shared `owb daemon` startup plus MCP adapter mode.
- Durable local state for agents, tasks, locks, transactions, conflicts, conflict telemetry, collaboration events, templates, regions, permissions, plans, and backup indexes.
- Confirmed rollback-chain preview/apply for related transactions.
- Parsed formula dependency graph reads for precedent/dependent tracing.
- Dependency-aware task schedule evaluation and ready-task resume.

## Current Runtime Shape

The current implementation runs coordination primitives inside the daemon-owned `RuntimeService`. This makes all batch and plan applies transaction-recorded and serialized through one coordinator when MCP adapters attach to the daemon.

Queued transaction progress is observable. Agents can call `excel.batch.preflight` before large generated writes, `excel.batch.submit` to enqueue mutating work without waiting for Excel execution, `excel.transaction.get` or `excel.transaction.list` for queue position and progress messages, `excel.transaction.wait` for a bounded wait until a terminal status, and `excel.transaction.cancel` to stop a queued transaction before it starts applying in Excel. When the writer is already busy, `excel.batch.apply` returns queued progress instead of waiting behind earlier mutations. Applying transactions are not interrupted mid-Office.js call.

Chunked work has a parent job record. `excel.job.get`, `excel.job.wait`, and `excel.job.cancel` aggregate progress for updates split into multiple child transactions, such as large style updates or row-chunked value/formula/number-format writes. Job cancellation cancels queued child transactions; chunks already applying in Excel are allowed to finish.

Runtime commands:

```bash
owb daemon start
owb mcp --agent-name tx-cleaner
owb mcp --agent-name dashboard-builder
```

`owb mcp` attaches to the daemon when it is available. `owb mcp --standalone` remains available for one-agent local testing.

Daemon state is persisted by default under:

```text
.open-workbook/state/collaboration-state.json
```

Set `OPEN_WORKBOOK_STATE_DIR` to move it. Set `OPEN_WORKBOOK_ADDIN_RPC_TIMEOUT_MS` to override the default 30000 ms backend-to-add-in request timeout. Set `OPEN_WORKBOOK_BATCH_DIRECT_OPERATION_THRESHOLD`, `OPEN_WORKBOOK_BATCH_DIRECT_PAYLOAD_BYTES`, and `OPEN_WORKBOOK_BATCH_DIRECT_CELL_THRESHOLD` to tune when preflight recommends queued work. Set `OPEN_WORKBOOK_STYLE_BATCH_CHUNK_SIZE` and `OPEN_WORKBOOK_MATRIX_CHUNK_ROWS` to control safe chunk sizes. On restart, active agents are marked disconnected, active locks are expired, and queued/applying transactions are marked blocked so agents do not mistake interrupted work for committed workbook changes. Job records, template registrations, region registrations, permission policy, plan records, and backup indexes are restored from the same state file.

## Conflict Model

The lock manager allows non-overlapping range work and blocks dangerous overlap:

- `Transactions!A1:F20` and `Transactions!H1:N20` can both be planned safely.
- `Transactions!A1:F20` and `Transactions!D5:H10` conflict.
- Sheet/workbook structure locks block child range writes.
- Expired or released locks stop blocking future work.

Pre-commit fingerprint checks remain active, so manual user edits after preview still force agents to refresh before applying.

Conflict classification is dependency-aware:

- `STRUCTURE_CONFLICT`: workbook/sheet structure scopes conflict with child work.
- `TABLE_CONFLICT`: table resize, structure, filter, sort, or table data changes conflict with dependent ranges.
- `FORMULA_DEPENDENCY_CONFLICT`: formula writes or formula-pattern repairs conflict with touched source/target ranges.
- `DERIVED_OBJECT_CONFLICT`: chart and pivot source/object changes conflict with dependent ranges.
- `NAMED_RANGE_CONFLICT`: named range changes conflict with ranges or formulas that may depend on those names.

Direct table, formula, chart, pivot, and named-range mutations now create transaction records and acquire typed locks, not just backups. This keeps object-level changes visible in `excel.transaction.*` and `excel.collab.*`.

## Public Tools

Tasks:

- `excel.task.create`
- `excel.task.claim`
- `excel.task.update`
- `excel.task.set_progress`
- `excel.task.add_blocker`
- `excel.task.resolve_blocker`
- `excel.task.evaluate_schedule`
- `excel.task.resume_ready`
- `excel.task.complete`
- `excel.task.fail`
- `excel.task.cancel`
- `excel.task.list`
- `excel.task.get`

Collaboration:

- `excel.collab.get_status`
- `excel.collab.list_agents`
- `excel.collab.list_tasks`
- `excel.collab.list_locks`
- `excel.collab.list_transactions`
- `excel.collab.get_conflicts`
- `excel.collab.get_recent_events`

Locks:

- `excel.lock.get_policy`
- `excel.lock.set_policy`
- `excel.lock.acquire`
- `excel.lock.renew`
- `excel.lock.release`

Transactions:

- `excel.transaction.list`
- `excel.transaction.get`
- `excel.transaction.preview_rollback`
- `excel.transaction.rollback`
- `excel.transaction.preview_rollback_chain`
- `excel.transaction.rollback_chain`

Plan safety:

- `excel.plan.refresh_preview`
- `excel.plan.rebase`

## Safe Plan Refresh

Plans store target range fingerprints during preview. Before an old plan is applied, agents can call `excel.plan.refresh_preview` or `excel.plan.rebase`.

- If the planned target ranges are unchanged, Open Workbook refreshes the preview fingerprints and the plan can continue to the normal transaction apply path.
- If any target range changed after preview, refresh/rebase is blocked with `TARGET_REGION_CHANGED`.
- Structure-level changes, table resizes, sheet deletes, and other high-risk conflicts should still be handled by creating a new plan instead of trying to merge.

## Transaction-Aware Rollback

Rollback is transaction-aware. Before a transaction rollback applies, Open Workbook compares the target transaction scopes with later applied transactions in the same workbook.

- If no later applied transaction overlaps the rollback scope, rollback can proceed through the plan rollback path.
- If later work overlaps, rollback is blocked with `ROLLBACK_CONFLICT`.
- If the transaction has no plan rollback metadata, rollback is blocked with `ROLLBACK_UNAVAILABLE` and the user should use backup repair or manual recovery.
- Rollback previews are exposed through `excel.transaction.preview_rollback` so agents can explain risk before changing Excel.
- When later related transactions must also be reverted, `excel.transaction.preview_rollback_chain` returns a newest-first rollback order and a confirmation token.
- `excel.transaction.rollback_chain` applies that order only after the exact confirmation token is supplied, and stops if any step fails.

## Task Progress And Blockers

Tasks expose `progress`, `currentStep`, and `blockers` so agent status can be visible in a Cursor/OpenCode-like workflow.

- Progress is a clamped `0..100` number with optional current-step text.
- Blockers can be `info`, `warning`, or `blocked`, and can optionally point to a workbook scope.
- A `blocked` blocker moves the task to `blocked` status without losing plan, transaction, or rollback history.

## Task Scheduling

`excel.task.evaluate_schedule` checks each task against explicit task dependencies, open blocking blockers, and active lock conflicts for the task's allowed scopes.

- Ready tasks return `state: "ready"` with `suggestedAction` of `start` or `resume`.
- Tasks waiting on incomplete dependencies return `state: "waiting_dependencies"`.
- Tasks blocked by active locks return `state: "waiting_locks"`.
- Lock-wait decisions include `nextRetryAt` when the active conflicting lock expiry is known.
- `excel.task.resume_ready` applies the scheduler result and moves blocked tasks whose waits cleared back to `open` or `claimed`.

## Lock Lease Policy

Manual and transaction locks use a runtime lease policy:

- `defaultTtlMs`: default TTL for manual locks.
- `transactionTtlMs`: TTL for transaction writer locks.
- `maxTtlMs`: upper bound applied to manual and transaction lock requests.
- `allowManualLocks`: disables `excel.lock.acquire` when false while preserving internal transaction locks.

The policy is persisted with daemon collaboration state and can also be initialized with `OPEN_WORKBOOK_LOCK_DEFAULT_TTL_MS`, `OPEN_WORKBOOK_LOCK_TRANSACTION_TTL_MS`, `OPEN_WORKBOOK_LOCK_MAX_TTL_MS`, and `OPEN_WORKBOOK_ALLOW_MANUAL_LOCKS=0`.

## Formula Dependency Safety

Formula graph parsing produces dependency nodes for local ranges, structured table references, and external workbook references. Structured table references resolve to precise header, data, totals, all, and column-span ranges when table metadata is available. Dynamic array spill references resolve to the spill range when metadata is available and otherwise fall back to the anchor cell with a warning.

Formula write transactions add parsed local range and table dependencies to their lock scopes, so source edits and dependent formula writes cannot commit concurrently through separate agents.

External workbook references remain graph-only because they cannot be locked inside the local workbook runtime.

## Conflict Guidance

Conflict records include structured guidance for agents and UI surfaces.

- Lock conflicts recommend waiting or retrying after `lockExpiresAt`, coordinating with the lock owner, and splitting scope when possible.
- Table, formula, chart, pivot, and named-range conflicts recommend task handoff and scope splitting before retry.
- Structure conflicts stay blocked behind manual review.
- Rollback conflicts point agents toward rollback-chain preview or backup/template repair.

Use `excel.conflict.get_guidance` for recent runtime conflicts, or `excel.conflict.explain` to explain a supplied conflict record.

## Conflict Telemetry

Conflict telemetry records each persisted runtime conflict and tracks whether lock waits clear when the blocking lock is released.

- `excel.conflict.get_telemetry` summarizes conflict counts, open/cleared counts, hot scopes, hot tasks, hot agents, codes, and primary actions.
- `excel.conflict.clear_telemetry` clears telemetry for one workbook or the whole runtime.
- Telemetry is persisted with daemon state and capped to the most recent runtime records.

## Remaining Production Hardening

The remaining production work is host verification around the implemented collaboration model:

1. Keep structure-level merge conflicts blocked by default.
2. Add real Excel sideload verification before declaring host-level production readiness.

See [Production Readiness](production-readiness.md) for the invariant checklist and release gates.
