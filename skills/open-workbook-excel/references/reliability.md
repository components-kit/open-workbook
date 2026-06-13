# Reliability

Open Workbook treats Excel as a transaction target. Preserve that contract.

## Mutation Contract

Every mutation should go through:

1. Target resolution.
2. Permission and locked-region checks.
3. Snapshot or backup capture.
4. Fingerprint checks.
5. Office.js execution through the backend.
6. Validation where relevant.
7. Transaction/audit output.
8. Rollback guidance.

If a tool returns backup IDs, transaction IDs, warnings, diffs, or rollback IDs, include the important ones in the final user-facing result.

## Validation Defaults

- Before risky changes: `excel.validate.workbook`, `excel.validate.sheet`, or a scoped validator.
- After formula edits: `excel.validate.no_formula_errors` and `excel.validate.no_broken_references`.
- After template work: `excel.template.validate_sheet_against_template`.
- After table/filter work: `excel.validate.tables` and `excel.validate.filters`.
- After broad changes: compare snapshots or use `excel.validate.no_unintended_changes`.

Validation reports have `ok`, issue counts, issue severities, stable codes, and targets. Surface the actual codes and affected ranges when they matter.

## Stale Plans

If a plan was previewed earlier, call `excel.plan.refresh_preview` or `excel.plan.rebase` before applying. If the tool returns `TARGET_REGION_CHANGED`, stop and create a new plan. Do not merge stale Excel edits by guessing.

## Conflicts And Locks

For blocked transactions or lock waits:

1. Call `excel.conflict.get_guidance`.
2. If a task is involved, update progress with `excel.task.set_progress`.
3. Add a blocker with `excel.task.add_blocker` when waiting for user input, another task, or a lock.
4. Retry only after the suggested wait or after splitting the scope.

Structure conflicts and broad workbook conflicts usually require manual review.

## Rollback

- Preview rollback with `excel.transaction.preview_rollback`.
- If later related transactions overlap, use `excel.transaction.preview_rollback_chain`.
- Apply rollback chains newest-first only with the returned confirmation token.
- If rollback metadata is unavailable, switch to backup/template repair and explain the limitation.

Never roll back a transaction without checking for later overlapping workbook work.

## Capability Limits

Treat these as real outcomes:

- `CAPABILITY_UNAVAILABLE`
- `unsupported`
- `partial`
- host-limited Office.js warnings
- disconnected add-in responses
- native file bridge unavailable responses

Choose a supported tool path or ask for user/host setup. Do not claim that unsupported pivot, chart, comment, note, save-as, or print-layout dimensions were changed.

## Disconnected Add-In

If Excel is not connected:

1. Ask the user to start the agent UI that has the Open Workbook MCP config.
2. If troubleshooting manually, ask the user to run `npx -y @component-kit/open-workbook@latest mcp`.
3. Ask the user to open Excel and load the Open Workbook add-in.
4. Retry `excel.runtime.get_status`.

Do not use stale snapshots as the source of truth for current workbook state.
