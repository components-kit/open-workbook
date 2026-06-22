# Reliability

Open Workbook treats Excel as a transaction target. Preserve that contract.

For live workbook tasks, Open Workbook MCP is the source of truth. Do not switch to shell scripts, Python, openpyxl, pandas, manual UI automation, or offline `.xlsx` parsing unless the user explicitly asked for offline file analysis, or MCP is unavailable and the user approves a non-live fallback. Saved-file parsing can be stale because it does not include unsaved Excel state. If Open Workbook is connected but returns an empty live-read diagnostic or `cannot_complete`, report that Open Workbook failure and stop; do not silently fall back to saved-file parsing.

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

If `excel.agent.run` returns backup IDs, transaction IDs, warnings, diffs, rollback IDs, compact proof, or resource links, include the important ones in the final user-facing result.

## Validation Defaults

- Before risky changes: use `excel.agent.run` `mode: "preview_update"` or `validate`.
- After formula edits: use `mode: "validate"` and report formula/broken-reference issues.
- After template work: validate the created sheet through `agent.run`.
- After table/filter work: validate through `agent.run`.
- After broad changes: ask for validation, diff, or rollback proof through `agent.run`.

Validation reports have `ok`, issue counts, issue severities, stable codes, and targets. Surface the actual codes and affected ranges when they matter.

## Stale Plans

If a preview was created earlier, apply only with the returned `operationId` and `confirmationToken`. If `agent.run` reports stale context or target-region drift, stop and create a fresh preview. Do not merge stale Excel edits by guessing.

## Conflicts And Locks

For blocked transactions or lock waits:

1. Report the returned conflict or lock guidance.
2. If a task is involved, explain the current blocker.
3. Ask the user or coordinating agent to wait, split scope, or retry after the lock clears.
4. Retry only after the suggested wait or after splitting the scope.

Structure conflicts and broad workbook conflicts usually require manual review.

## Rollback

- Preview rollback through `excel.agent.run` `mode: "rollback"` or the returned rollback guidance.
- If later related transactions overlap, report that rollback needs chain/dependency review.
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

Choose a supported `excel.agent.run` path or ask for user/host setup. Do not claim that unsupported pivot, chart, comment, note, save-as, or print-layout dimensions were changed.

## Disconnected Add-In

If Excel is not connected or `excel.agent.run mode=status` reports `connectionState: "stale"`:

1. Ask the user to start the agent UI that has the Open Workbook MCP config.
2. If troubleshooting manually, ask the user to run `npx -y @components-kit/open-workbook@latest mcp`.
3. Ask the user to open Excel and load or reload the OpenWorkbook Local taskpane.
4. Retry `excel.agent.run mode=status`.

`activeAddinConnected` is not enough for live workbook work. Prefer `connectionState: "ready"` plus `activeWorkbookAvailable: true`; stale sessions should fail fast with reload guidance instead of triggering offline parsing or repeated tool calls.

Do not use stale snapshots as the source of truth for current workbook state.
