# Session Diagnostics

Use the session diagnostic script when an agent seems to call `excel.agent.run` too many times, repeats previews/applies, or splits a single workbook edit into primitive operations.

```bash
corepack pnpm diagnose:session -- path/to/session.log
corepack pnpm diagnose:session -- path/to/session.jsonl --json
```

The input can be raw text, JSONL, or mixed logs. The script searches for `excel.agent.run` calls, extracts modes, intent actions, operation IDs, and confirmation-token use, then reports:

- total agent calls,
- mode and action counts,
- preview/apply/status/cancel counts,
- consecutive duplicate calls,
- likely missed batching opportunities,
- recommendations such as using `replace_range_with_styled_table` or `operation_status`.

## What To Look For

- Many `preview_update` calls for one logical table fill usually means the agent split values, styling, autofit, and clearing. Use one `replace_range_with_styled_table` preview/apply instead.
- Separate preview calls for filtering and sorting the same table usually mean the agent split one table-view request. Use one `apply_table_view` preview/apply instead.
- Repeated `apply_update` calls with the same operation should become `operation_status` unless the prior response was retryable.
- `cancel_operation` only applies to pending previews. After apply starts, use rollback or backup restore guidance.
- If a style batch times out, expect one parent transaction plus retry chunk transactions rather than many unrelated tool calls.

Session diagnostics are not a release gate. They are a fast way to turn real OpenCode/MCP logs into actionable workflow and prompt improvements.
