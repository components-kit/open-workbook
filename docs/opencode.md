# OpenCode Configuration

Open Workbook exposes Excel tools through MCP. OpenCode connects to an `owb mcp` adapter; the adapter attaches to the shared `owb daemon` when available. OpenRouter or another model provider is configured in OpenCode, not in Open Workbook.

## Installed CLI

```bash
owb opencode config --id open-workbook --agent-name finance-agent
```

Example:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["owb", "mcp", "--agent-name", "finance-agent"],
      "enabled": true
    }
  }
}
```

Add the snippet to your OpenCode MCP configuration.

## Source Checkout

```bash
corepack pnpm build
node packages/cli/dist/index.js opencode config --id open-workbook --command "node packages/cli/dist/index.js" --agent-name finance-agent
```

Example:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["node packages/cli/dist/index.js", "mcp", "--agent-name", "finance-agent"],
      "enabled": true
    }
  }
}
```

## Runtime

Start the shared daemon once:

```bash
owb daemon start
```

OpenCode launches `owb mcp`. That process:

- MCP stdio server for `excel.*` tools
- attaches to the shared daemon when available
- falls back to standalone single-process runtime only when no daemon is running

For one-agent local testing without a daemon:

```bash
owb mcp --standalone
```

Run the add-in asset server separately:

```bash
owb addin serve
```

Then sideload the generated Excel manifest once. See [Installation](installation.md) and [Sideloading](sideloading.md).

## Recommended First Calls

After Excel opens the add-in:

```text
excel.runtime.get_status
excel.runtime.get_active_context
excel.runtime.get_capabilities
excel.workbook.get_workbook_map
excel.collab.get_status
```

Use `excel.plan.*` or `excel.batch.*` for changes that should be previewed and rollback-aware. For stale plans, call `excel.plan.refresh_preview` or `excel.plan.rebase` before apply; both block if the target range changed since preview.

For multi-agent workflows, update status with `excel.task.set_progress`, use `excel.task.add_blocker` when waiting for locks or user input, and use `excel.transaction.preview_rollback_chain` before reverting work that may have later dependent transactions.

Use `excel.task.evaluate_schedule` before starting parallel work on the same workbook. It tells each agent whether its task is ready, waiting on another task, waiting on a lock, or blocked by an explicit blocker. Use `excel.formula.get_dependency_graph` or the trace tools before editing ranges that feed formulas, charts, pivots, or reports.

For long planning work, use `excel.lock.acquire` to reserve a workbook scope, renew it with `excel.lock.renew`, and release it with `excel.lock.release` when the plan is applied or cancelled.

When a transaction is blocked, call `excel.conflict.get_guidance`. It returns concrete next steps such as retry after lock expiry, split the task scope, hand off to the owner task, refresh the plan, or ask the user for manual review.

Use `excel.conflict.get_telemetry` during longer sessions to find repeated contention on the same range, table, task, or agent before assigning more parallel work.
