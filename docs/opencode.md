# OpenCode Configuration

Open Workbook exposes Excel tools through MCP. OpenCode connects to the local MCP server; OpenRouter or another model provider is configured in OpenCode, not in Open Workbook.

## Installed CLI

```bash
owb opencode config --id open-workbook
```

Example:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["owb", "mcp"],
      "enabled": true
    }
  }
}
```

Add the snippet to your OpenCode MCP configuration.

## Source Checkout

```bash
corepack pnpm build
node packages/cli/dist/index.js opencode config --id open-workbook --command "node packages/cli/dist/index.js"
```

Example:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["node packages/cli/dist/index.js", "mcp"],
      "enabled": true
    }
  }
}
```

## Runtime

OpenCode launches `owb mcp`. That process starts:

- MCP stdio server for `excel.*` tools
- local backend WebSocket for the Excel add-in

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
```

Use `excel.plan.*` or `excel.batch.*` for changes that should be previewed and rollback-aware.
