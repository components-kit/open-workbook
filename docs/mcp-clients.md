# Generic MCP Clients

Open Workbook is client-neutral. Any MCP-capable agent UI can launch the same command:

```json
{
  "mcpServers": {
    "open-workbook": {
      "command": "npx",
      "args": ["-y", "@components-kit/open-workbook@latest", "mcp"]
    }
  }
}
```

`npx` downloads and runs the public Open Workbook package. The `mcp` command starts the MCP stdio adapter, starts the local Excel add-in taskpane server when needed, and uses an embedded backend unless a shared daemon is already running.

Install the Excel skill with skills.sh:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

For OpenCode global install:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel -a opencode -g -y
```

## Source Checkout

For local development, build the repo and point the MCP client at the built CLI:

```json
{
  "mcpServers": {
    "open-workbook": {
      "command": "node",
      "args": ["packages/cli/dist/index.js", "mcp"]
    }
  }
}
```

Run setup from source to prepare the manifest:

```bash
node packages/cli/dist/index.js setup
```

## First Agent Calls

After Excel opens the Open Workbook add-in, start workbook sessions with:

```text
excel.runtime.get_status
excel.runtime.get_active_context
excel.runtime.get_capabilities
excel.workbook.get_workbook_map
excel.collab.get_status
```

Use `excel.plan.*` or `excel.batch.*` for writes so changes remain backup-aware, validated, and rollback-capable.
