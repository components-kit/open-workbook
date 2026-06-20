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
npx skills add components-kit/open-workbook --skill open-workbook-skills
```

For OpenCode global install:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-skills -a opencode -g -y
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
excel.agent.run mode=status request="Check Open Workbook status"
excel.agent.run mode=prepare request="Prepare workbook context"
excel.agent.run mode=find request="Find the sheet or table I need"
excel.agent.run mode=answer request="Summarize the active workbook"
```

Use `excel.agent.run` `mode=preview_update` and then `mode=apply_update` for writes so changes remain backup-aware, validated, and rollback-capable.
