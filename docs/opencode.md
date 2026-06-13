# OpenCode Example

OpenCode is one possible MCP client. The generic Open Workbook install flow is still:

```bash
npx -y @component-kit/open-workbook setup
```

Paste the generic MCP config printed by setup into OpenCode using the config shape OpenCode expects. If you prefer an OpenCode-shaped snippet from the local CLI, use:

```bash
owb opencode config --id open-workbook --agent-name finance-agent
```

Example output:

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

For `npx`-based config, keep the command equivalent to:

```json
{
  "mcpServers": {
    "open-workbook": {
      "command": "npx",
      "args": ["-y", "@component-kit/open-workbook@latest", "mcp"]
    }
  }
}
```

Install the Open Workbook Excel skill into OpenCode:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel -a opencode -g -y
```

After Excel opens the add-in, useful first calls are:

```text
excel.runtime.get_status
excel.runtime.get_active_context
excel.runtime.get_capabilities
excel.workbook.get_workbook_map
excel.collab.get_status
```
