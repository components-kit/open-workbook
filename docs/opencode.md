# OpenCode Example

OpenCode is one possible MCP client. The generic Open Workbook install flow is still:

```bash
npx -y @components-kit/open-workbook setup
```

OpenCode uses its own MCP config shape. Generate an OpenCode-shaped snippet from the local CLI with:

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

The generated OpenCode snippet uses the optimized compact-first MCP surface. There is no compact/full profile to configure.

For `npx`-based OpenCode config, keep the command equivalent to:

```json
{
  "mcp": {
    "open-workbook": {
      "type": "local",
      "command": ["npx", "-y", "@components-kit/open-workbook@latest", "mcp"],
      "enabled": true
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
