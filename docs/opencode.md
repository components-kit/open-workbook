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

The generated OpenCode snippet uses the public `excel.agent.run` workflow surface. Agents should call only `excel.agent.run`; the backend handles workbook discovery, target resolution, reads, previews, applies, validation, and compact proof internally. The primitive operation catalog is backend/test capability, not a normal OpenCode tool surface. Agents should not fall back to Python, openpyxl, pandas, shell scripts, or offline `.xlsx` parsing for a connected live workbook unless the user explicitly asks for offline file analysis or approves a non-live fallback after MCP is unavailable.

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

After Excel opens the add-in, useful default-surface calls are:

```text
excel.agent.run mode=status request="Check Open Workbook status"
excel.agent.run mode=prepare request="Prepare workbook context"
excel.agent.run mode=find request="Find the sheet or table I need"
excel.agent.run mode=answer request="Summarize the active workbook"
excel.agent.run mode=answer request="Compare January and February"
excel.agent.run mode=preview_update request="Change Sales!E2 to Reviewed"
excel.agent.run mode=apply_update request="Apply the previewed update"
```

For related edits across multiple ranges, send one grouped preview using `values.patches` and then apply that returned operation once. Do not issue one preview/apply pair per zone, column group, or row block unless the grouped apply returns a hard failure with actionable issue details.

Omitted mode or `mode=auto` remains compatible for casual prompts, but explicit modes are more predictable for agent UIs. The backend should either answer from compact proof in one call or return a precise `nextAction`; agents should not chain primitive compact tools.

`mode=status` reports workbook readiness with `connectionState`. `ready` means the add-in responded and an active workbook is available. `stale` means the backend saw an old or unresponsive taskpane session; reload or reopen the OpenWorkbook Local taskpane in Excel before retrying, and restart Excel only if the taskpane cannot reconnect.
