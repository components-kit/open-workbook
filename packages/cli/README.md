# @components-kit/open-workbook

User-facing package for installing, configuring, and running Open Workbook locally.

## Quickstart

```bash
npx -y @components-kit/open-workbook setup
```

Setup prepares the Excel add-in manifest and prints generic MCP config.

Install the Open Workbook Excel skill with skills.sh:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

Paste the printed MCP config into any MCP-capable agent UI:

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

`owb setup` also writes fallback instructions for clients that do not support skills.sh.

## Commands

```bash
owb setup
owb mcp
owb instructions
owb doctor
owb paths
owb sideload mac
owb sideload windows
owb sideload manifest
owb addin serve
owb daemon start
owb file-bridge start
owb service manifest --target macos --service addin
```

## Runtime

- `owb setup` initializes the generic install flow.
- `owb mcp` starts the MCP adapter, starts the local Excel add-in asset server when needed, and uses an embedded backend unless a shared daemon is available.
- `owb instructions` prints fallback generic Excel instructions.
- `owb doctor` checks packaged runtime assets.
- `owb sideload *` generates or installs Excel add-in manifests.
- `owb daemon start` is an advanced shared coordinator for multiple clients.
- `owb file-bridge start` is an optional native bridge for Save As and host file operations.

## Environment

- `OPEN_WORKBOOK_CONFIG_DIR`
- `OPEN_WORKBOOK_HOST`
- `OPEN_WORKBOOK_PORT`
- `OPEN_WORKBOOK_ADDIN_PATH`
- `OPEN_WORKBOOK_ADDIN_HOST`
- `OPEN_WORKBOOK_ADDIN_PORT`
- `OPEN_WORKBOOK_ADDIN_HTTPS=1`
- `OPEN_WORKBOOK_ADDIN_PROTOCOL=https`
- `OPEN_WORKBOOK_ADDIN_TLS_CERT`
- `OPEN_WORKBOOK_ADDIN_TLS_KEY`
- `OPEN_WORKBOOK_ADDIN_URL`
- `OPEN_WORKBOOK_BACKEND_URL`
- `OPEN_WORKBOOK_DAEMON_URL`
- `OPEN_WORKBOOK_FILE_BRIDGE_URL`
- `OPEN_WORKBOOK_STATE_DIR`
- `OPEN_WORKBOOK_BACKUP_DIR`
- `OPEN_WORKBOOK_EXPORT_DIR`
- `OPEN_WORKBOOK_PREVIEW_TOOLS=1`
