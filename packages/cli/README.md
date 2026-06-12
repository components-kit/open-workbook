# @open-workbook/cli

User-facing CLI for Open Workbook.

```bash
npm install -g @open-workbook/cli
owb doctor
```

## Commands

```bash
owb mcp
owb addin serve
owb sideload mac
owb sideload windows
owb sideload manifest
owb opencode config --id open-workbook
owb doctor
owb paths
```

## Runtime

- `owb mcp` starts the MCP stdio server and the local Excel add-in backend WebSocket.
- `owb addin serve` serves the sideloaded Excel taskpane, icons, and generated manifest.
- `owb sideload manifest` writes a manifest with the active taskpane and backend URLs.
- `owb sideload mac` copies a generated manifest into Excel for macOS' WEF sideload folder.
- `owb sideload windows` writes a generated manifest and prints trusted shared-folder catalog instructions.

## OpenCode

```bash
owb opencode config --id open-workbook
```

Example output:

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

## Asset Resolution

The CLI resolves source-checkout assets first, packaged assets second, and installed dependency assets last. This lets the same command shape work for contributors and installed users.

## Environment

- `OPEN_WORKBOOK_HOST`
- `OPEN_WORKBOOK_PORT`
- `OPEN_WORKBOOK_ADDIN_PATH`
- `OPEN_WORKBOOK_ADDIN_HOST`
- `OPEN_WORKBOOK_ADDIN_PORT`
- `OPEN_WORKBOOK_ADDIN_URL`
- `OPEN_WORKBOOK_BACKEND_URL`
- `OPEN_WORKBOOK_BACKUP_DIR`
- `OPEN_WORKBOOK_PREVIEW_TOOLS=1`
