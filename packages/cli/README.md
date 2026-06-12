# @open-workbook/cli

User-facing CLI for Open Workbook.

```bash
npm install -g @open-workbook/cli
owb doctor
```

## Commands

```bash
owb daemon start
owb daemon status
owb daemon stop
owb mcp
owb mcp --agent-name tx-cleaner
owb mcp --standalone
owb addin serve
owb service manifest --target macos --service addin
owb sideload mac
owb sideload windows
owb sideload manifest
owb opencode config --id open-workbook
owb doctor
owb paths
```

## Runtime

- `owb daemon start` starts the shared local coordinator used by the Excel add-in and multiple MCP adapters.
- `owb mcp` starts an MCP adapter and attaches to the daemon when available.
- `owb mcp --agent-name NAME` labels that adapter in collaboration status.
- `owb mcp --standalone` starts a single-process MCP server with an embedded backend for one-agent use.
- `owb addin serve` serves the sideloaded Excel taskpane, icons, and generated manifest.
- `owb addin serve --https --tls-cert ./cert.pem --tls-key ./key.pem` serves the add-in over HTTPS with a trusted local certificate.
- `owb service manifest` generates launchd, systemd user, or Windows Task Scheduler wrappers for `owb addin serve` or `owb daemon start`.
- `owb sideload manifest` writes a manifest with the active taskpane and backend URLs.
- `owb sideload mac` copies a generated manifest into Excel for macOS' WEF sideload folder.
- `owb sideload windows` writes a generated manifest and prints trusted shared-folder catalog instructions.

## OpenCode

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

## Asset Resolution

The CLI resolves source-checkout assets first, packaged assets second, and installed dependency assets last. This lets the same command shape work for contributors and installed users.

## Environment

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
- `OPEN_WORKBOOK_AGENT_NAME`
- `OPEN_WORKBOOK_MCP_STANDALONE=1`
- `OPEN_WORKBOOK_SERVICE_COMMAND`
- `OPEN_WORKBOOK_STATE_DIR`
- `OPEN_WORKBOOK_BACKUP_DIR`
- `OPEN_WORKBOOK_PREVIEW_TOOLS=1`
- `OPEN_WORKBOOK_LOCK_DEFAULT_TTL_MS`
- `OPEN_WORKBOOK_LOCK_TRANSACTION_TTL_MS`
- `OPEN_WORKBOOK_LOCK_MAX_TTL_MS`
- `OPEN_WORKBOOK_ALLOW_MANUAL_LOCKS=0`
