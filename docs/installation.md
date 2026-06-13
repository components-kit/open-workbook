# Installation

Open Workbook is distributed as one npm package for technical users:

```bash
npx -y @components-kit/open-workbook setup
```

The setup command prepares the local machine, but it does not edit every agent UI automatically. After setup, install the Open Workbook Excel skill with skills.sh and copy the printed MCP config into whichever MCP-capable agent you use.

## Requirements

- Node.js `>=20.11`
- Desktop Excel on macOS or Windows
- An MCP-capable agent UI
- Permission to sideload or trust an Office add-in manifest

## Quickstart

Run setup:

```bash
npx -y @components-kit/open-workbook setup
```

Install the Open Workbook Excel skill:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

It also prepares an Excel add-in manifest. On macOS, setup copies the manifest into Excel's local WEF sideload folder. On Windows, setup writes a manifest and prints the Trusted Add-in Catalog steps that Excel requires.

Paste the printed MCP config into your agent UI:

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

## Daily Use

1. Open your agent UI so it launches the Open Workbook MCP command.
2. Open Excel.
3. Open the Open Workbook add-in.
4. Ask the agent to inspect or edit the active workbook.

The `mcp` command starts the MCP adapter and, in the default flow, starts the local add-in taskpane server if it is not already running. The MCP server uses an embedded backend when no shared daemon is available, so users do not need to run `owb daemon start` for the simple flow.

## Useful Commands

```bash
npx -y @components-kit/open-workbook doctor
npx -y @components-kit/open-workbook instructions
npx -y @components-kit/open-workbook sideload manifest --out open-workbook.xml
```

`instructions` prints a fallback instruction bundle for clients that do not support skills.sh.

Advanced shared-daemon, service-wrapper, custom-port, HTTPS, and file-bridge setup is documented in [Advanced Runtime](advanced-runtime.md).

## Troubleshooting

- Run `npx -y @components-kit/open-workbook doctor` to confirm packaged assets are available.
- If the Excel add-in cannot load, start your agent UI first so `npx ... mcp` can start the local taskpane server.
- If Windows Excel does not show the add-in, confirm the manifest is in a trusted shared-folder catalog and `Show in Menu` is enabled.
- If the add-in loads but does not connect, confirm the manifest backend URL points at `ws://127.0.0.1:37845/addin`.
