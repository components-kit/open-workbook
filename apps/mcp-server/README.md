# @components-kit/open-workbook-mcp-server

MCP stdio server for Open Workbook.

This package exposes `excel.*` tools, resources, and prompts to MCP clients. It also starts the local backend WebSocket used by the sideloaded Excel add-in.

## Usage

Installed CLI users should normally run:

```bash
owb mcp
```

Direct package binary:

```bash
open-workbook-mcp
```

## Environment

- `OPEN_WORKBOOK_HOST`: backend host, default `127.0.0.1`
- `OPEN_WORKBOOK_PORT`: backend port, default `37845`
- `OPEN_WORKBOOK_ADDIN_PATH`: backend WebSocket path, default `/addin`
- `OPEN_WORKBOOK_PREVIEW_TOOLS=1`: expose preview tools in addition to stable tools

## Notes

Use `@components-kit/open-workbook` for end-user installs. This package is published separately so advanced integrators can embed or run the MCP server directly.
