# @components-kit/open-workbook-mcp-server

MCP stdio server for Open Workbook.

This package exposes the public `excel.agent.run` workflow tool, JSON resources, and prompts to MCP clients. It also starts or connects to the local backend WebSocket used by the sideloaded Excel add-in.

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
- `OPEN_WORKBOOK_DAEMON_URL`: existing backend daemon URL, default derived from host/port
- `OPEN_WORKBOOK_MCP_STANDALONE`: set `1` to force an in-process backend

## Source Layout

- `src/index.ts`: small stdio bootstrap
- `src/tools/agent-run.ts`: the single public MCP tool schema and registration
- `src/resources.ts`: JSON resource registrations
- `src/prompts.ts`: safe workflow prompts
- `src/runtime-facade.ts`: daemon proxy or standalone backend startup

## Notes

Use `@components-kit/open-workbook` for end-user installs. This package is published separately so advanced integrators can embed or run the MCP server directly.
