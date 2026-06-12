# Architecture

Open Workbook separates agent-facing protocol from Excel execution.

## Components

- **MCP server**: exposes tools, resources, and prompts to agents.
- **Backend broker**: owns sessions, permissions, backups, templates, snapshots, plans, diffs, and telemetry.
- **Excel add-in**: runs inside Excel and executes Office.js operations over WebSocket JSON-RPC.
- **Core packages**: define shared contracts and workbook-safe primitives.

## Runtime Flow

1. An MCP client calls an `excel.*` tool.
2. The MCP server validates the request and uses the embedded backend broker.
3. The backend creates or updates a plan, compiles batches, and records backup metadata.
4. The backend asks the add-in to snapshot, validate, or apply workbook operations over WebSocket JSON-RPC.
5. The add-in executes optimized Office.js calls inside `Excel.run`.
6. The backend stores operation records and returns a structured result.

The default `open-workbook-mcp` process starts both stdio MCP and the local add-in WebSocket endpoint so a desktop MCP client can launch one process.

## Engine Policy

Office.js is the primary engine for macOS and Windows desktop Excel. Other engines may be added behind the same compiled operation interface when Office.js cannot support a feature.

Engine adapters must report capabilities. Tools should fail honestly or degrade explicitly when a capability is unavailable.
