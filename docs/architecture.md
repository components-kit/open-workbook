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

The default `owb mcp` process starts stdio MCP and uses an embedded backend when no shared daemon is available. The user-facing CLI also starts the local add-in taskpane server when needed, so a desktop MCP client can launch one command for the simple flow.

## Engine Policy

Office.js is the primary engine for macOS and Windows desktop Excel. Other engines may be added behind the same compiled operation interface when Office.js cannot support a feature.

Engine adapters must report capabilities. The Office.js add-in reports host platform, Office version when available, supported `ExcelApi` versions, and derived feature statuses during its WebSocket hello. Tools should fail honestly or degrade explicitly when a capability is unavailable.

## Catalog Policy

The protocol package owns the full tool, resource, and prompt catalog. MCP registration is capability-gated:

- stable tools are registered by default
- preview tools, when present, require `OPEN_WORKBOOK_PREVIEW_TOOLS=1`
- unfinished tools are omitted from catalogs until their contract and implementation are ready
- unsupported host-specific capabilities are reported honestly by the relevant stable tools

This keeps agents from calling incomplete tools while still letting stable tools report host limitations explicitly.
