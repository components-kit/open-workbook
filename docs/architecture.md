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

The protocol package owns the public MCP tool contract and the internal backend capability catalog:

- MCP registers only `excel.agent.run` as a tool.
- The Excel capability catalog remains available to backend orchestration and tests.
- Unfinished capabilities are omitted from catalogs until their contract and implementation are ready.
- Unsupported host-specific capabilities are reported honestly by backend capability metadata and agent outputs.

This keeps agents on one deterministic workflow surface while still letting backend orchestration compose Excel capabilities internally.

Backend capability modules group the internal Excel catalog by operational domain, including runtime, workbook, backup, worksheet, range, lookup, batch, workflow, plan, job, task, collaboration, lock, conflict, transaction, diff, events, snapshot, compact resource, template, formatting, formula, table, pivot, chart, names, region, validation, repair, cleaning, permissions, and agent domains. The grouping is a planning and test inventory; it does not mean every internal capability is currently routable through `excel.agent.run`. MCP is a thin adapter over `excel.agent.run`; primitive Excel capabilities are not registered as MCP tools.

Capability coverage planning is tracked separately from public exposure. Each internal capability has an agent status (`agent_entrypoint`, `agent_action_handler`, or `internal_capability`) and a planning status (`covered`, `needs_unit_contract`, `future_orchestration_candidate`, `host_limited`, or `defer`). Run `corepack pnpm capabilities:report` for a grouped Markdown report. The matrix is for engineering planning only and does not enable additional agent routes.

## Backend Test Layout

Backend tests are grouped by behavior instead of keeping orchestration and runtime coverage in monolithic files. Agent orchestration tests cover prepare/cache, read/answer routing, target resolution, preview/apply safety, and structured intent routing. Runtime service tests cover persistence, transactions/jobs, capabilities/session readiness, range/table behavior, pivot/chart behavior, backups/native file bridge, and locks/tasks. Shared fake runtimes and helper fixtures live in `*.test-support.ts` files so new tests can target the closest subsystem without adding new public MCP tools or expanding orchestration coverage by accident.
