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

The backend never calls Office.js directly. Backend orchestration calls `AddinRpcClient.request(...)`, the WebSocket bridge carries JSON-RPC method names to the add-in, and `apps/excel-addin/src/host/registry.ts` dispatches those methods to grouped host operation modules. `apps/excel-addin/src/host/executor-core.ts` owns the Office.js implementation source for `Excel.run`, `Office.context`, workbook file export, and `context.sync` calls. `apps/excel-addin/src/excel-executor.ts` remains only as a compatibility export for existing imports.

## Catalog Policy

The protocol package owns the public MCP tool contract and the internal backend capability catalog:

- MCP registers only `excel.agent.run` as a tool.
- The Excel capability catalog remains available to backend orchestration and tests.
- Unfinished capabilities are omitted from catalogs until their contract and implementation are ready.
- Unsupported host-specific capabilities are reported honestly by backend capability metadata and agent outputs.

This keeps agents on one deterministic workflow surface while still letting backend orchestration compose Excel capabilities internally.

Backend capability modules group the internal Excel catalog by operational domain under `apps/backend/src/capabilities`. The registry maps every cataloged capability to its backend domain, implementation owner, optional runtime facade method, agent handler, host operation dependency, stateful manager dependency, and colocated unit test file. `apps/backend/src/excel-capabilities.ts` remains a compatibility wrapper over this registry. The grouping is a planning and test inventory; it does not mean every internal capability is currently routable through `excel.agent.run`. MCP is a thin adapter over `excel.agent.run`; primitive Excel capabilities are not registered as MCP tools.

Add-in host operation modules mirror that grouping under `apps/excel-addin/src/host`. The host registry maps every add-in JSON-RPC method to an implementation owner, related backend capabilities, batch operation kinds when applicable, host dependency class, and colocated unit test file. Backend registry tests verify declared host methods exist in the add-in registry so capability metadata cannot drift silently from the Office.js bridge.

Capability coverage planning is tracked separately from public exposure. Each internal capability has an agent status (`agent_entrypoint`, `agent_action_handler`, or `internal_capability`) and a planning status (`covered`, `needs_unit_contract`, `future_orchestration_candidate`, `host_limited`, or `defer`). Run `corepack pnpm capabilities:report` for a grouped Markdown report. The matrix is for engineering planning only and does not enable additional agent routes.

## Backend Test Layout

Backend tests are grouped by behavior instead of keeping orchestration and runtime coverage in monolithic files. Agent orchestration tests cover prepare/cache, read/answer routing, target resolution, preview/apply safety, and structured intent routing. Runtime service tests cover persistence, transactions/jobs, capabilities/session readiness, range/table behavior, pivot/chart behavior, backups/native file bridge, and locks/tasks. Capability domain tests live beside their domain files in `apps/backend/src/capabilities/domains/*.test.ts`, and the registry test verifies all 306 cataloged capabilities have exactly one backend registry entry and a colocated unit test file. Shared fake runtimes and helper fixtures live in `*.test-support.ts` files so new tests can target the closest subsystem without adding new public MCP tools or expanding orchestration coverage by accident.

Add-in host tests live beside grouped host modules in `apps/excel-addin/src/host/*.test.ts`. They validate registry coverage and dispatch metadata without requiring a live Excel process; live Office.js behavior remains covered by focused executor tests and e2e workflows as those paths are extracted further.
