# Open Workbook

Open Workbook is a local-first MCP runtime for fast, reversible, template-aware Excel automation.

The project goal is to let MCP clients such as OpenCode control live Excel workbooks through open or frontier models without locking teams into one AI vendor. The first implementation targets desktop Excel on macOS and Windows through an Office.js add-in connected to a local TypeScript backend.

## Design Priorities

- **Safe by default**: every write is planned, snapshotted, diffed, and rollback-aware.
- **Template-aware**: registered templates preserve headers, formulas, styles, filters, tables, print layout, and accounting/report structure.
- **Fast by default**: individual and batch writes both compile to batched Office.js operations with minimal `context.sync()` calls.
- **Local-first**: workbook data, snapshots, diffs, and backups stay on the user's machine unless an explicit integration sends them elsewhere.
- **Model-agnostic**: MCP is the external control interface; OpenRouter or any compatible agent can drive it.

## Architecture

```text
MCP client / agent
       |
       v
apps/mcp-server
       |
       v
apps/backend  <---- WebSocket JSON-RPC ---->  apps/excel-addin
       |
       v
packages/excel-core  +  packages/protocol  +  packages/office-js-engine
```

## Packages

- `apps/mcp-server`: MCP server exposing `excel.*` tools, resources, and prompts.
- `apps/backend`: local broker for add-in sessions, plans, backups, snapshots, diffs, permissions, and telemetry.
- `apps/excel-addin`: minimal Office.js add-in runtime; UI is limited to connection/status.
- `packages/protocol`: shared schemas, JSON-RPC envelopes, tool contracts, resources, and errors.
- `packages/excel-core`: workbook model, range addressing, batch compiler, backup lifecycle, templates, plans, diffs, and validation.
- `packages/office-js-engine`: Office.js adapter that executes compiled batches in Excel.

## Initial Safety Contract

All mutating operations go through the same lifecycle:

1. Resolve workbook, sheet, and range targets.
2. Validate permissions and destructive-action policy.
3. Capture affected-region snapshots.
4. Create or reuse a workbook-copy backup when the operation is structurally risky.
5. Compile individual writes into batch operations.
6. Check target-region fingerprints immediately before apply.
7. Apply through the Office.js engine.
8. Validate template, style, formula, filter, and table invariants.
9. Return a diff summary, telemetry, warnings, and rollback availability.
10. Auto-rollback if post-apply validation fails.

## Local Development

Install and verify:

```bash
corepack pnpm install
corepack pnpm check
corepack pnpm test
```

Run the MCP server for OpenCode or another MCP client:

```bash
corepack pnpm dev:mcp
```

The MCP process also starts the local add-in backend on:

```text
ws://127.0.0.1:37845/addin
```

Run only the backend broker when debugging the add-in connection separately:

```bash
corepack pnpm dev:backend
```

The Excel add-in scaffold currently connects to the backend, sends heartbeat/status messages, and can execute the first Office.js-backed batch operations.

## Status

This repository is in early implementation. The current milestone includes protocol contracts, safety primitives, a local backend broker, a real MCP stdio server, add-in WebSocket RPC, and first Office.js range/template execution primitives.

See:

- [Architecture](docs/architecture.md)
- [Backup Lifecycle](docs/backup-lifecycle.md)
- [Template System](docs/template-system.md)
- [Performance Contract](docs/performance.md)
- [Tool Surface](docs/tool-surface.md)
