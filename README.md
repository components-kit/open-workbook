# Open Workbook

Open Workbook is a local-first MCP runtime for fast, reversible, template-aware Excel automation. It connects any MCP-capable agent UI to live desktop Excel workbooks through an Office.js add-in and a local TypeScript backend, so teams can use their preferred agent without being locked into one client or model vendor.

## Why

Daily spreadsheet work usually does not need the largest model available, but it does need reliable workbook handling. Open Workbook focuses on the parts generic agents often break:

- preserving templates, headers, formulas, filters, tables, print layout, and styling
- batching reads and writes through Office.js instead of slow per-cell automation
- creating backups and rollback paths before changing workbooks
- keeping workbook data local unless a user explicitly sends it elsewhere
- exposing clear capability status for unsupported or host-limited Excel operations

## Current Status

The project is being prepared for npm distribution as `@components-kit/open-workbook`. It is not a Microsoft AppSource add-in and does not attempt to install itself into Excel without user or admin trust approval.

Stable areas include runtime connection, workbook/sheet/range operations, reversible batches, snapshots, rollback, templates, style fidelity, formula patterns and dependency tracing, tables, filters, sorting, named ranges, regions, validation, repair, cleaning, PivotTables, charts, multi-agent scheduling, permissions, packaging, generic MCP setup, and agent instructions. Some advanced Office.js-limited paths return explicit capability-unavailable results instead of pretending to work.

The simple flow is MCP-owned: `npx ... mcp` starts the MCP adapter, the local add-in taskpane server, and an embedded backend when no shared daemon is running. The shared `owb daemon` remains available for advanced multi-client coordination.

## Architecture

```text
MCP client / agent
       |
       v
apps/mcp-server  -- stdio MCP server
       |
       v
apps/backend     -- local WebSocket broker, plans, backups, snapshots
       |
       v
apps/excel-addin -- Office.js taskpane loaded by desktop Excel
       |
       v
Excel workbook
```

Shared packages:

- `packages/protocol`: tool catalog, JSON-RPC contracts, resources, prompts, errors, and workbook types
- `packages/excel-core`: range parsing, planning, backups, snapshots, templates, permissions, fingerprints, and diffs
- `packages/office-js-engine`: Office.js execution interface and defaults
- `packages/cli`: `owb` CLI for setup, running MCP, serving the add-in, generating manifests, sideloading, fallback instructions, and diagnostics

## Requirements

- Node.js `>=20.11`
- Desktop Microsoft Excel on macOS or Windows
- An MCP-capable agent UI
- Network access for `npx` package installation and whichever model provider your agent UI uses

## Quickstart

Run setup:

```bash
npx -y @components-kit/open-workbook setup
```

Setup prepares the Excel add-in manifest and prints the MCP config to add to your agent UI.
For an existing install, refresh local setup assets after a package update with:

```bash
npx -y @components-kit/open-workbook@latest upgrade
```

Install the agent skill with skills.sh:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

Paste the printed MCP config into any MCP-capable agent UI:

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

Start the agent UI before opening the Open Workbook add-in in Excel; the MCP command starts the local add-in asset server and backend for the simple flow.

## Source Development

```bash
git clone https://github.com/components-kit/open-workbook.git
cd open-workbook
corepack pnpm install
corepack pnpm build
node packages/cli/dist/index.js setup --dry-run
node packages/cli/dist/index.js doctor
```

Run setup from the checkout:

```bash
node packages/cli/dist/index.js setup
```

Use this MCP command in a local agent config:

```json
{
  "mcpServers": {
    "open-workbook": {
      "command": "node",
      "args": ["packages/cli/dist/index.js", "mcp"]
    }
  }
}
```

## Excel Add-in

`setup` prepares the Excel add-in manifest. On macOS it copies the manifest into Excel's local WEF sideload folder. On Windows it writes the manifest and prints the Trusted Add-in Catalog steps.

Manual sideloading and custom manifest generation are documented in [Sideloading](docs/sideloading.md).

## Runtime URLs

Defaults:

- Add-in taskpane: `http://127.0.0.1:37846/taskpane.html`
- Add-in backend: `ws://127.0.0.1:37845/addin`
- Native file bridge: `http://127.0.0.1:37847`

Environment overrides:

- `OPEN_WORKBOOK_ADDIN_HOST`
- `OPEN_WORKBOOK_ADDIN_PORT`
- `OPEN_WORKBOOK_ADDIN_URL`
- `OPEN_WORKBOOK_HOST`
- `OPEN_WORKBOOK_PORT`
- `OPEN_WORKBOOK_ADDIN_PATH`
- `OPEN_WORKBOOK_BACKUP_DIR`
- `OPEN_WORKBOOK_EXPORT_DIR`
- `OPEN_WORKBOOK_STATE_DIR`
- `OPEN_WORKBOOK_FILE_BRIDGE_URL`
- `OPEN_WORKBOOK_FILE_BRIDGE_PORT`
- `OPEN_WORKBOOK_PREVIEW_TOOLS=1`

## Common Commands

```bash
corepack pnpm check
corepack pnpm test
corepack pnpm build
corepack pnpm verify
corepack pnpm pack:dry-run
node packages/cli/dist/index.js paths
node packages/cli/dist/index.js daemon status
node packages/cli/dist/index.js file-bridge status
node packages/cli/dist/index.js instructions
node packages/cli/dist/index.js sideload manifest --out open-workbook.xml
```

## Agent Instructions

Open Workbook includes generic agent instruction source for fast, reliable live Excel automation through the MCP surface:

- `skills/open-workbook-excel/SKILL.md`
- `skills/open-workbook-excel/references/`

Install the skill with `npx skills add components-kit/open-workbook --skill open-workbook-excel`. The skill teaches agents to inspect runtime capabilities, choose the narrowest efficient MCP tool, batch workbook writes, preserve templates/formulas/styles, validate changes, and recover through snapshots, backups, transactions, and rollback previews. `owb instructions` remains available as a fallback for clients that do not support skills.sh.

## Safety Contract

Mutating operations should follow the same lifecycle:

1. Resolve workbook, sheet, and range targets.
2. Validate permissions, locked regions, and destructive-action policy.
3. Capture affected-region snapshots.
4. Create backups for rollback.
5. Compile writes into batch operations.
6. Check target fingerprints immediately before apply.
7. Apply through Office.js.
8. Validate template, style, formula, table, filter, and layout invariants where requested.
9. Return telemetry, warnings, diff summaries, and rollback IDs.
10. Roll back through stored snapshots when requested or when repair workflows fail validation.

## Documentation

- [Installation](docs/installation.md)
- [MCP Clients](docs/mcp-clients.md)
- [Generic Instructions](docs/instructions.md)
- [Advanced Runtime](docs/advanced-runtime.md)
- [Architecture](docs/architecture.md)
- [Tool Surface](docs/tool-surface.md)
- [Backup Lifecycle](docs/backup-lifecycle.md)
- [Template System](docs/template-system.md)
- [Style Fidelity](docs/style-fidelity.md)
- [Formula Intelligence](docs/formula-intelligence.md)
- [Tables, Filters, and Sorts](docs/table-filter-sort.md)
- [Names and Regions](docs/names-regions.md)
- [PivotTables and Charts](docs/pivot-chart.md)
- [Advanced Range Reads](docs/advanced-range-reads.md)
- [Validation and Repair](docs/validation-repair.md)
- [Permissions and Cleaning](docs/permissions-cleaning.md)
- [Workbook File Lifecycle](docs/workbook-file-lifecycle.md)
- [Performance Contract](docs/performance.md)
- [Multi-Agent Runtime](docs/multi-agent-runtime.md)
- [Production Readiness](docs/production-readiness.md)
- [Service Wrapper](docs/service-wrapper.md)
- [OpenCode Configuration](docs/opencode.md)
- [Packaging and Publishing](docs/packaging.md)
- [Release Process](docs/release.md)
- [Sideloading](docs/sideloading.md)
- [Roadmap](docs/roadmap.md)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for release notes.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). The most important rule is that write tools must not bypass planning, snapshots, permission checks, and rollback-aware batches.

## Security

Open Workbook handles sensitive spreadsheet data. See [SECURITY.md](SECURITY.md). Workbook content stays local by default; integrations that send data to external services must be explicit and documented.

## License

MIT
