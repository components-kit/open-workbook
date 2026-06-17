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

Stable areas include runtime connection, workbook/sheet/range operations, compact workbook/table/range discovery, workbook-wide lookup/search, compact paged reads with token telemetry, reversible batches, combined session-prep, formula-sheet, formula-repair, risky-edit, template-report, and pivot-chart workflows, snapshots, rollback, templates, style fidelity, formula patterns and dependency tracing, tables, filters, sorting, named ranges, regions, validation, repair, cleaning, PivotTables, charts, multi-agent scheduling, permissions, packaging, generic MCP setup, and agent instructions. Host-limited paths return explicit capability-unavailable results instead of pretending to work.

The simple flow is MCP-owned: `npx ... mcp` starts the MCP adapter, the local add-in taskpane server, and an embedded backend when no shared daemon is running. Agents see one public workflow tool, `excel.agent.run`, instead of hundreds of Excel primitives; Open Workbook handles workbook discovery, cached metadata, target resolution, preview/apply, validation, rollback, and compact proof internally. The full operation catalog remains backend capability for deterministic orchestration and test coverage. The shared `owb daemon` remains available for multi-client coordination.

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

Setup prepares the Excel add-in manifest and prints the MCP launch command to add to your agent UI's local stdio MCP configuration.
For an existing install, refresh local setup assets after a package update with:

```bash
npx -y @components-kit/open-workbook@latest upgrade
```

Install the agent skill with skills.sh:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

Use the printed MCP launch command in your agent UI:

```bash
npx -y @components-kit/open-workbook@latest mcp
```

Start the MCP adapter with the public agent workflow surface:

```bash
npx -y @components-kit/open-workbook@latest mcp
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

Source development sideloads should use the development variant so they do not replace the production add-in manifest:

```bash
node packages/cli/dist/index.js sideload mac --development
```

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
- `OPEN_WORKBOOK_ADDIN_RPC_TIMEOUT_MS`
- `OPEN_WORKBOOK_BATCH_DIRECT_OPERATION_THRESHOLD`
- `OPEN_WORKBOOK_BATCH_DIRECT_PAYLOAD_BYTES`
- `OPEN_WORKBOOK_BATCH_DIRECT_CELL_THRESHOLD`
- `OPEN_WORKBOOK_STYLE_BATCH_CHUNK_SIZE`
- `OPEN_WORKBOOK_MATRIX_CHUNK_ROWS`
- `OPEN_WORKBOOK_BACKUP_DIR`
- `OPEN_WORKBOOK_EXPORT_DIR`
- `OPEN_WORKBOOK_STATE_DIR`
- `OPEN_WORKBOOK_FILE_BRIDGE_URL`
- `OPEN_WORKBOOK_FILE_BRIDGE_PORT`

## Token-Saving Reads

Open Workbook exposes `excel.agent.run` so agents can send workbook intent without manually choosing discovery, lookup, read, write, validation, and rollback primitives. The backend builds a workbook metadata cache, resolves natural-language targets across sheets, tables, headers, named ranges, regions, summaries, and formulas, performs targeted compact reads internally, and returns compact structured answers, previews, proof ranges, telemetry, and resource links. Use explicit modes for deterministic workflows: `prepare`, `find`, `answer`, `preview_update`, `apply_update`, `validate`, and `rollback`. Omitted mode or `auto` remains compatible for casual prompts and can apply clearly scoped low-risk value edits after preview checks, while ambiguous, broad, formula-sensitive, structural, or destructive edits stop for review. Close matches return candidates instead of guessing; retry with `target.candidateId` from the chosen candidate to continue on the one-tool surface. Schema/header-only requests return cached table or range metadata when possible; requests for rows, samples, explicit A1 ranges, raw monthly sheets, or actual values perform a live read. Raw monthly transaction/invoice sections can resolve from detected header blocks even when no Excel Table exists. Table append requests preview/apply through the same agent tool while preserving the table mutation path internally.

When details would exceed a caller's budget, compact tools can store the full payload locally and return an `excel://compact/{resource_id}` handle for later retrieval through `excel.compact.get_resource`. Compact summary/schema cache entries are invalidated after workbook mutations. Use `excel.validate.compact` for validation proof that returns counts and examples inline while keeping the full issue report behind a resource handle.

## Common Commands

```bash
corepack pnpm check
corepack pnpm test
corepack pnpm test:e2e:agent-surface
corepack pnpm test:e2e:agent-workflow
corepack pnpm test:e2e:office-agent:behavior
corepack pnpm test:e2e:fast
corepack pnpm test:e2e:agent:quality:compare
corepack pnpm test:e2e:agent:quality:gate
corepack pnpm build
corepack pnpm verify
corepack pnpm pack:dry-run
node packages/cli/dist/index.js paths
node packages/cli/dist/index.js daemon status
node packages/cli/dist/index.js file-bridge status
node packages/cli/dist/index.js instructions
node packages/cli/dist/index.js sideload manifest --out open-workbook.xml
node packages/cli/dist/index.js sideload manifest --development --out open-workbook-local.xml
```

## Agent Instructions

Open Workbook includes generic agent instruction source for fast, reliable live Excel automation through the MCP surface:

- `skills/open-workbook-excel/SKILL.md`
- `skills/open-workbook-excel/references/`

Install the skill with `npx skills add components-kit/open-workbook --skill open-workbook-excel`. The skill teaches agents to inspect runtime capabilities, use lookup before broad reads, choose the narrowest efficient MCP tool, use combined session-prep, formula-sheet, formula-repair, risky-edit, template-report, and pivot-chart workflows when they match the task, avoid sparse null-padded overwrites, batch workbook writes, preserve templates/formulas/styles, validate changes, and recover through snapshots, backups, transactions, and rollback previews. `owb instructions` remains available as a fallback for clients that do not support skills.sh.

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
