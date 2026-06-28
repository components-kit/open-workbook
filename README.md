# Open Workbook

Open Workbook is a local Excel agent runtime. It gives any MCP-capable agent one safe tool, `excel.agent.run`, for live desktop Excel workbooks while a local TypeScript backend handles workbook discovery, target resolution, previews, validation, backups, rollback, and multi-agent coordination.

## Why

Daily spreadsheet work needs more than raw cell writes. Open Workbook focuses on the parts generic agents often break:

- one public MCP tool instead of hundreds of spreadsheet primitives
- live Office.js reads and writes through the open desktop workbook
- preview/apply safety with permissions, fingerprints, backups, validation, and rollback
- compact workbook context so agents can work without flooding the model with cells
- multi-agent coordination through a shared daemon, locks, tasks, transactions, and conflict guidance
- honest host capability results when Excel cannot safely perform an operation

## Current Status

Open Workbook is tagged for releases as `@components-kit/open-workbook`; npm publishing is a separate release step. It is not a Microsoft AppSource add-in and does not install itself into Excel without user or admin trust approval.

Packaged MCP clients see only `excel.agent.run`. The internal backend catalog covers workbook, sheet, range, table, formula, style, template, validation, repair, cleaning, backup, snapshot, transaction, permission, and collaboration capabilities for deterministic orchestration and tests.

## Architecture

```text
MCP client / agent
       |
       v
apps/mcp-server  -- stdio MCP server
       |
       v
apps/backend     -- local broker, orchestration, safety, snapshots, transactions
       |
       v
apps/excel-addin -- Office.js taskpane loaded by desktop Excel
       |
       v
Excel workbook
```

Shared packages:

- `packages/protocol`: public agent schema and internal backend capability catalog
- `packages/excel-core`: range parsing, planning, backups, snapshots, templates, permissions, locks, fingerprints, and diffs
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

Setup prepares the Excel add-in manifest and prints the MCP launch command for your agent UI.
For an existing install, refresh local setup assets after a package update with:

```bash
npx -y @components-kit/open-workbook@latest upgrade
```

Install the agent skill with skills.sh:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-skills
```

Use the printed MCP launch command in your agent UI. It will look like:

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
- `OPEN_WORKBOOK_BACKUP_RETENTION_DAYS`
- `OPEN_WORKBOOK_BACKUP_RETENTION_COUNT`
- `OPEN_WORKBOOK_BACKUP_RETENTION_BYTES`
- `OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED`
- `OPEN_WORKBOOK_EXPORT_DIR`
- `OPEN_WORKBOOK_STATE_DIR`
- `OPEN_WORKBOOK_FILE_BRIDGE_URL`
- `OPEN_WORKBOOK_FILE_BRIDGE_PORT`

## Agent Flow

Agents call `excel.agent.run` with natural language plus optional structured fields:

- `status`: check Excel/add-in readiness and collaboration state
- `prepare`: cache workbook structure and return `workbookContextId`
- `find`: locate sheets, tables, headers, named ranges, regions, and candidate targets
- `answer`: read live values or answer from cached metadata when enough
- `preview_update`: prepare a scoped change and return a confirmation token
- `apply_update`: apply a previewed change
- `validate`: validate workbook, sheet, table, formula, style, or template state
- `rollback`: recover through stored backup/transaction guidance

The backend keeps verbose workbook context local and returns compact proof, resource links, telemetry, warnings, and next actions. Caller LLMs may provide canonical `intent.action`, `intent.targetHints`, explicit `target`, structured `values`, and optional `context` policy hints, but the backend still owns ambiguity checks, stale-context checks, permissions, locks, backups, validation, and rollback metadata. Use `context.strategy/scope/include` to describe why and where context is needed; use `detailLevel` only for the returned preset shape such as `workbook_summary`, `sheet_summary`, `semantic_index`, `table_sample`, or `full_table`.

For styling review, agents should use `intent.action: "style_overview"` or `detailLevel: "style_overview"` with `mode: "answer"` to get current style context, column groups, grouped-header suggestions, and workflow hints without full data reads. For workbook design review, such as deciding which columns should be free text, dates, money, ID/text codes, dropdowns, or lookups/references from related sheets, agents should use `intent.action: "workbook_design_overview"` with `mode: "answer"` once before reading related sheets manually. It returns column-by-column recommendations, related-sheet hints, and next workflows without broad-reading empty data rows. For broad styling/readability work, agents should use `intent.action: "improve_visual_readability"` with `mode: "preview_update"` rather than issuing many primitive style calls. Options live under `values.visualReadability`; standard mode compiles safe column-first layout/formatting/highlight rules, comprehensive mode can include preview-only validation/formula suggestions, `stylePreservationMode` defaults to `protected_regions` so summary/template areas and grouped header bands stay guarded while ordinary table body styling, widths, alignment, and date/money formats can still be intentionally improved, `strict` preserves every detected existing style, `none` allows an explicit redesign, `referenceStyle` can preview adaptation from another sheet, and `presentationMode` can preview print/export suggestions. Apply still requires `apply_update` with the returned operation token, `nextAction: "call_apply_update"`, and `operationCount > 0`; if a preview reports `operationCount: 0` or `nextAction: "answer_now"`, agents should explain the skipped reasons instead of applying or decomposing the work into primitive style calls. Use `intent.action: "grouped_header"` for the separate structural preview that inserts a visual group row, merges group labels, and restyles the shifted table header. Grouped-header groups should use `{ "label": "...", "startColumn": "A", "endColumn": "B" }`; `{ "columns": ["A", "B"] }` and `{ "range": "A:B" }` are also accepted. Do not reuse an `operationId` from visual readability when creating a grouped-header preview.

Grouped headers are structure-level styling. If apply is blocked by `DESTRUCTIVE_ACTION_BLOCKED` or `PERMISSION_DENIED`, the public agent path can enable the required policy with `intent.action: "set_permissions"` and `values.permissions` such as `{ "allowWrites": true, "allowDestructiveActions": true, "scopeToWorkbook": true, "requireConfirmationFor": [] }`; after that, create and apply a fresh grouped-header preview.

Example OpenCode prompts:

```text
Use open-workbook. Inspect the active sheet with a style overview first, without reading every data cell. Suggest visual readability improvements including grouped headers, one consistent palette, safe widths, alignment, filters, number formats, and highlights. Do not apply yet.
```

```text
Preview a grouped_header workflow for this sheet. Add a higher-level grouped header row above the existing column headers, merge group labels, and use matching group colors. Wait for approval before apply_update.
```

```json
{
  "mode": "preview_update",
  "intent": { "action": "grouped_header" },
  "target": { "sheetName": "Invoices", "tableName": "InvoicesTable" },
  "values": {
    "stylePreservationMode": "none",
    "groupedHeader": {
      "groups": [
        { "label": "สถานะ", "startColumn": "A", "endColumn": "B" },
        { "label": "ข้อมูลงาน", "startColumn": "C", "endColumn": "E" }
      ]
    }
  }
}
```

```text
Apply the safe visual readability preview in one apply_update. Include opt-in buckets layout, validation, and freeze_panes only if they were present in the preview.
```

With the shared daemon, multiple MCP sessions get distinct trusted agent identities. `status` and `prepare` include compact collaboration summaries for active agents, open tasks, locks, queued/applying transactions, conflicts, and recent events.

## Common Commands

```bash
corepack pnpm check
corepack pnpm test
corepack pnpm capabilities:report
corepack pnpm test:regression:opencode
corepack pnpm test:e2e:mcp-contract
corepack pnpm test:e2e:workbook
corepack pnpm test:e2e:scenarios
corepack pnpm test:e2e:scenarios:regression
corepack pnpm test:e2e:scenarios:departments
corepack pnpm test:e2e:agent-surface
corepack pnpm test:e2e:agent-workflow
corepack pnpm test:e2e:office-agent:behavior
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

The current test strategy and next coverage lanes are documented in [Test Strategy](docs/test-strategy.md).

## Agent Instructions

Open Workbook includes generic agent instruction source for fast, reliable live Excel automation through the MCP surface:

- `skills/open-workbook-skills/SKILL.md`
- `skills/open-workbook-skills/references/`

Install the skill with `npx skills add components-kit/open-workbook --skill open-workbook-skills`. The skill teaches agents to use the public `excel.agent.run` workflow, pass structured intent and target hints when available, normalize multilingual requests into canonical routing fields while preserving the user's language, avoid sparse null-padded overwrites, preserve templates/formulas/styles, validate changes, and recover through snapshots, backups, transactions, and rollback previews. `owb instructions` remains available as a fallback for clients that do not support skills.sh.

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
- [Runtime Configuration](docs/advanced-runtime.md)
- [Architecture](docs/architecture.md)
- [Tool Surface](docs/tool-surface.md)
- [Backup Lifecycle](docs/backup-lifecycle.md)
- [Template System](docs/template-system.md)
- [Style Fidelity](docs/style-fidelity.md)
- [Formula Intelligence](docs/formula-intelligence.md)
- [Tables, Filters, and Sorts](docs/table-filter-sort.md)
- [Names and Regions](docs/names-regions.md)
- [PivotTables and Charts](docs/pivot-chart.md)
- [Range Metadata Reads](docs/advanced-range-reads.md)
- [Validation and Repair](docs/validation-repair.md)
- [Permissions and Cleaning](docs/permissions-cleaning.md)
- [Workbook File Lifecycle](docs/workbook-file-lifecycle.md)
- [Performance Contract](docs/performance.md)
- [Operation Authoring](docs/operation-authoring.md)
- [Session Diagnostics](docs/session-diagnostics.md)
- [Multi-Agent Runtime](docs/multi-agent-runtime.md)
- [Production Readiness](docs/production-readiness.md)
- [Service Wrapper](docs/service-wrapper.md)
- [OpenCode Configuration](docs/opencode.md)
- [Packaging and Publishing](docs/packaging.md)
- [Release Process](docs/release.md)
- [Sideloading](docs/sideloading.md)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for release notes.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). The most important rule is that write tools must not bypass planning, snapshots, permission checks, and rollback-aware batches.

## Security

Open Workbook handles sensitive spreadsheet data. See [SECURITY.md](SECURITY.md). Workbook content stays local by default; integrations that send data to external services must be explicit and documented.

## License

MIT
