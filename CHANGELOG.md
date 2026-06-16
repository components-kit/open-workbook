# Changelog

All notable changes to Open Workbook will be documented in this file.

The format is based on Keep a Changelog, and this project follows semantic versioning for published packages.

## [0.1.12] - 2026-06-16

### Added

- Added the default `excel.agent.run` MCP workflow surface so normal agents can use one compact intent interface while the backend orchestrates workbook discovery, cached metadata, target resolution, preview/apply, validation, rollback, and proof generation internally.
- Added server-side `workbookContextId` reuse for agent prepare/find/answer/update flows, plus context-aware advanced compact routes for workbook summary, sheet summary, range compact reads, table compact reads, and compact resource fetches.
- Added natural-language target resolution across sheets, tables, headers, named ranges, regions, summary blocks, and formula regions with structured ambiguity candidates instead of guessing.
- Added safe `auto` mode application for clearly scoped low-risk value edits after preview checks, while formula-sensitive, broad, sparse, structural, destructive, stale, or ambiguous edits stop before mutation.
- Added agent surface and workflow E2E coverage for default tool exposure, metadata cache reuse, compact target reads, preview/apply confirmation, safe auto-apply, formula-sensitive blocking, and context-aware compact routes.

### Changed

- Default MCP `tools/list` now exposes only `excel.agent.run`; set `OPEN_WORKBOOK_MCP_SURFACE=advanced` to expose the previous optimized primitive surface for debugging and compatibility.
- Updated Open Workbook agent skill guidance, tool-surface docs, release E2E report, and generated `llms-full.txt` for the agent workflow and context ID routing model.

## [0.1.11] - 2026-06-15

### Changed

- MCP now exposes one optimized compact-first tool surface by default, removing compact/full/read-only profile and allow/deny list configuration from the user-facing CLI.
- Raw high-token range/table reads, full snapshot/diff detail tools, and duplicate filter/sort wrappers are hidden from the public MCP surface while compact reads and advanced workbook tools remain available.

## [0.1.10] - 2026-06-15

### Added

- Added compact context garbage collection and context statistics tools for pruning stored workbook detail and debugging token savings.
- Added `excel.workflow.inspect_analyze` for deterministic local table/range profiling with compact proof and stored full detail.
- Added `excel.workflow.rollback_validate` to combine rollback or backup restore, recalculation, and compact workbook validation in one recovery workflow.
- Added compact-profile validation that keeps the advertised tool surface aligned with the protocol catalog.

### Changed

- Compact profile responses now support brief/standard/verbose response modes, bounded proof output, answer-now recommendations, confidence metadata, and stale context invalidation after mutations.
- Mutating batch and workflow tools now accept `idempotencyKey` so retries can return prior proof instead of applying workbook edits twice.
- Synthetic benchmarks now include compact response-size and token-savings coverage for large reads, validation failures, and mutation/diff proof.

## [0.1.9] - 2026-06-15

### Added

- Added generic MCP tool profiles with `full`, `compact`, and `read-only` surfaces plus explicit allow/deny list controls for agent UIs that need lower context overhead.
- Added guarded compact resource inspection with metadata, preview, page, and full modes so stored detail can be inspected without pulling large payloads into model context by default.

### Changed

- Compact mode now hides raw high-context range and table reads, trims empty compact-read output, and applies a hard response budget that stores oversized results behind `resourceUri`.
- Compact capabilities and workflow preflight responses now return profile summaries unless the full catalog is explicitly requested.
- Snapshot, backup, and mutation responses now keep workbook recovery payloads local and return metadata, rollback proof, and resource handles instead of inline cell matrices.

## [0.1.8] - 2026-06-15

### Added

- Added compact token-saving discovery and read tools: `excel.workbook.get_summary`, `excel.workbook.get_used_range_summary`, `excel.sheet.get_summary`, `excel.table.get_schema`, `excel.range.get_summary`, `excel.range.read_compact`, and `excel.table.read_compact`.
- Added lookup-first token-saving tools: `excel.lookup.search_workbook`, `excel.lookup.find_headers`, `excel.lookup.find_tables_by_columns`, `excel.lookup.find_entity`, `excel.lookup.resolve_range`, and `excel.lookup.inspect_match`.
- Added compact read telemetry with `payloadBytes`, rough `estimatedTokens`, truncation status, and continuation metadata so agents can avoid sending broad workbook payloads to models by default.
- Added local compact detail resources, compact cache lifecycle tools, `excel.validate.compact`, and additive mutation `compactProof` metadata so large details can be expanded only when needed.
- Added compact snapshot and diff access through `excel.snapshot.get_compact`, `excel.snapshot.compare_compact`, and `excel.diff.get_compact`.

## [0.1.7] - 2026-06-15

### Added

- Added a development sideload manifest variant with a separate Office add-in ID and `OpenWorkbook Local` display name so source installs can coexist with the production sideload.

## [0.1.6] - 2026-06-15

### Added

- Added transaction progress controls with `excel.transaction.wait` and `excel.transaction.cancel`, plus queued transaction progress metadata for agents coordinating long-running workbook mutations.
- Added `excel.batch.preflight` and `excel.batch.submit_chunked` so agents can estimate batch size and choose synchronous apply, queued submit, or chunked parent-job submit before Excel receives writes.
- Added parent job progress tools with `excel.job.list`, `excel.job.get`, `excel.job.wait`, and `excel.job.cancel` for large updates split across multiple queued transactions.
- Added `excel.range.write_styles_many` for applying grouped report styles through one reversible batch transaction.
- Added `excel.batch.submit` for queueing batch mutations and returning transaction progress immediately.

### Changed

- Made the backend-to-add-in RPC timeout configurable with `OPEN_WORKBOOK_ADDIN_RPC_TIMEOUT_MS`.
- Large `excel.range.write_styles_many` payloads now split into queued chunks controlled by `OPEN_WORKBOOK_STYLE_BATCH_CHUNK_SIZE`.
- Large style, value, formula, and number-format batches now preflight into safe chunk plans where possible instead of relying on parallel MCP calls or timeout recovery.
- Style-only batches now retry add-in timeouts by queueing smaller chunks when safe.

## [0.1.5] - 2026-06-14

### Added

- Added `excel.workflow.preview_risky_edit`, a combined scoped risky-edit workflow that returns before/after snapshots, diff, transaction metadata, and rollback preview.
- Added `excel.workflow.prepare_session`, a read-only combined discovery workflow for runtime status, active context, capabilities, workbook map, and collaboration state.
- Added `excel.workflow.create_formula_sheet`, a combined formula-sheet workflow with value writes, formula writes, number formats, and formula validation.
- Added `excel.workflow.create_template_report`, a combined template report workflow with declared region fill, style comparison/repair, and template validation.
- Added `excel.workflow.create_pivot_chart_summary`, a combined PivotTable/chart workflow with refresh and source validation.
- Added `excel.workflow.repair_formula_errors`, a combined formula repair workflow with validation, pattern read, dependency graph inspection, scoped repair, and after-validation.
- Combined mutating workflows now return an internal preflight payload so compact or low-cost agents can still establish workbook identity and capability context when they select a matched workflow directly.
- Added release-gate E2E script entries for deterministic fake-host coverage, Codex agent decision coverage, live Excel host gates, and generated E2E reports.

### Changed

- Split Codex agent E2E into a blocking core safety lane and report-only workflow quality lane, with bundled `open-workbook-excel` skill guidance loaded into agent prompts.
- Added report-only cheap/frontier Codex agent quality comparison and failure-category reporting for workflow diagnostics.
- Replaced the live Excel E2E placeholder with an opt-in backend/add-in connectivity smoke contract that writes JSON and Markdown artifacts.
- Strengthened skill and MCP tool guidance for formula writes, large table operations, formula repair, snapshots, diffs, and multi-agent locking.
- Blocked sparse/null-padded risky workflow value writes by default and added strict quality gate scripts for cheap/frontier agent comparison.
- Changed setup and upgrade output to print the MCP local stdio launch command instead of a client-specific config wrapper.

## [0.1.4] - 2026-06-13

### Added

- Added a Vite-built React Excel taskpane with updated styling and bundled add-in assets.
- Added packaged add-in icon assets and hosted manifest icon URLs for Excel ribbon and Developer Add-ins branding.

### Changed

- Kept the local taskpane server on loopback HTTP while using hosted HTTPS assets only for static branding icons.
- Updated sideloading and advanced runtime documentation to match the local taskpane serving model.

## [0.1.3] - 2026-06-13

### Added

- Added `owb upgrade` for refreshing local manifests and fallback instruction assets after package updates.
- Added `excel.table.reorder_columns` for preserving Excel table identity while changing column order.
- Added projected `excel.table.read` options for column selection, row windows, and payload facets on large tables.

### Changed

- Improved large-table guidance so agents avoid full-table rewrites and use table-native operations.
- Limited `excel.clean.normalize_headers` writes to the normalized header row instead of rewriting the full source range.

## [0.1.2] - 2026-06-13

### Added

- Added version-aware local runtime checks so `owb mcp` fails fast instead of silently reusing an older running Excel taskpane asset server.
- Added add-in server `/status` metadata with runtime version, process id, taskpane/backend URLs, and packaged workspace module health.
- Added `setup` and `doctor` update notices that warn when a newer `@components-kit/open-workbook` npm version is available and remind users to restart their MCP host after upgrading.

### Changed

- Runtime status now includes backend package/version/process metadata, and MCP server metadata follows the CLI-launched package version.
- CLI smoke coverage now verifies packaged add-in server health, stale-runtime rejection, and upgrade notice output.

## [0.1.1] - 2026-06-13

### Fixed

- Fixed npm-packaged Excel taskpane runtime module loading by resolving browser workspace module imports from installed `@components-kit` dependency packages.
- Added a CLI smoke check that starts the packaged add-in server and verifies taskpane workspace module endpoints return JavaScript instead of 404 responses.

## [0.1.0] - 2026-06-13

### Added

- Initial public release candidate for Open Workbook.
- Local-first MCP runtime for connecting MCP-capable agents to desktop Excel through a sideloaded Office.js add-in.
- User-facing `@components-kit/open-workbook` CLI with `setup`, `mcp`, `doctor`, `paths`, `instructions`, `addin`, `daemon`, `file-bridge`, and `sideload` commands.
- Publishable protocol, core workbook utilities, Office.js engine contracts, backend runtime, MCP server, and CLI packages.
- Excel tool surface covering workbook, worksheet, range, formula, table, filter, sort, chart, PivotTable, template, snapshot, backup, validation, repair, cleanup, permissions, and file lifecycle workflows.
- Reversible workbook mutation flow with planning, permission checks, snapshots, backups, fingerprints, validation, diff summaries, and rollback support.
- Packaged agent instruction skill under `skills/open-workbook-excel` and bundled CLI fallback instructions.
- Documentation for installation, MCP clients, architecture, runtime, sideloading, packaging, safety contracts, workbook lifecycle, performance, and production readiness.
- Release validation scripts for builds, package metadata, MCP catalog consistency, documentation surface, skill assets, CLI smoke checks, package dry runs, tests, and synthetic benchmarks.

### Notes

- This release prepares the repository and packages for open-source distribution. npm publishing is intentionally deferred.
- Open Workbook is not a Microsoft AppSource add-in; Excel manifest trust is handled through user or administrator sideloading.
