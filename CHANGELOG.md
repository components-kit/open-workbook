# Changelog

All notable changes to Open Workbook will be documented in this file.

The format is based on Keep a Changelog, and this project follows semantic versioning for published packages.

## [Unreleased]

## [0.1.22] - 2026-06-23

### Added

- Added first-class `read_formulas` support through `excel.agent.run`, including exact formula proof, displayed values, R1C1/pattern evidence when available, and formula/hardcoded/blank status.
- Added formula-safe response shaping so formula reads and formula-pattern answers preserve bounded inline proof in brief and standard MCP responses.
- Added row-aware `formula_like` derivations and a `settle_reconciliation` workflow for Payment Variance, Reconciliation Note, and Detail Notes updates without reading whole columns into model context.

### Fixed

- Fixed exact formula questions that were routed through generic find/read paths or compacted into value-only answers.
- Fixed formula/reference workflows that could infer formula existence from displayed numbers instead of live formula proof.
- Fixed semantic-index budget compaction so role-aware entries and next-step hints remain available after compacting large workbook maps.

### Changed

- Improved formula, reconciliation, and search guidance in the MCP description and Open Workbook skills so agents use `read_formulas`, `derive_values`, `settle_reconciliation`, `search_range`, or `find_similar_rows` instead of broad reads or Python/openpyxl fallback.
- Improved formula/reconciliation semantic hints so formula regions advertise formula reads, validation, repair, and settlement workflows.
- Refreshed the local development launcher so the file bridge URL is exported before the daemon starts.

## [0.1.21] - 2026-06-22

### Added

- Added regression coverage for selection-aware live reads, targeted table/range samples, continuation-only result fetches, and no-Python fallback behavior when Open Workbook is connected.
- Added cross-sheet context hints and similar-row discovery so agents can find prior-period labels, style/template candidates, dropdown columns, and validation rules without reading whole worksheets.
- Added compact-response safeguards for wide rows, column-role projections, dictionary-encoded repeated values, workbook-context freshness handles, and task-completion telemetry.

### Fixed

- Fixed bloated runtime state by migrating inline workbook backup payloads to `payloadRef` JSON files and repairing existing oversized state on startup.
- Fixed live read routing so targeted sheet/table/range requests return exact values instead of empty snapshots or active-selection data.
- Fixed patch target handling, stale workbook-only permission scopes, object-placeholder apply warnings, and small exact value edits that were being forced through preview/apply loops.
- Fixed OpenCode workflow guidance so connected live workbooks use `excel.agent.run` first, avoid Webfetch/Python/openpyxl fallbacks, and stop after `taskOutcome: final_answer` with `maxRecommendedFollowupCalls: 0`.

### Changed

- Changed safe explicit value edits to use `mode: "auto"` by default after session write access is allowed, while keeping preview/apply for formulas, styles, tables, templates, structural changes, broad edits, and user-reviewable changes.
- Improved dropdown workflows so agents read validation/source-list proof before correcting exact source-list cells.
- Refreshed packaged and installed Open Workbook skill guidance for selection-first targeting, dropdowns, small-edit auto-apply, context reuse, and multi-agent coordination.

## [0.1.20] - 2026-06-22

### Added

- Added stricter worksheet operation regression coverage for styles, inserted columns, column reorder, dropdown validation, formula conditional formatting, closed-workbook recovery, and repeated-call budgets across operation, MCP scenario, `.xlsx`, and live Excel lanes.
- Added department workflow scenario coverage for finance, sales/ops, logistics, HR, executive reporting, and data cleanup, including workbook artifact reload assertions for mutating scenarios.
- Added opt-in live Excel regression groups for worksheet operation fixes, template-backed formula repair, PivotTable/chart creation, PivotTable template repair, and chart template copy.

### Fixed

- Fixed agent mutation routing for worksheet styling, formulas, validation, conditional formatting, inserted columns, and column reorder so `excel.agent.run` can preview/apply the correct operation payloads.
- Fixed batch compiler and Office.js executor handling for whole-row/whole-column structural targets, unsupported operation reporting, PivotTable validation casing, synthetic `Values` hierarchy compatibility, and readonly conditional-formatting rules.

### Changed

- Replaced misleading registry-only coverage scaffolding with behavior-focused test lanes and documented the current release gates for deterministic, workbook-file, scenario, and live Excel claims.

## [0.1.19] - 2026-06-21

### Fixed

- Fixed `apply_update` responses being rejected by strict MCP clients after successful workbook mutations by declaring invalidated workbook context and resource handles in the MCP output schema.
- Improved mutation preview diagnostics so agents that embed rows only in request text are told to send structured `values` payloads.
- Updated OpenCode and Open Workbook skill guidance for structured `write_values` payloads and top-level `operationId`/`confirmationToken` apply calls.

## [0.1.18] - 2026-06-20

### Added

- Added a semantic workbook index and deterministic intent routing so agents can find workbook, sheet, table, region, formula, and template targets from cached metadata before reading cells.
- Added first-class resource-handle reuse for workbook contexts, semantic indexes, operation status, stored agent results, and compact result aliases.
- Added mutation invalidation evidence with `invalidatedContextIds` and `invalidatedResourceUris` so agents do not reuse stale workbook context or result handles after edits.
- Added targeted handle, resource, and invalidation tests covering continuation reuse, copied resource URIs, result summary/full views, and apply-time cache invalidation.

### Changed

- Improved style and format orchestration with style summaries, format diagnostics, precise style-dimension clearing, and grouped style operations that avoid repeated tool calls.
- Updated Open Workbook Skills guidance, packaged CLI instructions, and docs to prefer handle-first follow-ups and fetch full result resources only for explicit audit, raw-value, or all-row requests.
- Aligned the MCP resource catalog, resource registration, and documentation for semantic-index, agent-result, and compact compatibility resources.

## [0.1.17] - 2026-06-19

### Added

- Added compact MCP result resources, continuation metadata, and brief default result text so large workbook answers stay retrievable without being duplicated into model context.
- Added backend aggregate answers for unique/count-style range questions so agents can get compact value counts without reading every cell into context.

### Changed

- Reduced default `excel.agent.run` payloads by compacting schemas, profiles, sparse rows, candidates, and MCP tool text while keeping full detail available through result resources.
- Improved read target resolution so semantic column names can correct conflicting single-column ranges before live reads.
- Tightened broad workbook/range read guards and header detection to avoid accidental full-sheet dumps and false headers from data rows.

## [0.1.16] - 2026-06-19

### Added

- Added stricter office-agent E2E coverage reporting that verifies all 294 internal capabilities, 67 add-in host methods, and 37 operation kinds are represented by production scenarios or unit-contract coverage.
- Added expanded production office-agent fixtures for workbook, worksheet, range, table, formula, template, style, cleaning, repair, backup, snapshot, collaboration, and workflow behaviors.
- Added colocated backend capability registry tests and add-in host registry tests so internal capability groups stay aligned with the one-tool public MCP surface.
- Added package, app, script, docs, and skill READMEs that document the current repository layout and backend-owned operation model.

### Changed

- Refactored backend capability metadata into domain modules while preserving `excel.agent.run` as the only public MCP tool.
- Refactored the Excel add-in host executor into grouped host method modules with a central registry and focused unit tests.
- Refactored MCP server internals into catalog, runtime facade, prompt, resource, result, and agent-run tool modules to make the public surface easier to audit.
- Refactored package internals, app runtime boundaries, and script organization for clearer ownership and less duplication.
- Simplified root documentation and removed stale examples, roadmap content, and personalized agent notes from the committed repo.
- Modernized bundled skill guidance for routing through `excel.agent.run`, multilingual structured intent, capability maps, and common workbook workflows.

## [0.1.15] - 2026-06-19

### Added

- Added explicit `excel.agent.run` route telemetry, operation-risk telemetry, target fingerprint status, and ambiguity candidate hints so backend orchestration decisions are easier to audit without exposing primitive tools.
- Added optional caller-provided `intent.action` hints for `excel.agent.run` so the user's existing LLM/client can parse complex requests into structured agent intent while the backend remains deterministic.
- Added an internal agent action-handler registry and `actionHandlerId` telemetry for preview routing observability.
- Added fine-grained internal Excel capability grouping and catalog-level agent status classification for future orchestration and test planning.
- Added an internal capability coverage matrix and report command for tracking covered, unit-contract, future orchestration, host-limited, and deferred capability work.
- Added skill guidance and tests for multilingual agent requests normalized through structured `excel.agent.run` intent fields.
- Added `excel.agent.run` orchestration for workbook info, open-workbook listing, snapshot refresh/read/change detection, workbook backup restore, local config export/import/embed, embedded config read/import, and workbook close actions.
- Added `excel.agent.run` orchestration for template registry reads, template detection, region inference, sheet/template validation, and preview/apply template registry or repair mutations.
- Added `excel.agent.run` orchestration for template data-region clearing and filling, with filling constrained to explicit region value matrices.
- Added `excel.agent.run` orchestration for style fingerprint reads, style comparisons, template style copy, and style consistency repair while keeping style primitives backend-owned.
- Added structured `excel.agent.run` theme capability reports for workbook theme read/apply requests when the connected Excel runtime cannot safely provide theme APIs.
- Added `excel.agent.run` orchestration for all cleaning capabilities, with read-only inspections in answer mode and mutating transforms behind preview/apply.
- Added `excel.agent.run` orchestration for repair capabilities, including preview/apply style, formula, and table-structure repair plus structured capability reports for host-limited repair categories.
- Added `excel.agent.run` orchestration for the full Range capability group, including compact/metadata reads and preview/apply structural range mutations.
- Added `excel.agent.run` workflow plan reports for combined workbook workflows so agents can route formula-sheet, template-report, pivot/chart, risky-edit, inspect/analyze, and rollback/validate requests without exposing primitive workflow tools.
- Removed unsupported Office.js placeholder capabilities for circular-reference enumeration and deep style/layout copy dimensions, and reclassified implemented file-bridge, PivotTable, and chart capabilities as internal contract-tested coverage instead of host-limited placeholders.
- Added backend contract coverage for runtime, batch, plan, job, task, collaboration, lock, conflict, transaction, event, and permission capabilities, leaving only compact-resource and diff wrappers as unit-contract follow-up work.

### Changed

- Strengthened agent preview/apply protection by storing target-specific fingerprints for pending operations and compacting verbose candidate hints when tight response budgets are requested.
- Routed structured caller intent through the same target-resolution, preview/apply, risk, stale-check, and validation flow as natural-language-only requests.
- Refactored agent preview intent dispatch through registered handlers while preserving existing preview/apply behavior.
- Used caller-provided `intent.targetHints` as a bounded deterministic target-ranking signal and added target-hint telemetry to `excel.agent.run`.
- Refactored catalog reporting so MCP exposes only `excel.agent.run` while the full Excel catalog is reported as internal backend capabilities.
- Removed inactive primitive MCP tool registration groups from the MCP server path and grouped internal Excel capabilities for backend orchestration.
- Removed unimplemented lookup/table placeholder capabilities (`excel.lookup.find_headers`, `excel.lookup.find_tables_by_columns`, `excel.lookup.find_entity`, and `excel.table.preserve_filters`) from the internal catalog so advertised capabilities map to implemented backend logic.
- Slimmed the default E2E gate to the supported one-tool public agent surface and removed the obsolete primitive-tool fake-host sweep.
- Split monolithic backend orchestration and runtime tests into behavior-focused suites with shared test-support fixtures.
- Updated bundled and packaged agent instructions to emphasize the one-tool public surface instead of primitive MCP tool selection.

## [0.1.14] - 2026-06-18

### Added

- Added faster default `excel.agent.run` paths for workbook overviews, active/current sheet prompts, explicit two-sheet comparisons, style previews, formula previews, and template cleanup previews.
- Added `target.candidateId` recovery for `excel.agent.run` so default-surface agents can retry an ambiguous workbook target without switching to primitive tools.
- Added cached schema/header answers for `excel.agent.run` schema-style requests to avoid broad reads when table or range metadata is already available.
- Added live-read routing and table-append previews for `excel.agent.run` so default-surface agents can read actual rows or append table data without using primitive tools.
- Added raw monthly sheet range resolution for `excel.agent.run`, including exact sheet/range parsing and transaction/invoice header-block reads when no Excel Table exists.
- Added a logging-first office-agent behavior harness for normal workbook overview, sheet reading, comparison, simple edit, and formula prompts without making the run a deploy gate.
- Added the production office-agent scenario fixture to the behavior harness, expanding coverage to connection, targeting, table reads/edits, formula safety, token guards, multilingual prompts, save/recalc actions, and multi-step workflows.

### Changed

- Made vague workbook-level `.xlsx` review prompts return a complete lightweight workbook overview before target-specific reads, avoiding cold broad sheet sampling and misleading narrow-sheet summaries.
- Reframed compact read/resource tools as backend-owned internals while keeping normal agent guidance centered on the backend-composed `excel.agent.run` workflow.
- Removed the public advanced MCP surface path; primitive operation exposure is now reserved for internal test/development harnesses while packaged MCP clients see one public tool.
- Reduced normal office-agent behavior call count and payload size by auto-applying small explicit value edits, compacting apply results, and updating the behavior harness to avoid unnecessary prepare/read chains.
- Updated Open Workbook Skills skill guidance so default-surface clients start with `excel.agent.run` `mode: "prepare"` instead of unavailable advanced workflow tools.
- Clarified OpenCode guidance so default-surface clients call `excel.agent.run` instead of runtime/workbook primitives.

## [0.1.13] - 2026-06-16

### Fixed

- Fixed strict MCP clients rejecting `excel.agent.run` structured results when agent telemetry included auto-apply, safety-decision, preview-operation, or validation-status fields not declared in the published output schema.

### Added

- Added agent E2E coverage that checks `structuredContent` mirrors text JSON and emitted telemetry keys are declared by `tools/list.outputSchema`.

## [0.1.12] - 2026-06-16

### Added

- Added the default `excel.agent.run` MCP workflow surface so normal agents can use one compact intent interface while the backend orchestrates workbook discovery, cached metadata, target resolution, preview/apply, validation, rollback, and proof generation internally.
- Added server-side `workbookContextId` reuse for agent prepare/find/answer/update flows, plus context-aware internal compact routes for workbook summary, sheet summary, range compact reads, table compact reads, and compact resource fetches.
- Added natural-language target resolution across sheets, tables, headers, named ranges, regions, summary blocks, and formula regions with structured ambiguity candidates instead of guessing.
- Added safe `auto` mode application for clearly scoped low-risk value edits after preview checks, while formula-sensitive, broad, sparse, structural, destructive, stale, or ambiguous edits stop before mutation.
- Added agent surface and workflow E2E coverage for default tool exposure, metadata cache reuse, compact target reads, preview/apply confirmation, safe auto-apply, formula-sensitive blocking, and context-aware compact routes.

### Changed

- Default MCP `tools/list` now exposes only `excel.agent.run`; primitive operation exposure is no longer documented as a public MCP client mode.
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
- Added lookup-first token-saving tools: `excel.lookup.search_workbook`, `excel.lookup.resolve_range`, and `excel.lookup.inspect_match`.
- Added compact read telemetry with `payloadBytes`, rough `estimatedTokens`, truncation status, and continuation metadata so agents can avoid sending broad workbook payloads to models by default.
- Added local compact detail resources, compact cache lifecycle internals, `excel.validate.compact`, and additive mutation `compactProof` metadata so large details can be expanded only when needed.
- Added compact snapshot and diff access through `excel.snapshot.get_compact` and `excel.snapshot.compare_compact`.

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

- Split Codex agent E2E into a blocking core safety lane and report-only workflow quality lane, with bundled `open-workbook-skills` skill guidance loaded into agent prompts.
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
- Packaged agent instruction skill under `skills/open-workbook-skills` and bundled CLI fallback instructions.
- Documentation for installation, MCP clients, architecture, runtime, sideloading, packaging, safety contracts, workbook lifecycle, performance, and production readiness.
- Release validation scripts for builds, package metadata, MCP catalog consistency, documentation surface, skill assets, CLI smoke checks, package dry runs, tests, and synthetic benchmarks.

### Notes

- This release prepares the repository and packages for open-source distribution. npm publishing is intentionally deferred.
- Open Workbook is not a Microsoft AppSource add-in; Excel manifest trust is handled through user or administrator sideloading.
