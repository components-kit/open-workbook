# Changelog

All notable changes to Open Workbook will be documented in this file.

The format is based on Keep a Changelog, and this project follows semantic versioning for published packages.

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
