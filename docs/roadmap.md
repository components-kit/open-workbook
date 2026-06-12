# Roadmap

Open Workbook is ready for local source installs, packaged CLI testing, MCP integration, and sideloaded desktop Excel development. The roadmap below tracks what is implemented and what still needs hardening before broader production use.

## Implemented Foundations

- Monorepo package structure and publish metadata.
- Shared protocol contracts and tool catalog.
- MCP stdio server.
- Local backend broker.
- Excel Office.js add-in runtime.
- CLI for running MCP, serving the add-in, generating manifests, sideloading, diagnostics, and OpenCode config.
- Backup, snapshot, plan, permission, template, diff, and rollback primitives.
- Runtime capability reporting by connected Office host, platform, Office version, and supported `ExcelApi` versions.
- HTTPS add-in asset serving with user-provided local certificate and manifest URL generation.
- Large matrix write chunking for values, formulas, and number formats, with chunk telemetry.
- Persistent local state for templates, regions, permissions, plans, and backup indexes.
- Workbook local config export/import for portable templates, regions, and permission metadata.
- Workbook custom XML embedding/import for portable templates, regions, and permission metadata when the connected Excel host exposes the Office.js API.
- Service wrapper manifest generation for add-in and daemon auto-start on macOS, Linux systemd user services, and Windows Task Scheduler.
- Deterministic chart template metadata replay for chart type, style, title, and geometry.

## Implemented Excel Surfaces

- Runtime status, active workbook context, selection, active sheet switching.
- Workbook map, snapshots, external-change detection, calculate, save, backup, restore, export fallback, close.
- Sheet create/copy/rename/delete/hide/protect/clear/tab color operations.
- Range reads, writes, formulas, formats, clears, copy/move, rows/columns, autofit, merge/unmerge, search, blanks, errors.
- Batch validate, dry run, apply, result tracking.
- Template register, infer, create sheet, previous-period sheet, clear/fill regions, validate, repair.
- Style fingerprints, comparisons, repair, and granular style copy.
- Formula pattern reads, copy, fill, validation, repair, convert to values, recalculate, explain.
- Tables, filters, sorts, names, regions, validation, repair, cleaning, PivotTables, charts, events, snapshots, permissions.

## Hardening Before Broader Production

- Real Excel E2E matrix on macOS and Windows.
- Larger workbook performance profiling and SLO calibration from real workbook traces.
- Deeper PivotTable field-layout/template copying where Office.js exposes deterministic APIs.

## Explicit Non-Goals For Now

- Microsoft AppSource submission.
- Macro execution.
- Cloud storage of workbook data by default.
- OCR and vertical reconciliation workflows.
