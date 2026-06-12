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
- Runtime capability checks by Office API set and platform.
- HTTPS/dev-certificate option for add-in serving.
- Larger workbook performance profiling and true large-range chunk execution.
- Persistent local stores for templates, regions, permissions, plans, and backup indexes.
- Native installer or service wrapper for auto-starting the add-in asset server.
- Deeper chart and PivotTable template copying where Office.js exposes deterministic APIs.
- Workbook custom XML or local config integration for portable template/region metadata.

## Explicit Non-Goals For Now

- Microsoft AppSource submission.
- Macro execution.
- Cloud storage of workbook data by default.
- OCR and vertical reconciliation workflows.
