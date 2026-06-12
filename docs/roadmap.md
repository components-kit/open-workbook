# Roadmap

## Milestone 1: Local Runtime Foundation

- Monorepo scaffold and contributor docs.
- Shared protocol contracts.
- Backend session registry.
- MCP tool/resource catalog.
- Add-in WebSocket connection.
- Core backup, template, batch, fingerprint, and plan primitives.

## Milestone 2: Live Excel Execution

- Bind add-in requests to `Excel.run`. Initial implementation is in place.
- Implement range read/write operations with batched Office.js calls. Initial values, formulas, read-full, clear-values, and template-copy paths are in place.
- Capture real region snapshots. Initial value/formula/format/style snapshots are in place.
- Add target-region conflict checks. Initial strict target fingerprint checks are in place.
- Return real telemetry and diff summaries.
  Initial sync count, range count, chunk count, cells touched, warnings, and duration are returned.

## Milestone 3: Template Fidelity

- Register workbook and local templates.
- Create sheets from templates. Initial sheet copy and declared data-region clearing are in place.
- Validate formulas, styles, filters, tables, and print layout.
- Repair style/formula drift from templates.

## Milestone 4: Hardening

- Full MCP SDK transport.
- Windows and macOS manual test matrix.
- Backup retention controls.
- Capability matrix by Excel API set and platform.
- Performance budgets from real workbooks.
