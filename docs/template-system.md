# Template System

Registered templates are authoritative for template workflows.

## Registry

Templates live in two places:

- **Workbook metadata** for workbook-specific structure and portability.
- **Local registry** for reusable team templates across workbooks.

Workbook templates win over local templates with the same ID unless the caller explicitly selects a local template version.

The runtime stores both hashes and the captured fingerprint payload. Hashes are used for fast comparisons; payloads are retained so validation and repair can explain which component differs.

Template metadata can be moved between environments with workbook local config:

- `excel.workbook.export_local_config` exports registered templates, regions, and optional permission metadata as versioned JSON.
- `excel.workbook.import_local_config` imports the JSON into another local daemon registry, either merging or overwriting matching ids.
- `excel.workbook.embed_local_config` writes that JSON into the workbook's custom XML part when the connected Excel host supports the Office.js custom XML API.
- `excel.workbook.import_embedded_local_config` reads embedded workbook metadata and imports it into the local daemon registry.

Local config does not modify workbook cells or save the `.xlsx`; it only moves Open Workbook's registry layer. Embedded local config modifies workbook metadata, not cells, and is guarded by workbook-level permissions and the transaction queue.

## Template Fingerprint

A template captures:

- sheet structure and named regions
- headers and merged cells
- row heights and column widths
- formulas and formula patterns
- number formats, fills, fonts, borders, and alignment
- tables, filters, sorts, conditional formatting, and data validation
- freeze panes, print settings, page layout, and hidden rows or columns

## Template Workflows

Creating a new accounting or reporting sheet must:

1. Copy layout and styles from the registered template.
2. Copy formula patterns separately from values.
3. Clear only declared data-entry regions.
4. Preserve filters, table structure, print layout, freeze panes, and named ranges.
5. Validate against the template before returning success.

If no registered template exists, the runtime may infer candidates but mutating template workflows must warn and require confirmation.

## Implemented Flow

Current stable MCP support covers:

- detecting candidate template sheets
- registering, listing, getting, and unregistering templates
- creating a sheet from a registered template
- clearing declared data regions while preserving formats
- filling target regions while preserving existing formats
- validating a target sheet against template structure, formulas, styles, filters, tables, and print-layout fingerprints
- repairing target sheet styles, formulas, declared data regions, or full layout from the source template sheet

Repair creates a backup before it mutates a target sheet. Validation ignores the sheet name when comparing structure, so a new period sheet can be named differently from the template without failing solely on identity.
