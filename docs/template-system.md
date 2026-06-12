# Template System

Registered templates are authoritative for template workflows.

## Registry

Templates live in two places:

- **Workbook metadata** for workbook-specific structure and portability.
- **Local registry** for reusable team templates across workbooks.

Workbook templates win over local templates with the same ID unless the caller explicitly selects a local template version.

The runtime stores both hashes and the captured fingerprint payload. Hashes are used for fast comparisons; payloads are retained so validation and repair can explain which component differs.

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
