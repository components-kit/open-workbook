# Template System

Registered templates are authoritative for template workflows.

## Registry

Templates live in two places:

- **Workbook metadata** for workbook-specific structure and portability.
- **Local registry** for reusable team templates across workbooks.

Workbook templates win over local templates with the same ID unless the caller explicitly selects a local template version.

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
