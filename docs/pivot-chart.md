# PivotTables and Charts

Open Workbook exposes native Office.js PivotTable and chart tools for reporting workflows.

## PivotTables

Supported tools:

- `excel.pivot.list`: list PivotTables in the workbook.
- `excel.pivot.get_info`: return PivotTable metadata and source details where Office.js exposes them.
- `excel.pivot.create`: create a PivotTable from a source range or structured table at a destination cell, optionally applying row fields, column fields, filter fields, data fields, aggregation, number formats, and layout flags.
- `excel.pivot.refresh`: refresh one PivotTable.
- `excel.pivot.refresh_all`: refresh all PivotTables in the workbook.
- `excel.pivot.copy_from_template`: replay deterministic PivotTable settings from a template PivotTable to a target PivotTable.
- `excel.pivot.delete`: delete a PivotTable after capturing a transaction-backed backup of the reported output range when available.
- `excel.pivot.validate_source`: check whether source, source type, output range, and data-field metadata are available.
- `excel.pivot.get_capability_matrix`: report supported, partial, unsupported, and unknown PivotTable dimensions for the active host.
- `excel.pivot.get_fingerprint`: hash deterministic PivotTable source/layout/output metadata.
- `excel.pivot.compare_fingerprint`: compare two PivotTable fingerprints.
- `excel.pivot.diff`: return a review-friendly PivotTable diff.
- `excel.pivot.repair_from_template`: replay a template PivotTable and return before/after comparison.
- `excel.pivot.rebuild_with_source`: create a new PivotTable from a desired source and optionally replay a template.

Capability-status tools:

- `excel.pivot.update_source`: Office.js does not expose a safe in-place source reassignment path in this runtime. Create a new PivotTable from the desired source.

`excel.pivot.update_source`, `excel.pivot.copy_from_template`, `excel.pivot.repair_from_template`, and `excel.pivot.rebuild_with_source` return machine-readable `capabilityStatus` metadata plus `warnings` when the operation is intentionally partial. Agents should inspect these fields before claiming a PivotTable is fully identical to a template. In-place source reassignment returns `CAPABILITY_UNAVAILABLE` with fallback `excel.pivot.rebuild_with_source`.

`excel.pivot.get_info` returns source metadata, the PivotTable output range, layout settings, row/column/filter/data hierarchy summaries, field settings, and source hierarchy names where Office.js exposes them.

`excel.pivot.create` accepts optional `rowFields`, `columnFields`, `filterFields`, `dataFields`, `layout`, and `refresh`. This lets agents create a usable summary PivotTable in one transaction instead of creating a blank PivotTable and requiring a second mutation. Field names must match the source table/range fields exposed by Excel.

`excel.pivot.copy_from_template` requires `templatePivotTableName`. Before mutating, it compares the template pivot's required row, column, filter, and data source fields with the target pivot source fields when Office.js exposes both sides. If required fields are missing, it returns `TEMPLATE_MISMATCH` with `PIVOT_TEMPLATE_SOURCE_FIELD_MISSING` issues before creating a backup or changing Excel. Compatible copies create a backup and transaction record, then replay settable PivotTable options, layout flags, axis membership/order, data hierarchy aggregation and number formats, and basic field settings. Agents can restrict replay with `dimensions`, using any of `metadata`, `layout`, `fields`, `dataFields`, `numberFormats`, `filters`, and `refresh`. When Office.js reports the target PivotTable range, the backup is scoped to that range instead of the whole sheet. The response includes `PIVOT_TEMPLATE_COPY_PARTIAL` because source reassignment, PivotChart-specific settings, slicers/timelines, item-level manual filters/sorts, grouping details, calculated fields/items, and host-specific settings are replayed only when Office.js exposes deterministic read/write APIs.

`excel.pivot.validate_source` returns a summary plus structured issues such as `PIVOT_SOURCE_UNAVAILABLE`, `PIVOT_OUTPUT_RANGE_UNAVAILABLE`, and `PIVOT_HAS_NO_DATA_FIELDS`, so agents can distinguish a missing pivot from a partially configured pivot. It also accepts optional `expectedFields`, `expectedRowFields`, `expectedColumnFields`, `expectedFilterFields`, `expectedDataFields`, `expectedDataFieldSettings`, and `expectedLayout`. When Office.js exposes source hierarchy metadata, missing fields return `PIVOT_EXPECTED_FIELD_MISSING`; fields that exist but are not currently on the expected axis return `PIVOT_EXPECTED_LAYOUT_MISMATCH`. Aggregation, number-format, and layout mismatches return `PIVOT_EXPECTED_AGGREGATION_MISMATCH`, `PIVOT_EXPECTED_NUMBER_FORMAT_MISMATCH`, and `PIVOT_EXPECTED_LAYOUT_SETTING_MISMATCH`. If source field metadata is unavailable, validation returns `PIVOT_SOURCE_FIELDS_UNAVAILABLE` as a warning instead of guessing.

`excel.pivot.get_fingerprint` captures source type/value, source fields, row/column/filter/data fields, aggregation, number formats, layout flags, and output shape where Office.js exposes those dimensions. Missing dimensions become warnings, not fabricated values. `excel.pivot.compare_fingerprint` and `excel.pivot.diff` compare those captured dimensions so agents can review template drift before mutation.

`excel.pivot.repair_from_template` runs fingerprint comparison, applies deterministic template copy, and compares again. Strict mode blocks when required fingerprint dimensions are unavailable. Successful repairs still return the same capability warnings as template copy when deep template dimensions cannot be proven or replayed. `excel.pivot.rebuild_with_source` is the safe source-change strategy: create a new PivotTable from the desired source and optionally replay template settings. With `replaceExisting`, it performs an explicit delete/create flow with separate backups and transactions instead of trying unsafe in-place source reassignment, and returns `PIVOT_REBUILD_NOT_IN_PLACE` so agents can explain the lifecycle accurately.

## Charts

Supported tools:

- `excel.chart.list`: list charts across worksheets.
- `excel.chart.get_info`: return chart metadata.
- `excel.chart.create`: create a chart from a source range.
- `excel.chart.update_data_source`: reset a chart's source range.
- `excel.chart.copy_from_template`: copy deterministic chart metadata from a template chart to a target chart.
- `excel.chart.refresh`: return current chart metadata. Excel updates chart visuals from source data.
- `excel.chart.delete`: delete a chart.
- `excel.chart.validate_against_template`: verify target and optional template chart metadata availability.

`excel.chart.create` supports chart type, source address, series orientation, title, position, and style.

`excel.chart.copy_from_template` currently copies chart type, style, title text, and geometry. It creates a backup and transaction record before mutating the target chart. Deep series formatting and PivotChart-specific controls remain capability-limited until Office.js exposes a deterministic replay path.

## Safety

Pivot/chart creation and source changes are structure-level operations. Pivot/chart template copy is a format/layout operation. These tools respect permission policy, sheet/region scope, and locked-region checks before calling Excel. The backend creates a range snapshot backup for known source and destination or used ranges before mutating where a range target is available.

## Current Limits

Direct PivotTable source reassignment, PivotChart-specific controls, unsupported host-specific PivotTable settings, and deep series-level chart fingerprints are future milestones.
