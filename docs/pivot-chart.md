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
- `excel.pivot.validate_source`: check whether source metadata is available.

Capability-status tools:

- `excel.pivot.update_source`: Office.js does not expose a safe in-place source reassignment path in this runtime. Create a new PivotTable from the desired source.

`excel.pivot.get_info` returns source metadata, the PivotTable output range, layout settings, row/column/filter/data hierarchy summaries, field settings, and source hierarchy names where Office.js exposes them.

`excel.pivot.create` accepts optional `rowFields`, `columnFields`, `filterFields`, `dataFields`, `layout`, and `refresh`. This lets agents create a usable summary PivotTable in one transaction instead of creating a blank PivotTable and requiring a second mutation. Field names must match the source table/range fields exposed by Excel.

`excel.pivot.copy_from_template` requires `templatePivotTableName`. It creates a backup and transaction record, then replays settable PivotTable options, layout flags, axis membership/order, data hierarchy aggregation and number formats, and basic field settings when the target pivot exposes matching source field names. When Office.js reports the target PivotTable range, the backup is scoped to that range instead of the whole sheet. It intentionally does not claim source reassignment or PivotChart-specific styling.

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
