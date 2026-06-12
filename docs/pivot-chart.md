# PivotTables and Charts

Open Workbook exposes native Office.js PivotTable and chart tools for reporting workflows.

## PivotTables

Supported tools:

- `excel.pivot.list`: list PivotTables in the workbook.
- `excel.pivot.get_info`: return PivotTable metadata and source details where Office.js exposes them.
- `excel.pivot.create`: create a PivotTable from a source range or structured table at a destination cell.
- `excel.pivot.refresh`: refresh one PivotTable.
- `excel.pivot.refresh_all`: refresh all PivotTables in the workbook.
- `excel.pivot.validate_source`: check whether source metadata is available.

Capability-status tools:

- `excel.pivot.update_source`: Office.js does not expose a safe in-place source reassignment path in this runtime. Create a new PivotTable from the desired source.
- `excel.pivot.copy_from_template`: deep field-layout and style replay is not implemented yet.

## Charts

Supported tools:

- `excel.chart.list`: list charts across worksheets.
- `excel.chart.get_info`: return chart metadata.
- `excel.chart.create`: create a chart from a source range.
- `excel.chart.update_data_source`: reset a chart's source range.
- `excel.chart.refresh`: return current chart metadata. Excel updates chart visuals from source data.
- `excel.chart.delete`: delete a chart.
- `excel.chart.validate_against_template`: verify target and optional template chart metadata availability.

`excel.chart.create` supports chart type, source address, series orientation, title, position, and style.

## Safety

Pivot/chart creation and source changes are structure-level operations. They respect permission policy, sheet/region scope, and locked-region checks before calling Excel. The backend creates a range snapshot backup for known source and destination ranges before mutating where a range target is available.

## Current Limits

Pivot field layout authoring, PivotTable source reassignment, PivotChart-specific controls, and deep chart style fingerprints are future milestones.
