# Performance Contract

Open Workbook treats speed as a correctness requirement because slow workbook automation encourages unsafe shortcuts.

## Rules

- All writes route through the batch compiler.
- Individual write tools are syntactic sugar over batch operations.
- Do not call `context.sync()` inside per-cell loops.
- Use 2D array assignments for values and formulas.
- Load only requested Office.js properties.
- Group operations by workbook, sheet, and contiguous range.
- Chunk large payloads before Office.js limits.
- Suspend calculation and screen updating only around large operations.

## Telemetry

Every read/write result records:

- duration in milliseconds
- sync count
- payload bytes
- cells read or written
- range count
- chunk count
- engine name and version
- warning count

Initial releases collect telemetry first, then set hard SLOs from real workbook tests.
