# Formula Intelligence

Open Workbook treats formulas as reusable patterns, not only as raw cell text.

## Pattern Model

`excel.formula.read_patterns` reads a sheet or range and returns:

- A1 formulas.
- R1C1 formulas when Excel exposes them.
- A pattern hash per formula cell.
- Pattern groups with counts and relative cell positions.

R1C1 patterns let agents validate that copied accounting/report formulas keep the same relative structure after they move to a new period, sheet, row, or column.

## Mutations

- `excel.formula.copy_patterns` copies formula patterns from a source range or template sheet to a target range and validates the result.
- `excel.formula.fill_down` and `excel.formula.fill_right` fill target ranges with R1C1 formula semantics.
- `excel.formula.repair_patterns` repairs formulas from a registered template.
- `excel.formula.convert_to_values` replaces formulas with their calculated values after creating a rollback backup.
- `excel.formula.recalculate` runs workbook recalculation.

All formula mutations validate permissions and create backups before Excel receives writes.

## Validation

- `excel.formula.validate` and `excel.formula.find_errors` scan workbook, sheet, or range targets for formula error cells.
- `excel.formula.validate_against_template` compares the target sheet against a registered template fingerprint.
- Formula pattern comparisons report missing formulas, unexpected formulas, shape mismatches, and pattern mismatches.

## Dependency Graph

- `excel.formula.get_dependency_graph` parses formulas in a sheet or range and returns precedent edges.
- `excel.formula.trace_precedents` returns parsed precedent ranges for a formula cell.
- `excel.formula.trace_dependents` scans formulas on the sheet and returns formula cells that reference the target range.

The parser handles normal A1 references, quoted sheet names, absolute references, rectangular ranges, whole-column references, dynamic array spill anchors, structured table references, and external workbook references. Structured references are represented as table dependency nodes and are also resolved to precise local ranges when table metadata is available, including data body, headers, totals, `#All`, and bounded column spans. Dynamic array spill references expand to the spill range when spill metadata is available and fall back to the anchor cell with a warning when it is not. External workbook references are represented as external dependency nodes and are not resolved to local workbook ranges.

Formula writes use parsed local dependencies during pre-commit lock checks. If an agent writes `=SUM(A1:A10)` while another agent has a write lock on `A1:A10`, the formula write is blocked before Excel receives the mutation.

## Current Limits

Circular-reference enumeration is not advertised as an internal capability until Excel host support can be normalized safely.

`excel.formula.explain` is intentionally lightweight. It summarizes functions, references, structured references, external references, and volatile functions; it is not a full Excel formula parser.
