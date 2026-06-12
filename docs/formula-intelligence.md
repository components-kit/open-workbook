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

## Current Limits

`excel.formula.find_circular_references`, `excel.formula.trace_precedents`, and `excel.formula.trace_dependents` are exposed as deterministic capability-status tools until dependency graph results can be normalized safely across Excel hosts.

`excel.formula.explain` is intentionally lightweight. It summarizes functions, references, structured references, external references, and volatile functions; it is not a full Excel formula parser.
