# Validation and Repair

Open Workbook exposes validation and repair capabilities so agents can check workbook integrity before and after writes, then recover safely when template-backed repair is available. Normal MCP clients request these through `excel.agent.run`; primitive names below describe backend-owned capability groups.

## Validation Reports

All `excel.validate.*` tools return a `ValidationReport`:

- `ok`: false when any issue has `severity: "error"`.
- `summary`: counts errors, warnings, and info messages.
- `issues`: categorized findings with stable codes and optional targets.
- `data`: supporting context such as workbook maps, table metadata, diffs, or template validation output.

Implemented checks:

- Workbook and sheet checks read workbook maps and scan used ranges for formula errors.
- Formula checks use `range.find_errors` over the workbook, sheet, or explicit range.
- Broken-reference checks search for `#REF!` over the workbook, sheet, or explicit range.
- Template consistency checks compare a target sheet against a registered template fingerprint.
- Style checks capture a sheet fingerprint or filter template consistency issues down to style/template categories.
- Table and filter checks return current structured-table metadata for agent review.
- Unintended-change checks compare two snapshots or detect current changes since a snapshot.
- Print-layout validation currently reports Office.js capability limits unless template fingerprint comparison is requested.

## Repair Reports

All `excel.repair.*` tools return a `RepairReport`:

- `ok`: true only when an implemented repair path ran successfully.
- `backups`: backup IDs created before mutation.
- `validation`: post-repair validation when available.
- `error`: capability or execution error when the repair cannot run.

Implemented repair paths:

- `excel.repair.style_from_template` uses registered template styles and creates a rollback backup.
- `excel.repair.formulas_from_template` uses registered template formulas and creates a rollback backup.
- `excel.repair.table_structure` copies a table structure to a target range through the table mutation lifecycle.

Capability-status repair paths:

- `excel.repair.filters_from_template`
- `excel.repair.print_layout`
- `excel.repair.named_ranges`
- `excel.repair.formula_errors`
- `excel.repair.merged_cells`

These return `CAPABILITY_UNAVAILABLE` until the add-in has safe, deterministic Office.js or host-bridge implementations for those repairs.

## Agent Flow

Recommended lifecycle for risky workbook changes through `excel.agent.run`:

1. Ask for `mode: "validate"` or a scoped preview before risky edits.
2. Let the backend create snapshots/backups during preview or apply.
3. Apply only with the returned confirmation token.
4. Run `mode: "validate"` for unintended changes, formula errors, and template-specific checks.
5. Use `mode: "rollback"` or template-backed repair guidance if validation fails.
