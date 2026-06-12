# Backup Lifecycle

Backup is a core runtime invariant, not an optional feature.

## Backup Tiers

- **Region snapshot**: fast rollback for changed values, formulas, formats, comments, notes, validation, filters, table data, row heights, and column widths.
- **Sheet snapshot**: rollback for sheet create, copy, rename, move, hide, and clear workflows.
- **Workbook-copy backup**: disaster recovery for structural changes, broad clears, save-as, restore, or operations that risk workbook integrity.

## Mutation Lifecycle

Every write operation creates an operation record with:

- operation ID
- workbook ID
- base snapshot ID
- target ranges
- before fingerprints
- backup references
- compiled batch
- validation result
- diff summary
- rollback status
- telemetry

## Conflict Policy

Plans use a target-region strict policy:

- Stop if workbook structure changed.
- Stop if a targeted sheet, table, or range fingerprint changed.
- Continue with a warning when unrelated regions changed.
- Revalidate template and formula invariants after apply.
- Auto-rollback if post-apply validation fails and rollback data is available.

## Retention

Region snapshots are retained for the active session by default. Workbook-copy backups persist until explicit cleanup or configured retention removes them. Normal rollback must not delete backups.

## Persistent Snapshot Backups

Workbook-copy backups are persisted as JSON snapshot payloads under `OPEN_WORKBOOK_BACKUP_DIR` when set, otherwise `.open-workbook/backups` under the MCP process working directory.

These backups are restorable through `excel.workbook.restore_backup` because they contain captured range values, formulas, number formats, display text, and basic styles. They are not full `.xlsx` file copies. True file-copy backup/export needs a native host bridge because Office.js does not expose local file-system save-as/export APIs.

## Implemented Rollback

The current rollback path supports:

- restoring values, formulas, number formats, and basic style metadata from captured range snapshots
- deleting sheets created by `template.create_sheet_from_template`
- applying rollback through the same batch and telemetry path as normal writes

Full workbook-copy restore remains a disaster-recovery layer because replacing an open workbook file usually requires user or OS-level coordination.
