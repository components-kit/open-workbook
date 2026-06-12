# Workbook File Lifecycle

Open Workbook separates workbook state safety from local file-system control.

## Supported Now

- `excel.workbook.save` saves the active workbook through Office.js.
- `excel.workbook.close` closes the workbook through Office.js with `Save` or `SkipSave`.
- `excel.workbook.create_backup` creates a persistent JSON snapshot backup.
- `excel.workbook.restore_backup` restores a persistent or in-memory snapshot backup.
- `excel.workbook.export_copy` creates a persistent snapshot backup and reports that true `.xlsx` export is unavailable.

## Not Supported By Office.js

Office.js does not expose a local file path API for `save_as` or exporting a workbook copy as `.xlsx`. Those operations need a future native host bridge or user-driven Excel UI.

The MCP tools return explicit `CAPABILITY_UNAVAILABLE` errors instead of silently pretending a file was written.

## Backup Directory

Set `OPEN_WORKBOOK_BACKUP_DIR` to control where persistent snapshot backups are written. By default, backups are stored in `.open-workbook/backups` under the MCP process working directory.
