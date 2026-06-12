# Workbook File Lifecycle

Open Workbook separates workbook state safety from local file-system control.

## Supported Now

- `excel.workbook.save` saves the active workbook through Office.js.
- `excel.workbook.close` closes the workbook through Office.js with `Save` or `SkipSave`.
- `excel.workbook.create_backup` creates a persistent JSON snapshot backup.
- `excel.workbook.restore_backup` restores a persistent or in-memory snapshot backup.
- `excel.workbook.export_copy` creates a persistent snapshot backup and reports that true `.xlsx` export is unavailable.
- `excel.workbook.export_local_config` exports versioned JSON metadata for Open Workbook templates, regions, and optional permissions.
- `excel.workbook.import_local_config` imports that JSON into the local daemon registry.
- `excel.workbook.embed_local_config` stores local config JSON in a namespaced workbook custom XML part when the Excel host supports it.
- `excel.workbook.read_embedded_local_config` reads embedded local config from workbook custom XML.
- `excel.workbook.import_embedded_local_config` imports embedded local config into the local daemon registry.

## Not Supported By Office.js

Office.js does not expose a local file path API for `save_as` or exporting a workbook copy as `.xlsx`. Those operations need a future native host bridge or user-driven Excel UI.

The MCP tools return explicit `CAPABILITY_UNAVAILABLE` errors instead of silently pretending a file was written.

Local config export/import is different from workbook file export. It does not create or modify an `.xlsx`; it moves Open Workbook registry metadata so teams can version templates, semantic regions, and permission defaults alongside a project. Embedded local config modifies workbook metadata through Office.js custom XML parts and is guarded by workbook-level permissions.

## Backup Directory

Set `OPEN_WORKBOOK_BACKUP_DIR` to control where persistent snapshot backups are written. By default, backups are stored in `.open-workbook/backups` under the MCP process working directory.
