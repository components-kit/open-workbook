# Workbook File Lifecycle

Open Workbook separates workbook state safety from local file-system control.

## Supported Now

- `excel.workbook.save` saves the active workbook through Office.js.
- `excel.workbook.close` closes the workbook through Office.js with `Save` or `SkipSave`.
- `excel.workbook.save_as` uses the configured native file bridge when `OPEN_WORKBOOK_FILE_BRIDGE_URL` is set; otherwise it reports capability status.
- `excel.workbook.create_backup` creates a persistent JSON snapshot backup.
- `excel.workbook.restore_backup` restores a persistent or in-memory snapshot backup.
- `excel.workbook.export_copy` creates a persistent snapshot backup, then uses the configured native file bridge for true `.xlsx` export when available.
- `excel.workbook.export_local_config` exports versioned JSON metadata for Open Workbook templates, regions, and optional permissions.
- `excel.workbook.import_local_config` imports that JSON into the local daemon registry.
- `excel.workbook.embed_local_config` stores local config JSON in a namespaced workbook custom XML part when the Excel host supports it.
- `excel.workbook.read_embedded_local_config` reads embedded local config from workbook custom XML.
- `excel.workbook.import_embedded_local_config` imports embedded local config into the local daemon registry.

## Not Supported By Office.js

Office.js does not expose a local file path API for `save_as` or exporting a workbook copy as `.xlsx`. Those operations need a native host bridge or user-driven Excel UI.

The MCP tools return explicit `CAPABILITY_UNAVAILABLE` errors instead of silently pretending a file was written when no bridge is configured.

## Native File Bridge Contract

Set `OPEN_WORKBOOK_FILE_BRIDGE_URL` to a local helper base URL. The backend sends `POST /v1/workbook-file` with JSON:

```json
{
  "operation": "workbook.save_as",
  "workbookId": "workbook_...",
  "targetPath": "/path/to/report.xlsx"
}
```

Supported operations are `workbook.save_as`, `workbook.export_copy`, and the reserved `workbook.restore_file_backup`. `export_copy` includes the snapshot `sourceBackupId` so the helper can audit which safety backup was captured before writing a true file copy.

The helper should return:

```json
{
  "ok": true,
  "operation": "workbook.save_as",
  "workbookId": "workbook_...",
  "targetPath": "/path/to/report.xlsx",
  "filePath": "/path/to/report.xlsx"
}
```

Set `OPEN_WORKBOOK_FILE_BRIDGE_TIMEOUT_MS` to override the default 30000 ms bridge timeout.

Local config export/import is different from workbook file export. It does not create or modify an `.xlsx`; it moves Open Workbook registry metadata so teams can version templates, semantic regions, and permission defaults alongside a project. Embedded local config modifies workbook metadata through Office.js custom XML parts and is guarded by workbook-level permissions.

## Backup Directory

Set `OPEN_WORKBOOK_BACKUP_DIR` to control where persistent snapshot backups are written. By default, backups are stored in `.open-workbook/backups` under the MCP process working directory.
