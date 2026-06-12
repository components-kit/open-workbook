# Workbook File Lifecycle

Open Workbook separates workbook state safety from local file-system control.

## Supported Now

- `excel.workbook.save` saves the active workbook through Office.js.
- `excel.workbook.close` closes the workbook through Office.js with `Save` or `SkipSave`.
- `excel.workbook.save_as` uses the configured native file bridge when `OPEN_WORKBOOK_FILE_BRIDGE_URL` is set; otherwise it reports capability status.
- `excel.workbook.create_backup` creates a persistent JSON snapshot backup.
- `excel.workbook.restore_backup` restores a persistent or in-memory snapshot backup.
- `excel.workbook.export_copy` creates a persistent snapshot backup, then writes a true `.xlsx` file. If `OPEN_WORKBOOK_FILE_BRIDGE_URL` is configured, the built-in native bridge uses Excel desktop `SaveCopyAs` automation first. If no bridge is configured or the bridge fails, Open Workbook falls back to Office.js compressed workbook slices on supported Excel desktop hosts. If no `targetPath` is provided, the file is written under `.open-workbook/exports` or `OPEN_WORKBOOK_EXPORT_DIR`.
- `excel.workbook.export_local_config` exports versioned JSON metadata for Open Workbook templates, regions, and optional permissions.
- `excel.workbook.import_local_config` imports that JSON into the local daemon registry.
- `excel.workbook.embed_local_config` stores local config JSON in a namespaced workbook custom XML part when the Excel host supports it.
- `excel.workbook.read_embedded_local_config` reads embedded local config from workbook custom XML.
- `excel.workbook.import_embedded_local_config` imports embedded local config into the local daemon registry.

## Not Supported By Office.js

Office.js does not expose a local file path API for `save_as`. That operation needs a native host bridge or user-driven Excel UI.

Office.js does expose compressed document file slices on Excel desktop hosts. Open Workbook uses that path for `export_copy` by asking the add-in for the workbook bytes and writing them from the local backend. Excel on the web does not expose compressed Excel workbook export through this API.

The MCP tools return explicit `CAPABILITY_UNAVAILABLE` errors instead of silently pretending a file was written when no bridge is configured.

## Native File Bridge Contract

Run `owb file-bridge start` to start Open Workbook's built-in native bridge. It listens on `http://127.0.0.1:37847` by default. Set `OPEN_WORKBOOK_FILE_BRIDGE_URL=http://127.0.0.1:37847` for the backend daemon so `excel.workbook.save_as` and bridge-first `excel.workbook.export_copy` can use it. If you override the route with `OPEN_WORKBOOK_FILE_BRIDGE_PATH`, set the same value for the backend daemon and file bridge process.

The bridge exposes:

- `GET /status`
- `POST /shutdown`
- `POST /v1/workbook-file`

The backend sends `POST /v1/workbook-file` with JSON:

```json
{
  "operation": "workbook.save_as",
  "workbookId": "workbook_...",
  "targetPath": "/path/to/report.xlsx"
}
```

Supported operations are `workbook.save_as`, `workbook.export_copy`, and the reserved `workbook.restore_file_backup`. `export_copy` includes the snapshot `sourceBackupId` so the helper can audit which safety backup was captured before writing a true file copy. If the bridge is not configured or fails, Open Workbook falls back to the add-in compressed-file export path where supported.

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

The built-in bridge uses AppleScript on macOS and PowerShell COM automation on Windows for `workbook.save_as` and `workbook.export_copy`. `export_copy` calls Excel `SaveCopyAs`, which saves a copy of the workbook without changing the open workbook's file identity. It matches `workbookId` against the open Excel workbook name or full path. Linux returns an explicit unsupported result.

Set `OPEN_WORKBOOK_FILE_BRIDGE_TIMEOUT_MS` to override the default 30000 ms backend-to-bridge timeout. Set `OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS` to a path-delimited allowlist of output directories for native Save As and Export Copy. Set `OPEN_WORKBOOK_EXPORT_DIR` to control the default output directory for add-in compressed-file exports.

`excel.runtime.get_status` and `excel.runtime.get_capabilities` include `fileBridge` status, including the configured URL and route path, so agents can check whether native Save As or Export Copy is configured before requesting it.

Local config export/import is different from workbook file export. It does not create or modify an `.xlsx`; it moves Open Workbook registry metadata so teams can version templates, semantic regions, and permission defaults alongside a project. Embedded local config modifies workbook metadata through Office.js custom XML parts and is guarded by workbook-level permissions.

## Backup Directory

Set `OPEN_WORKBOOK_BACKUP_DIR` to control where persistent snapshot backups are written. By default, backups are stored in `.open-workbook/backups` under the MCP process working directory.
