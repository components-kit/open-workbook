# Workbook File Lifecycle

Open Workbook separates workbook state safety from local file-system control.

## Supported Now

- `excel.workbook.save` saves the active workbook through Office.js.
- `excel.workbook.close` closes the workbook through Office.js with `Save` or `SkipSave`.
- `excel.workbook.save_as` uses the configured native file bridge when `OPEN_WORKBOOK_FILE_BRIDGE_URL` is set; otherwise it reports capability status.
- `excel.workbook.create_backup` creates a persistent JSON snapshot backup.
- `excel.workbook.restore_backup` restores a persistent or in-memory snapshot backup.
- `excel.workbook.export_copy` creates a persistent snapshot backup, then writes a true `.xlsx` file. If `OPEN_WORKBOOK_FILE_BRIDGE_URL` is configured, the built-in native bridge uses Excel desktop `SaveCopyAs` automation first. If no bridge is configured or the bridge fails, Open Workbook falls back to Office.js compressed workbook slices on supported Excel desktop hosts. If no `targetPath` is provided, the file is written under `.open-workbook/exports` or `OPEN_WORKBOOK_EXPORT_DIR`.
- `excel.backup.create_file` creates a durable full-file backup manifest with path, size, checksum, source snapshot backup id when available, pin state, verification time, and restore status.
- `excel.backup.list` and `excel.backup.get` inspect persisted backup records, including full-file backups and JSON snapshot backups.
- `excel.backup.verify` verifies durable full-file backups and detects missing or checksum-mismatched backup files.
- `excel.backup.pin`, `excel.backup.unpin`, `excel.backup.prune`, and `excel.backup.delete` manage persisted backup retention. Pinned backups are not pruned or deleted.
- `excel.backup.restore_file` defaults to safe `open-as-new` recovery. `replace-open-workbook` requires explicit confirmation, verifies the source backup, creates a pinned emergency file backup, then calls the native bridge to close, replace, reopen, and verify through host automation. `restore-into-open-workbook` remains unsupported because it is not deterministic for full workbook fidelity.
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

After opening a workbook in desktop Excel, run `owb file-bridge smoke --workbook Book1.xlsx --target ./book-copy.xlsx` to verify the running bridge against the real host with non-destructive `SaveCopyAs` export-copy automation. Use `--operation save-as --confirm-save-as` only for intentional Save As testing because it changes the open workbook's file identity.

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

Supported operations are `workbook.save_as`, `workbook.export_copy`, and `workbook.restore_file_backup`. `export_copy` includes the snapshot `sourceBackupId` so the helper can audit which safety backup was captured before writing a true file copy. `restore_file_backup` accepts `backupPath`, `restoreMode`, and optional `restoreTargetPath`; the built-in bridge supports `open-as-new` and `replace-open-workbook` on macOS and Windows, and rejects `restore-into-open-workbook`. If the bridge is not configured or fails, Open Workbook falls back to the add-in compressed-file export path where supported for export copy only.

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

The built-in bridge uses AppleScript on macOS and PowerShell COM automation on Windows for `workbook.save_as`, `workbook.export_copy`, and supported `workbook.restore_file_backup` modes. `export_copy` calls Excel `SaveCopyAs`, which saves a copy of the workbook without changing the open workbook's file identity. It matches `workbookId` against the open Excel workbook name or full path. Linux returns an explicit unsupported result.

Before invoking Excel automation, the bridge resolves `targetPath` to an absolute path, validates it against `OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS` when configured, and creates missing parent directories. Successful `export_copy` responses preserve `sourceBackupId` so audit logs can connect the exported file to the safety snapshot captured before the file write.

## Durable File Backup Lifecycle

Durable file backups are separate from snapshot backups. Snapshot backups are used for range-level rollback through `excel.workbook.restore_backup`. File backups are disaster-recovery artifacts for the entire `.xlsx` file.

`excel.backup.create_file` first captures the existing snapshot safety backup used by `excel.workbook.export_copy`, then writes a true `.xlsx` file through the native bridge or supported Office.js compressed-file fallback. The runtime records a `file-copy` backup with a manifest containing the backup file path, byte size, SHA-256 checksum, source snapshot backup id when present, verification status, pin state, and bridge/export metadata.

`excel.backup.verify` re-checks that the backup file exists and matches its recorded checksum. Missing files are marked `missing`; checksum mismatches are marked `checksum_mismatch`. Agents should verify a file backup before presenting it as restorable.

`excel.backup.restore_file` has explicit restore modes:

- `open-as-new`: default safe mode. The runtime verifies the backup and returns the file path for recovery as a separate workbook.
- `replace-open-workbook`: destructive mode. Requires confirmation. The runtime verifies the selected backup, creates a pinned emergency file backup, and asks the native bridge to close the matched workbook, copy the backup over the target path, and reopen it.
- `restore-into-open-workbook`: unsupported. It would require deterministic sheet/table/name/pivot/chart replacement inside an already-open workbook, which is not safe enough to claim as full-file restore.

The built-in bridge uses AppleScript on macOS and PowerShell COM automation on Windows for `replace-open-workbook`. It validates backup and target paths against `OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS` when configured. Production agents should still prefer snapshot rollback for in-workbook undo and use full-file replace restore only for explicit disaster recovery.

Backup lifecycle operations emit collaboration audit events: `backup.created`, `backup.verified`, `backup.restored`, `backup.deleted`, `backup.pruned`, and `backup.updated`. These events appear in `excel.collab.get_status` so multi-agent clients can show which agent or workflow created, verified, restored, pinned, or removed a persisted backup.

`excel.backup.prune` covers both durable full-file backups and persisted JSON snapshot backups. It accepts `kind` (`all`, `file-copy`, or `snapshot-json`), `maxAgeDays`, `maxBackupsPerWorkbook`, `maxTotalBytes`, and `dryRun`. Automatic retention uses balanced defaults of 30 days, 20 backups per workbook, and 1 GiB total payload bytes. Override with `OPEN_WORKBOOK_BACKUP_RETENTION_DAYS`, `OPEN_WORKBOOK_BACKUP_RETENTION_COUNT`, and `OPEN_WORKBOOK_BACKUP_RETENTION_BYTES`; set `OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED=1` to disable automatic pruning.

Set `OPEN_WORKBOOK_FILE_BRIDGE_TIMEOUT_MS` to override the default 30000 ms backend-to-bridge timeout. Set `OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS` to a path-delimited allowlist of output directories for native Save As and Export Copy. Set `OPEN_WORKBOOK_EXPORT_DIR` to control the default output directory for add-in compressed-file exports.

`excel.runtime.get_status` and `excel.runtime.get_capabilities` include `fileBridge` status, including the configured URL and route path, so agents can check whether native Save As or Export Copy is configured before requesting it. Pass `probeFileBridge: true` to `excel.runtime.get_status` when an agent needs a live `/status` health check with reachability, route, adapter platform, and supported-operation metadata.

## Real Host Smoke

Use the CLI smoke command to verify the native bridge against a real open Excel desktop workbook:

```bash
owb file-bridge start
owb file-bridge smoke --workbook Book1.xlsx --target ./open-workbook-smoke-copy.xlsx
```

The default smoke uses `workbook.export_copy`, which calls Excel `SaveCopyAs` and should not change the open workbook's file identity. The command first checks `GET /status`, then posts to the bridge operation route and fails if Excel is not running, the workbook cannot be matched by name/full path, or the target file cannot be written.

To test file identity-changing Save As behavior, use an explicit confirmation flag:

```bash
owb file-bridge smoke --workbook Book1.xlsx --operation save-as --confirm-save-as --target ./open-workbook-save-as.xlsx
```

Local config export/import is different from workbook file export. It does not create or modify an `.xlsx`; it moves Open Workbook registry metadata so teams can version templates, semantic regions, and permission defaults alongside a project. Embedded local config modifies workbook metadata through Office.js custom XML parts and is guarded by workbook-level permissions.

## Backup Directory

Set `OPEN_WORKBOOK_BACKUP_DIR` to control where persistent snapshot backups are written. By default, backups are stored in `.open-workbook/backups` under the MCP process working directory.
