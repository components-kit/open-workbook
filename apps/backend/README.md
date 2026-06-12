# @open-workbook/backend

Local backend runtime for Open Workbook.

The backend coordinates Excel add-in sessions, workbook snapshots, backups, plans, permissions, events, diffs, validation, repair, and rollback. It exposes a WebSocket endpoint for the Office.js add-in and is normally started by `@open-workbook/mcp-server`.

## Exports

```ts
import { RuntimeService } from "@open-workbook/backend/runtime";
import { startBackendServer } from "@open-workbook/backend/server";
```

## Environment

- `OPEN_WORKBOOK_HOST`
- `OPEN_WORKBOOK_PORT`
- `OPEN_WORKBOOK_ADDIN_PATH`
- `OPEN_WORKBOOK_BACKUP_DIR`

## Notes

Most users should install `@open-workbook/cli`. This package is published for embedding, testing, and advanced integrations.
