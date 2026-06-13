# @components-kit/open-workbook-backend

Local backend runtime for Open Workbook.

The backend coordinates Excel add-in sessions, workbook snapshots, backups, plans, permissions, events, diffs, validation, repair, and rollback. It exposes a WebSocket endpoint for the Office.js add-in and is normally started by `@components-kit/open-workbook-mcp-server`.

## Exports

```ts
import { RuntimeService } from "@components-kit/open-workbook-backend/runtime";
import { startBackendServer } from "@components-kit/open-workbook-backend/server";
```

## Environment

- `OPEN_WORKBOOK_HOST`
- `OPEN_WORKBOOK_PORT`
- `OPEN_WORKBOOK_ADDIN_PATH`
- `OPEN_WORKBOOK_BACKUP_DIR`

## Notes

Most users should install `@components-kit/open-workbook`. This package is published for embedding, testing, and advanced integrations.
