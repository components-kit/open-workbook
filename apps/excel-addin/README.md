# @components-kit/open-workbook-excel-addin

Office.js Excel add-in runtime for Open Workbook.

This workspace package contains the taskpane HTML, manifest template, static asset server, and Office.js executor. It is private because the built add-in assets are bundled into `@components-kit/open-workbook`.

The taskpane WebSocket lifecycle is handled by `src/connection.ts`. Backend JSON-RPC method dispatch lives in `src/rpc-methods.ts`, which delegates to the Office.js executor functions in `src/excel-executor.ts`.

## Development

```bash
corepack pnpm --filter @components-kit/open-workbook-excel-addin build
corepack pnpm --filter @components-kit/open-workbook-excel-addin dev
```

The development server serves:

- `taskpane.html`
- `/manifest.xml`
- add-in icons
- compiled JavaScript under `/dist`

## Runtime

The taskpane connects to the backend URL embedded in the generated manifest:

```text
ws://127.0.0.1:37845/addin
```

Use `owb sideload mac`, `owb sideload windows`, or `owb sideload manifest` to generate the manifest that Excel should load.
