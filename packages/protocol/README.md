# @open-workbook/protocol

Shared protocol package for Open Workbook.

It contains:

- workbook, range, operation, snapshot, template, table, chart, pivot, and permission types
- JSON-RPC envelope types
- tool catalog and exposure rules
- MCP resource catalog
- prompt catalog
- runtime error helpers

## Usage

```ts
import { ToolCatalog, getToolCatalogSummary, runtimeError } from "@open-workbook/protocol";
```

Most users should install `@open-workbook/cli`; this package is published for integrators building clients, tests, or alternate engines.
