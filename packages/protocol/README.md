# @components-kit/open-workbook-protocol

Shared protocol package for Open Workbook.

It contains:

- workbook, range, operation, snapshot, template, table, chart, pivot, and permission types
- JSON-RPC envelope types
- public MCP tool catalog and internal Excel capability metadata
- MCP resource catalog
- prompt catalog
- runtime error helpers

## Usage

```ts
import { ToolCatalog, getToolCatalogSummary, runtimeError } from "@components-kit/open-workbook-protocol";
```

Most users should install `@components-kit/open-workbook`; this package is published for integrators building clients, tests, or alternate engines.
