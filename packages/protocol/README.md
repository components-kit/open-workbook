# @components-kit/open-workbook-protocol

Shared protocol package for Open Workbook.

It contains:

- workbook, range, operation, snapshot, template, table, chart, pivot, and permission types
- JSON-RPC envelope types
- public MCP agent tool contract exposing only `excel.agent.run`
- internal backend capability catalog for orchestration and tests
- MCP resource catalog
- prompt catalog
- runtime error helpers

## Usage

```ts
import { PublicAgentToolCatalog, InternalCapabilityCatalog, runtimeError } from "@components-kit/open-workbook-protocol";
```

Normal MCP clients should only see `excel.agent.run`. The internal capability catalog is backend-owned metadata, not an agent tool surface.

Most users should install `@components-kit/open-workbook`; this package is published for integrators building clients, tests, or alternate engines.
