# @components-kit/open-workbook-excel-core

Core workbook utilities for Open Workbook.

It contains:

- A1 range parsing and cell counting
- stable fingerprints
- batch compilation
- safety utilities for backups, snapshots, permissions, plans, and transactions
- coordination utilities for locks, tasks, and conflicts
- template registry
- formula dependency parsing and tracing

## Usage

```ts
import { BatchCompiler, BackupManager, SnapshotManager, TemplateRegistry } from "@components-kit/open-workbook-excel-core";
```

This package is runtime-agnostic. It does not depend on Office.js and can be reused by future Excel engines.

Source is grouped by domain under `src/range`, `src/safety`, `src/coordination`, `src/formula`, and `src/templates`; root exports stay compatible for consumers.
