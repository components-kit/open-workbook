# @open-workbook/excel-core

Core workbook utilities for Open Workbook.

It contains:

- A1 range parsing and cell counting
- stable fingerprints
- batch compilation
- backup records
- snapshot comparison
- template registry
- plan manager
- default permission policy

## Usage

```ts
import { BatchCompiler, BackupManager, SnapshotManager, TemplateRegistry } from "@open-workbook/excel-core";
```

This package is runtime-agnostic. It does not depend on Office.js and can be reused by future Excel engines.
