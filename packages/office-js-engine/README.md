# @open-workbook/office-js-engine

Office.js engine interface and defaults for Open Workbook.

The live Office.js execution currently lives in the Excel add-in runtime. This package defines engine contracts and default options that future host bindings can share.

## Usage

```ts
import { DefaultOfficeJsEngineOptions, OfficeJsEngine } from "@open-workbook/office-js-engine";
```

Most users should install `@open-workbook/cli`. This package is published for engine integration work and tests.
