# @components-kit/open-workbook-office-js-engine

Office.js engine interface and defaults for Open Workbook.

The live Office.js execution currently lives in the Excel add-in runtime. This package defines engine contracts, defaults, and a placeholder implementation that future host bindings can share.

## Usage

```ts
import { DefaultOfficeJsEngineOptions, OfficeJsEngine } from "@components-kit/open-workbook-office-js-engine";
```

Most users should install `@components-kit/open-workbook`. This package is published for engine integration work and tests.
