# Workbook Fixtures

This directory is reserved for curated `.xlsx` fixtures used by `pnpm test:e2e:workbook`.

Current coverage starts with a generated scratch workbook in `tests/e2e/workbook-fixture.mjs` so the lane can run without binary fixture churn or desktop Excel. Future checked-in workbooks should stay small and must have explicit assertions for the OOXML parts they are meant to protect: values, formulas, styles, data validation, conditional formatting, table definitions, sheet metadata, and structure changes.
