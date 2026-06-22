# Tests

The root `package.json` keeps stable command names for all test lanes. This folder contains test runners and fixtures that sit outside package-local unit tests.

## Folders

- `e2e/`: MCP contract, MCP agent surface, fake-host workflow, Codex agent, model matrix, live Excel smoke, and office behavior runners.
- `e2e/fixtures/`: shared production-like scenario data and workbook fixture notes for E2E behavior runs.
- `smoke/`: CLI installation and command smoke tests that do not require Excel.
- `benchmarks/`: synthetic performance checks used by `pnpm verify` and `pnpm benchmark:synthetic`.

Default CI uses `pnpm verify`, deterministic `pnpm test:e2e` lanes, and `pnpm pack:dry-run`. Live Excel and model-quality E2E lanes are opt-in or diagnostic unless a release explicitly gates on them.

Package-local `pnpm test` includes operation-level checks before MCP scenarios run. In particular, `apps/excel-addin/src/host/executor-core.batch.test.ts` mocks the Office.js host and asserts `operation.execute_batch` routes recent worksheet regression operations to the expected Office.js APIs, reports unsupported operations explicitly, and counts internal syncs.

`pnpm test:regression:opencode` is the focused backend regression lane for recent OpenCode failures. It asserts the preview/apply batch payloads for header styling, grouped styling, inserted columns, column reorder, dropdown validation, and formula conditional formatting.
`pnpm test:e2e:scenarios:regression` runs those recent failures through MCP `excel.agent.run` preview/apply/status calls against the fake Office host and asserts route, host calls, workbook mutation, call budgets, closed-workbook handling, and reloaded `.xlsx` artifact state.
`pnpm test:e2e:scenarios:departments` runs finance, sales/ops, logistics, HR, executive reporting, and data-cleanup workflows through the same strict MCP runner, with read-only and mutating cases for each department. Mutating cases cover values, formulas, number formats, styles, grouped styles, inserted columns, range column reorder, dropdown validation, conditional formatting, sheet protection options, template-backed formula repair, trim-whitespace cleanup, date parsing, numeric parsing, currency normalization, and duplicate cleanup. The same lane also asserts formula auto-repair without a template is rejected with a manual-review repair report instead of mutating the workbook, and pivot/chart executive reporting requests return a host-limited workflow plan with required capabilities.

## E2E Lanes

- `test:e2e:mcp-contract`: direct JSON-RPC contract tests for MCP lifecycle, tools, schemas, structured content, and protocol/tool error boundaries.
- `test:e2e:workbook`: CI-safe `.xlsx` fixture assertions for OOXML workbook parts. It writes before/after workbook artifacts, applies a local operation plan, and checks final values, formulas, styles, data validation, conditional formatting, column metadata, column order, and table definitions.
- `test:e2e:scenario-contract`: fixture contract guard for regression and department scenario packs. It enforces budgets, mutation/safety expectations, route assertions, workbook-output assertions for mutating cases, disconnected-workbook setup for recovery cases, and required department/operation coverage.
- `test:e2e:scenarios`: strict LLM-free MCP scenario gate over a stable default subset. It fails on wrong result type, expectation issues, workbook mutation mismatch, read/payload/latency budget overages, or thrown tool calls. The runner supports single-scenario selection with `--scenarios id`, category selection with `--category name`, scenario-pack selection with `--scenario-file file`, and full-catalog selection with `--scenarios all`; empty selections fail before the MCP server starts.
- `test:e2e:scenarios:category`: strict LLM-free category smoke for the production `formatting` category.
- `test:e2e:scenarios:full`: strict LLM-free full production scenario catalog. This is intentionally separate from default E2E because it is broader and slower.
- `test:e2e:scenarios:regression`: strict LLM-free MCP scenario pack for recent worksheet-operation regressions and closed-workbook recovery, including per-scenario workbook artifact files and ZIP/XML assertions.
- `test:e2e:scenarios:departments`: strict LLM-free department workflow pack with per-scenario workbook artifact assertions for mutating values, formulas, number formats, styles, structure, validation, conditional formatting, sheet protection options, cleanup, date parsing, numeric/currency normalization, duplicate cleanup, template-backed formula repair, formula repair boundaries, pivot/chart workflow boundaries, and column-order cases.
- `test:e2e:agent:*`: real Codex/LLM diagnostics. These are intentionally separate from deterministic default E2E because they depend on an installed, signed-in agent runtime.
- `test:e2e:live:*`: opt-in desktop Excel fidelity tests against scratch workbook copies. The `live:mac:regression` and `live:windows:regression` scripts run the `scratch-core`, `regression-pack`, `template-formula-repair`, `pivot-chart-core`, `pivot-template-repair`, and `chart-template-copy` scenario groups after Excel and the add-in are connected.

The detailed milestone plan lives in `docs/test-strategy.md`.
