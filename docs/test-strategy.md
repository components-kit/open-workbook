# Test Strategy

Open Workbook should test behavior at the layer where failures actually happen. Registry metadata checks are useful, but they are not operation coverage.

## Current Lanes

- `pnpm test`: package-local unit and contract tests for backend runtime behavior, MCP result handling, protocol metadata, core planning utilities, add-in registry wiring, and mocked Office.js executor behavior.
- `pnpm verify`: build, catalog checks, operation manifest checks, docs/skills/package checks, CLI smoke, unit tests, and synthetic benchmarks.
- `pnpm test:regression:opencode`: focused backend preview/apply regression pack for recent OpenCode failures: black header style, grouped styles, inserted columns, range/table column reorder, dropdown validation, and formula conditional formatting.
- `pnpm test:e2e:mcp-contract`: direct JSON-RPC contract checks for MCP lifecycle, tool schemas, structured content, and malformed request behavior.
- `pnpm test:e2e:workbook`: CI-safe `.xlsx` fixture assertion lane that writes a before workbook, applies a local OOXML operation plan, and validates final workbook parts for values, styles, data validation, conditional formatting, column metadata, and table definitions.
- `pnpm test:e2e:scenario-contract`: fixture contract guard for regression and department MCP scenario packs. It fails if scenarios omit required budgets, mutation/safety expectations, route assertions, final workbook assertions for mutating cases, disconnected-workbook setup for recovery cases, or required department/operation coverage.
- `pnpm test:e2e:scenarios`: strict LLM-free MCP scenario gate over the stable default subset, currently status, formatting preview, and value preview.
- `pnpm test:e2e:scenarios:category`: strict LLM-free MCP scenario gate for a selected production category, currently `formatting`.
- `pnpm test:e2e:scenarios:full`: strict LLM-free MCP scenario gate over the full production scenario catalog.
- `pnpm test:e2e:scenarios:regression`: strict LLM-free MCP regression pack for recent worksheet failures and connection recovery, asserting preview/apply routing, host calls, workbook mutation, call budgets, closed-workbook handling, and per-scenario `.xlsx` artifact reload checks for the regression workbook state.
- `pnpm test:e2e:scenarios:departments`: strict LLM-free department workflow pack with read-only and mutating scenarios for finance, sales/ops, logistics, HR, executive reporting, and data cleanup. Mutating scenarios write `.xlsx` artifacts and reload ZIP/XML parts to assert final values, formulas, styles, inserted columns, dropdown validation, conditional formatting, column order, date parsing, numeric/currency cleanup, duplicate cleanup, sheet protection options, and template-backed formula repair. The same lane also asserts unsupported formula auto-repair returns an actionable manual-review boundary without mutating the workbook, and that pivot/chart executive reporting requests return a host-limited workflow plan with required capabilities.
- `pnpm test:e2e:agent-surface`: MCP stdio smoke for the public one-tool surface.
- `pnpm test:e2e:agent-workflow`: MCP plus fake Excel add-in workflow checks without a real LLM.
- `pnpm test:e2e:office-agent:behavior`: MCP-direct production-like scenario runner; useful as broad diagnostics across the full fixture catalog.
- `pnpm test:e2e:agent:*`: real Codex/LLM decision diagnostics; intentionally separate from default deterministic E2E.
- `pnpm test:e2e:live:*`: opt-in desktop Excel host smoke, including the release regression group for scratch formulas, worksheet regressions, template-backed formula repair, supported PivotTable/chart creation, PivotTable template repair/diff, and chart template copy.

## Cleanup Policy

The suite should not keep one-test wrapper files that only assert a registry entry exists. Central registry tests should own metadata consistency. Behavior tests should prove reads, writes, routing, error handling, or workbook state changes.

## Target Lanes

- **Operation correctness**: every host-backed or batch-backed operation should have deterministic fake Office.js behavior tests, especially writes through `operation.execute_batch`.
- **MCP contract**: direct JSON-RPC tests for accepted shapes, rejected shapes, structured errors, and non-looping guidance.
- **MCP scenario simulation**: LLM-free scenarios that call MCP and assert selected action/operation, workbook delta, safety result, and call budget.
- **Department workflows**: finance, sales/ops, inventory/logistics, HR, project/admin, executive reporting, and data-cleaning fixtures that model real workbook tasks.
- **Optional quality gates**: real-agent and live-Excel lanes stay separate from default correctness checks because they are slower and environment-dependent.

## Milestones

### Milestone 0: Baseline and Taxonomy

Goal: keep the suite honest after pruning fake coverage.

Deliverables:

- Keep registry metadata tests limited to catalog and wiring consistency.
- Classify future coverage by lane: operation correctness, MCP contract, MCP scenario, real workbook, department workflow, live Excel, or real-agent quality.
- Classify capabilities by implementation path: batch-backed, host-backed, backend-only, agent-routing, or diagnostic.
- Document which commands are default CI, release gates, diagnostics, and opt-in host checks.

Acceptance criteria:

- Every new test belongs to exactly one lane.
- Registry metadata checks are never counted as operation behavior coverage.

### Milestone 1: Operation Correctness Harness

Goal: prove supported host-backed and batch-backed operations before MCP or agent routing is involved.

Status: expanded. `pnpm test:regression:opencode` verifies the backend preview/apply boundary for the recent production failures and asserts the exact batch or table request emitted on apply. Package tests now include mocked Office.js executor coverage for every protocol `ExcelOperation` kind: supported operations must hit a focused fake Office.js behavior path, and host-limited operations must be listed as explicit unsupported-warning cases. This is still mocked Office.js coverage, not desktop Excel save/export fidelity.

Deliverables:

- Add a deterministic operation harness around `operation.execute_batch`.
- Cover operation input normalization, executor route, affected ranges, result telemetry, warnings, and workbook model state.
- Start with production failures: `range.write_styles`, `range.write_styles_many`, `range.insert_columns`, `range.reorder_columns`, `range.write_data_validation`, and `range.write_conditional_formatting`.
- Keep compiler coverage for whole-row and whole-column structural targets such as `F:F` and `3:3`, because agents commonly produce these addresses for insert/delete operations.

Acceptance criteria:

- A broken operation fails in a focused operation test before scenario tests run.
- Unsupported or host-limited operations return explicit capability warnings or errors, never silent success.

### Milestone 2: Real `.xlsx` Fixture and Assertion Framework

Goal: verify final workbook files, not only in-memory fake workbook state.

Status: expanded. `pnpm test:e2e:workbook` generates a small valid `.xlsx`, writes a before copy, applies a local OOXML operation plan for header style, inserted column, formulas, dropdown validation, conditional formatting, table resize, and column reorder, then parses the after file without desktop Excel and asserts final workbook parts. `pnpm test:e2e:scenarios:regression` now also writes per-scenario `.xlsx` artifacts from the MCP-applied fake workbook and reloads the ZIP/XML parts to assert values, styles, inserted columns, validations, conditional formatting, and table column order. This is still fake-host file output, not desktop Excel save/export output.

Deliverables:

- Add `tests/e2e/fixtures/workbooks/` with small curated `.xlsx` files.
- Add workbook assertion helpers that inspect Office Open XML for values, formulas, styles, row/column structure, data validation, conditional formatting, table definitions, and sheet metadata.
- Use direct ZIP/XML inspection for style, validation, and conditional-format fidelity; use a workbook reader only for convenient value/formula checks.
- Save artifacts per run: before workbook, after workbook, assertion report, and concise failure summary.

Acceptance criteria:

- A test can copy a fixture, apply an operation or MCP scenario, save/export the result, reopen the `.xlsx`, and assert exact expected state.
- File-level tests run in CI without desktop Excel.

### Milestone 3: MCP Contract Compliance

Goal: catch protocol failures like `-32602`, malformed payloads, missing structured content, and retry-loop ambiguity.

Status: expanded. `pnpm test:e2e:mcp-contract` covers the public MCP lifecycle, tool/resource listing, unknown tools/resources, malformed `tools/call` params, missing request fields, invalid modes, bad target shapes, malformed patch targets, and recent agent-facing bad update payloads for styles, dropdown validation, conditional formatting, and column order. Errors must mention the relevant field/path and must not dump input/output schemas or tell agents to resend schemas.

Deliverables:

- Add direct JSON-RPC tests for `initialize`, `notifications/initialized`, `tools/list`, `tools/call`, unknown tools, malformed arguments, `resources/list`, and `resources/read`.
- Validate public surface, input schemas, output schemas, `structuredContent`, compact text fallback, telemetry schema, resource links, and request-id handling.
- Separate protocol errors from tool execution errors: invalid JSON-RPC params should use protocol errors; workbook/runtime failures should return structured tool results with actionable `isError` semantics.

Acceptance criteria:

- Bad style, validation, conditional-format, and missing-field payloads fail with clear field/path guidance.
- No error response encourages the agent to resend schema repeatedly.

### Milestone 4: Strict MCP Scenario Runner

Goal: replace logging-style behavior reports with pass/fail MCP scenarios.

Status: expanded. `pnpm test:e2e:scenarios` runs `tests/e2e/office-agent-behavior.mjs --strict` against a stable subset and fails on expectation issues, thrown tool calls, wrong result type, workbook mutation mismatch, read budget overages, payload budget overages, or latency budget overages. The runner now has explicit selection modes for single scenarios (`--scenarios id`), categories (`--category category-name`), scenario packs (`--scenario-file file`), and the full production catalog (`--scenarios all`), and it fails before startup when a selector matches no scenarios. `pnpm test:e2e:scenario-contract` guards the regression and department fixture schema so new scenarios cannot be added without budgets, route expectations, safety expectations, and final workbook assertions where required. `pnpm test:e2e:scenarios:category` and `pnpm test:e2e:scenarios:full` expose the category and full-catalog paths. `pnpm test:e2e:scenarios:regression` uses the same runner for the recent worksheet operation failures and is included in the deterministic `pnpm test:e2e` gate.

Deliverables:

- Refactor `tests/e2e/office-agent-behavior.mjs` into a strict LLM-free runner.
- Scenario definitions include request, mode, optional intent/target/values, expected status, expected action handler, expected operation kinds, expected workbook delta, and budgets.
- Preserve MCP transcript, before/after state, workbook artifacts when available, and a compact failure report.

Acceptance criteria:

- A scenario fails on wrong route, wrong operation, wrong workbook mutation, excessive tool calls, excessive reads, excessive payload, wrong safety result, or missing actionable error guidance.
- Runner can execute a single scenario, category, regression pack, or full suite.

### Milestone 5: Production Regression Pack

Goal: encode known production failures so they cannot regress.

Status: expanded. `tests/e2e/fixtures/office-agent-regression-scenarios.json` covers the current regression pack through MCP `excel.agent.run` preview/apply/status calls against the fake Office host. It asserts route, operation kind or host method, final fake workbook state, call budgets, and reloaded `.xlsx` artifact state for values, styles, inserted columns, validations, conditional formatting, and table column order. It also simulates a no-active-workbook state and verifies the agent returns a graceful setup response without reading cells or mutating workbook state. Real desktop Excel save/export assertions remain part of Milestone 7.

Scenarios:

- Header color to black.
- Add a new column.
- Swap or reorder columns.
- Add select-list dropdown to cells.
- Add formula-based conditional color.
- Apply styling once without repeated/confused calls.
- Reject malformed validation/style payloads cleanly.
- Recover cleanly from disconnected or closed workbook responses.

Acceptance criteria:

- Each scenario asserts both MCP route and final workbook state.
- Failures identify the broken layer: MCP contract, routing, operation compile, executor, or workbook output.
- Regression pack runs without a real LLM.

### Milestone 6: Department Workflow Fixtures

Goal: test real office tasks instead of only technical operations.

Departments:

- Finance: formulas, summaries, conditional formatting, and monthly reports.
- Sales/Ops: pipeline table updates, filters, status columns, and dropdowns.
- Logistics: shipment status, ETA formulas, and column reordering.
- HR: employee tracker cleanup, validation lists, and protected headers.
- Executive reporting: styled summary sheets, charts, and pivots where supported.
- Data cleanup: headers, dates, numbers, duplicates, and whitespace.

Status: expanded. `tests/e2e/fixtures/office-agent-department-scenarios.json` covers read-only and mutating MCP scenarios for finance, sales/ops, logistics, HR, executive reporting, and data cleanup. The mutating scenarios apply through `excel.agent.run`, assert host route and operation kind, then reload generated `.xlsx` artifacts to verify final workbook state. Coverage now includes `range.write_values`, `range.write_formulas`, `range.write_number_formats`, `range.write_styles`, `range.write_styles_many`, `range.insert_columns`, `range.reorder_columns`, `range.write_data_validation`, `range.write_conditional_formatting`, `sheet.protect` with reviewer permission options, trim-whitespace cleanup, date parsing with number-format persistence, numeric parsing, currency normalization, duplicate cleanup through batch read/write, and supported template-backed formula repair through `template.capture`, `template.repair`, and `template.capture_sheet`. Formula auto-repair without a template is currently an unsupported capability, so the lane asserts `repair_formula_errors` returns `VALIDATION_FAILED` with `manual_review` and no workbook mutation. Pivot/chart executive reporting is covered at the public MCP workflow layer: `create_pivot_chart_summary` must return a non-mutating `workflow_plan`, manual review next action, host-limit warning, and the expected pivot/chart capabilities. Broader department-specific operations such as live host PivotTable/chart mutation fidelity still need additional fixtures.

Acceptance criteria:

- Each department has at least one read-only scenario and one mutating scenario.
- Every mutating scenario validates final `.xlsx` state after reload.
- Scenarios include performance budgets and safety expectations.

### Milestone 7: Live Excel Fidelity Gate

Goal: confirm Office.js behavior matches deterministic file-level tests.

Status: expanded. `tests/e2e/live-smoke.mjs` now supports named opt-in scenario groups. `scratch-core` preserves the existing scratch-sheet write/formula/snapshot/rollback smoke, `regression-pack` applies the high-risk worksheet operation classes against a scratch sheet in desktop Excel: styles, inserted columns, range column reorder, dropdown validation, and conditional formatting, `template-formula-repair` registers a live template sheet, repairs formulas on a target sheet through `template.repair`, then reads formulas back from Excel, `pivot-chart-core` creates a live PivotTable from scratch data, validates/refreshes it, creates a chart, and reads chart metadata back from Excel, `pivot-template-repair` creates template/target PivotTables, proves their fingerprint diff, repairs target layout/fields/data settings from the template, then verifies the target row field and cleared row-field diff, and `chart-template-copy` copies chart type, style, title, and position from a template chart to a target chart, then reads chart metadata back from Excel. These groups are exposed through `pnpm test:e2e:live:mac:regression` and `pnpm test:e2e:live:windows:regression`; they still require `OPEN_WORKBOOK_LIVE_E2E=1` or `--run`.

Deliverables:

- Extend `test:e2e:live:mac` and `test:e2e:live:windows` into scenario groups.
- Run against scratch workbook copies opened in desktop Excel with the Open Workbook add-in connected.
- Cover the Milestone 5 regression pack, template-backed formula repair, supported PivotTable/chart creation, PivotTable template repair/diff, and chart template copy in live mode after deterministic scenarios are stable.

Acceptance criteria:

- Live tests are opt-in through `OPEN_WORKBOOK_LIVE_E2E=1`.
- Tests never mutate user workbooks directly.
- Release notes can distinguish file-level verified behavior from desktop Excel verified behavior.

### Milestone 8: CI and Release Policy

Goal: make the right tests block the right claims.

Policy:

- `pnpm verify`: fast unit, contract, smoke, docs, catalog, and synthetic checks.
- `pnpm test:e2e`: deterministic, LLM-free E2E checks: MCP contract, workbook fixture, scenario fixture contract, agent surface, fake-host workflow, strict stable scenarios, strict regression scenarios, and strict department workflow scenarios.
- `pnpm test:e2e:workbook`: CI-safe real `.xlsx` gate.
- `pnpm test:e2e:mcp-contract`: required before MCP contract claims.
- `pnpm test:e2e:scenarios`: required before agent workflow claims.
- `pnpm test:e2e:live:*`: manual or release host-readiness gate.

Acceptance criteria:

- Release docs map each production claim to the command that supports it.
- No production-readiness claim depends only on fake-host tests.

## Next Work

Continue broadening operation correctness coverage toward the full capability set and add deeper department fixtures for host-limited PivotTable/chart dimensions such as PivotChart-specific behavior, slicers/timelines, and filter/item grouping details not exposed consistently by Office.js.
