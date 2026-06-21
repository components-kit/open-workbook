# Test Strategy

Open Workbook should test behavior at the layer where failures actually happen. Registry metadata checks are useful, but they are not operation coverage.

## Current Lanes

- `pnpm test`: package-local unit and contract tests for backend runtime behavior, MCP result handling, protocol metadata, core planning utilities, and add-in registry wiring.
- `pnpm verify`: build, catalog checks, operation manifest checks, docs/skills/package checks, CLI smoke, unit tests, and synthetic benchmarks.
- `pnpm test:e2e:agent-surface`: MCP stdio smoke for the public one-tool surface.
- `pnpm test:e2e:agent-workflow`: MCP plus fake Excel add-in workflow checks without a real LLM.
- `pnpm test:e2e:office-agent:behavior`: MCP-direct production-like scenario runner; useful as diagnostics, but not yet strict operation coverage.
- `pnpm test:e2e:agent:*`: real Codex/LLM decision diagnostics.
- `pnpm test:e2e:live:*`: opt-in desktop Excel host smoke.

## Cleanup Policy

The suite should not keep one-test wrapper files that only assert a registry entry exists. Central registry tests should own metadata consistency. Behavior tests should prove reads, writes, routing, error handling, or workbook state changes.

## Target Lanes

- **Operation correctness**: every host-backed or batch-backed operation should have deterministic fake Office.js behavior tests, especially writes through `operation.execute_batch`.
- **MCP contract**: direct JSON-RPC tests for accepted shapes, rejected shapes, structured errors, and non-looping guidance.
- **MCP scenario simulation**: LLM-free scenarios that call MCP and assert selected action/operation, workbook delta, safety result, and call budget.
- **Department workflows**: finance, sales/ops, inventory/logistics, HR, project/admin, executive reporting, and data-cleaning fixtures that model real workbook tasks.
- **Optional quality gates**: real-agent and live-Excel lanes stay separate from default correctness checks because they are slower and environment-dependent.

## Next Work

Refactor `tests/e2e/office-agent-behavior.mjs` into a stricter MCP scenario runner, then add focused production regressions for header styling, column insertion, column reorder, dropdown validation, conditional formula formatting, and repeated style-apply behavior.
