# Tests

The root `package.json` keeps stable command names for all test lanes. This folder contains test runners and fixtures that sit outside package-local unit tests.

## Folders

- `e2e/`: MCP agent surface, fake-host workflow, Codex agent, model matrix, live Excel smoke, office behavior, operation coverage, and E2E report runners.
- `e2e/fixtures/`: shared production-like scenario data for E2E behavior runs. The production scenario fixture also includes coverage metadata for representative internal capability, host-method, and batch operation groups.
- `e2e/lib/`: shared report utilities for validating fixture coverage against the built backend and add-in registries.
- `smoke/`: CLI installation and command smoke tests that do not require Excel.
- `benchmarks/`: synthetic performance checks used by `pnpm verify` and `pnpm benchmark:synthetic`.

Default CI uses `pnpm verify`, `pnpm test:e2e:agent-surface`, and `pnpm pack:dry-run`. `pnpm test:e2e:report` validates the operation coverage fixture and writes coverage artifacts. Live Excel and model-quality E2E lanes are opt-in or diagnostic unless a release explicitly gates on them.
