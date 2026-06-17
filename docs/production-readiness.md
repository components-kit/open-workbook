# Production Readiness

This checklist defines what must remain true before Open Workbook is published or tagged as production ready.

## Multi-Agent Invariants

- Excel receives workbook mutations through one serialized writer queue per daemon.
- Agents may read and plan in parallel, but commits must pass through task, lock, transaction, permission, backup, and validation paths.
- Every mutating transaction records `agentId`, `taskId`, `workbookId`, operation scopes, lock ids, backups, warnings, and final status when that metadata is available.
- Manual locks and transaction locks use bounded leases. Expired locks must not block future work.
- Lock conflicts return owner/task/expiry metadata where possible, plus structured conflict guidance.
- Formula writes include parsed local range and table dependencies in lock scopes.
- Table, chart, pivot, named range, sheet, and workbook-structure operations use typed scopes, not only broad workbook locks.
- User/manual edits after preview must be detected through fingerprint checks before commit.
- Rollback must be previewed before apply. Later overlapping transactions block single rollback and require rollback-chain preview.
- Rollback chains apply newest-first and require the preview confirmation token when more than one transaction is affected.
- Daemon restart recovery must not present queued or applying transactions as committed work.
- Local daemon state must restore templates, regions, permissions, plans, backup indexes, and collaboration records without requiring a real Excel reconnect.

## Safety Lifecycle

Every stable mutating MCP tool should follow this lifecycle:

1. Validate permissions and destructive-action policy.
2. Resolve workbook/sheet/range/object scopes.
3. Capture snapshot or fingerprint data needed for stale-preview detection.
4. Create a backup for rollback or repair.
5. Acquire scoped transaction locks.
6. Commit through the serialized writer queue.
7. Validate result and record warnings.
8. Release locks and persist task, event, and transaction state.

Capability-unavailable responses are preferred over simulated success when Office.js cannot perform an operation safely across supported hosts.

## Release Gates

- `corepack pnpm verify` passes. This runs build, tests, and synthetic core benchmarks.
- `corepack pnpm test:e2e` passes before release-gate E2E claims. This runs the generated E2E report, default agent surface smoke, default agent workflow smoke, deterministic fake-host MCP sweep, and `test:e2e:agent:core`.
- `corepack pnpm test:e2e:agent-surface` passes before default MCP exposure claims. It validates that default `tools/list` exposes only `excel.agent.run` and that status mode returns structured telemetry without Excel.
- `corepack pnpm test:e2e:agent-workflow` passes before token-saving agent workflow claims. It validates metadata cache reuse, bounded find payloads, targeted answer reads, preview confirmation tokens, apply, and validate on a fake Excel host.
- `corepack pnpm test:e2e:office-agent:behavior` is reviewed before changing default-surface office workflow guidance. It is logging-first and produces behavior reports from the production office-agent scenario fixture, covering connection, workbook overview, targeting, sheet/table reads, comparisons, formula safety, edits, token guards, multilingual prompts, and multi-step workflows without acting as a release gate.
- `corepack pnpm test:e2e:agent:core` passes with the default signed-in Codex model (`gpt-5.4-mini`) before large-workbook or multi-agent agent-safety claims. It validates skill-guided decisions for large table reorder/filter/sort/append/update and lock/conflict handling.
- `corepack pnpm test:e2e:agent:quality` is reviewed before complex workflow quality claims. It is report-only by default and covers sheet/formula creation, formula repair, snapshot/diff/rollback preview, template/style repair, and pivot/chart decisions.
- `corepack pnpm test:e2e:agent:quality:compare` is reviewed before comparing cheap-model baseline quality with a frontier model. Set `OPEN_WORKBOOK_E2E_CHEAP_MODEL` and `OPEN_WORKBOOK_E2E_FRONTIER_MODEL` to control the profiles.
- `corepack pnpm test:e2e:agent:quality:gate` runs the cheap/frontier quality matrix in strict mode when a release claims workflow quality is passing rather than diagnostic.
- MCP callable tool names stay in sync with the shared protocol catalog. `corepack pnpm verify` runs `scripts/validate-mcp-catalog.mjs` after build and fails if a stable or preview callable tool is missing from the MCP server or if the server registers a tool that is not in the callable catalog.
- Tool surface docs stay in sync with the protocol catalog and public one-tool MCP surface. `corepack pnpm verify` runs `scripts/validate-docs-surface.mjs` and fails if `docs/tool-surface.md` omits a stable or preview catalog capability.
- Package metadata stays publishable. `corepack pnpm verify` runs `scripts/validate-package-metadata.mjs` and fails if package versions, repository metadata, public/private publish intent, README presence, or `dist` entrypoints drift.
- CLI install smoke stays healthy. `corepack pnpm verify` runs `scripts/smoke-cli.mjs` and checks `doctor`, `paths`, setup dry-run output, skills.sh guidance, fallback instruction generation, service wrapper generation, sideload manifest generation, and concise disconnected-runtime errors without requiring Excel or a bound local daemon.
- `git diff --check` passes.
- GitHub Actions CI passes the same non-E2E verification plus `corepack pnpm pack:dry-run` for every publishable package.
- Tool catalog, README, installation docs, MCP client docs, and instruction docs match the public `excel.agent.run` MCP surface.
- Fresh local install can run `npx -y @components-kit/open-workbook setup`, install `open-workbook-excel` with `npx skills add components-kit/open-workbook --skill open-workbook-excel`, generate/install a manifest, launch `npx -y @components-kit/open-workbook@latest mcp`, and report runtime status. If the runtime cannot bind or cannot be reached, CLI commands must fail with concise user-facing errors.
- Excel add-in can be sideloaded on macOS and Windows using the generated manifest.
- Native file bridge host smoke passes on macOS and Windows with `owb file-bridge smoke --workbook <open-workbook-name> --target <copy.xlsx>`.
- Real Excel smoke starts with `OPEN_WORKBOOK_LIVE_E2E=1 corepack pnpm test:e2e:live:mac` or `OPEN_WORKBOOK_LIVE_E2E=1 corepack pnpm test:e2e:live:windows` and must confirm the backend is reachable and the Excel add-in has an active workbook. For deeper host claims, add `-- --deep` to run a scratch-sheet range write, formula validation, snapshot diff, and rollback preview.
- Multi-agent smoke test covers two MCP clients planning concurrently, one writer waiting on a lock, and telemetry reporting the contention.
- Known Office.js host limitations are documented as `CAPABILITY_UNAVAILABLE` or warnings.

## Host Verification Matrix

- macOS Excel desktop: sideload manifest, connect add-in, read/write ranges, batch apply, backup, rollback.
- Windows Excel desktop: sideload manifest, connect add-in, read/write ranges, batch apply, backup, rollback.
- Excel on the web: optional until HTTPS hosting and tenant sideload instructions are finalized.

## Structure-Level Conflicts

Open Workbook should keep structure-level merge conflicts blocked by default. Sheet deletes, sheet renames, row/column insertions, table resizes, named-range changes, pivot source changes, and chart source changes can invalidate later operations in ways that are not equivalent to code text merges. Agents should create a new plan after refresh instead of auto-merging these changes.

## Publish Notes

Before publishing, bump versions consistently, regenerate CLI assets, confirm the public package is `@components-kit/open-workbook`, and run the release gates above from a clean checkout. Do not claim production host readiness until the real Excel smoke tests pass on both macOS and Windows.
