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
- MCP callable tool names stay in sync with the shared protocol catalog. `corepack pnpm verify` runs `scripts/validate-mcp-catalog.mjs` after build and fails if a stable or preview callable tool is missing from the MCP server or if the server registers a tool that is not in the callable catalog.
- Tool surface docs stay in sync with the callable catalog. `corepack pnpm verify` runs `scripts/validate-docs-surface.mjs` and fails if `docs/tool-surface.md` omits an exposed stable or preview tool.
- Package metadata stays publishable. `corepack pnpm verify` runs `scripts/validate-package-metadata.mjs` and fails if package versions, repository metadata, public/private publish intent, README presence, or `dist` entrypoints drift.
- CLI install smoke stays healthy. `corepack pnpm verify` runs `scripts/smoke-cli.mjs` and checks `doctor`, `paths`, OpenCode config generation, service wrapper generation, sideload manifest generation, and concise disconnected-daemon errors without requiring Excel or a bound local daemon.
- `git diff --check` passes.
- GitHub Actions CI passes the same non-E2E verification plus `corepack pnpm pack:dry-run` for every publishable package.
- Tool catalog, README, installation docs, and OpenCode docs match the exposed MCP surface.
- Fresh local install can run `owb doctor`, generate a manifest, start `owb daemon`, connect `owb mcp`, and report runtime status. If the daemon cannot bind or cannot be reached, CLI commands must fail with concise user-facing errors.
- Excel add-in can be sideloaded on macOS and Windows using the generated manifest.
- Real Excel smoke test covers read, write, batch apply, template repair, rollback preview/apply, and conflict detection.
- Multi-agent smoke test covers two MCP clients planning concurrently, one writer waiting on a lock, and telemetry reporting the contention.
- Known Office.js host limitations are documented as `CAPABILITY_UNAVAILABLE` or warnings.

## Host Verification Matrix

- macOS Excel desktop: sideload manifest, connect add-in, read/write ranges, batch apply, backup, rollback.
- Windows Excel desktop: sideload manifest, connect add-in, read/write ranges, batch apply, backup, rollback.
- Excel on the web: optional until HTTPS hosting and tenant sideload instructions are finalized.

## Structure-Level Conflicts

Open Workbook should keep structure-level merge conflicts blocked by default. Sheet deletes, sheet renames, row/column insertions, table resizes, named-range changes, pivot source changes, and chart source changes can invalidate later operations in ways that are not equivalent to code text merges. Agents should create a new plan after refresh instead of auto-merging these changes.

## Publish Notes

Before publishing, bump versions consistently, regenerate CLI assets, and run the release gates above from a clean checkout. Do not claim production host readiness until the real Excel smoke tests pass on both macOS and Windows.
