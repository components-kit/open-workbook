# Performance Contract

Open Workbook treats speed as a correctness requirement because slow workbook automation encourages unsafe shortcuts.

## Rules

- All writes route through the batch compiler.
- Individual write tools are syntactic sugar over batch operations.
- Do not call `context.sync()` inside per-cell loops.
- Use 2D array assignments for values and formulas.
- Load only requested Office.js properties.
- Group operations by workbook, sheet, and contiguous range.
- Chunk large payloads before Office.js limits.
- Use `excel.batch.preflight` before applying large generated batches. It tells the agent whether to apply directly, submit to the queue, or use `excel.batch.submit_chunked`.
- Use `excel.range.write_styles_many` for grouped report styling instead of many parallel single-range style writes; large style entry lists are split into queued parent jobs.
- Surface queued or long-running mutations through transaction progress instead of hiding them behind parallel write calls.
- For chunked work, surface `excel.job.get`/`excel.job.wait` progress so the user sees one update with chunk progress.
- Suspend calculation and screen updating only around large operations.

## Large Range Execution

Large matrix writes are split into row chunks before assignment to Office.js ranges:

- values
- formulas
- number formats

The chunker preserves whole rows and uses the configured cell limit to avoid sending very large matrices through a single Office.js range assignment. Structural operations such as row insertion, table resize, chart creation, and PivotTable creation remain native Office.js object calls because splitting them would change semantics.

Style-only batches can be retried adaptively after an add-in timeout because reapplying the same range style is idempotent. Value, formula, and number-format matrices are split before execution when preflight can safely chunk by rows. Structural mutations do not auto-retry after a timeout because Excel may have applied part or all of the original request.

## Telemetry

Every read/write result records:

- duration in milliseconds
- sync count
- payload bytes
- cells read or written
- range count
- chunk count
- engine name and version
- warning count

Initial releases collect telemetry first, then set hard SLOs from real workbook tests.

## Synthetic Core Benchmark

Run:

```bash
corepack pnpm benchmark:synthetic
```

This builds the repo, then runs `scripts/benchmark-synthetic.mjs` against compiled core modules. The benchmark emits JSON for:

- large value matrix chunking
- large formula matrix chunking
- batch compilation across many range writes

These numbers are useful for detecting obvious core regressions in chunking and compilation overhead. They are not real Excel SLOs because they do not include Office.js, workbook calculation, add-in websocket latency, Excel rendering, or platform-specific host behavior.

Real workbook SLOs should be calibrated from telemetry captured on macOS and Windows with representative workbook sizes before hard latency budgets are enforced in CI.
