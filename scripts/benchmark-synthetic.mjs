#!/usr/bin/env node
import { performance } from "node:perf_hooks";
import { BatchCompiler, chunkMatrixRows } from "../packages/excel-core/dist/index.js";

const workbookId = "workbook_benchmark";
const scenarios = [
  {
    name: "chunk_100k_x_10_values",
    run: () => {
      const matrix = makeMatrix(10_000, 10, "value");
      const chunks = timed(() => chunkMatrixRows(matrix, 50_000));
      return {
        durationMs: chunks.durationMs,
        rows: 10_000,
        columns: 10,
        cells: 100_000,
        chunkCount: chunks.value.length,
        firstChunkRows: chunks.value[0]?.rows.length ?? 0
      };
    }
  },
  {
    name: "chunk_250k_x_25_formulas",
    run: () => {
      const matrix = makeMatrix(10_000, 25, "=SUM(A1:A10)");
      const chunks = timed(() => chunkMatrixRows(matrix, 50_000));
      return {
        durationMs: chunks.durationMs,
        rows: 10_000,
        columns: 25,
        cells: 250_000,
        chunkCount: chunks.value.length,
        firstChunkRows: chunks.value[0]?.rows.length ?? 0
      };
    }
  },
  {
    name: "compile_500_range_writes",
    run: () => {
      const compiler = new BatchCompiler({ now: () => "2026-01-01T00:00:00.000Z" });
      const operations = Array.from({ length: 500 }, (_unused, index) => writeValuesOperation(index));
      const compiled = timed(() => compiler.compile({ workbookId, mode: "apply", operations }));
      return {
        durationMs: compiled.durationMs,
        operationCount: operations.length,
        estimatedCellsTouched: compiled.value.estimatedCellsTouched,
        targetFingerprintCount: compiled.value.targetFingerprints.length,
        backupKinds: compiled.value.requiredBackups
      };
    }
  },
  {
    name: "compact_large_range_read",
    run: () => compactScenario("range_read", makeRangePayload(5_000, 12), {
      ok: true,
      contextId: "compact_range_read",
      resourceUri: "excel://compact/compact_range_read",
      summary: "Read 5,000 rows x 12 columns. Full matrix stored locally.",
      shape: { rows: 5_000, columns: 12 },
      sampleRows: 5,
      nextActionRecommendation: "answer_now"
    })
  },
  {
    name: "compact_validation_failure",
    run: () => compactScenario("validation_failure", makeValidationPayload(250), {
      ok: false,
      contextId: "compact_validation_failure",
      resourceUri: "excel://compact/compact_validation_failure",
      issueCount: 250,
      examples: makeValidationPayload(5).issues,
      nextActionRecommendation: "fetch_more_context"
    })
  },
  {
    name: "compact_mutate_validate_diff",
    run: () => compactScenario("mutate_validate_diff", makeMutationPayload(1_000), {
      ok: true,
      contextId: "compact_mutation",
      resourceUri: "excel://compact/compact_mutation",
      changedRanges: ["Data!D2:D1001"],
      changedCells: 1_000,
      validationSummary: { ok: true, issueCount: 0 },
      diffSummary: { cellsChanged: 1_000, formulasChanged: 0, stylesChanged: 0 },
      nextActionRecommendation: "answer_now"
    })
  }
];

const results = scenarios.map((scenario) => ({
  name: scenario.name,
  ...scenario.run()
}));

const summary = {
  generatedAt: new Date().toISOString(),
  node: process.version,
  kind: "synthetic",
  note: "Synthetic core benchmarks do not measure Office.js or real Excel latency.",
  results
};

console.log(JSON.stringify(summary, null, 2));

function timed(fn) {
  const started = performance.now();
  const value = fn();
  return {
    value,
    durationMs: Number((performance.now() - started).toFixed(3))
  };
}

function makeMatrix(rows, columns, value) {
  return Array.from({ length: rows }, () => Array.from({ length: columns }, () => value));
}

function writeValuesOperation(index) {
  const startRow = index * 20 + 1;
  const endRow = startRow + 19;
  return {
    kind: "range.write_values",
    operationId: `op_bench_${index}`,
    workbookId,
    destructiveLevel: "values",
    reason: "Synthetic benchmark range write",
    target: {
      workbookId,
      sheetName: "Data",
      address: `A${startRow}:J${endRow}`
    },
    values: makeMatrix(20, 10, index),
    preserveFormats: true
  };
}

function compactScenario(name, fullPayload, compactPayload) {
  const fullBytes = byteLength(fullPayload);
  const compact = {
    ...compactPayload,
    detailLevel: "summary",
    responseMode: "brief",
    telemetry: {
      responseBytes: byteLength(compactPayload),
      estimatedResponseTokens: Math.ceil(byteLength(compactPayload) / 4),
      storedPayloadBytes: fullBytes,
      estimatedStoredTokens: Math.ceil(fullBytes / 4),
      estimatedTokensSaved: Math.max(0, Math.ceil(fullBytes / 4) - Math.ceil(byteLength(compactPayload) / 4)),
      cacheHit: false
    }
  };
  const responseBytes = byteLength(compact);
  const maxResponseBytes = 4_000;
  if (responseBytes > maxResponseBytes) {
    throw new Error(`${name} compact benchmark exceeded ${maxResponseBytes} bytes: ${responseBytes}`);
  }
  return {
    responseBytes,
    estimatedResponseTokens: Math.ceil(responseBytes / 4),
    fullPayloadBytes: fullBytes,
    estimatedFullPayloadTokens: Math.ceil(fullBytes / 4),
    estimatedTokensSaved: Math.max(0, Math.ceil(fullBytes / 4) - Math.ceil(responseBytes / 4)),
    compressionRatio: Number((responseBytes / fullBytes).toFixed(6)),
    maxResponseBytes
  };
}

function makeRangePayload(rows, columns) {
  return {
    workbookId,
    sheetName: "Data",
    address: `A1:${columnName(columns)}${rows}`,
    values: makeMatrix(rows, columns, "invoice-value"),
    text: makeMatrix(rows, columns, "invoice text")
  };
}

function makeValidationPayload(issueCount) {
  return {
    ok: false,
    issues: Array.from({ length: issueCount }, (_unused, index) => ({
      range: `Data!D${index + 2}`,
      severity: index % 3 === 0 ? "error" : "warning",
      category: index % 2 === 0 ? "formula" : "blankValues",
      message: `Synthetic validation issue ${index + 1}`
    }))
  };
}

function makeMutationPayload(changedCells) {
  return {
    ok: true,
    transactionId: "txn_benchmark",
    beforeSnapshot: makeRangePayload(changedCells, 4),
    afterSnapshot: makeRangePayload(changedCells, 4),
    diff: {
      changedRanges: ["Data!D2:D1001"],
      cellsChanged: changedCells,
      changes: Array.from({ length: changedCells }, (_unused, index) => ({
        address: `Data!D${index + 2}`,
        before: index,
        after: index + 1
      }))
    },
    validation: { ok: true, issueCount: 0 }
  };
}

function byteLength(value) {
  return Buffer.byteLength(JSON.stringify(value), "utf8");
}

function columnName(index) {
  let name = "";
  let value = index;
  while (value > 0) {
    const modulo = (value - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    value = Math.floor((value - modulo) / 26);
  }
  return name;
}
