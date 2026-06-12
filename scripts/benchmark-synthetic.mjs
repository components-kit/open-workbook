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
