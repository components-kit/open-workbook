import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const OPERATION_KIND_RE = /"(range|sheet|workbook|template)\.[a-z_]+"/g;
const CASE_RE = /case\s+"((?:range|sheet|workbook|template)\.[a-z_]+)"/g;
const UNSUPPORTED_EXECUTOR_OPERATIONS = ["range.write_hyperlinks", "range.write_comments", "sheet.move"];

describe("Excel operation matrix", () => {
  it("keeps protocol, batch compiler, and add-in executor operation coverage aligned", () => {
    const protocolKinds = operationKinds(readFileSync(new URL("../../protocol/src/operations.ts", import.meta.url), "utf8"));
    const compilerCases = caseKinds(readFileSync(new URL("./range/batch-compiler.ts", import.meta.url), "utf8"));
    const executorSource = readFileSync(new URL("../../../apps/excel-addin/src/host/executor-core.ts", import.meta.url), "utf8");
    const executorCases = caseKinds(executorSource);

    expect([...compilerCases].sort()).toEqual([...protocolKinds].sort());
    expect([...protocolKinds].filter((kind) => !executorCases.has(kind))).toEqual([]);
    for (const kind of UNSUPPORTED_EXECUTOR_OPERATIONS) {
      expect(protocolKinds.has(kind), kind).toBe(true);
      expect(executorCases.has(kind), kind).toBe(true);
    }
    expect(executorSource).toContain("OPERATION_NOT_SUPPORTED");
    expect(executorSource).toContain("ok: !warnings.some");
  });
});

function operationKinds(source: string): Set<string> {
  const operationDefinitions = source.slice(0, source.indexOf("export type ExcelOperation"));
  return new Set([...operationDefinitions.matchAll(OPERATION_KIND_RE)].map((match) => match[0].slice(1, -1)));
}

function caseKinds(source: string): Set<string> {
  return new Set([...source.matchAll(CASE_RE)].map((match) => match[1]!));
}
