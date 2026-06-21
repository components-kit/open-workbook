import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { BATCH_OPERATION_KINDS, getHostMethod, HOST_METHOD_NAMES, HOST_METHOD_REGISTRY } from "./registry.js";

const repoRoot = path.resolve(process.cwd(), "../..");

describe("add-in host method registry", () => {
  it("keeps every JSON-RPC method unique and backed by a handler", () => {
    expect(HOST_METHOD_REGISTRY.length).toBeGreaterThan(0);
    expect(new Set(HOST_METHOD_NAMES).size).toBe(HOST_METHOD_NAMES.length);
    for (const entry of HOST_METHOD_REGISTRY) {
      expect(getHostMethod(entry.method)?.handler).toBe(entry.handler);
      expect(typeof entry.handler).toBe("function");
      expect(entry.implementationOwner.length).toBeGreaterThan(0);
    }
  });

  it("tracks the central registry contract test for every host method", () => {
    for (const entry of HOST_METHOD_REGISTRY) {
      expect(entry.unitTestFile).toBe("apps/excel-addin/src/host/registry.test.ts");
      expect(existsSync(path.resolve(repoRoot, entry.unitTestFile)), entry.method).toBe(true);
    }
  });

  it("documents batch operation kinds handled by the Office.js executor", () => {
    expect(BATCH_OPERATION_KINDS).toContain("range.write_values");
    expect(BATCH_OPERATION_KINDS).toContain("sheet.create");
    expect(BATCH_OPERATION_KINDS).toContain("template.create_sheet_from_template");
    expect(getHostMethod("operation.execute_batch")?.operationKinds).toEqual([...BATCH_OPERATION_KINDS]);
  });

  it("keeps every batch operation kind represented in executor core", () => {
    const source = readFileSync(path.resolve(repoRoot, "apps/excel-addin/src/host/executor-core.ts"), "utf8");
    for (const kind of BATCH_OPERATION_KINDS) {
      expect(source.includes(`case "${kind}"`), kind).toBe(true);
    }
    expect(source).toContain("range.write_comments");
    expect(source).toContain("sheet.move");
    expect(source).toContain("OPERATION_NOT_SUPPORTED");
    expect(source).toContain("warning.code === \"OPERATION_NOT_SUPPORTED\"");
  });
});
