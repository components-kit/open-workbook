import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("batch host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/batch.test.ts");
    coveredHostFunction("operation.execute_batch");
  });
});
