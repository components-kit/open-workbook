import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("formula host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/formula.test.ts");
    coveredHostFunction("formula.read_patterns");
    coveredHostFunction("formula.copy_patterns");
    coveredHostFunction("formula.fill_pattern");
    coveredHostFunction("formula.convert_to_values");
  });
});
