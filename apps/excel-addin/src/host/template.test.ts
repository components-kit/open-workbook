import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("template host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/template.test.ts");
    coveredHostFunction("template.capture");
    coveredHostFunction("template.capture_sheet");
    coveredHostFunction("template.repair");
  });
});
