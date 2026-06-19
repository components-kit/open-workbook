import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("style host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/style.test.ts");
    coveredHostFunction("style.capture_fingerprint");
    coveredHostFunction("style.copy_dimensions");
  });
});
