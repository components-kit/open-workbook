import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("runtime host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/runtime.test.ts");
    coveredHostFunction("runtime.ping");
    coveredHostFunction("runtime.get_active_context");
    coveredHostFunction("runtime.get_selection");
    coveredHostFunction("runtime.set_active_sheet");
  });
});
