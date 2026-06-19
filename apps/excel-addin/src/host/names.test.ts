import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("names host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/names.test.ts");
    coveredHostFunction("names.list");
    coveredHostFunction("names.get");
    coveredHostFunction("names.create");
    coveredHostFunction("names.update");
    coveredHostFunction("names.delete");
  });
});
