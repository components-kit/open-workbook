import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("range host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/range.test.ts");
    coveredHostFunction("range.read_hyperlinks");
    coveredHostFunction("range.read_comments");
    coveredHostFunction("range.read_notes");
    coveredHostFunction("range.read_merged_cells");
    coveredHostFunction("range.read_data_validation");
    coveredHostFunction("range.read_conditional_formatting");
    coveredHostFunction("range.search");
    coveredHostFunction("range.find_blank_cells");
    coveredHostFunction("range.find_errors");
  });
});
