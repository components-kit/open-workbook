import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("workbook host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/workbook.test.ts");
    coveredHostFunction("workbook.get_info");
    coveredHostFunction("workbook.get_map");
    coveredHostFunction("workbook.calculate");
    coveredHostFunction("workbook.save");
    coveredHostFunction("workbook.get_file");
    coveredHostFunction("workbook.close");
    coveredHostFunction("workbook.snapshot_ranges");
    coveredHostFunction("workbook.embed_local_config");
    coveredHostFunction("workbook.read_embedded_local_config");
  });
});
