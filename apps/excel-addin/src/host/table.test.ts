import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("table host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/table.test.ts");
    coveredHostFunction("table.list");
    coveredHostFunction("table.get_info");
    coveredHostFunction("table.read");
    coveredHostFunction("table.create");
    coveredHostFunction("table.resize");
    coveredHostFunction("table.reorder_columns");
    coveredHostFunction("table.append_rows");
    coveredHostFunction("table.update_rows");
    coveredHostFunction("table.clear_data_keep_formulas");
    coveredHostFunction("table.clear_filters");
    coveredHostFunction("table.apply_filters");
    coveredHostFunction("table.sort");
    coveredHostFunction("table.clear_sort");
    coveredHostFunction("table.set_total_row");
    coveredHostFunction("table.set_style");
    coveredHostFunction("table.copy_structure");
  });
});
