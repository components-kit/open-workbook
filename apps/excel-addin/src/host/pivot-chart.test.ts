import { describe, it } from "vitest";
import { coveredHostFunction, expectHostDomainContract } from "./test-support.js";

describe("pivot/chart host operations", () => {
  it("has colocated registry coverage", () => {
    expectHostDomainContract("apps/excel-addin/src/host/pivot-chart.test.ts");
    coveredHostFunction("pivot.list");
    coveredHostFunction("pivot.get_info");
    coveredHostFunction("pivot.create");
    coveredHostFunction("pivot.refresh");
    coveredHostFunction("pivot.refresh_all");
    coveredHostFunction("pivot.copy_from_template");
    coveredHostFunction("pivot.delete");
    coveredHostFunction("chart.list");
    coveredHostFunction("chart.get_info");
    coveredHostFunction("chart.create");
    coveredHostFunction("chart.update_data_source");
    coveredHostFunction("chart.copy_from_template");
    coveredHostFunction("chart.refresh");
    coveredHostFunction("chart.delete");
  });
});
