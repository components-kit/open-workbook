import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("chart capability domain", () => {
  it("has colocated registry coverage for all chart operations", () => {
    expectDomainContract("chart", "apps/backend/src/capabilities/domains/chart.test.ts");
  });
});
