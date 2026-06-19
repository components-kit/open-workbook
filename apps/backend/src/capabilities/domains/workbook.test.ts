import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("workbook capability domain", () => {
  it("has colocated registry coverage for all workbook operations", () => {
    expectDomainContract("workbook", "apps/backend/src/capabilities/domains/workbook.test.ts");
  });
});
