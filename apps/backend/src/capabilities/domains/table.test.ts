import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("table capability domain", () => {
  it("has colocated registry coverage for all table operations", () => {
    expectDomainContract("table", "apps/backend/src/capabilities/domains/table.test.ts");
  });
});
