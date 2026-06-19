import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("formula capability domain", () => {
  it("has colocated registry coverage for all formula operations", () => {
    expectDomainContract("formula", "apps/backend/src/capabilities/domains/formula.test.ts");
  });
});
