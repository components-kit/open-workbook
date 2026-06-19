import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("lookup capability domain", () => {
  it("has colocated registry coverage for all lookup operations", () => {
    expectDomainContract("lookup", "apps/backend/src/capabilities/domains/lookup.test.ts");
  });
});
