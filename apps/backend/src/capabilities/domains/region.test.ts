import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("region capability domain", () => {
  it("has colocated registry coverage for all region operations", () => {
    expectDomainContract("region", "apps/backend/src/capabilities/domains/region.test.ts");
  });
});
