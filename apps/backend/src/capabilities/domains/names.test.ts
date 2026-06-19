import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("names capability domain", () => {
  it("has colocated registry coverage for all names operations", () => {
    expectDomainContract("names", "apps/backend/src/capabilities/domains/names.test.ts");
  });
});
