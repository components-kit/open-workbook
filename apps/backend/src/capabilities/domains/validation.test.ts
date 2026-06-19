import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("validation capability domain", () => {
  it("has colocated registry coverage for all validation operations", () => {
    expectDomainContract("validation", "apps/backend/src/capabilities/domains/validation.test.ts");
  });
});
