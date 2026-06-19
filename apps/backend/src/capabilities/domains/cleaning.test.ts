import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("cleaning capability domain", () => {
  it("has colocated registry coverage for all cleaning operations", () => {
    expectDomainContract("cleaning", "apps/backend/src/capabilities/domains/cleaning.test.ts");
  });
});
