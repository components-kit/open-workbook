import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("repair capability domain", () => {
  it("has colocated registry coverage for all repair operations", () => {
    expectDomainContract("repair", "apps/backend/src/capabilities/domains/repair.test.ts");
  });
});
