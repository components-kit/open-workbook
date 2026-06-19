import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("permissions capability domain", () => {
  it("has colocated registry coverage for all permissions operations", () => {
    expectDomainContract("permissions", "apps/backend/src/capabilities/domains/permissions.test.ts");
  });
});
