import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("runtime capability domain", () => {
  it("has colocated registry coverage for all runtime operations", () => {
    expectDomainContract("runtime", "apps/backend/src/capabilities/domains/runtime.test.ts");
    expectDomainContract("events", "apps/backend/src/capabilities/domains/runtime.test.ts");
  });
});
