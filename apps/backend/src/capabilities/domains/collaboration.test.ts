import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("collaboration capability domain", () => {
  it("has colocated registry coverage for collaboration, task, lock, and conflict operations", () => {
    expectDomainContract("collaboration", "apps/backend/src/capabilities/domains/collaboration.test.ts");
    expectDomainContract("task", "apps/backend/src/capabilities/domains/collaboration.test.ts");
    expectDomainContract("lock", "apps/backend/src/capabilities/domains/collaboration.test.ts");
    expectDomainContract("conflict", "apps/backend/src/capabilities/domains/collaboration.test.ts");
  });
});
