import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("batch plan capability domain", () => {
  it("has colocated registry coverage for batch, plan, job, and transaction operations", () => {
    expectDomainContract("batch", "apps/backend/src/capabilities/domains/batch-plan.test.ts");
    expectDomainContract("plan", "apps/backend/src/capabilities/domains/batch-plan.test.ts");
    expectDomainContract("job", "apps/backend/src/capabilities/domains/batch-plan.test.ts");
    expectDomainContract("transaction", "apps/backend/src/capabilities/domains/batch-plan.test.ts");
  });
});
