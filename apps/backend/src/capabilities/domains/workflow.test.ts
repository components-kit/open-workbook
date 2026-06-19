import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("workflow capability domain", () => {
  it("has colocated registry coverage for all workflow operations", () => {
    expectDomainContract("workflow", "apps/backend/src/capabilities/domains/workflow.test.ts");
  });
});
