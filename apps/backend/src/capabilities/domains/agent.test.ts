import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("agent capability domain", () => {
  it("has colocated registry coverage for all agent operations", () => {
    expectDomainContract("agent", "apps/backend/src/capabilities/domains/agent.test.ts");
  });
});
