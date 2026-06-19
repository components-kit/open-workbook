import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("snapshot capability domain", () => {
  it("has colocated registry coverage for all snapshot operations", () => {
    expectDomainContract("snapshot", "apps/backend/src/capabilities/domains/snapshot.test.ts");
  });
});
