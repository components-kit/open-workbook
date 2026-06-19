import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("pivot capability domain", () => {
  it("has colocated registry coverage for all pivot operations", () => {
    expectDomainContract("pivot", "apps/backend/src/capabilities/domains/pivot.test.ts");
  });
});
