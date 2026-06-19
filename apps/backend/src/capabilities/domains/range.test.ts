import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("range capability domain", () => {
  it("has colocated registry coverage for all range operations", () => {
    expectDomainContract("range", "apps/backend/src/capabilities/domains/range.test.ts");
  });
});
