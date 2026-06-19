import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("style capability domain", () => {
  it("has colocated registry coverage for all formatting operations", () => {
    expectDomainContract("formatting", "apps/backend/src/capabilities/domains/style.test.ts");
  });
});
