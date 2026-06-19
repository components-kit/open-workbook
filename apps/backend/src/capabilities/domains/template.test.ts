import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("template capability domain", () => {
  it("has colocated registry coverage for all template operations", () => {
    expectDomainContract("template", "apps/backend/src/capabilities/domains/template.test.ts");
  });
});
