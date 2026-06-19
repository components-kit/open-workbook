import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("worksheet capability domain", () => {
  it("has colocated registry coverage for all worksheet operations", () => {
    expectDomainContract("worksheet", "apps/backend/src/capabilities/domains/worksheet.test.ts");
  });
});
