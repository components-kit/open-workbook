import { describe, it } from "vitest";
import { expectDomainContract } from "./test-support.js";

describe("backup capability domain", () => {
  it("has colocated registry coverage for all backup operations", () => {
    expectDomainContract("backup", "apps/backend/src/capabilities/domains/backup.test.ts");
  });
});
