import { expect } from "vitest";
import { listBackendCapabilityRegistry } from "../registry.js";
import type { ExcelCapabilityGroup } from "../types.js";

export function expectDomainContract(group: ExcelCapabilityGroup, unitTestFile: string): void {
  const entries = listBackendCapabilityRegistry().filter((entry) => entry.group === group);
  expect(entries.length).toBeGreaterThan(0);
  expect(entries.every((entry) => entry.implementationOwner.length > 0)).toBe(true);
  expect(entries.every((entry) => entry.unitTestFile === unitTestFile)).toBe(true);
  expect(new Set(entries.map((entry) => entry.name)).size).toBe(entries.length);
}
