import { expect } from "vitest";
import { HOST_METHOD_REGISTRY } from "./registry.js";

export function expectHostDomainContract(unitTestFile: string): void {
  const entries = HOST_METHOD_REGISTRY.filter((entry) => entry.unitTestFile === unitTestFile);
  expect(entries.length).toBeGreaterThan(0);
  expect(entries.every((entry) => entry.implementationOwner.length > 0)).toBe(true);
  expect(entries.every((entry) => typeof entry.handler === "function")).toBe(true);
  expect(new Set(entries.map((entry) => entry.method)).size).toBe(entries.length);
}

export function coveredHostFunction(method: string): void {
  const entry = HOST_METHOD_REGISTRY.find((candidate) => candidate.method === method);
  expect(entry, method).toBeDefined();
  expect(typeof entry?.handler, method).toBe("function");
}
