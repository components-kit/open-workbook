import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { getPublicAgentToolCatalog } from "@components-kit/open-workbook-protocol";
import { describe, expect, it } from "vitest";
import { AGENT_ACTION_HANDLERS } from "../agent-action-handlers.js";
import { RuntimeService } from "../runtime-service.js";
import {
  EXCEL_CAPABILITY_GROUPS,
  listBackendCapabilityRegistry,
  listExcelCapabilities
} from "./registry.js";

const repoRoot = path.resolve(process.cwd(), "../..");

describe("backend capability registry", () => {
  it("maps every cataloged capability to exactly one backend registry entry", () => {
    const catalog = listExcelCapabilities();
    const registry = listBackendCapabilityRegistry();

    expect(catalog).toHaveLength(306);
    expect(registry).toHaveLength(catalog.length);
    expect(new Set(registry.map((entry) => entry.name)).size).toBe(catalog.length);
    expect(registry.map((entry) => entry.name).sort()).toEqual(catalog.map((capability) => capability.name).sort());
  });

  it("keeps registry domain metadata complete and backed by colocated tests", () => {
    const validGroups = new Set(EXCEL_CAPABILITY_GROUPS.map((definition) => definition.group));
    for (const entry of listBackendCapabilityRegistry()) {
      expect(validGroups.has(entry.group)).toBe(true);
      expect(entry.implementationOwner.length).toBeGreaterThan(0);
      expect(entry.unitTestFile.endsWith(".test.ts")).toBe(true);
      expect(existsSync(path.resolve(repoRoot, entry.unitTestFile))).toBe(true);
    }
  });

  it("references real runtime facade methods and agent action handlers", () => {
    const runtime = new RuntimeService({ persistState: false });
    const handlerIds = new Set(AGENT_ACTION_HANDLERS.map((handler) => handler.id));

    for (const entry of listBackendCapabilityRegistry()) {
      if (entry.runtimeMethod) {
        expect(typeof (runtime as any)[entry.runtimeMethod], entry.name).toBe("function");
      }
      for (const handlerId of entry.agentHandlerIds) {
        expect(handlerIds.has(handlerId), entry.name).toBe(true);
      }
    }
  });

  it("keeps the public MCP surface to the single agent run tool", () => {
    expect(getPublicAgentToolCatalog().map((tool) => tool.name)).toEqual(["excel.agent.run"]);
  });

  it("declares only add-in host methods that exist in the add-in host registry", () => {
    const source = readFileSync(path.resolve(repoRoot, "apps/excel-addin/src/host/registry.ts"), "utf8");
    const hostMethods = new Set([...source.matchAll(/(?:method: |\[)"([a-z]+(?:\.[a-z_]+)+)"/g)].map((match) => match[1]));
    for (const entry of listBackendCapabilityRegistry()) {
      for (const hostMethod of entry.hostMethods) {
        expect(hostMethods.has(hostMethod), `${entry.name} -> ${hostMethod}`).toBe(true);
      }
    }
  });
});
