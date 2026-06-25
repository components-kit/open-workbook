import { getPublicAgentToolCatalog } from "@components-kit/open-workbook-protocol";
import { describe, expect, it } from "vitest";
import { AGENT_ACTION_HANDLERS } from "./agent-action-handlers.js";
import {
  EXCEL_CAPABILITY_GROUPS,
  getExcelCapability,
  getExcelCapabilityAgentStatus,
  getExcelCapabilityGroup,
  getExcelCapabilitySummary,
  listExcelCapabilityCoverage,
  listExcelCapabilities,
  listExcelCapabilityGroups,
  summarizeExcelCapabilityCoverage
} from "./excel-capabilities.js";

describe("excel capabilities", () => {
  it("keeps Excel operations available as internal backend capabilities", () => {
    const capabilities = listExcelCapabilities();

    expect(capabilities.length).toBe(308);
    expect(getExcelCapability("excel.agent.run")).toBeTruthy();
    expect(getExcelCapability("excel.range.write_values")?.mutatesWorkbook).toBe(true);
    expect(getExcelCapability("excel.range.write_data_validation")?.mutatesWorkbook).toBe(true);
    expect(getExcelCapability("excel.range.write_conditional_formatting")?.mutatesWorkbook).toBe(true);
    expect(getExcelCapability("excel.range.reorder_columns")?.mutatesWorkbook).toBe(true);
    expect(getExcelCapability("excel.range.hide_columns")?.destructiveLevel).toBe("structure");
    expect(getExcelCapability("excel.range.unhide_columns")?.destructiveLevel).toBe("structure");
    expect(getExcelCapability("excel.workflow.inspect_analyze")?.mutatesWorkbook).toBe(false);
  });

  it("does not advertise executor-limited operations as stable", () => {
    for (const capabilityName of ["excel.range.write_hyperlinks", "excel.range.write_comments", "excel.sheet.move"]) {
      const capability = getExcelCapability(capabilityName);
      if (!capability) {
        expect(listExcelCapabilityCoverage().some((entry) => entry.capability.name === capabilityName), capabilityName).toBe(false);
        continue;
      }
      expect(capability.status, capabilityName).toBe("planned");
      expect(listExcelCapabilityCoverage().find((entry) => entry.capability.name === capabilityName)?.planningStatus, capabilityName).toBe("defer");
    }
  });

  it("keeps internal capabilities separate from the public MCP tool surface", () => {
    const summary = getExcelCapabilitySummary();
    const exposed = getPublicAgentToolCatalog();

    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(summary.total).toBe(308);
    expect(summary.exposed).toBe(0);
    expect(summary.capabilities.some((capability) => capability.name === "excel.range.read_compact")).toBe(true);
  });

  it("keeps catalog metadata complete enough for grouped planning", () => {
    const capabilities = listExcelCapabilities();
    const names = new Set<string>();

    for (const capability of capabilities) {
      expect(capability.name).toMatch(/^excel\.[a-z_]+\.[a-z0-9_]+$/);
      expect(names.has(capability.name), capability.name).toBe(false);
      names.add(capability.name);
      expect(capability.title.length, capability.name).toBeGreaterThan(0);
      expect(capability.description.length, capability.name).toBeGreaterThan(0);
      expect(capability.namespace.length, capability.name).toBeGreaterThan(0);
      expect(capability.status, capability.name).toMatch(/^(stable|preview|planned|unsupported)$/);
      expect(capability.destructiveLevel, capability.name).toMatch(/^(none|values|format|structure|workbook)$/);
      expect(capability.requiresConfirmation, capability.name).toBe(capability.mutatesWorkbook);
      expect(Array.isArray(capability.requiredCapabilities), capability.name).toBe(true);
    }

    expect(names.size).toBe(capabilities.length);
  });

  it("assigns every internal capability to one backend capability group", () => {
    const capabilities = listExcelCapabilities();
    const grouped = listExcelCapabilityGroups();
    const groupedCount = grouped.reduce((total, group) => total + group.capabilities.length, 0);
    const groupedNames = new Set(grouped.flatMap((group) => group.capabilities.map((capability) => capability.name)));

    expect(groupedCount).toBe(capabilities.length);
    expect(groupedNames.size).toBe(capabilities.length);
    expect(EXCEL_CAPABILITY_GROUPS.every((group) => group.label.length > 0 && group.description.length > 0)).toBe(true);
    expect(getExcelCapabilityGroup("excel.range.write_values")).toBe("range");
    expect(getExcelCapabilityGroup("excel.workflow.inspect_analyze")).toBe("workflow");
    expect(getExcelCapabilityGroup("excel.snapshot.get_compact")).toBe("snapshot");
    expect(getExcelCapabilityGroup("excel.backup.prune")).toBe("backup");
    expect(getExcelCapabilityGroup("excel.permissions.set_scope")).toBe("permissions");
  });

  it("records current agent status without treating it as operation coverage", () => {
    const grouped = listExcelCapabilityGroups();
    const agentActionHandlerCount = grouped.reduce((total, group) => total + group.agentActionHandlers, 0);

    expect(getExcelCapabilityAgentStatus("excel.agent.run")).toBe("agent_entrypoint");
    expect(getExcelCapabilityAgentStatus("excel.range.write_values")).toBe("agent_action_handler");
    expect(getExcelCapabilityAgentStatus("excel.table.get_schema")).toBe("agent_action_handler");
    expect(getExcelCapabilityAgentStatus("excel.range.read_compact")).toBe("agent_action_handler");
    expect(agentActionHandlerCount).toBeGreaterThan(new Set(AGENT_ACTION_HANDLERS.map((handler) => handler.capabilityName)).size);
    expect(grouped.reduce((total, group) => total + group.agentEntrypoint, 0)).toBe(1);
  });

  it("summarizes planning status without asserting behavior coverage", () => {
    const capabilities = listExcelCapabilities();
    const coverage = listExcelCapabilityCoverage();
    const summary = summarizeExcelCapabilityCoverage();

    expect(coverage.length).toBe(capabilities.length);
    expect(summary.total).toBe(capabilities.length);
    expect(summary.byGroup.reduce((total, group) => total + group.total, 0)).toBe(capabilities.length);
    expect(Object.values(summary.byPlanningStatus).reduce((total, count) => total + count, 0)).toBe(capabilities.length);
  });
});
