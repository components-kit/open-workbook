import { getExposedToolCatalog } from "@components-kit/open-workbook-protocol";
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

    expect(capabilities.length).toBeGreaterThan(300);
    expect(getExcelCapability("excel.agent.run")).toBeTruthy();
    expect(getExcelCapability("excel.range.write_values")?.mutatesWorkbook).toBe(true);
    expect(getExcelCapability("excel.workflow.inspect_analyze")?.mutatesWorkbook).toBe(false);
  });

  it("keeps internal capabilities separate from the public MCP tool surface", () => {
    const summary = getExcelCapabilitySummary();
    const exposed = getExposedToolCatalog();

    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(summary.total).toBeGreaterThan(300);
    expect(summary.exposed).toBe(0);
    expect(summary.capabilities.some((capability) => capability.name === "excel.range.read_compact")).toBe(true);
  });

  it("keeps catalog metadata complete enough for grouped coverage planning", () => {
    const capabilities = listExcelCapabilities();
    const names = new Set<string>();

    for (const capability of capabilities) {
      expect(capability.name).toMatch(/^excel\.[a-z_]+\.[a-z0-9_]+$/);
      expect(names.has(capability.name)).toBe(false);
      names.add(capability.name);
      expect(capability.title.length).toBeGreaterThan(0);
      expect(capability.description.length).toBeGreaterThan(0);
      expect(capability.namespace.length).toBeGreaterThan(0);
      expect(capability.status).toMatch(/^(stable|preview|planned|unsupported)$/);
      expect(capability.destructiveLevel).toMatch(/^(none|values|format|structure|workbook)$/);
      expect(capability.requiresConfirmation).toBe(capability.mutatesWorkbook);
      expect(Array.isArray(capability.requiredCapabilities)).toBe(true);
    }

    expect(names.size).toBe(capabilities.length);
  });

  it("assigns every internal capability to one stable backend capability group", () => {
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

  it("records current agent status without expanding orchestration coverage", () => {
    const handlerCapabilities = new Set(AGENT_ACTION_HANDLERS.map((handler) => handler.capabilityName));
    const grouped = listExcelCapabilityGroups();
    const agentActionHandlerCount = grouped.reduce((total, group) => total + group.agentActionHandlers, 0);

    expect(getExcelCapabilityAgentStatus("excel.agent.run")).toBe("agent_entrypoint");
    expect(getExcelCapabilityAgentStatus("excel.range.write_values")).toBe("agent_action_handler");
    expect(getExcelCapabilityAgentStatus("excel.range.read_compact")).toBe("internal_capability");
    expect(agentActionHandlerCount).toBe(handlerCapabilities.size);
    expect(grouped.reduce((total, group) => total + group.agentEntrypoint, 0)).toBe(1);
  });

  it("assigns every internal capability a coverage planning status", () => {
    const capabilities = listExcelCapabilities();
    const coverage = listExcelCapabilityCoverage();
    const summary = summarizeExcelCapabilityCoverage();

    expect(coverage.length).toBe(capabilities.length);
    expect(summary.total).toBe(capabilities.length);
    expect(summary.byPlanningStatus.covered).toBe(new Set(AGENT_ACTION_HANDLERS.map((handler) => handler.capabilityName)).size + 1);
    expect(summary.byPlanningStatus.future_orchestration_candidate).toBeGreaterThan(0);
    expect(summary.byPlanningStatus.needs_unit_contract).toBeGreaterThan(0);
    expect(summary.byPlanningStatus.host_limited).toBeGreaterThan(0);
    expect(summary.byGroup.reduce((total, group) => total + group.total, 0)).toBe(capabilities.length);
    expect(coverage.find((entry) => entry.capability.name === "excel.range.write_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.copy")?.planningStatus).toBe("future_orchestration_candidate");
    expect(coverage.find((entry) => entry.capability.name === "excel.runtime.get_status")?.planningStatus).toBe("needs_unit_contract");
    expect(coverage.find((entry) => entry.capability.name === "excel.pivot.update_source")?.planningStatus).toBe("host_limited");
  });
});
