import { describe, expect, it } from "vitest";
import { AGENT_ACTION_HANDLERS, findAgentActionHandler } from "./agent-action-handlers.js";
import type { AgentRunInput } from "@components-kit/open-workbook-protocol";
import { getExcelCapability } from "./excel-capabilities.js";

describe("agent action handlers", () => {
  it("declares stable handler metadata", () => {
    expect(AGENT_ACTION_HANDLERS.length).toBeGreaterThanOrEqual(8);
    for (const handler of AGENT_ACTION_HANDLERS) {
      expect(handler.id).toMatch(/^[a-z0-9_]+$/);
      expect(handler.riskKind).toBeTruthy();
      expect(getExcelCapability(handler.capabilityName), handler.capabilityName).toBeTruthy();
      expect(typeof handler.matches).toBe("function");
    }
  });

  it("matches caller intent and natural language to the same handler", () => {
    const hinted: AgentRunInput = {
      request: "Do it",
      intent: { action: "format_range" },
      target: { sheetName: "Data", range: "A1:D1" }
    };
    const natural: AgentRunInput = {
      request: "Format the header row on Data",
      target: { sheetName: "Data", range: "A1:D1" }
    };

    expect(findAgentActionHandler(hinted, "format_range", true)?.id).toBe("format_range");
    expect(findAgentActionHandler(natural, undefined, true)?.id).toBe("format_range");
  });

  it("keeps workbook-level and target-level handlers separate", () => {
    const save: AgentRunInput = { request: "Save the workbook", intent: { action: "save" } };
    const filter: AgentRunInput = { request: "Add filters", intent: { action: "filter_range" }, target: { sheetName: "Data", range: "A1:D4" } };

    expect(findAgentActionHandler(save, "save", false)?.id).toBe("save_workbook");
    expect(findAgentActionHandler(save, "save", true)).toBeUndefined();
    expect(findAgentActionHandler(filter, "filter_range", false)).toBeUndefined();
    expect(findAgentActionHandler(filter, "filter_range", true)?.id).toBe("filter_range");
  });
});
