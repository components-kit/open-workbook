import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Read Answer Routing", () => {
  it("fails fast with reload guidance when the Excel add-in session is stale", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.readiness = {
        ok: false,
        connectionState: "stale",
        status: {
          activeAddinConnected: false,
          connectionState: "stale",
          activeWorkbookAvailable: false,
          sessions: [{ connectionId: "conn_stale", stale: true }]
        }
      };
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({ request: "Read headers and first 5 rows from Data A1:D4", mode: "answer" });
  
      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.summary).toContain("stale Excel add-in session");
      expect(result.warnings[0]).toContain("Reload or reopen the OpenWorkbook Local taskpane");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("expands selected cell to its used-range column for selected column prompts", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C2", { row: 2, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({ request: "Analyze this column.", mode: "answer" });
  
      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "C1:C4", label: "selected column" });
      expect(result.telemetry.internalReadCount).toBe(1);
    });
});
