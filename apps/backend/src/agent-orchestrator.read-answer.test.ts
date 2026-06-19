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

  it("uses compact table reads for table-target value answers", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_table_read");
      metadata.tables[0]!.columns = [
        { index: 0, letter: "A", name: "Date", normalizedName: "date", inferredType: "date" },
        { index: 1, letter: "B", name: "Account", normalizedName: "account", inferredType: "text" },
        { index: 2, letter: "C", name: "Amount", normalizedName: "amount", inferredType: "number" },
        { index: 3, letter: "D", name: "Status", normalizedName: "status", inferredType: "status" }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Read the first two Amount and Status rows from Transactions",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { tableName: "Transactions" },
        values: { columns: ["Amount", "Status"], rowLimit: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_compact_read");
      expect((result.answer as any).values).toEqual([[123, "Open"], [456, "Closed"]]);
      expect((result.answer as any).projectedColumns.map((column: any) => column.name)).toEqual(["Amount", "Status"]);
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(1);
      expect(runtime.readBatchCount).toBe(0);
    });
});
