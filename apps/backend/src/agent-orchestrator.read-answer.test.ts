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

  it("summarizes large range requests without reading cell values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read Data!A1:XFD1048576",
        mode: "answer",
        target: { sheetName: "Data", range: "A1:XFD1048576" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("large_range_guard");
      expect((result.answer as any).requestedCells).toBeGreaterThan(10000);
      expect(result.warnings[0]).toContain("Large range read was summarized");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(result.telemetry.fullReadCellCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("asks callers to narrow full workbook cell dumps", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Print every cell in every sheet.", mode: "answer" });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("workbook_dump_guard");
      expect((result.answer as any).refusedFullCellDump).toBe(true);
      expect((result.answer as any).alternatives).toContain("specific range");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("returns workbook and sheet detail summaries from metadata without live cell reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const workbookSummary = await agent.run({
        request: "Summarize workbook",
        mode: "answer",
        detailLevel: "workbook_summary"
      });
      const sheetSummary = await agent.run({
        request: "Summarize Data sheet",
        mode: "answer",
        detailLevel: "sheet_summary",
        workbookContextId: workbookSummary.workbookContextId,
        target: { sheetName: "Data" }
      });

      expect((workbookSummary.answer as any).kind).toBe("workbook_summary");
      expect((sheetSummary.answer as any).kind).toBe("sheet_summary");
      expect((sheetSummary.answer as any).sheet.name).toBe("Data");
      expect(workbookSummary.telemetry.metadataDetailLevel).toBe("structure");
      expect(sheetSummary.telemetry.cacheHit).toBe(true);
      expect(runtime.readBatchCount).toBe(0);
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
      expect((result.answer as any).values).toBeUndefined();
      expect((result.answer as any).resultUri).toMatch(/^excel:\/\/agent\/results\/agentres_/);
      expect((result.answer as any).fullResultUri).toMatch(/\?view=full$/);
      expect(result.continuation).toMatchObject({
        workbookContextId: metadata.workbookContextId,
        resultUri: (result.answer as any).resultUri,
        fullResultUri: (result.answer as any).fullResultUri,
        responseMode: "brief"
      });
      expect((result.answer as any).projectedColumns.map((column: any) => column.name)).toEqual(["Amount", "Status"]);
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      expect((agent.getResultResource(resultId) as any).answer.values).toEqual([[123, "Open"], [456, "Closed"]]);
      const summary = agent.getResultResource(resultId, { view: "summary" }) as any;
      expect(summary.answer.values).toBeUndefined();
      expect(summary.answer.resultUri).toBe((result.answer as any).resultUri);
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(1);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("uses deterministic table sample and full table detail levels", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_table_detail_levels");
      agent.metadataCache.set(metadata);

      const sample = await agent.run({
        request: "Read Transactions table sample",
        mode: "answer",
        detailLevel: "table_sample",
        workbookContextId: metadata.workbookContextId,
        target: { tableName: "Transactions" }
      });
      const full = await agent.run({
        request: "Read full Transactions table",
        mode: "answer",
        detailLevel: "full_table",
        workbookContextId: metadata.workbookContextId,
        target: { tableName: "Transactions" }
      });

      expect((sample.answer as any).kind).toBe("table_compact_read");
      expect((sample.answer as any).rowLimit).toBe(20);
      expect((full.answer as any).kind).toBe("table_compact_read");
      expect((full.answer as any).rowLimit).toBe(10000);
      expect((full.answer as any).fullResultUri).toMatch(/\?view=full$/);
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(2);
    });

  it("preserves compact table row details in verbose mode", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_table_read_verbose");
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
        responseMode: "verbose",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { tableName: "Transactions" },
        values: { columns: ["Amount", "Status"], rowLimit: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).values).toEqual([[123, "Open"], [456, "Closed"]]);
    });
});
