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

  it("downgrades vague full table detail requests to sheet summaries", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Look into Data sheet, how is it?",
        mode: "answer",
        detailLevel: "full_table",
        target: { sheetName: "Data" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_summary");
      expect(result.warnings.join(" ")).toContain("full_table");
      expect(result.telemetry.fullReadCellCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("keeps explicit target range reads exact", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read exact values from A2:B2",
        mode: "answer",
        target: { sheetName: "Data", range: "A2:B2" },
        responseMode: "verbose"
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A2:B2" });
      expect((result.answer as any).shape).toEqual({ rows: 1, columns: 2 });
      expect((result.answer as any).valuesPreview).toEqual([["20/6/26", "20/6/26"]]);
      expect((result.answer as any).previewRange).toBe("A2:B2");
    });

  it("diagnoses date-like text formatting issues on the exact range", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Diagnose formatting error in A2:B2",
        mode: "answer",
        intent: { action: "format_diagnostics" },
        target: { sheetName: "Data", range: "A2:B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("format_diagnostics");
      expect((result.answer as any).range).toBe("A2:B2");
      expect((result.answer as any).issues.map((issue: any) => issue.code)).toContain("DATE_TEXT_TWO_DIGIT_YEAR");
      expect(result.telemetry.fullReadUsed).toBe(true);
    });

  it("returns current style summaries without falling back to value profiles", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "What is the current styling of A1:D1?",
        mode: "answer",
        intent: { action: "read_style_summary" },
        target: { sheetName: "Data", range: "A1:D1" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_summary");
      expect((result.answer as any).fills).toEqual({ hash: "fills" });
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A1:D1" });
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBe(1);
    });

  it("infers style summary reads from auto-mode styling questions", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the style summary of the Data sheet. Tell me about fonts, colors, borders, alignment, fills, and number formats.",
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_summary");
      expect((result.answer as any).fills).toEqual({ hash: "fills" });
      expect(result.nextAction).toBe("answer_now");
      expect(result.telemetry.routeMode).toBe("answer");
      expect(result.telemetry.workflowRoute).toBe("style.inspect");
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBe(1);
    });

  it("keeps compact style summaries useful and retrieves full result handles", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.getStyleFingerprint = async (request: any) => {
        runtime.runtimeMethodCalls["style.get_fingerprint"] = (runtime.runtimeMethodCalls["style.get_fingerprint"] ?? 0) + 1;
        return {
          ok: true,
          fingerprint: {
            workbookId: request.workbookId,
            sheetName: request.sheetName,
            address: request.address,
            rowCount: 7,
            columnCount: 24,
            dimensions: {
              fills: { hash: "fills", samples: Array.from({ length: 60 }, (_, index) => ({ index, color: `#00${String(index).padStart(4, "0")}` })) },
              fonts: { hash: "fonts", samples: Array.from({ length: 60 }, (_, index) => ({ index, name: "Aptos", bold: index === 0 })) },
              borders: { hash: "borders", samples: Array.from({ length: 60 }, (_, index) => ({ index, style: "Continuous" })) },
              alignment: { hash: "alignment", horizontal: "Center" },
              numberFormats: { hash: "numberFormats", samples: Array.from({ length: 60 }, (_, index) => ({ index, format: "General" })) }
            },
            warnings: []
          }
        };
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the existing styles on the Data sheet.",
        mode: "answer",
        intent: { action: "read_style_summary" },
        target: { sheetName: "Data", range: "A1:X7" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_summary");
      expect((result.answer as any).fills.hash).toBe("fills");
      expect((result.answer as any).fills.samples).toHaveLength(5);
      expect((result.answer as any).fills.samplesTruncated).toBe(true);
      expect((result.answer as any).resultUri).toMatch(/^excel:\/\/agent\/results\/agentres_/);
      expect(result.continuation?.fullResultUri).toMatch(/\?view=full$/);

      const callsAfterRead = runtime.runtimeMethodCalls["style.get_fingerprint"];
      const full = await agent.run({
        request: "I need the full uncompacted style summary for Data!A1:X7 including fonts, colors, borders, alignment, fills, and number formats.",
        continuation: {
          resultUri: result.continuation!.resultUri!,
          fullResultUri: result.continuation!.fullResultUri!,
          responseMode: "verbose"
        }
      });

      expect(full.status).toBe("SUCCESS");
      expect((full.answer as any).kind).toBe("agent_result_resource");
      expect((full.answer as any).view).toBe("full");
      expect((full.answer as any).result.answer.fills.samples).toHaveLength(60);
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBe(callsAfterRead);
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
      expect((result.answer as any).valuesPreview).toEqual([[123, "Open"], [456, "Closed"]]);
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
      const callsAfterRead = runtime.runtimeMethodCalls["table.read"];
      const continued = await agent.run({
        request: "Continue using the stored result",
        mode: "answer",
        continuation: result.continuation
      });
      expect(continued.status).toBe("SUCCESS");
      expect((continued.answer as any).kind).toBe("agent_result_resource");
      expect((continued.answer as any).view).toBe("summary");
      expect((continued.answer as any).result.answer.values).toBeUndefined();
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(callsAfterRead);
      const pastedFull = await agent.run({
        request: `Show full details from ${(result.answer as any).fullResultUri}`,
        mode: "answer"
      });
      expect((pastedFull.answer as any).view).toBe("full");
      expect((pastedFull.answer as any).result.answer.values).toEqual([[123, "Open"], [456, "Closed"]]);
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(callsAfterRead);
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
