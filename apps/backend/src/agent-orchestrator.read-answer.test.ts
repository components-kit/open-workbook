import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets, workbookId } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Read Answer Routing", () => {
  it("answers style overview when caller omits mode but provides structured style_overview intent", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "style overview of Data sheet",
        intent: { action: "style_overview" },
        target: { sheetName: "Data" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.mode).toBe("auto");
      expect((result.answer as any).kind).toBe("style_overview");
      expect(result.nextAction).toBe("answer_now");
      expect(result.operationId).toBeUndefined();
    });

  it("summarizes grouped header row 1 without chasing design overview or full results", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_grouped_header_summary");
      metadata.workbook = { ...metadata.workbook, activeSheet: "Invoices" };
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Invoices",
          name: "Invoices",
          index: 3,
          usedRange: "A1:O1002",
          rowCount: 1002,
          columnCount: 15,
          kind: "transaction",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Okay, can you look at grouped header at row 1 on Invoices, please summarize it",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Invoices", range: "A1:O1002" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.nextAction).toBe("answer_now");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect((result.answer as any)).toMatchObject({
        kind: "grouped_header_summary",
        sheetName: "Invoices",
        range: "A1:O1",
        mergedRangeCount: 3,
        mergeStatus: "merged_spans_detected",
        spans: [
          { range: "A1:B1", label: "สถานะ", merged: true },
          { range: "C1:F1", label: "ข้อมูลการจอง", merged: true },
          { range: "G1:N1", label: "ค่าใช้จ่าย", merged: true }
        ],
        unmergedLabels: [
          { cell: "O1", label: "งานจ้างช่วง", merged: false }
        ]
      });
      expect(runtime.runtimeMethodCalls["range.read_merged_cells"]).toBe(1);
      expect(runtime.readBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toEqual([
        expect.objectContaining({
          kind: "range.read_full",
          target: expect.objectContaining({ sheetName: "Invoices", address: "A1:O1" })
        })
      ]);
    });

  it("updates permission policy through public agent intents", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const updated = await agent.run({
        request: "Allow workbook styling structure changes for this workbook",
        intent: { action: "set_permissions" },
        values: {
          permissions: {
            allowWrites: true,
            allowDestructiveActions: true,
            scopeToWorkbook: true,
            requireConfirmationFor: []
          }
        }
      });
      const readBack = await agent.run({
        request: "Read permissions",
        intent: { action: "get_permissions" }
      });

      expect(updated.status).toBe("SUCCESS");
      expect((updated.answer as any).kind).toBe("permissions_update");
      expect((updated.answer as any).result.permissions).toMatchObject({
        allowWrites: true,
        allowDestructiveActions: true,
        scope: { workbookId }
      });
      expect(readBack.status).toBe("SUCCESS");
      expect((readBack.answer as any).result.permissions.allowDestructiveActions).toBe(true);
    });

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

  it("prefers the active selected range for generic read requests", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = {
        ...selectionInfo("Data", "A2:D3", { row: 2, column: 1 }),
        endCell: { ...selectionInfo("Data", "A2:D3", { row: 3, column: 4 }).endCell },
        rowCount: 2,
        columnCount: 4,
        cellCount: 8,
        isSingleCell: false
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Can you check this?", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A2:D3", label: "selected range" });
      expect(runtime.runtimeMethodCalls["table.read"]).toBeUndefined();
      expect(runtime.lastSnapshotRanges[0]).toMatchObject({ sheetName: "Data", address: "A2:D3" });
      expect((result.answer as any).rowMetadata[0]).toMatchObject({ rowIndex: 0, sheetRowNumber: 2, address: "A2:D2" });
      expect((result.answer as any).rows).toEqual([
        ["2026-06-01", "A-100", 123, "Open"],
        ["2026-06-02", "A-101", 456, "Closed"]
      ]);
    });

  it("uses the active table row for generic reads from a single selected table cell", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "What do you think about this?", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A3:D3", label: "active table row" });
      expect(runtime.runtimeMethodCalls["table.read"]).toBeUndefined();
      expect(runtime.lastSnapshotRanges[0]).toMatchObject({ sheetName: "Data", address: "A3:D3" });
      expect((result.answer as any).rowMetadata[0]).toMatchObject({ rowIndex: 0, sheetRowNumber: 3, address: "A3:D3" });
      expect((result.answer as any).rows).toEqual([["2026-06-02", "A-101", 456, "Closed"]]);
    });

  it("does not let an incidental selected cell hijack workbook overview requests", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "What is this workbook about?", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("workbook_overview");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("does not let an incidental selected cell hijack worksheet overview requests", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Look at the Data worksheet", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_summary");
      expect((result.answer as any).sheet.name).toBe("Data");
      expect(result.taskOutcome).toBe("final_answer");
      expect(result.finalAnswer).toContain("Data");
      expect(result.finalAnswer).toContain("Transactions");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("keeps sheet overview useful under a tight payload budget", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Look at the Data worksheet",
        mode: "answer",
        budget: { maxPayloadBytes: 1800, maxExamples: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_summary");
      expect((result.answer as any).sheet.name).toBe("Data");
      expect((result.answer as any).tables[0]).toMatchObject({
        name: "Transactions",
        range: "A1:D4"
      });
      expect(result.finalAnswer).toContain("Transactions");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("uses an inferred record row for vague prompts from a single selected cell in a header-shaped sheet", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Apr 2026", "D2", { row: 2, column: 4 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "what do you think about this?", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Apr 2026", range: "A2:AJ2", label: "active record row" });
      expect(result.proof[1]).toMatchObject({ sheetName: "Apr 2026", range: "D2", label: "current Excel selection" });
      expect(runtime.lastSnapshotRanges[0]).toMatchObject({ sheetName: "Apr 2026", address: "A2:AJ2" });
      expect((result.answer as any).rowMetadata[0]).toMatchObject({ rowIndex: 0, sheetRowNumber: 2, address: "A2:AJ2" });
      expect((result.answer as any).rows[0][3]).toBe("Company gas top-up");
    });

  it("uses a small active-cell neighborhood outside tables for here-context reads", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Report", "B5", { row: 5, column: 2 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Can you check here?", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Report", range: "A4:B6", label: "active cell neighborhood" });
      expect(runtime.lastSnapshotRanges[0]).toMatchObject({ sheetName: "Report", address: "A4:B6" });
    });

  it("keeps explicit table reads broad instead of overriding with selection", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read Transactions table",
        mode: "answer",
        target: { tableName: "Transactions" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_compact_read");
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(1);
      expect((result.answer as any).rowMetadata[0]).toMatchObject({ rowIndex: 0, sheetRowNumber: 2, address: "A2:D2" });
    });

  it("does not let active selection hijack table names mentioned in the request", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read all data from the Transactions table and show the first 20 rows",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_compact_read");
      expect((result.answer as any).tableName).toBe("Transactions");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A2:D4", label: "Transactions" });
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(1);
      expect(runtime.lastSnapshotRanges).toEqual([]);
    });

  it("analyzes a reference sheet as compact patterns instead of row chunks", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Please skim quickly by tx type from Apr 2026, count types, explain relation between Desc, Tx Type, Detail Note, Cash Amount, Actual Amount, formulas and header styling.",
        mode: "answer",
        intent: { action: "analyze_reference_sheet" },
        target: { sheetName: "Apr 2026" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("reference_sheet_analysis");
      expect((result.answer as any).reference).toMatchObject({ sheetName: "Apr 2026", range: "A1:AJ244" });
      expect((result.answer as any).columnProfiles.some((profile: any) => profile.name === "Transaction Type" && profile.uniqueCount >= 1)).toBe(true);
      expect((result.answer as any).relationships.some((relationship: any) => relationship.columns.includes("Description") && relationship.columns.includes("Transaction Type"))).toBe(true);
      expect((result.answer as any).formulaPatterns.some((pattern: any) => pattern.column === "Payment Variance")).toBe(true);
      expect((result.answer as any).stylePatterns.some((style: any) => style.label === "header_style")).toBe(true);
      expect(result.agentInstruction).toContain("Do not broad-read or chunk-read");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.read_full",
        workbookId,
        target: { sheetName: "Apr 2026", address: "A1:AJ244" }
      });
      expect(runtime.snapshotRangesHistory).toEqual([]);
    });

  it("auto-routes reference convention prompts to reference analysis", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Look into Apr 2026 as reference and learn the same conventions for notes, formulas, header styling, and row conditions.",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("reference_sheet_analysis");
      expect((result.answer as any).objectives).toEqual(expect.arrayContaining(["formula patterns", "style conventions", "row conditions"]));
      expect(result.maxRecommendedFollowupCalls).toBe(0);
    });

  it("uses row windows from table read requests instead of the active selected row", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read rows 2-3 from the Transactions table. Show actual values.",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_compact_read");
      expect((result.answer as any).rowOffset).toBe(1);
      expect((result.answer as any).rowLimit).toBe(2);
      expect((result.answer as any).valuesPreview).toEqual([
        ["2026-06-02", "A-101", 456, "Closed"],
        ["2026-06-03", "A-102", 789, "Open"]
      ]);
    });

  it("routes read-only preview_update requests to answer mode instead of asking for write values", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C3", { row: 3, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read ALL cell values from range A1:D4 on Data sheet. Return every row.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.mode).toBe("answer");
      expect((result.answer as any).kind).toBe("range_profile");
      expect((result.answer as any).rows[0]).toEqual(["Date", "Account", "Amount", "Status"]);
    });

  it("includes a tiny exact table preview so overview reads can answer without a follow-up", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the Transactions table data",
        mode: "answer",
        target: { tableName: "Transactions" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_compact_read");
      expect((result.answer as any).valuesPreview).toEqual([
        ["2026-06-01", "A-100", 123, "Open"],
        ["2026-06-02", "A-101", 456, "Closed"],
        ["2026-06-03", "A-102", 789, "Open"]
      ]);
      expect(result.finalAnswer).toContain("3 exact preview rows are inline");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
    });

  it("routes explicit targeted value reads through workbook snapshots", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Report", "B5", { row: 5, column: 2 });
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_snapshot_read_route");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Read Data!A1:D4",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls["workbook.snapshot_ranges"]).toBe(1);
      expect(runtime.lastSnapshotRanges[0]).toMatchObject({ sheetName: "Data", address: "A1:D4" });
      expect(runtime.readBatchCount).toBe(0);
      expect((result.answer as any).rows[0]).toEqual(["Date", "Account", "Amount", "Status"]);
    });

  it("keeps exact narrow route lookup rows inline under brief budget", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_quotation_narrow_rows");
      metadata.sheets.push({
        id: "sheet:Vendor Propose",
        name: "Vendor Propose",
        index: 99,
        usedRange: "A1:P6",
        rowCount: 6,
        columnCount: 16,
        kind: "unknown",
        headers: [],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Show me ALL rows with Item No, Transport Mode, and Route name columns B C D from rows 3 to 28",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { sheetName: "Vendor Propose", range: "B3:D28" },
        budget: { maxPayloadBytes: 2600, maxExamples: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).inlineRowsReason).toBe("narrow_exact_read");
      expect((result.answer as any).rows).toEqual([
        ["Item No.", "Transport Mode", "Route"],
        [1, "Export", "Nongkhae - Klongtoey Port"],
        [2, "Export", "Nongkhae - Ladkrabang Port"],
        [3, "Export", "Nongkhae - Sahathai Port"]
      ]);
      expect(result.warnings.join(" ")).not.toContain("fullResultUri");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(runtime.runtimeMethodCalls["workbook.snapshot_ranges"]).toBe(1);
    });

  it("returns a diagnostic instead of empty success when metadata proves a live range should contain data", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.snapshotRangesOverride = {
        ok: true,
        workbookId: "workbook_agent_unit",
        rangeSnapshots: [{ workbookId: "workbook_agent_unit", sheetName: "Data", address: "A1:D4", values: [] }]
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read Data!A1:D4",
        mode: "answer",
        intent: { action: "read_values" },
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("ERROR");
      expect(result.taskOutcome).toBe("cannot_complete");
      expect(result.finalAnswer).toContain("Open Workbook tried to read");
      expect(result.agentInstruction).toContain("Do not use Python/openpyxl");
      expect(result.warnings.join(" ")).toContain("live Excel read returned empty data");
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
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("large_range_guard");
      expect((result.answer as any).requestedCells).toBeGreaterThan(10000);
      expect(result.warnings[0]).toContain("Large full-data read was blocked");
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
      expect(workbookSummary.nextAction).toBe("answer_now");
      expect(workbookSummary.taskOutcome).toBe("final_answer");
      expect(workbookSummary.maxRecommendedFollowupCalls).toBe(0);
      expect(workbookSummary.agentInstruction).toContain("complete for an overview request");
      expect(workbookSummary.agentInstruction).toContain("do not fetch fullResultUri");
      expect((workbookSummary.answer as any).fullResultUri).toBeUndefined();
      expect(workbookSummary.continuation?.fullResultUri).toBeUndefined();
      expect(workbookSummary.resourceLinks.some((resource) => resource.uri.startsWith("excel://agent/results/"))).toBe(false);
      expect(workbookSummary.telemetry.contextDecision).toMatchObject({
        strategy: "overview",
        scope: "workbook",
        level: 2,
        source: "inferred"
      });
      expect(workbookSummary.contextUsed).toMatchObject({
        strategy: "overview",
        scope: "workbook",
        levelUsed: 2,
        levelReason: expect.stringContaining("overview"),
        stagesPlanned: expect.arrayContaining(["metadata", "schema"]),
        stagesUsed: expect.arrayContaining(["metadata", "schema"]),
        stopReason: expect.stringContaining("lightweight workbook structure"),
        included: expect.arrayContaining(["metadata", "schema"]),
        source: "mixed"
      });
      expect(workbookSummary.telemetry.contextDecision?.include).toEqual(expect.arrayContaining(["metadata", "schema"]));
      expect(sheetSummary.nextAction).toBe("answer_now");
      expect(sheetSummary.maxRecommendedFollowupCalls).toBe(0);
      expect(sheetSummary.agentInstruction).toContain("chunk-read the sheet");
      expect(workbookSummary.telemetry.metadataDetailLevel).toBe("structure");
      expect(sheetSummary.telemetry.cacheHit).toBe(true);
      expect(runtime.readBatchCount).toBe(0);
      expect(JSON.stringify(workbookSummary.answer)).toContain("dropdown rules");
      expect(JSON.stringify(workbookSummary.answer)).not.toContain("Company gas top-up");
      expect(JSON.stringify(workbookSummary.answer).length).toBeLessThan(20000);
    });

  it("reports caller-supplied context policy in telemetry", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Check dropdown issue on Data",
        mode: "answer",
        target: { sheetName: "Data" },
        context: {
          strategy: "audit",
          include: ["validation"]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.telemetry.contextDecision).toMatchObject({
        strategy: "audit",
        scope: "active_sheet",
        include: ["validation"],
        source: "caller"
      });
      expect(result.contextUsed).toMatchObject({
        strategy: "audit",
        scope: "active_sheet",
        included: expect.arrayContaining(["validation"])
      });
    });

  it("reports required cached facets and stale-only refresh plans in contextUsed", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_context_refresh_plan");
      agent.metadataCache.set(metadata);
      agent.metadataCache.markFacetsStale(metadata.workbookContextId, ["values", "aggregates"], ["Data!B2"]);

      const result = await agent.run({
        request: "Analyze Data values",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Data" },
        context: {
          strategy: "analysis",
          include: ["schema", "values"]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.contextUsed).toMatchObject({
        strategy: "analysis",
        requiredFacets: expect.arrayContaining(["schema", "headers", "values"]),
        cachedFacetsUsed: expect.arrayContaining(["schema", "headers"]),
        staleFacets: ["values"],
        facetsToRefresh: ["values"],
        freshnessRequiresRead: true,
        refreshReason: expect.stringContaining("stale context facets")
      });
      expect(result.contextUsed?.missingFacets).toBeUndefined();
      expect(result.telemetry.contextRefresh).toMatchObject({
        requiredFacets: expect.arrayContaining(["schema", "headers", "values"]),
        cachedFacets: expect.arrayContaining(["schema", "headers"]),
        staleFacets: ["values"],
        facetsToRefresh: ["values"],
        readStrategy: "read_stale_facets",
        requiresRead: true
      });
    });

  it("executes query_rows as a read-only table query distinct from visible filters", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_query_rows_contract");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Show unpaid invoice rows",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "query_rows" },
        target: { sheetName: "Data", tableName: "Transactions" },
        values: {
          where: [{ column: "Status", op: "=", value: "Open" }],
          return: ["Date", "Status"],
          updateColumn: "Status",
          updateValue: "Reviewed",
          limit: 10,
          format: "json_rows"
        },
        responseMode: "verbose"
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any)).toMatchObject({
        kind: "query_rows_result",
        matchedRows: 2,
        returnedRows: 2,
        predicates: [{ column: "Status", op: "=", value: "Open" }],
        columns: ["Date", "Status"],
        rows: [
          { Date: "2026-06-01", Status: "Open" },
          { Date: "2026-06-03", Status: "Open" }
        ],
        rowAddresses: ["Data!2:2", "Data!4:4"],
        fieldCandidates: expect.arrayContaining([
          expect.objectContaining({
            term: "Status",
            candidates: expect.arrayContaining([expect.objectContaining({ field: "Status", columnLetter: "D" })])
          })
        ])
      });
      expect(result.telemetry.intentAction).toBe("query_rows");
      expect(result.telemetry.workflowRoute).toBe("rows.query");
      expect(result.agentInstruction).toContain("do not apply visible Excel filters");
      expect(result.suggestedOperation).toMatchObject({
        mode: "preview_update",
        intent: { action: "write_values" },
        values: {
          patches: [
            { target: { sheetName: "Data", range: "D2" }, values: [["Reviewed"]] },
            { target: { sheetName: "Data", range: "D4" }, values: [["Reviewed"]] }
          ]
        }
      });
      expect(runtime.runtimeMethodCalls["table.read"]).toBe(1);
    });

  it("includes section header and data anchors in sheet summaries", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_sheet_section_anchors");
      metadata.sheets.push({
        id: "sheet:Vendor Propose",
        name: "Vendor Propose",
        index: 99,
        usedRange: "A1:P6",
        rowCount: 6,
        columnCount: 16,
        kind: "unknown",
        headers: [],
        tableIds: [],
        sectionIds: ["section:quote:main"],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      metadata.sections.push({
        id: "section:quote:main",
        sheetName: "Vendor Propose",
        label: "quotation section",
        kind: "table-like",
        range: "A1:P6",
        headerRange: "B3:P3",
        headerRow: 3,
        columns: [
          { name: "Route", normalizedName: "route", inferredType: "text", role: "dimension", importance: 0.95, index: 3, letter: "D" },
          { name: "Truck Available", normalizedName: "truck_available", inferredType: "number", role: "measure", importance: 0.8, index: 14, letter: "O" },
          { name: "Vendor Propose", normalizedName: "vendor_propose", inferredType: "currency", role: "amount", importance: 0.9, index: 15, letter: "P" }
        ],
        labels: ["YLTH_CTG_Zone BKK TT _Y2026"],
        rowCount: 6,
        columnCount: 16,
        nonEmptyCellCount: 42,
        confidence: 0.91
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Look at the Vendor Propose worksheet",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        detailLevel: "sheet_summary",
        target: { sheetName: "Vendor Propose" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_summary");
      expect((result.answer as any).sections[0]).toMatchObject({
        id: "section:quote:main",
        headerRange: "B3:P3",
        dataRange: "A4:P6"
      });
      expect((result.answer as any).sections[0].fingerprint).toMatch(/^[a-f0-9]{16}$/);
      expect((result.answer as any).sections[0].editableColumns.map((column: any) => column.letter)).toEqual(["O", "P"]);
      expect(result.finalAnswer).toContain("quotation section");
      expect(result.finalAnswer).toContain("header B3:P3");
      expect(result.finalAnswer).toContain("data A4:P6");
      expect(runtime.readBatchCount).toBe(0);
    });

  it("returns semantic index context hints for dropdowns, styles, and historical labels without row payloads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_semantic_hints");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Map workbook context",
        mode: "answer",
        detailLevel: "semantic_index",
        workbookContextId: metadata.workbookContextId
      });

      expect(result.status).toBe("SUCCESS");
      const entries = (result.answer as any).entries;
      const apr = entries.find((entry: any) => entry.sheetName === "Apr 2026");
      expect(apr.nextRequestHints.join(" ")).toContain("find_similar_rows");
      expect(apr.nextRequestHints.join(" ")).toContain("read_data_validation");
      expect(JSON.stringify(result.answer)).not.toContain("Company gas top-up");
      expect(result.telemetry.fullReadCellCount).toBe(0);
    });

  it("reads exact dropdown validation rules for a selected/status column", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "What dropdown values are allowed in this status column?",
        mode: "answer",
        intent: { action: "read_data_validation" },
        target: { sheetName: "Data", range: "D1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls["range.read_data_validation"]).toBe(1);
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        sheetName: "Data",
        range: "D1:D4",
        options: ["Open", "Closed", "Pending"],
        optionCount: 3,
        sourceComplete: true,
        fieldContext: [
          expect.objectContaining({
            field: "Status",
            range: "D1:D4",
            headerRange: "D1",
            semanticType: "status",
            dataType: "status",
            hasValidation: true,
            allowedValues: ["Open", "Closed", "Pending"],
            allowedValueCount: 3,
            validation: expect.objectContaining({
              type: "list",
              sourceType: "inline",
              options: ["Open", "Closed", "Pending"],
              optionCount: 3
            })
          })
        ]
      });
      expect(result.taskOutcome).toBe("final_answer");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.agentInstruction).toContain("do not fetch fullResultUri");
    });

  it("routes natural dropdown inspection to data-validation metadata without reading values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Check data validation / dropdown on these specific cells. Is there any list validation?",
        mode: "answer",
        target: { sheetName: "Data", range: "D1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        method: "range.read_data_validation",
        options: ["Open", "Closed", "Pending"],
        sourceComplete: true
      });
      expect(runtime.runtimeMethodCalls["range.read_data_validation"]).toBe(1);
      expect(result.telemetry.fullReadCellCount).toBe(0);
    });

  it("reads values from a sheet named Dropdown Lists without treating the sheet name as validation intent", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_dropdown_list_values");
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Dropdown Lists",
          name: "Dropdown Lists",
          index: 9,
          usedRange: "A1:B29",
          rowCount: 29,
          columnCount: 2,
          kind: "reference",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Read all values in column B of the Dropdown Lists sheet, from B1 downwards. Show me the full list of all transaction types listed.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Dropdown Lists", range: "B1:B28" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).not.toBe("data_validation_summary");
      expect((result.answer as any).kind).not.toBe("workbook_design_overview");
      expect((result.answer as any).valuesPreview).toContainEqual(["owner_cash_topup"]);
      expect(runtime.runtimeMethodCalls["range.read_data_validation"] ?? 0).toBe(0);
      expect(result.telemetry.fullReadCellCount).toBeGreaterThan(0);
    });

  it("checks source-list text containment as cell values, not validation metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_dropdown_text_contains");
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Dropdown Lists",
          name: "Dropdown Lists",
          index: 9,
          usedRange: "A1:B29",
          rowCount: 29,
          columnCount: 2,
          kind: "reference",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Check Dropdown Lists sheet, column B, rows 1 to 28. Tell me if any of the cells contain the text \"owner_cash_topup\". Just read the live values.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).not.toBe("data_validation_summary");
      expect(JSON.stringify((result.answer as any).valuesPreview)).toContain("owner_cash_topup");
      expect(runtime.runtimeMethodCalls["range.read_data_validation"] ?? 0).toBe(0);
    });

  it("inlines small targeted source-list reads so agents do not fetch full result resources", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_dropdown_list_inline");
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Dropdown Lists",
          name: "Dropdown Lists",
          index: 9,
          usedRange: "A1:B29",
          rowCount: 29,
          columnCount: 2,
          kind: "reference",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "",
        mode: "auto",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Dropdown Lists", range: "B1:B28" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).valuesPreview).toContainEqual(["owner_cash_topup"]);
      expect((result.answer as any).inlineIsComplete).toBe(true);
      expect(result.nextAction).toBe("answer_now");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.agentInstruction).toContain("do not call workbook tools again");
      expect((result.answer as any).fullResultUri).toBeUndefined();
      expect(result.continuation?.fullResultUri).toBeUndefined();
      expect(result.resourceLinks.map((link) => link.uri).some((uri) => uri.includes("/results/"))).toBe(false);
    });

  it("honors explicit actual-cell-value wording even when the request says not data validation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_dropdown_actual_values");
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Dropdown Lists",
          name: "Dropdown Lists",
          index: 9,
          usedRange: "A1:B29",
          rowCount: 29,
          columnCount: 2,
          kind: "reference",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Actually read the cell values, not the data validation. I want to see the text strings stored in cells B1:B28 on the Dropdown Lists sheet. Give me each cell value.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Dropdown Lists", range: "B1:B28" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).not.toBe("data_validation_summary");
      expect((result.answer as any).kind).not.toBe("workbook_design_overview");
      expect((result.answer as any).valuesPreview).toContainEqual(["owner_cash_topup"]);
      expect(runtime.runtimeMethodCalls["range.read_data_validation"] ?? 0).toBe(0);
    });

  it("routes natural validation formula/source requests to data-validation metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the data validation rule on cell May 2026!E2. What is the validation formula, type, and inline list or range reference?",
        mode: "answer",
        target: { sheetName: "May 2026", range: "E2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        method: "range.read_data_validation"
      });
      expect((result.answer as any).kind).not.toBe("workbook_design_overview");
      expect((result.answer as any).kind).not.toBe("formula_read");
      expect((result.answer as any).kind).not.toBe("reference_sheet_analysis");
      expect(runtime.runtimeMethodCalls["range.read_data_validation"]).toBe(1);
      expect(result.maxRecommendedFollowupCalls).toBe(0);
    });

  it("treats inconsistent multi-cell validation reads as inconclusive, not broken dropdown proof", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      runtime.readRangeMetadata = async (method: string, request: any) => {
        runtime.runtimeMethodCalls[method] = (runtime.runtimeMethodCalls[method] ?? 0) + 1;
        return {
          ok: true,
          method,
          request,
          data: {
            address: request.address,
            type: "Inconsistent",
            inCellDropDown: false
          }
        };
      };

      const result = await agent.run({
        request: "Read the data validation rule on May 2026 column E (E2:E244). Does it include owner_cash_topup?",
        mode: "answer",
        target: { sheetName: "May 2026", range: "E2:E244" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        type: "Inconsistent",
        validationRangeStatus: "mixed_or_inconsistent_range"
      });
      expect((result.answer as any).guidance).toContain("Do not conclude");
      expect(result.finalAnswer).toContain("inconclusive");
      expect(result.finalAnswer).toContain("do not say the dropdown is broken");
      expect(runtime.runtimeMethodCalls["range.read_data_validation"]).toBe(1);
    });

  it("does not large-range guard full-column validation inspection", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Check data validation (dropdown lists) across the entire Data sheet. Which columns have dropdown validation?",
        mode: "answer",
        target: { sheetName: "Data" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        method: "range.read_data_validation",
        optionCount: 3,
        sourceComplete: true
      });
      expect((result.answer as any).kind).not.toBe("large_range_guard");
      expect(runtime.runtimeMethodCalls["range.read_data_validation"]).toBe(1);
      expect(result.telemetry.fullReadCellCount).toBe(0);
    });

  it("summarizes long dropdown validation inline without recommending full-result retries", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const options = Array.from({ length: 60 }, (_value, index) => `tx_type_${index + 1}`);
      runtime.readRangeMetadata = async (method: string, request: any) => {
        runtime.runtimeMethodCalls[method] = (runtime.runtimeMethodCalls[method] ?? 0) + 1;
        return {
          ok: true,
          method,
          request,
          data: {
            address: request.address,
            rules: [{
              address: request.address,
              type: "list",
              source: options,
              inCellDropDown: true
            }]
          }
        };
      };

      const result = await agent.run({
        request: "What transaction type dropdown values are allowed?",
        mode: "answer",
        intent: { action: "read_data_validation" },
        target: { sheetName: "May 2026", range: "E:E" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.nextAction).toBe("answer_now");
      expect(result.taskOutcome).toBe("final_answer");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        optionCount: 60,
        sourceComplete: true
      });
      expect((result.answer as any).fullResultUri).toBeUndefined();
      expect(result.continuation?.fullResultUri).toBeUndefined();
      expect(result.agentInstruction).toContain("do not fetch fullResultUri");
    });

  it("exposes dropdown source ranges so missing options can be added by source-list writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      runtime.readRangeMetadata = async (method: string, request: any) => {
        runtime.runtimeMethodCalls[method] = (runtime.runtimeMethodCalls[method] ?? 0) + 1;
        return {
          ok: true,
          method,
          request,
          data: {
            address: request.address,
            type: "List",
            rule: { list: { source: "=Lists!$A$2:$A$40", inCellDropDown: true } },
            ignoreBlanks: true,
            valid: true
          }
        };
      };

      const result = await agent.run({
        request: "Check transaction type dropdown source for May 2026.",
        mode: "answer",
        intent: { action: "read_data_validation" },
        target: { sheetName: "May 2026", range: "E:E" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any)).toMatchObject({
        kind: "data_validation_summary",
        sourceFormula: "=Lists!$A$2:$A$40",
        sourceRange: "Lists!$A$2:$A$40",
        sourceComplete: false,
        fieldContext: [
          expect.objectContaining({
            validation: expect.objectContaining({
              type: "List",
              sourceType: "range",
              sourceRange: "Lists!$A$2:$A$40",
              optionsResolved: false
            })
          })
        ]
      });
      expect((result.answer as any).guidance).toContain("source-list cells");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
    });

  it("marks dynamic dropdown formulas as unresolved field context", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      runtime.readRangeMetadata = async (method: string, request: any) => {
        runtime.runtimeMethodCalls[method] = (runtime.runtimeMethodCalls[method] ?? 0) + 1;
        return {
          ok: true,
          method,
          request,
          data: {
            address: request.address,
            rules: [{
              address: request.address,
              type: "list",
              source: "=INDIRECT($B2)",
              inCellDropDown: true
            }]
          }
        };
      };

      const result = await agent.run({
        request: "Check dependent subcategory dropdown.",
        mode: "answer",
        intent: { action: "read_data_validation" },
        target: { sheetName: "Data", range: "D2:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).fieldContext[0].validation).toMatchObject({
        type: "list",
        sourceType: "formula",
        formula: "=INDIRECT($B2)",
        optionsResolved: false,
        reason: expect.stringContaining("Dynamic or dependent dropdown")
      });
    });

  it("finds similar prior-period rows across related workbook sheets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_similar_rows");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Look back at last month and show how this kind of transaction was labeled.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "find_similar_rows" },
        target: { sheetName: "Apr 2026", range: "A2:AJ2" },
        budget: { maxExamples: 3 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("similar_rows");
      expect((result.answer as any).rows[0]).toMatchObject({ sheetName: "Mar 2026", range: "A2:AJ2" });
      expect((result.answer as any).rows[0].values).toContain("company_gas_topup");
      expect(result.proof[0]).toMatchObject({ sheetName: "Mar 2026", range: "A2:AJ2" });
    });

  it("returns exact row numbers for bounded text searches without full result handles", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_exact_search_rows");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Search column K in May 2026 for all cells containing WITSARUT and return row numbers.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "May 2026", range: "K1:K244" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("exact_search_rows");
      expect((result.answer as any).matchedRows).toEqual([25, 27, 28, 31, 33, 37, 38, 39, 40, 44, 48]);
      expect((result.answer as any).fullResultUri).toBeUndefined();
      expect(result.nextAction).toBe("answer_now");
      expect(result.agentInstruction).toContain("Do not fetch full rows");
    });

  it("searches requested prior sheets for exact reference rows instead of returning compact broad ranges", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_reference_search");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Look more into Apr 2026 for how we label adding fund from X1183 / PRACH with amount 10,000.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Apr 2026", range: "A:L" },
        budget: { maxPayloadBytes: 1400, maxExamples: 4 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("similar_rows");
      expect((result.answer as any).rows[0]).toMatchObject({
        sheetName: "Apr 2026",
        sheetRowNumber: 3
      });
      expect((result.answer as any).rows[0].columns).toEqual(expect.arrayContaining([
        expect.objectContaining({ letter: "E", name: "Transaction Type", value: "owner_fund_added" }),
        expect.objectContaining({ letter: "F", name: "Direction", value: "Inflow" }),
        expect.objectContaining({ letter: "K", name: "Transfer From/To", value: "From X1183 MR. PRACH YOTHAPRA++" })
      ]));
      expect((result.answer as any).comparedRanges.every((range: any) => range.sheetName === "Apr 2026")).toBe(true);
      expect(result.proof.every((proof) => proof.sheetName === "Apr 2026")).toBe(true);
      expect(JSON.stringify(result.answer)).not.toContain("agent_result_resource");
      expect(result.warnings.join(" ")).not.toContain("fullResultUri");
    });

  it("keeps multilingual reference signals available for similar-row search", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_reference_thai");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "หา reference Apr 2026 เติมเงินเข้าบริษัท 10000 จาก PRACH",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Apr 2026", range: "A:L" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).rows[0].matchedSignals.join(" ")).toContain("เติมเงิน");
      expect((result.answer as any).rows[0].columns.map((column: any) => column.value)).toContain("เติมเงินเข้าบริษัท");
    });

  it("returns style reference candidates without falling back to value range profiles", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_style_refs");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Find a prior month style reference so this sheet can look like last month.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "find_style_references" },
        target: { sheetName: "Apr 2026" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_reference_candidates");
      expect((result.answer as any).candidates[0]).toEqual(expect.objectContaining({
        sheetName: expect.any(String),
        range: expect.any(String),
        styleSummary: expect.any(Object)
      }));
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBeGreaterThan(0);
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

  it("answers freeze pane status questions without starting a mutation preview", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Which column is frozen on Data?",
        mode: "answer",
        target: { sheetName: "Data" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("freeze_panes_status");
      expect((result.answer as any).readable).toBe(true);
      expect((result.answer as any).lastFrozenColumn).toBe("C");
      expect((result.answer as any).firstUnfrozenColumn).toBe("D");
      expect(result.summary).toContain("columns A:C are frozen");
      expect(result.operationId).toBeUndefined();
      expect(result.nextAction).toBe("answer_now");
    });

  it("uses intent.reason to answer freeze pane status even when caller chose the wrong read action", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Check Invoices",
        mode: "answer",
        target: { sheetName: "Data" },
        intent: { action: "read_style_summary", reason: "Check freeze panes on Data sheet" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("freeze_panes_status");
      expect((result.answer as any).lastFrozenColumn).toBe("C");
      expect((result.answer as any).firstUnfrozenColumn).toBe("D");
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBe(1);
    });

  it("does not treat legacy freeze pane notes as no frozen panes", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.getStyleFingerprint = async (request: any) => {
        runtime.runtimeMethodCalls["style.get_fingerprint"] = (runtime.runtimeMethodCalls["style.get_fingerprint"] ?? 0) + 1;
        return {
          ok: true,
          fingerprint: {
            workbookId: request.workbookId,
            sheetName: request.sheetName,
            address: request.address,
            dimensions: { freezePanes: { note: "Office.js freeze pane capture is tracked as a layout capability." } },
            warnings: []
          }
        };
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Which column is frozen on Data?",
        mode: "answer",
        target: { sheetName: "Data" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).readable).toBe(false);
      expect(result.summary).toContain("cannot be read");
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

  it("returns a compact style overview with grouped header suggestions without reading table values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_style_overview");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Give me a styling overview and best-practice suggestions for Apr 2026.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "style_overview" },
        target: { sheetName: "Apr 2026" },
        detailLevel: "style_overview"
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_overview");
      expect((result.answer as any).detected.headerRange).toBe("A1:AJ1");
      expect((result.answer as any).freezePanes).toMatchObject({
        readable: true,
        frozen: true,
        rows: 2,
        columns: 3,
        lastFrozenColumn: "C",
        firstUnfrozenColumn: "D"
      });
      expect((result.answer as any).columnRoles).toEqual(expect.arrayContaining([
        expect.objectContaining({ column: "A", freezePane: expect.objectContaining({ isFrozen: true }) }),
        expect.objectContaining({ column: "C", freezePane: expect.objectContaining({ isFrozen: true, isLastFrozenColumn: true }) }),
        expect.objectContaining({ column: "D", freezePane: expect.objectContaining({ isFrozen: false, isFirstUnfrozenColumn: true }) })
      ]));
      expect((result.answer as any).groupedHeaderSuggestion).toMatchObject({
        kind: "grouped_header_suggestion",
        requiresStructuralPreview: true,
        defaultApplyBehavior: "suggest_only"
      });
      expect((result.answer as any).groupedHeaderSuggestion.groups.length).toBeGreaterThan(1);
      expect((result.answer as any).recommendations.map((item: any) => item.id)).toContain("grouped_header");
      expect(result.telemetry.fullReadCellCount).toBe(0);
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBe(1);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("returns workbook design overview with column recommendations without sampling values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_design_overview");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Before applying, please do a workbook design review for this sheet. For each column decide free text, date, number/money, ID/text code, dropdown list, or lookup/reference from another sheet.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "workbook_design_overview" },
        target: { sheetName: "Data" }
      });

      const answer = result.answer as any;
      expect(result.status).toBe("SUCCESS");
      expect(answer.kind).toBe("workbook_design_overview");
      expect(answer.inspectionPolicy.fullReadCellCount).toBe(0);
      expect(answer.columnRecommendations).toEqual(expect.arrayContaining([
        expect.objectContaining({ column: "A", header: "Date", recommendedBehavior: "date" }),
        expect.objectContaining({ column: "B", header: "Account", recommendedBehavior: "lookup_reference" }),
        expect.objectContaining({ column: "C", header: "Amount", recommendedBehavior: "number_money" }),
        expect.objectContaining({ column: "D", header: "Status", recommendedBehavior: "dropdown_list" })
      ]));
      expect(answer.relatedSheets).toEqual(expect.arrayContaining([
        expect.objectContaining({ sheetName: "Customer Master" })
      ]));
      expect(answer.nextWorkflows.map((workflow: any) => workflow.intentAction)).toEqual(expect.arrayContaining(["improve_visual_readability", "write_data_validation"]));
      expect(result.telemetry.fullReadCellCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBeUndefined();
    });

  it("routes natural column-by-column design review prompts to workbook_design_overview", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_design_natural");
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Please review Apr 2026 column-by-column and recommend which columns should be dates, money, dropdowns, IDs, or lookups. Do not apply yet.",
        mode: "answer",
        workbookContextId: metadata.workbookContextId
      });

      const answer = result.answer as any;
      expect(result.status).toBe("SUCCESS");
      expect(answer.kind).toBe("workbook_design_overview");
      expect(answer.sheet.name).toBe("Apr 2026");
      expect(answer.columnRecommendations).toEqual(expect.arrayContaining([
        expect.objectContaining({ header: "Transaction Date", recommendedBehavior: "date" }),
        expect.objectContaining({ header: "Payment Variance", recommendedBehavior: "number_money" }),
        expect.objectContaining({ header: "Container Size", recommendedBehavior: "dropdown_list" })
      ]));
      expect(answer.inspectionPolicy.guidance).toContain("Do not broad-read empty data rows");
      expect(runtime.readBatchCount).toBe(0);
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
      expect((result.answer as any).values).toEqual([[123, "Open"], [456, "Closed"]]);
      expect((result.answer as any).valuesPreview).toEqual([[123, "Open"], [456, "Closed"]]);
      expect((result.answer as any).resultUri).toMatch(/^excel:\/\/agent\/results\/agentres_/);
      expect((result.answer as any).fullResultUri).toMatch(/\?view=full$/);
      expect(result.continuation).toMatchObject({
        workbookContextId: metadata.workbookContextId,
        resultUri: (result.answer as any).resultUri,
        fullResultUri: (result.answer as any).fullResultUri,
        responseMode: "brief",
        freshness: {
          workbookId: metadata.workbook.workbookId,
          workbookStructureHash: metadata.fingerprint.structureHash
        }
      });
      expect((result.answer as any).projectedColumns.map((column: any) => column.name)).toEqual(["Amount", "Status"]);
      expect((result.answer as any).fieldContext).toEqual([
        expect.objectContaining({
          field: "Amount",
          range: "C2:C4",
          headerRange: "C1",
          dataType: "number",
          currentDistinctValues: [123, 456],
          blankCount: 0,
          examples: [123, 456]
        }),
        expect.objectContaining({
          field: "Status",
          range: "D2:D4",
          headerRange: "D1",
          dataType: "status",
          currentDistinctValues: ["Open", "Closed"],
          blankCount: 0,
          examples: ["Open", "Closed"]
        })
      ]);
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      expect((agent.getResultResource(resultId) as any).answer.values).toEqual([[123, "Open"], [456, "Closed"]]);
      const summary = agent.getResultResource(resultId, { view: "summary" }) as any;
      expect(summary.answer.values).toBeUndefined();
      expect(summary.answer.fieldContext).toEqual((result.answer as any).fieldContext);
      expect(summary.answer.resultUri).toBe((result.answer as any).resultUri);
      expect(summary.freshness).toMatchObject({
        workbookStructureHash: metadata.fingerprint.structureHash
      });
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
      expect(continued.warnings.join(" ")).toContain("never use Webfetch/browser");
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

  it("projects wide table rows inline using cached column role importance", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_wide_projection");
      const wideColumns = [
        { name: "Transaction Date", role: "date", importance: 0.97 },
        { name: "Internal Seq", role: "identifier", importance: 0.55 },
        { name: "Upload Batch", role: "dimension", importance: 0.45 },
        { name: "Description", role: "description", importance: 1 },
        { name: "Vendor", role: "vendor", importance: 0.92 },
        { name: "Department", role: "dimension", importance: 0.58 },
        { name: "Account", role: "account", importance: 0.82 },
        { name: "Memo", role: "note", importance: 0.7 },
        { name: "Cash Amount", role: "amount", importance: 0.99 },
        { name: "Actual Amount", role: "amount", importance: 0.96 },
        { name: "Variance", role: "formula", importance: 0.78 },
        { name: "Category", role: "category", importance: 0.94 },
        { name: "Status", role: "status", importance: 0.93 },
        { name: "Reviewer", role: "dimension", importance: 0.5 },
        { name: "Comment", role: "note", importance: 0.72 },
        { name: "Source File", role: "identifier", importance: 0.68 },
        { name: "Import Row", role: "identifier", importance: 0.6 },
        { name: "Audit Flag", role: "status", importance: 0.88 },
        { name: "Unused 1", role: "unknown", importance: 0.2 },
        { name: "Unused 2", role: "unknown", importance: 0.2 }
      ].map((column, index) => ({
        ...column,
        normalizedName: column.name.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, ""),
        inferredType: column.role === "date" ? "date" : column.role === "amount" ? "currency" : column.role === "status" ? "status" : "text",
        index,
        letter: String.fromCharCode(65 + index)
      }));
      metadata.tables[0]!.range = "A1:T4";
      metadata.tables[0]!.dataRange = "A2:T4";
      metadata.tables[0]!.columns = wideColumns as any;
      agent.metadataCache.set(metadata);
      runtime.readTable = async () => {
        runtime.runtimeMethodCalls["table.read"] = (runtime.runtimeMethodCalls["table.read"] ?? 0) + 1;
        return {
          ok: true,
          table: {
            tableName: "Transactions",
            headers: wideColumns.map((column) => column.name),
            values: [wideColumns.map((column) => `value:${column.name}`)]
          }
        };
      };

      const result = await agent.run({
        request: "Read the first transaction row",
        mode: "answer",
        responseMode: "brief",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { tableName: "Transactions" },
        values: { rowLimit: 1 },
        budget: { maxPayloadBytes: 20_000 }
      });

      expect(result.status).toBe("SUCCESS");
      const answer = result.answer as any;
      expect(answer.inlineColumnProjection).toMatchObject({
        reason: "role_aware_wide_row_projection",
        omittedColumnCount: 8
      });
      expect(answer.values[0]).toHaveLength(12);
      expect(answer.headers).toHaveLength(12);
      expect(answer.headers).toEqual(expect.arrayContaining([
        "Transaction Date",
        "Description",
        "Vendor",
        "Cash Amount",
        "Category",
        "Status"
      ]));
      expect(JSON.stringify(answer.values)).not.toContain("value:Unused 2");
      expect(answer.schemaSummary.columns[0]).toMatchObject({ name: "Transaction Date", role: "date", importance: 0.97 });
      expect(result.telemetry.internalReadCount).toBe(1);
    });

  it("uses domain dictionary encoding for repeated compact table values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_domain_encoding");
      metadata.tables[0]!.columns = [
        { name: "Date", normalizedName: "date", inferredType: "date", role: "date", importance: 0.97, index: 0, letter: "A" },
        { name: "Description", normalizedName: "description", inferredType: "text", role: "description", importance: 1, index: 1, letter: "B" },
        { name: "Amount", normalizedName: "amount", inferredType: "currency", role: "amount", importance: 0.99, index: 2, letter: "C" },
        { name: "Status", normalizedName: "status", inferredType: "status", role: "status", importance: 0.93, index: 3, letter: "D" }
      ];
      agent.metadataCache.set(metadata);
      const descriptions = [
        "Company fuel top-up requiring month-end operating expense label",
        "Customer invoice settlement requiring accounts receivable label"
      ];
      runtime.readTable = async () => {
        runtime.runtimeMethodCalls["table.read"] = (runtime.runtimeMethodCalls["table.read"] ?? 0) + 1;
        return {
          ok: true,
          table: {
            tableName: "Transactions",
            headers: ["Date", "Description", "Amount", "Status"],
            values: Array.from({ length: 12 }, (_value, index) => [
              `2026-06-${String((index % 4) + 1).padStart(2, "0")}`,
              descriptions[index % descriptions.length],
              100 + index,
              index % 3 === 0 ? "Needs review" : "Approved"
            ])
          }
        };
      };

      const result = await agent.run({
        request: "Read recent transaction labels",
        mode: "answer",
        responseMode: "brief",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { tableName: "Transactions" },
        values: { rowLimit: 12 }
      });

      expect(result.status).toBe("SUCCESS");
      const answer = result.answer as any;
      expect(answer.values).toBeUndefined();
      expect(answer.encodedValues).toHaveLength(12);
      expect(answer.encodedValues[0][2]).toBe(100);
      expect(answer.valueEncoding).toMatchObject({ kind: "domain_dictionary_by_column" });
      expect(answer.valueEncoding.columns.map((column: any) => column.name)).toEqual(expect.arrayContaining(["Description", "Status"]));
      expect(JSON.stringify(answer.encodedValues)).not.toContain("Company fuel top-up");
      expect(JSON.stringify(answer.valueEncoding.columns)).toContain("Company fuel top-up");
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

  it("requires fetching stored detail when an explicit full-table request is compacted to preview rows", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_full_table_compacted");
      metadata.tables[0]!.range = "A1:D121";
      metadata.tables[0]!.dataRange = "A2:D121";
      agent.metadataCache.set(metadata);
      runtime.readTable = async (request: any) => {
        runtime.runtimeMethodCalls["table.read"] = (runtime.runtimeMethodCalls["table.read"] ?? 0) + 1;
        return {
          ok: true,
          table: {
            tableName: request.tableName,
            headers: ["Date", "Account", "Amount", "Status"],
            values: Array.from({ length: 120 }, (_value, index) => [`2026-06-${String((index % 28) + 1).padStart(2, "0")}`, `A-${index}`, index, "Open"])
          }
        };
      };

      const result = await agent.run({
        request: "Show all rows from Transactions table",
        mode: "answer",
        detailLevel: "full_table",
        workbookContextId: metadata.workbookContextId,
        target: { tableName: "Transactions" },
        budget: { maxPayloadBytes: 1800, maxExamples: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.nextAction).toBe("fetch_resource");
      expect(result.taskOutcome).toBe("cannot_complete");
      expect(result.requiredFollowup?.nextAction).toBe("fetch_resource");
      expect((result.answer as any).inlineIsComplete).toBe(false);
      expect((result.answer as any).inlineRowCount).toBeLessThan((result.answer as any).totalRowCount);
      expect(result.agentInstruction).toContain("fetch the stored full result");
    });

  it("marks small full-table reads as complete inline", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Show all rows from Transactions table",
        mode: "answer",
        detailLevel: "full_table",
        target: { tableName: "Transactions" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.nextAction).toBe("answer_now");
      expect((result.answer as any).inlineIsComplete).toBe(true);
      expect((result.answer as any).inlineRowCount).toBe(3);
      expect((result.answer as any).totalRowCount).toBe(3);
    });

  it("asks for predicates instead of broad-reading large search-like ranges", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Find METRO PARTICLE in Data!A1:XFD1048576",
        mode: "answer",
        target: { sheetName: "Data", range: "A1:XFD1048576" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("large_range_guard");
      expect(result.warnings.join(" ")).toContain("Do not broad-read follow-up chunks");
    });

  it("asks users to narrow explicit huge full-sheet reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Show all rows from Data!A1:XFD1048576",
        mode: "answer",
        target: { sheetName: "Data", range: "A1:XFD1048576" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("large_range_guard");
      expect((result.answer as any).recommendation).toContain("smaller range");
      expect(result.taskOutcome).toBe("needs_user_input");
      expect(result.warnings.join(" ")).toContain("Large full-data read was blocked");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("asks the user after a targeted search returns no exact match", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.readRangeMetadata = async (method: string, request: any) => {
        runtime.runtimeMethodCalls[method] = (runtime.runtimeMethodCalls[method] ?? 0) + 1;
        return { ok: true, method, request, data: { address: request.address, count: 0, matches: [] } };
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Search Data for Missing Customer",
        mode: "answer",
        intent: { action: "search_range" },
        target: { sheetName: "Data", range: "A1:D4" },
        values: { text: "Missing Customer" }
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("range_search_no_match");
      expect((result.answer as any).searchedRangeWasComplete).toBe(true);
      expect(result.warnings.join(" ")).toContain("Do not broad-read adjacent chunks");
    });

  it("labels fresh empty sheets as successful empty results instead of live-read failures", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.snapshotRangesOverride = {
        ok: true,
        workbookId: "workbook_agent_unit",
        rangeSnapshots: [{ workbookId: "workbook_agent_unit", sheetName: "Fresh", address: "A1:D20", values: [] }]
      };
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_fresh_empty_sheet");
      metadata.sheets.push({
        id: "sheet:Fresh",
        name: "Fresh",
        index: 99,
        usedRange: "A1:D20",
        rowCount: 20,
        columnCount: 4,
        kind: "transaction",
        headers: [],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Read Fresh!A1:D20",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { sheetName: "Fresh", range: "A1:D20" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).emptyResultKind).toBe("fresh_sheet");
      expect(result.taskOutcome).toBe("final_answer");
      expect(result.warnings.join(" ")).toContain("No non-empty cells");
    });

  it("normalizes short-year date writes and adds targeted date number formats", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update booking dates",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A2:B2" },
        values: { patches: [{ target: { sheetName: "Data", range: "A2:B2" }, values: [["25/6/26", "27/6/26"]] }] }
      });

      expect(result.status).toBe("PREVIEW_READY");
      const operations = (agent.dumpOperations()[0] as any).action.operations;
      expect(operations[0]).toMatchObject({ kind: "range.write_values_many" });
      expect(operations[0].entries[0].values).toEqual([[46198, 46200]]);
      expect(operations[1]).toMatchObject({ kind: "range.write_number_formats_many" });
      expect(operations[1].entries).toHaveLength(2);
      expect(operations[1].entries[0].numberFormat).toEqual([["dd/mm/yyyy"]]);
      expect(result.warnings.join(" ")).toContain("Short-year date text was normalized");
    });

  it("normalizes multi-row short-year booking dates without formatting non-date cells", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Write booking dates",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A2:C3" },
        values: { patches: [{ target: { sheetName: "Data", range: "A2:C3" }, values: [["25/6/26", "Customer A", "27/6/26"], ["26/6/26", "Customer B", "28/6/26"]] }] }
      });

      expect(result.status).toBe("PREVIEW_READY");
      const operations = (agent.dumpOperations()[0] as any).action.operations;
      expect(operations[0].entries[0].values).toEqual([[46198, "Customer A", 46200], [46199, "Customer B", 46201]]);
      expect(operations[1]).toMatchObject({ kind: "range.write_number_formats_many" });
      expect(operations[1].entries.map((entry: any) => entry.target.address)).toEqual(["A2", "C2", "A3", "C3"]);
    });

  it("does not normalize impossible short-year date text", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update booking dates",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A2:A2" },
        values: { patches: [{ target: { sheetName: "Data", range: "A2:A2" }, values: [["31/2/26"]] }] }
      });

      expect(result.status).toBe("PREVIEW_READY");
      const operations = (agent.dumpOperations()[0] as any).action.operations;
      expect(operations).toHaveLength(1);
      expect(operations[0].entries[0].values).toEqual([["31/2/26"]]);
      expect(result.warnings.join(" ")).not.toContain("Short-year date text was normalized");
    });

  it("asks for scope before building unclear broad mutation payloads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update all rows in the worksheet",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:Z100" },
        values: { value: "Reviewed" }
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("broad_mutation_scope_guard");
      expect(result.warnings.join(" ")).toContain("Broad update scope is unclear");
    });

  it("asks for scope before broad mutations even when values are not provided yet", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Color all rows in the worksheet",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:Z100" }
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.nextAction).toBe("ask_user");
      expect((result.answer as any).kind).toBe("broad_mutation_scope_guard");
      expect((result.answer as any).alternatives).toContain("matching rows only");
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
