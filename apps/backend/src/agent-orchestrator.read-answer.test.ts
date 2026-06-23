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
      expect(JSON.stringify(workbookSummary.answer)).toContain("dropdown rules");
      expect(JSON.stringify(workbookSummary.answer)).not.toContain("Company gas top-up");
      expect(JSON.stringify(workbookSummary.answer).length).toBeLessThan(20000);
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
      expect((result.answer as any).result.data.rules[0].source).toEqual(["Open", "Closed", "Pending"]);
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
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      expect((agent.getResultResource(resultId) as any).answer.values).toEqual([[123, "Open"], [456, "Closed"]]);
      const summary = agent.getResultResource(resultId, { view: "summary" }) as any;
      expect(summary.answer.values).toBeUndefined();
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
