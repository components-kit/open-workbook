import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Target Resolution", () => {
  it("answers with one targeted read instead of repeated broad reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Answer amount from Transactions table", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(result.telemetry.fullReadCellCount).toBeLessThanOrEqual(16);
      expect(runtime.readBatchCount).toBeLessThanOrEqual(sheets.length);
    });

  it("reads live values when batch results use the protocol data field", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.returnDataOnly = true;
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read headers and first 5 rows from Data A1:D4",
        mode: "answer",
        responseMode: "verbose",
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).source).toBe("live_read");
      expect((result.answer as any).sample[1][1]).toBe("A-100");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(result.telemetry.fullReadCellCount).toBeGreaterThan(0);
    });

  it("resolves explicit row deletion wording to a row range instead of the sheet used range", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Please delete row 2 on Data",
        mode: "preview_update"
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.proof[0]).toMatchObject({ sheetName: "Data", range: "2:2" });

      await agent.run({
        request: "Apply row delete",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.delete_rows",
        target: { sheetName: "Data", address: "2:2" }
      });
    });

  it("resolves this row deletion to the current selected worksheet row", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "C5", { row: 5, column: 3 });
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Delete this row",
        mode: "preview_update"
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.proof[0]).toMatchObject({ sheetName: "Data", range: "5:5" });
    });

  it("treats headers plus rows as a live value read instead of schema metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read headers and first 5 rows from Data A1:D4",
        mode: "answer",
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("range_profile");
      expect((result.answer as any).source).toBe("live_read");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(result.taskOutcome).toBe("final_answer");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.agentInstruction).toContain("do not call workbook tools again");
    });

  it("returns coordinate-aware sparse rows for mostly empty ranges", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the sparse planning area.",
        mode: "answer",
        target: { sheetName: "Data", range: "A1:J10" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("range_profile");
      expect((result.answer as any).rows).toBeUndefined();
      expect((result.answer as any).emptySummary).toBeUndefined();
      expect((result.answer as any).sparseRows).toBeUndefined();
      expect((result.answer as any).resultUri).toMatch(/^excel:\/\/agent\/results\/agentres_/);
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      const stored = agent.getResultResource(resultId) as any;
      expect(stored.answer.emptySummary.emptyCells).toBe(97);
      expect(stored.answer.sparseRows).toEqual([
        { row: 1, cells: [{ column: "A", address: "A1", value: "Owner" }, { column: "J", address: "J1", value: "Status" }] },
        { row: 10, cells: [{ column: "J", address: "J10", value: "Ready" }] }
      ]);
    });

  it("keeps selected blank cells explicit without inventing row data", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "B3", { row: 3, column: 2 });
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Analyze the selected cell.", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "B3", label: "selected cell" });
      expect((result.answer as any).rows).toBeUndefined();
      expect((result.answer as any).emptySummary).toBeUndefined();
      expect(result.warnings[0]).toContain("No non-empty cells");
    });

  it("corrects a conflicting single-column range when the request names a different header", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Count unique values in the Status column.",
        mode: "answer",
        target: { sheetName: "Data", range: "C1:C4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "D1:D4", label: "Status" });
      expect(result.warnings[0]).toContain("Adjusted target range from C1:C4 to D1:D4");
      expect((result.answer as any).kind).toBe("range_value_counts");
      expect((result.answer as any).uniqueCount).toBe(2);
      expect((result.answer as any).topValues).toEqual([
        { value: "Open", count: 2 },
        { value: "Closed", count: 1 }
      ]);
      expect((result.answer as any).resultUri).toMatch(/^excel:\/\/agent\/results\/agentres_/);
    });

  it("does not rewrite a literal column-letter read without a semantic header name", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Count unique values in column C.",
        mode: "answer",
        target: { sheetName: "Data", range: "C1:C4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "C1:C4" });
      expect(result.warnings).toEqual([]);
      expect((result.answer as any).kind).toBe("range_value_counts");
      expect((result.answer as any).uniqueCount).toBe(4);
    });

  it("uses target.column to narrow a sheet read to the matching header column", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "How many unique status values are there?",
        mode: "answer",
        target: { sheetName: "Data", column: "Status" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "D1:D4", label: "Status" });
      expect(result.warnings[0]).toContain("Adjusted target range from A1:D4 to D1:D4");
      expect((result.answer as any).kind).toBe("range_value_counts");
      expect((result.answer as any).uniqueCount).toBe(2);
    });

  it("resolves exact raw sheet targets to the used range when no table exists", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read actual values from Apr 2026",
        mode: "answer",
        target: { sheetName: "Apr 2026" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Apr 2026");
      expect(result.proof[0]?.range).toBe("A1:AJ244");
      expect((result.answer as any).kind).toBe("range_profile");
      expect((result.answer as any).warning).toBeUndefined();
      expect(result.telemetry.internalReadCount).toBe(1);
    });

  it("parses quoted raw sheet A1 references before fuzzy target matching", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read 'Apr 2026'!O1:AE3 actual values",
        responseMode: "verbose",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Apr 2026");
      expect(result.proof[0]?.range).toBe("O1:AE3");
      expect((result.answer as any).sample[0][0]).toBe("Invoice No");
    });

  it("offers raw invoice header blocks as candidates for non-table sheets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read Apr 2026 invoice rows",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Apr 2026");
      expect(result.proof[0]?.range).toBe("O1:AE244");
    });

  it("answers messy worksheet section inventory from cached sampled metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "What sections are on the Operations sheet?",
        mode: "answer",
        target: { sheetName: "Operations" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_sections");
      expect((result.answer as any).sectionCount).toBeGreaterThanOrEqual(4);
      expect((result.answer as any).sections.map((section: any) => section.label)).toContain("KPI summary section");
      expect((result.answer as any).sections.map((section: any) => section.label)).toContain("invoice section");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("builds sampled section metadata when sample reads use the protocol data field", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.returnDataOnly = true;
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "What sections are on the Operations sheet?",
        mode: "answer",
        target: { sheetName: "Operations" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_sections");
      expect((result.answer as any).sectionCount).toBeGreaterThanOrEqual(4);
      expect((result.answer as any).sections.map((section: any) => section.label)).toContain("invoice section");
    });

  it("samples offset used ranges from their true top-left cell", async () => {
      sheets.push({ workbookId: "workbook_agent_unit" as any, worksheetId: "sheet_Offset", name: "Offset", usedRange: { address: "E5001:AN5100", rowCount: 100, columnCount: 36 }, tables: [] });
      try {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);

        const result = await agent.run({
          request: "What sections are on the Offset sheet?",
          mode: "answer",
          target: { sheetName: "Offset" }
        });

        expect(result.status).toBe("SUCCESS");
        expect(runtime.lastBatchOperations.some((operation) => operation.target?.sheetName === "Offset" && operation.target.address === "E5001:AN5020")).toBe(true);
      } finally {
        sheets.pop();
      }
    });

  it("targets one section on an unstructured worksheet instead of reading the whole used range", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Show example rows from the invoice section on Operations",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Operations");
      expect(result.proof[0]?.range).toBe("A10:F13");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(result.telemetry.fullReadCellCount).toBeLessThan(40);
    });

  it("keeps schema-only header requests on cached metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the table header schema",
        mode: "answer",
        target: { tableName: "Transactions" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_schema");
      expect((result.answer as any).source).toBe("cached_metadata");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("answers workbook overview and table list questions from metadata without target ambiguity", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const about = await agent.run({ request: "Can you look into Agent Unit.xlsx file, what is it about?" });
      const tables = await agent.run({ request: "Which tables are in this workbook?" });

      expect(about.status).toBe("SUCCESS");
      expect((about.answer as any).kind).toBe("workbook_overview");
      expect(about.telemetry.internalReadCount).toBe(0);
      expect(tables.status).toBe("SUCCESS");
      expect((tables.answer as any).tables.map((table: any) => table.name)).toContain("Transactions");
      expect(tables.nextAction).toBe("answer_now");
    });

  it("uses active sheet metadata for active/current/this sheet prompts", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Analyze the active sheet." });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("sheet_summary");
      expect(result.proof[0]?.sheetName).toBe("Data");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("reads the active Excel selection without an extra outer tool call", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "B2");
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Analyze the selected cell.", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "B2", label: "selected cell" });
      expect((result.answer as any).rows).toEqual([["A-100"]]);
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      expect((agent.getResultResource(resultId) as any).answer.rows).toEqual([["A-100"]]);
      expect(result.telemetry.internalReadCount).toBe(1);
    });

  it("returns needs input for selection prompts when selection is unavailable", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Analyze the selected cell.", mode: "answer" });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.summary).toContain("tried to read the current live Excel selection");
      expect(result.warnings[0]).toContain("Reload or reopen the OpenWorkbook Local taskpane");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("keeps structural context cached while refreshing selection between runs", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.selection = selectionInfo("Data", "B2");
      const agent = new AgentOrchestrator(runtime as any);

      const first = await agent.run({ request: "Prepare workbook", mode: "prepare" });
      runtime.selection = selectionInfo("Data", "C2", { row: 2, column: 3 });
      const second = await agent.run({ request: "Prepare workbook again", mode: "prepare", workbookContextId: first.workbookContextId });

      expect(second.telemetry.cacheHit).toBe(true);
      expect((second.answer as any).selection.address).toBe("C2");
    });

  it("compares two explicitly named sheets in one agent call", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Compare Financials - May 2026 and Financials - June 2026." });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("comparison_profile");
      expect((result.answer as any).sheets.map((sheet: any) => sheet.sheetName)).toEqual(["Financials - June 2026", "Financials - May 2026"]);
      expect((result.answer as any).numericComparison.highestSumSheet).toBe("Financials - June 2026");
      expect(result.telemetry.internalReadCount).toBe(2);
      expect(runtime.readBatchCount).toBeLessThanOrEqual(sheets.length + 1);
    });

  it("compares monthly performance from KPI sections instead of whole used ranges", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Can you compare Mar and Apr, how our company perform", mode: "answer", responseMode: "verbose" });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("comparison_profile");
      expect((result.answer as any).sheets.map((sheet: any) => sheet.sheetName)).toEqual(["Mar 2026", "Apr 2026"]);
      expect((result.answer as any).sheets.map((sheet: any) => sheet.range)).toEqual(["AG1:AJ20", "AG1:AJ20"]);
      expect((result.answer as any).sheets[0].rows.length).toBeGreaterThan(4);
      expect((result.answer as any).alignedRows.some((row: any) => row.key.includes("cash received"))).toBe(true);
      expect(result.telemetry.internalReadCount).toBe(2);
      expect(result.telemetry.fullReadCellCount).toBeLessThan(200);
    });

  it("normalizes whole-column summary reads and returns complete small rows", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Read all data from columns AG to AJ in the Apr 2026 sheet.", mode: "answer", responseMode: "verbose" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Apr 2026");
      expect(result.proof[0]?.range).toBe("AG1:AJ244");
      expect((result.answer as any).kind).toBe("range_profile");
      expect((result.answer as any).rows.length).toBeGreaterThan(10);
      expect((result.answer as any).rows.some((row: any[]) => row[0] === "Management takeaway")).toBe(true);
    });

  it("resolves natural-language sheet names before answering", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Analyze the June financial sheet", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Financials - June 2026");
      expect(result.summary).toContain("Financials - June 2026");
    });

  it("handles non-English requests safely when the workbook target is clear", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "ช่วยอ่านข้อมูลจาก Financials - June 2026",
        mode: "answer"
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Financials - June 2026");
      expect(result.telemetry.intentSource).toBe("deterministic_fallback");
    });

  it("returns ambiguity instead of guessing when natural language matches competing targets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Analyze financial 2026", mode: "answer" });

      expect(result.status).toBe("AMBIGUOUS_TARGET");
      expect(result.candidates?.map((candidate) => candidate.sheetName)).toContain("Financials - June 2026");
      expect(result.candidates?.map((candidate) => candidate.sheetName)).toContain("Financials - May 2026");
    });

  it("lets agents recover from ambiguous table targets with a returned candidateId", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_table_ambiguity");
      metadata.tables = [
        {
          id: "table:Transactions_January",
          sheetName: "Data",
          name: "Transactions",
          range: "A1:D4",
          headerRange: "A1:D1",
          dataRange: "A2:D4",
          columns: [
            { index: 0, letter: "A", name: "Date", normalizedName: "date", inferredType: "date" },
            { index: 1, letter: "B", name: "Account", normalizedName: "account", inferredType: "text" },
            { index: 2, letter: "C", name: "Amount", normalizedName: "amount", inferredType: "currency" },
            { index: 3, letter: "D", name: "Status", normalizedName: "status", inferredType: "status" }
          ]
        },
        {
          id: "table:Transactions_February",
          sheetName: "Data",
          name: "Transactions",
          range: "F1:I4",
          headerRange: "F1:I1",
          dataRange: "F2:I4",
          columns: [
            { index: 5, letter: "F", name: "Date", normalizedName: "date", inferredType: "date" },
            { index: 6, letter: "G", name: "Account", normalizedName: "account", inferredType: "text" },
            { index: 7, letter: "H", name: "Amount", normalizedName: "amount", inferredType: "currency" },
            { index: 8, letter: "I", name: "Status", normalizedName: "status", inferredType: "status" }
          ]
        }
      ];
      agent.metadataCache.set(metadata);

      const ambiguous = await agent.run({
        request: "Read the transactions table schema",
        mode: "answer",
        workbookContextId: metadata.workbookContextId
      });
      const candidateId = ambiguous.candidates?.find((candidate) => candidate.id === "table:Transactions_February")?.id;
      const resolved = await agent.run({
        request: "Read the selected table schema",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { candidateId }
      });

      expect(ambiguous.status).toBe("AMBIGUOUS_TARGET");
      expect(candidateId).toBe("table:Transactions_February");
      expect(resolved.status).toBe("SUCCESS");
      expect((resolved.answer as any).kind).toBe("table_schema");
      expect((resolved.answer as any).range).toBe("F1:I4");
      expect((resolved.answer as any).schemaSummary.columns.map((column: any) => column.name)).toEqual(["Date", "Account", "Amount", "Status"]);
      expect(resolved.telemetry.internalReadCount).toBe(0);
    });

  it("honors exact tableName targets for schema requests before fuzzy ambiguity", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the schema columns",
        mode: "answer",
        target: { tableName: "Transactions" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_schema");
      expect((result.answer as any).tableName).toBe("Transactions");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("reports missing candidateId without fuzzy-reading a different target", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read the selected table schema",
        mode: "answer",
        target: { candidateId: "table:missing" }
      });

      expect(result.status).toBe("NOT_FOUND");
      expect(result.nextAction).toBe("call_with_target");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("canonicalizes fuzzy sheet names before previewing updates", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update june financial report input",
        mode: "preview_update",
        target: { sheetName: "Financial Jun 2026", range: "B2" },
        values: { values: [[1234]] }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect(result.proof[0]?.sheetName).toBe("Financials - June 2026");
      expect(result.changes?.[0]?.sheetName).toBe("Financials - June 2026");
    });

  it("adds candidate reasons and retry hints for ambiguous targets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Analyze financial 2026",
        mode: "answer",
        responseMode: "verbose"
      });

      expect(result.status).toBe("AMBIGUOUS_TARGET");
      expect(result.candidates?.[0]?.reason).toContain("match");
      expect(result.candidates?.[0]?.nextRequestHint).toContain("target.candidateId");
    });

  it("uses caller target hints to resolve vague read targets deterministically", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Analyze financial 2026",
        mode: "answer",
        responseMode: "verbose",
        intent: { action: "read_values", targetHints: ["Financials - June 2026"] }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Financials - June 2026");
      expect(result.candidates?.[0]?.reason).toContain("caller target hint");
      expect(result.telemetry.targetHintCount).toBe(1);
      expect(result.telemetry.targetHintUsed).toBe(true);
    });

  it("keeps conflicting caller target hints ambiguous", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Analyze financial 2026",
        mode: "answer",
        intent: { action: "read_values", targetHints: ["Financials - June 2026", "Financials - May 2026"] }
      });

      expect(result.status).toBe("AMBIGUOUS_TARGET");
      expect(result.candidates?.slice(0, 2).map((candidate) => candidate.sheetName)).toEqual([
        "Financials - June 2026",
        "Financials - May 2026"
      ]);
      expect(result.telemetry.targetHintCount).toBe(2);
      expect(result.telemetry.targetHintUsed).toBe(false);
    });

  it("keeps explicit targets ahead of misleading caller target hints", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read this target",
        mode: "answer",
        intent: { action: "read_values", targetHints: ["Financials - June 2026"] },
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Data");
      expect(result.proof[0]?.range).toBe("A1:D4");
      expect(result.telemetry.targetHintCount).toBe(1);
      expect(result.telemetry.targetHintUsed).toBe(false);
    });
});
