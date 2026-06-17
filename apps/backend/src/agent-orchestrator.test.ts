import { describe, expect, it } from "vitest";
import type { BatchRequest, WorkbookId } from "@components-kit/open-workbook-protocol";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { createMetadataFingerprint, type WorkbookMetadata } from "./workbook-metadata-cache.js";

const workbookId = "workbook_agent_unit" as WorkbookId;
const activeWorkbook = { workbookId, name: "Agent Unit.xlsx", path: "/tmp/Agent Unit.xlsx" };
const sheets = [
  { workbookId, worksheetId: "sheet_Data", name: "Data", usedRange: { address: "A1:D4", rowCount: 4, columnCount: 4 }, tables: [{ name: "Transactions" }] },
  { workbookId, worksheetId: "sheet_Report", name: "Report", usedRange: { address: "A1:B20", rowCount: 20, columnCount: 2 }, tables: [] },
  { workbookId, worksheetId: "sheet_Operations", name: "Operations", usedRange: { address: "A1:J24", rowCount: 24, columnCount: 10 }, tables: [] },
  { workbookId, worksheetId: "sheet_Mar", name: "Mar 2026", usedRange: { address: "A1:AJ206", rowCount: 206, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_Apr", name: "Apr 2026", usedRange: { address: "A1:AJ244", rowCount: 244, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_FinJun", name: "Financials - June 2026", usedRange: { address: "A1:C4", rowCount: 4, columnCount: 3 }, tables: [] },
  { workbookId, worksheetId: "sheet_FinMay", name: "Financials - May 2026", usedRange: { address: "A1:C4", rowCount: 4, columnCount: 3 }, tables: [] }
];

describe("AgentOrchestrator", () => {
  it("reuses prepared metadata and reports cache telemetry", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const first = await agent.run({ request: "Prepare workbook", mode: "prepare" });
    const readCallsAfterFirstPrepare = runtime.readBatchCount;
    const second = await agent.run({ request: "Prepare workbook again", mode: "prepare", workbookContextId: first.workbookContextId });

    expect(first.status).toBe("SUCCESS");
    expect(first.telemetry.cacheHit).toBe(false);
    expect(first.telemetry.metadataCacheStatus).toBe("miss");
    expect(readCallsAfterFirstPrepare).toBe(0);
    expect(second.telemetry.cacheHit).toBe(true);
    expect(second.telemetry.metadataCacheStatus).toBe("hit");
    expect(runtime.readBatchCount).toBe(readCallsAfterFirstPrepare);
  });

  it("answers vague workbook file reviews from complete structure metadata without sheet sampling", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({ request: "Can you look into transactions.xlsx?" });

    expect(result.status).toBe("SUCCESS");
    expect((result.answer as any).kind).toBe("workbook_overview");
    expect((result.answer as any).sheets.map((sheet: any) => sheet.name)).toEqual(sheets.map((sheet) => sheet.name));
    expect(result.telemetry.internalReadCount).toBe(0);
    expect(runtime.readBatchCount).toBe(0);
  });

  it("upgrades prepared structure metadata when a later answer needs sheet samples", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const prepared = await agent.run({ request: "Prepare workbook", mode: "prepare" });
    const answered = await agent.run({
      request: "Read Apr 2026 invoice rows",
      mode: "answer",
      workbookContextId: prepared.workbookContextId
    });

    expect(prepared.status).toBe("SUCCESS");
    expect(answered.status).toBe("SUCCESS");
    expect(answered.proof[0]?.sheetName).toBe("Apr 2026");
    expect(answered.proof[0]?.range).toBe("O1:AE244");
    expect(runtime.readBatchCount).toBeGreaterThan(0);
  });

  it("bounds candidates and reports token savings when a response budget is provided", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Find amount status region report data total formula",
      mode: "find",
      budget: { maxExamples: 2, maxPayloadBytes: 850 }
    });

    expect(result.status).toBe("SUCCESS");
    expect(result.candidates?.length).toBeLessThanOrEqual(2);
    expect(result.proof.length).toBeLessThanOrEqual(2);
    expect(result.telemetry.candidateCount).toBe(result.candidates?.length ?? 0);
    expect(result.telemetry.payloadBytes).toBeLessThanOrEqual(1000);
    expect(result.telemetry.estimatedTokensSaved).toBeGreaterThanOrEqual(0);
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

  it("requires preview confirmation token before apply", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Update Data B2",
      mode: "preview_update",
      target: { sheetName: "Data", range: "B2" },
      values: { values: [[999]] }
    });
    const missingToken = await agent.run({
      request: "Apply update",
      mode: "apply_update",
      operationId: preview.operationId
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.confirmationToken).toBeTruthy();
    expect(missingToken.status).toBe("NEEDS_INPUT");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews and applies grouped range patches as one batch operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Update related pricing zones",
      mode: "preview_update",
      values: {
        patches: [
          { target: { sheetName: "Data", range: "B2:C2" }, values: [[9000, 9000]], reason: "Zone A price" },
          { target: { sheetName: "Data", range: "B3:C3" }, values: [[6000, 6000]], reason: "Zone B price" }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply grouped pricing update",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("multi_range_preview");
    expect((preview.answer as any).patchCount).toBe(2);
    expect(preview.summary).toContain("Apply this grouped preview once");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.writeBatchCount).toBe(1);
    expect(runtime.lastWriteOperations.map((operation) => operation.target.address)).toEqual(["B2:C2", "B3:C3"]);
  });

  it("blocks unsafe grouped patches before creating a pending batch", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Update related cells",
      mode: "preview_update",
      values: {
        patches: [
          { target: { sheetName: "Data", range: "B2:C2" }, values: [[1, 2]] },
          { target: { sheetName: "Data", range: "B3:C3" }, values: [["=SUM(A1:A2)", 3]] }
        ]
      }
    });

    expect(result.status).toBe("VALIDATION_FAILED");
    expect(result.summary).toContain("formula-like");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("treats apply results without explicit ok as success when validation passes", async () => {
    const runtime = new FakeAgentRuntime();
    runtime.omitOkOnWrite = true;
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Update Data B2",
      mode: "preview_update",
      target: { sheetName: "Data", range: "B2" },
      values: { values: [[999]] }
    });
    const applied = await agent.run({
      request: "Apply update",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(applied.status).toBe("SUCCESS");
    expect((applied.answer as any).ok).toBe(true);
    expect((applied.answer as any).validationOk).toBe(true);
  });

  it("uses parsed A1 overlap checks for formula-protected regions", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_formula_overlap");
    metadata.formulaRegions = [{ id: "formula:manual", sheetName: "Report", range: "A10:A20", formulaCount: 11 }];
    agent.metadataCache.set(metadata);

    const allowed = await agent.run({
      request: "Update report input",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "B1" },
      values: { values: [["ok"]] }
    });
    const formulaPreview = await agent.run({
      request: "Update formula area",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "A12" },
      values: { values: [["bad"]] }
    });

    expect(allowed.status).toBe("PREVIEW_READY");
    expect(formulaPreview.status).toBe("PREVIEW_READY");
    expect(formulaPreview.warnings).toContain("Target overlaps detected formula regions. Review carefully before applying.");
  });

  it("answers with one targeted read instead of repeated broad reads", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({ request: "Answer amount from Transactions table", mode: "answer" });

    expect(result.status).toBe("SUCCESS");
    expect(result.telemetry.internalReadCount).toBe(1);
    expect(result.telemetry.fullReadCellCount).toBeLessThanOrEqual(16);
    expect(runtime.readBatchCount).toBeLessThanOrEqual(8);
  });

  it("reads live values when batch results use the protocol data field", async () => {
    const runtime = new FakeAgentRuntime();
    runtime.returnDataOnly = true;
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Read headers and first 5 rows from Data A1:D4",
      mode: "answer",
      target: { sheetName: "Data", range: "A1:D4" }
    });

    expect(result.status).toBe("SUCCESS");
    expect((result.answer as any).source).toBe("live_read");
    expect((result.answer as any).sample[1][1]).toBe("A-100");
    expect(result.telemetry.internalReadCount).toBe(1);
    expect(result.telemetry.fullReadCellCount).toBeGreaterThan(0);
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
    expect((result.answer as any).emptySummary.emptyCells).toBe(97);
    expect((result.answer as any).sparseRows).toEqual([
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
    expect((result.answer as any).emptySummary).toMatchObject({ sourceCells: 1, nonEmptyCells: 0, emptyCells: 1 });
    expect((result.answer as any).warning).toContain("No non-empty cells");
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
    expect(result.proof[0]?.sheetName).toBe("Data");
    expect(result.telemetry.internalReadCount).toBe(1);
  });

  it("reads the active Excel selection without an extra outer tool call", async () => {
    const runtime = new FakeAgentRuntime();
    runtime.selection = selectionInfo("Data", "B2");
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({ request: "Analyze the selected cell.", mode: "answer" });

    expect(result.status).toBe("SUCCESS");
    expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "B2", label: "selected cell" });
    expect((result.answer as any).rows).toEqual([["A-100"]]);
    expect(result.telemetry.internalReadCount).toBe(1);
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

  it("returns needs input for selection prompts when selection is unavailable", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({ request: "Analyze the selected cell.", mode: "answer" });

    expect(result.status).toBe("NEEDS_INPUT");
    expect(result.summary).toContain("selection is unavailable");
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
    expect(runtime.readBatchCount).toBeLessThanOrEqual(8);
  });

  it("compares monthly performance from KPI sections instead of whole used ranges", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({ request: "Can you compare Mar and Apr, how our company perform", mode: "answer" });

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

    const result = await agent.run({ request: "Read all data from columns AG to AJ in the Apr 2026 sheet.", mode: "answer" });

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
    expect((resolved.answer as any).columns.map((column: any) => column.name)).toEqual(["Date", "Account", "Amount", "Status"]);
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

  it("returns previews for scoped auto value edits by default", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Change Data B2 to 999",
      target: { sheetName: "Data", range: "B2" },
      values: { values: [[999]] }
    });

    expect(result.status).toBe("PREVIEW_READY");
    expect(result.mode).toBe("auto");
    expect(result.confirmationToken).toBeTruthy();
    expect(result.telemetry.autoApplied).toBeUndefined();
    expect(result.telemetry.safetyDecision).toBe("manual_review:auto_apply_disabled");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews small add requests with explicit values and range", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Add a notes block to Report",
      target: { sheetName: "Report", range: "B1:B2" },
      values: { values: [["Owner"], ["Finance"]] }
    });

    expect(result.status).toBe("PREVIEW_READY");
    expect(result.telemetry.safetyDecision).toBe("manual_review:auto_apply_disabled");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews header formatting as a style mutation instead of reading values", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Format the header row on the Data sheet.",
      target: { sheetName: "Data", range: "A1:D1" }
    });

    expect(result.status).toBe("PREVIEW_READY");
    expect((result.answer as any).kind).toBe("style_preview");
    expect(result.proof[0]?.range).toBe("A1:D1");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews formula writes through a formula-aware path", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Duplicate the formula from A12 down",
      mode: "preview_update",
      target: { sheetName: "Report", range: "A13:A14" },
      values: { values: [["=SUM(B1:B10)"], ["=SUM(B2:B11)"]] }
    });

    expect(result.status).toBe("PREVIEW_READY");
    expect((result.answer as any).kind).toBe("formula_preview");
    expect(result.nextAction).toBe("call_apply_update");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews duplicate sheet template cleanup as one pending operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Can you duplicate Report sheet, remove data and keep only template?",
      target: { sheetName: "Report" }
    });

    expect(result.status).toBe("PREVIEW_READY");
    expect((result.answer as any).kind).toBe("template_cleanup_preview");
    expect((result.answer as any).sourceSheetName).toBe("Report");
    expect(result.nextAction).toBe("call_apply_update");
  });

  it("previews and applies table appends through the agent surface", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Append transaction rows to the Transactions table",
      mode: "preview_update",
      target: { tableName: "Transactions" },
      values: { rows: [["2026-06-01", 204, "71-4653", "Company gas top-up"]] }
    });
    const applied = await agent.run({
      request: "Apply append",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("table_append_preview");
    expect((preview.answer as any).tableName).toBe("Transactions");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.appendTableRowCount).toBe(1);
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("does not auto-apply ambiguous natural-language updates", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Update financial 2026",
      values: { values: [[999]] }
    });

    expect(result.status).toBe("AMBIGUOUS_TARGET");
    expect(result.nextAction).toBe("call_with_target");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("does not auto-apply formula-sensitive requests as value writes", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Fix formula in Report A12",
      target: { sheetName: "Report", range: "A12" },
      values: { values: [[100]] }
    });

    expect(result.status).toBe("NEEDS_INPUT");
    expect(result.nextAction).toBe("manual_review");
    expect(result.telemetry.safetyDecision).toBe("manual_review:advanced_workflow");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("blocks formula-like values in generic value previews", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Change Data B2",
      mode: "preview_update",
      target: { sheetName: "Data", range: "B2" },
      values: { values: [["=SUM(A1:A2)"]] }
    });

    expect(result.status).toBe("VALIDATION_FAILED");
    expect(result.nextAction).toBe("manual_review");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("does not auto-apply sparse broad writes", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Change Report B1:C4",
      target: { sheetName: "Report", range: "B1:C4" },
      values: { values: [[1]] }
    });

    expect(result.status).toBe("VALIDATION_FAILED");
    expect(runtime.writeBatchCount).toBe(0);
  });
});

class FakeAgentRuntime {
  readBatchCount = 0;
  writeBatchCount = 0;
  appendTableRowCount = 0;
  returnDataOnly = false;
  omitOkOnWrite = false;
  lastWriteOperations: Array<Extract<BatchRequest["operations"][number], { target: any }>> = [];
  selection: any;
  readiness: any;

  sessions = { getActive: () => ({ activeWorkbook }) };

  getStatus() {
    return { activeAddinConnected: true, connectionState: "ready", activeWorkbookAvailable: true, activeWorkbook };
  }

  async getConnectionReadiness() {
    if (this.readiness) return this.readiness;
    return { ok: true, connectionState: "ready", activeWorkbook, status: this.getStatus() };
  }

  async getActiveContext() {
    return { ok: true, activeWorkbook };
  }

  async getSelection() {
    return this.selection ? { workbook: activeWorkbook, selection: this.selection } : { ok: false };
  }

  async getWorkbookMap() {
    return { ok: true, map: { workbook: activeWorkbook, activeSheet: "Data", sheets } };
  }

  async listTables() {
    return { ok: true, tables: [{ name: "Transactions" }] };
  }

  async getTableInfo() {
    return {
      ok: true,
      info: {
        tableName: "Transactions",
        sheetName: "Data",
        address: "A1:D4",
        headerAddress: "A1:D1",
        dataRange: "A2:D4",
        columns: [{ name: "Date" }, { name: "Account" }, { name: "Amount" }, { name: "Status" }]
      }
    };
  }

  async listNames() {
    return { ok: true, names: [{ workbookId, name: "RevenueTotal", scope: "workbook", address: "Report!B2" }] };
  }

  listRegions() {
    return { ok: true, regions: [{ name: "InputRegion", sheetName: "Report", address: "B1:B3" }] };
  }

  async applyBatch(request: BatchRequest) {
    if (request.operations.every((operation) => operation.kind === "range.read_full")) {
      this.readBatchCount += 1;
      const readData = request.operations.map((operation) => {
        const sheetName = operation.target.sheetName;
        const address = operation.target.address;
        return {
          operationId: operation.operationId,
          snapshot: {
            values: valuesFor(sheetName, address),
            formulas: formulasFor(sheetName, address),
            text: valuesFor(sheetName, address).map((row) => row.map((value) => value === null || value === undefined ? "" : String(value)))
          }
        };
      });
      return {
        ok: true,
        ...(this.returnDataOnly ? { data: readData } : { readData }),
        telemetry: { cellsRead: request.operations.reduce((total, operation) => total + valuesFor(operation.target.sheetName, operation.target.address).flat().length, 0) }
      };
    }
    this.writeBatchCount += 1;
    this.lastWriteOperations = request.operations.filter((operation): operation is Extract<BatchRequest["operations"][number], { target: any }> => "target" in operation);
    return {
      ...(this.omitOkOnWrite ? {} : { ok: true }),
      backups: [],
      warnings: [],
      telemetry: { cellsWritten: 1 }
    };
  }

  async appendTableRows() {
    this.appendTableRowCount += 1;
    return { ok: true, backups: [], warnings: [], telemetry: { rowsWritten: 1 } };
  }

  async validateWorkbook() {
    return { ok: true, issues: [] };
  }
}

function selectionInfo(sheetName: string, address: string, position = { row: 2, column: 2 }) {
  return {
    workbookId,
    sheetName,
    address,
    startCell: {
      workbookId,
      sheetName,
      address,
      row: position.row,
      column: position.column,
      rowIndex: position.row - 1,
      columnIndex: position.column - 1
    },
    endCell: {
      workbookId,
      sheetName,
      address,
      row: position.row,
      column: position.column,
      rowIndex: position.row - 1,
      columnIndex: position.column - 1
    },
    rowCount: 1,
    columnCount: 1,
    cellCount: 1,
    isSingleCell: true
  };
}

function valuesFor(sheetName: string, address: string) {
  if (sheetName === "Data") {
    if (address === "A1:J10") {
      return [
        ["Owner", "", "", "", "", "", "", "", "", "Status"],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "Ready"]
      ];
    }
    if (address === "B2") return [["A-100"]];
    if (address === "B3") return [[""]];
    if (address === "C2") return [[100]];
    if (address === "C1:C4") return [["Amount"], [100], [200], [300]];
    return [
      ["Date", "Account", "Amount", "Status"],
      ["2026-01-01", "A-100", 100, "Open"],
      ["2026-01-02", "A-200", 200, "Closed"],
      ["2026-01-03", "A-300", 300, "Open"]
    ];
  }
  if (sheetName === "Financials - June 2026") {
    if (address === "B2") return [[1200]];
    return [["Metric", "Jun 2026", "Variance"], ["Revenue", 1200, 50], ["Expense", 700, -20], ["Profit", 500, 70]];
  }
  if (sheetName === "Financials - May 2026") {
    return [["Metric", "May 2026", "Variance"], ["Revenue", 1100, 30], ["Expense", 680, 10], ["Profit", 420, 20]];
  }
  if (sheetName === "Mar 2026" || sheetName === "Apr 2026") {
    return monthlySheetValues(sheetName, address);
  }
  if (sheetName === "Operations") {
    const rows = [
      ["Operations Review", "", "", "", "", "", "", "", "", ""],
      ["Period", "Jun 2026", "Prepared by", "Ops", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["Metric", "Value", "Status", "", "", "", "", "Owner", "Action", "Due"],
      ["Revenue", 7200, "On Track", "", "", "", "", "A. Chen", "Review invoices", "2026-06-20"],
      ["Expense", 2800, "Watch", "", "", "", "", "M. Lee", "Check variance", "2026-06-21"],
      ["Profit", 4400, "On Track", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["Invoice No", "Customer", "Job", "Amount", "Status", "Owner", "", "", "", ""],
      ["INV-100", "Acme", "Lift", 1200, "Open", "A. Chen", "", "", "", ""],
      ["INV-101", "Northwind", "Haul", 900, "Paid", "M. Lee", "", "", "", ""],
      ["INV-102", "Contoso", "Storage", 450, "Open", "A. Chen", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["Reconciliation", "", "", "", "", "", "", "", "", ""],
      ["Expected", "Actual", "Variance", "Check", "", "", "", "", "", ""],
      [2500, 2550, "=B17-A17", "Review", "", "", "", "", "", ""]
    ];
    if (address === "A10:F13") {
      return rows.slice(9, 13).map((row) => row.slice(0, 6));
    }
    return rows;
  }
  if (address === "B1") return [["input"]];
  if (address === "A12") return [["=SUM(B1:B10)"]];
  return [
    ["Metric", "Value"],
    ["Revenue", 1000],
    ["Total", 1000]
  ];
}

function monthlySheetValues(sheetName: string, address: string) {
  const month = sheetName.startsWith("Mar") ? "March" : "April";
  const cashReceived = sheetName.startsWith("Mar") ? 280000 : 333881.72;
  const cashSpent = sheetName.startsWith("Mar") ? 240000 : 363263.96;
  const profit = cashReceived - cashSpent;
  const summaryRows = [
    [`${month} 2026 Summary`, "", "", ""],
    ["", "", "", ""],
    ["Cash Summary", "", "", ""],
    ["Metric", "Amount (THB)", "View", "Notes"],
    [`Cash received in ${month}`, cashReceived, "Cash in", `Actual cash/bank inflows received in ${month}.`],
    [`Cash spent in ${month}`, cashSpent, "Cash out", `Actual cash/bank outflows paid in ${month}.`],
    ["Net cash movement", profit, "Cash net", "Cash received less cash spent."],
    ["", "", "", ""],
    ["P&L / Credit Float", "", "", ""],
    ["Billed / earned transport revenue", cashReceived + 45000, "Revenue", "Invoice revenue for jobs in the month."],
    ["Operating spend", cashSpent, "Expense", "Cash operating spend."],
    ["Profit / loss", profit, "Profit", "Revenue less spend."],
    ["", "", "", ""],
    ["Spend Breakdown", "", "", ""],
    ["Company gas top-up", sheetName.startsWith("Mar") ? 58000 : 71000, "Expense", "Fuel-related spend."],
    ["Truck repair", sheetName.startsWith("Mar") ? 18000 : 42000, "Expense", "Maintenance spend."],
    ["Driver salary", sheetName.startsWith("Mar") ? 64000 : 65000, "Expense", "Payroll spend."],
    ["", "", "", ""],
    ["Checks / Interpretation", "", "", ""],
    ["Management takeaway", profit >= 0 ? "Profitable" : "Loss", "Status", profit >= 0 ? "Month remained profitable." : "Cash out exceeded cash in."]
  ];
  if (address.startsWith("AG")) {
    return summaryRows;
  }
  const row1 = padToSummary([
    "Transaction Date", "Job ID", "Truck ID", "Description", "Transaction Type", "Direction", "Cash Amount", "Actual Amount", "Payment Variance", "Reconciliation Note", "Transfer From/To", "Proof File", "Detail Notes", "",
    "Invoice No", "Job ID", "Invoice Date", "Billed To", "Booking No", "Customer", "Job", "Container No", "Container Size", "Job Price", "Lifting On", "Lifting Off", "Total Lifting", "Other Fees", "Gross Billed", "W/H Tax", "Net Collect"
  ], summaryRows[0]);
  const row2 = padToSummary([
    sheetName.startsWith("Mar") ? "2026-03-01" : "2026-04-01", "204", "71-4653", "Company gas top-up", "company_gas_topup", "Outflow", "2211.21", "2211.21", "0", "", "Bank", "proof.pdf", "text note", "",
    "INV-001", "204", sheetName.startsWith("Mar") ? "2026-03-01" : "2026-04-01", "ACME", "BK-001", "Customer A", "Job 204", "CONT-1", "20GP", "10000", "1000", "1000", "2000", "0", "12000", "360", "11640"
  ], summaryRows[1]);
  const rows = [row1, row2, ...summaryRows.slice(2).map((summary) => padToSummary([], summary))];
  if (address.startsWith("O")) {
    return rows.map((row) => row.slice(14, 31));
  }
  return rows;
}

function padToSummary(left: unknown[], summary: unknown[]) {
  const row = Array.from({ length: 36 }, (_cell, index) => left[index] ?? "");
  summary.forEach((value, index) => {
    row[32 + index] = value;
  });
  return row;
}

function formulasFor(sheetName: string, address: string) {
  const values = valuesFor(sheetName, address);
  return values.map((row) => row.map((value) => typeof value === "string" && value.startsWith("=") ? value : null));
}

function createCachedMetadata(workbookContextId: string): WorkbookMetadata {
  const fingerprint = createMetadataFingerprint({ workbookId, workbook: activeWorkbook, sheets });
  return {
    workbookContextId,
    workbookKey: `id:${workbookId}`,
    detailLevel: "sampled",
    workbook: { workbookId, name: activeWorkbook.name, path: activeWorkbook.path, sheetCount: sheets.length, activeSheet: "Data" },
    sheets: [
      { id: "sheet:0", name: "Data", index: 0, usedRange: "A1:D4", rowCount: 4, columnCount: 4, kind: "transaction", headers: [], tableIds: ["table:Transactions"], sectionIds: [], summaryBlockIds: [], formulaRegionIds: [] },
      { id: "sheet:1", name: "Report", index: 1, usedRange: "A1:B20", rowCount: 20, columnCount: 2, kind: "summary", headers: [], tableIds: [], sectionIds: [], summaryBlockIds: [], formulaRegionIds: ["formula:manual"] }
    ],
    tables: [{ id: "table:Transactions", sheetName: "Data", name: "Transactions", range: "A1:D4", columns: [] }],
    namedRanges: [],
    sections: [],
    summaryBlocks: [],
    formulaRegions: [],
    fingerprint,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 60_000
  };
}
