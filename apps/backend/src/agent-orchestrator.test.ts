import { describe, expect, it } from "vitest";
import type { BatchRequest, WorkbookId } from "@components-kit/open-workbook-protocol";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { createMetadataFingerprint, type WorkbookMetadata } from "./workbook-metadata-cache.js";

const workbookId = "workbook_agent_unit" as WorkbookId;
const activeWorkbook = { workbookId, name: "Agent Unit.xlsx", path: "/tmp/Agent Unit.xlsx" };
const sheets = [
  { workbookId, worksheetId: "sheet_Data", name: "Data", usedRange: { address: "A1:D4", rowCount: 4, columnCount: 4 }, tables: [{ name: "Transactions" }] },
  { workbookId, worksheetId: "sheet_Report", name: "Report", usedRange: { address: "A1:B20", rowCount: 20, columnCount: 2 }, tables: [] },
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
    expect(second.telemetry.cacheHit).toBe(true);
    expect(second.telemetry.metadataCacheStatus).toBe("hit");
    expect(runtime.readBatchCount).toBe(readCallsAfterFirstPrepare);
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
    const blocked = await agent.run({
      request: "Update formula area",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "A12" },
      values: { values: [["bad"]] }
    });

    expect(allowed.status).toBe("PREVIEW_READY");
    expect(blocked.status).toBe("VALIDATION_FAILED");
  });

  it("answers with one targeted read instead of repeated broad reads", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({ request: "Answer amount from Transactions table", mode: "answer" });

    expect(result.status).toBe("SUCCESS");
    expect(result.telemetry.internalReadCount).toBe(1);
    expect(result.telemetry.fullReadCellCount).toBeLessThanOrEqual(16);
    expect(runtime.readBatchCount).toBeLessThanOrEqual(6);
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

  it("auto-applies clear scoped value edits after safe preview", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "Change Data B2 to 999",
      target: { sheetName: "Data", range: "B2" },
      values: { values: [[999]] }
    });

    expect(result.status).toBe("SUCCESS");
    expect(result.mode).toBe("auto");
    expect(result.confirmationToken).toBeUndefined();
    expect(result.metrics?.autoApplied).toBe(true);
    expect(result.telemetry.autoApplied).toBe(true);
    expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
    expect(runtime.writeBatchCount).toBe(1);
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

  sessions = { getActive: () => ({ activeWorkbook }) };

  getStatus() {
    return { activeAddinConnected: true };
  }

  async getActiveContext() {
    return { ok: true, activeWorkbook };
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
    const operation = request.operations[0];
    if (operation?.kind === "range.read_full") {
      this.readBatchCount += 1;
      const sheetName = operation.target.sheetName;
      const address = operation.target.address;
      return {
        ok: true,
        readData: [{
          operationId: operation.operationId,
          snapshot: {
            values: valuesFor(sheetName, address),
            formulas: formulasFor(sheetName, address),
            text: valuesFor(sheetName, address).map((row) => row.map((value) => value === null || value === undefined ? "" : String(value)))
          }
        }],
        telemetry: { cellsRead: valuesFor(sheetName, address).flat().length }
      };
    }
    this.writeBatchCount += 1;
    return { ok: true, backups: [], warnings: [], telemetry: { cellsWritten: 1 } };
  }

  async validateWorkbook() {
    return { ok: true, issues: [] };
  }
}

function valuesFor(sheetName: string, address: string) {
  if (sheetName === "Data") {
    if (address === "B2") return [["A-100"]];
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
  if (address === "B1") return [["input"]];
  if (address === "A12") return [["=SUM(B1:B10)"]];
  return [
    ["Metric", "Value"],
    ["Revenue", 1000],
    ["Total", 1000]
  ];
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
    workbook: { workbookId, name: activeWorkbook.name, path: activeWorkbook.path, sheetCount: sheets.length, activeSheet: "Data" },
    sheets: [
      { id: "sheet:0", name: "Data", index: 0, usedRange: "A1:D4", rowCount: 4, columnCount: 4, kind: "transaction", headers: [], tableIds: ["table:Transactions"], summaryBlockIds: [], formulaRegionIds: [] },
      { id: "sheet:1", name: "Report", index: 1, usedRange: "A1:B20", rowCount: 20, columnCount: 2, kind: "summary", headers: [], tableIds: [], summaryBlockIds: [], formulaRegionIds: ["formula:manual"] }
    ],
    tables: [{ id: "table:Transactions", sheetName: "Data", name: "Transactions", range: "A1:D4", columns: [] }],
    namedRanges: [],
    summaryBlocks: [],
    formulaRegions: [],
    fingerprint,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 60_000
  };
}
