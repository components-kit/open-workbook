import { describe, expect, it } from "vitest";
import {
  NativeFileBridge,
  RuntimeService,
  appliedTransaction,
  existsSync,
  mkdirSync,
  mkdtempSync,
  operationOk,
  path,
  readFileSync,
  runtimeWithDynamicSnapshotHash,
  runtimeWithExecutingAddin,
  runtimeWithFormulaGraph,
  runtimeWithPersistentAddin,
  runtimeWithSnapshotHash,
  sleepForTest,
  snapshotResponse,
  tmpdir,
  writeFileSync,
  writeFormulaOperation,
  writeStyleOperation,
  writeValuesOperation
} from "./runtime-service.test-support.js";
import type { AgentId, OperationId, PlanId, RuntimeCapabilities, WorkbookId } from "./runtime-service.test-support.js";

describe("RuntimeService formula dependency graph", () => {
  it("resolves structured references with table metadata from the add-in", async () => {
    const workbookId = "workbook_runtime_formula_graph" as WorkbookId;
    const runtime = runtimeWithFormulaGraph(workbookId);

    const result = await runtime.getFormulaDependencyGraph({
      workbookId,
      sheetName: "Report",
      address: "B2"
    });

    expect(result.ok).toBe(true);
    if (result.ok) {
      const rangeNodes = result.graph.nodes.filter((node) => node.kind === "range" && node.formula === undefined);
      expect(rangeNodes.some((node) => node.sheetName === "Transactions" && node.address === "C2:C10")).toBe(true);
    }
  });
});

describe("RuntimeService performance-oriented scoping", () => {
  it("forwards projected table read options to the add-in", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_table_read_projection" as WorkbookId;
    const session = runtime.sessions.createSession();
    let forwarded: any;
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "table.read") {
          forwarded = params;
          return {
            ok: true,
            table: {
              info: { workbookId, tableName: params.tableName, rowCount: 100, columnCount: 3, columns: [] },
              headers: [["Account", "Opportunity", "Amount"]],
              values: [["Acme"]]
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.readTable({
      workbookId,
      tableName: "Opportunities",
      columns: ["Account"],
      rowOffset: 10,
      rowLimit: 5,
      includeFormulas: false,
      includeText: false,
      includeNumberFormats: false
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect(forwarded).toMatchObject({
      workbookId,
      tableName: "Opportunities",
      columns: ["Account"],
      rowOffset: 10,
      rowLimit: 5,
      includeFormulas: false,
      includeText: false,
      includeNumberFormats: false
    });
  });

  it("normalizes only the detected header row instead of rewriting the full clean range", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_header_scope" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: Array<{ method: string; params: any }> = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push({ method, params });
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "header_scope_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              fingerprint: { range, hash: `hash_${range.address}`, cellCount: 9, capturedAt: new Date().toISOString() },
              values: [
                ["Account Name", "Opportunity Name", "Close Date"],
                ["Acme", "Renewal", "2026-01-01"],
                ["Globex", "Expansion", "2026-02-01"]
              ]
            }))
          };
        }
        if (method === "operation.execute_batch") {
          return { ok: true, rollbackAvailable: true, backups: [], warnings: [], telemetry: {} };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.cleanNormalizeHeaders({
      workbookId,
      sheetName: "Opportunity Tracking",
      address: "A5:C7",
      headerRowIndex: 0
    });

    const executed = calls.find((call) => call.method === "operation.execute_batch");
    expect(result.ok).toBe(true);
    expect(executed?.params.request.operations[0].target.address).toBe("A5:C5");
    expect(executed?.params.request.operations[0].values).toEqual([["account_name", "opportunity_name", "close_date"]]);
  });

  it("reorders table columns through a table mutation with a scoped backup", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_table_reorder" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: Array<{ method: string; params: any }> = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push({ method, params });
        if (method === "table.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              tableName: "Opportunities",
              sheetName: "Opportunity Tracking",
              address: "A1:C10",
              rowCount: 9,
              columnCount: 3,
              columns: [
                { index: 0, name: "Opportunity Name" },
                { index: 1, name: "Account / Company" },
                { index: 2, name: "Stage" }
              ]
            }
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: { workbookId, workbookHash: "table_reorder_workbook", structureHash: "structure", capturedAt: new Date().toISOString() },
            rangeSnapshots: params.ranges.map((range: any) => ({
              fingerprint: { range, hash: "table_reorder_range", cellCount: 30, capturedAt: new Date().toISOString() },
              values: [["snapshot"]]
            }))
          };
        }
        if (method === "table.reorder_columns") {
          return { ok: true, info: { workbookId, tableName: params.tableName, rowCount: 9, columnCount: 3, columns: [] }, warnings: [] };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.reorderTableColumns({
      workbookId,
      tableName: "Opportunities",
      columnOrder: ["Account / Company", "Opportunity Name", "Stage"]
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect(calls.some((call) => call.method === "workbook.snapshot_ranges" && call.params.ranges[0].address === "A1:C10")).toBe(true);
    expect(calls.some((call) => call.method === "table.reorder_columns")).toBe(true);
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
  });

  it("blocks table reorder when the requested column order is incomplete", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_table_reorder_invalid" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: string[] = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        calls.push(method);
        if (method === "table.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              tableName: "Opportunities",
              rowCount: 9,
              columnCount: 2,
              columns: [
                { index: 0, name: "Opportunity Name" },
                { index: 1, name: "Account / Company" }
              ]
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.reorderTableColumns({
      workbookId,
      tableName: "Opportunities",
      columnOrder: ["Account / Company"]
    });

    expect((result as { ok?: boolean }).ok).toBe(false);
    expect((result as { error?: { code?: string } }).error?.code).toBe("INVALID_ARGUMENT");
    expect(calls).toEqual(["table.get_info"]);
    expect(runtime.transactions.list(workbookId)).toHaveLength(0);
  });
});
