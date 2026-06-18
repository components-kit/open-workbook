import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import type { AgentId, ExcelOperation, OperationId, PlanId, RuntimeCapabilities, TransactionRecord, WorkbookId, WorkbookScope } from "@components-kit/open-workbook-protocol";
import { NativeFileBridge } from "./native-file-bridge.js";
import { RuntimeService } from "./runtime-service.js";

export { existsSync, mkdirSync, mkdtempSync, NativeFileBridge, path, readFileSync, RuntimeService, tmpdir, writeFileSync };
export type { AgentId, ExcelOperation, OperationId, PlanId, RuntimeCapabilities, TransactionRecord, WorkbookId, WorkbookScope };
export function appliedTransaction(runtime: RuntimeService, workbookId: WorkbookId, planId: PlanId | undefined, scopes: WorkbookScope[]): TransactionRecord {
  const transaction = runtime.transactions.create({
    workbookId,
    planId,
    goal: "Applied transaction",
    scopes,
    destructiveLevel: "values"
  });
  runtime.transactions.markApplying(transaction.transactionId, []);
  runtime.transactions.markApplied(transaction.transactionId, { backups: [], warnings: [] });
  return transaction;
}

export function runtimeWithSnapshotHash(hash: string): RuntimeService {
  return runtimeWithDynamicSnapshotHash(() => hash);
}

export function runtimeWithDynamicSnapshotHash(hash: () => string): RuntimeService {
  const runtime = new RuntimeService({ persistState: false });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (_method: string, params: any) => ({
      workbookFingerprint: {
        workbookId: params.workbookId,
        workbookHash: `workbook_${hash()}`,
        structureHash: "structure",
        capturedAt: new Date().toISOString()
      },
      rangeSnapshots: params.ranges.map((range: any) => ({
        fingerprint: {
          range,
          hash: hash(),
          cellCount: 1,
          capturedAt: new Date().toISOString()
        }
      }))
    })
  } as any);
  return runtime;
}

export function runtimeWithPersistentAddin(stateDir: string, workbookId: WorkbookId): RuntimeService {
  const runtime = new RuntimeService({ stateDir });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (method: string, params: any) => {
      if (method === "template.capture") {
        return {
          sourceSheetName: params.sourceSheetName,
          dataRegions: params.dataRegions,
          fingerprintPayload: {
            structure: { sheets: ["Template", "Report"] },
            formulas: { range: "B2:D20" },
            styles: { theme: "default" },
            filters: {},
            tables: {},
            printLayout: {}
          }
        };
      }
      if (method === "workbook.snapshot_ranges") {
        return {
          workbookFingerprint: {
            workbookId,
            workbookHash: "workbook_persist_hash",
            structureHash: "structure",
            capturedAt: new Date().toISOString()
          },
          rangeSnapshots: params.ranges.map((range: any) => ({
            range,
            values: [["snapshot"]],
            fingerprint: {
              range,
              hash: "range_persist_hash",
              cellCount: 1,
              capturedAt: new Date().toISOString()
            }
          }))
        };
      }
      throw new Error(`Unexpected method ${method}`);
    }
  } as any);
  return runtime;
}

export function runtimeWithExecutingAddin(workbookId: WorkbookId): RuntimeService {
  const runtime = new RuntimeService({ persistState: false });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (method: string, params: any) => {
      if (method === "workbook.snapshot_ranges") {
        return snapshotResponse(workbookId, params.ranges);
      }
      if (method === "operation.execute_batch") {
        return operationOk();
      }
      throw new Error(`Unexpected method ${method}`);
    }
  } as any);
  return runtime;
}

export function snapshotResponse(workbookId: WorkbookId, ranges: any[]) {
  return {
    workbookFingerprint: {
      workbookId,
      workbookHash: "workbook_transaction_hash",
      structureHash: "structure",
      capturedAt: new Date().toISOString()
    },
    rangeSnapshots: ranges.map((range: any) => ({
      range,
      values: [["snapshot"]],
      fingerprint: {
        range,
        hash: "range_transaction_hash",
        cellCount: 1,
        capturedAt: new Date().toISOString()
      }
    }))
  };
}

export function operationOk() {
  return {
    ok: true,
    rollbackAvailable: false,
    backups: [],
    warnings: [],
    telemetry: { warningCount: 0 }
  };
}

export function sleepForTest(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function runtimeWithFormulaGraph(workbookId: WorkbookId): RuntimeService {
  const runtime = new RuntimeService({ persistState: false });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (method: string) => {
      if (method === "formula.read_patterns") {
        return {
          workbookId,
          sheetName: "Report",
          address: "B2:B2",
          capturedAt: new Date().toISOString(),
          rowCount: 1,
          columnCount: 1,
          formulaCount: 1,
          formulas: [["=SUM(Transactions[Amount])"]],
          patternMatrix: [["hash"]],
          patterns: [],
          cells: [{ rowIndex: 0, columnIndex: 0, formula: "=SUM(Transactions[Amount])", patternHash: "hash" }],
          warnings: []
        };
      }
      if (method === "table.list") {
        return {
          ok: true,
          tables: [
            {
              workbookId,
              tableName: "Transactions",
              sheetName: "Transactions",
              address: "A1:D11",
              rowCount: 11,
              columnCount: 4,
              showHeaders: true,
              showTotals: true,
              columns: [
                { index: 0, name: "Date" },
                { index: 1, name: "Status" },
                { index: 2, name: "Amount" },
                { index: 3, name: "Memo" }
              ]
            }
          ]
        };
      }
      throw new Error(`Unexpected method ${method}`);
    }
  } as any);
  return runtime;
}

export function writeValuesOperation(workbookId: WorkbookId): ExcelOperation {
  return {
    kind: "range.write_values",
    operationId: "op_refresh" as OperationId,
    workbookId,
    destructiveLevel: "values",
    reason: "Write values",
    target: { workbookId, sheetName: "Sheet1", address: "A1" },
    values: [["ok"]],
    preserveFormats: true
  };
}

export function writeStyleOperation(workbookId: WorkbookId, address: string): ExcelOperation {
  return {
    kind: "range.write_styles",
    operationId: `op_style_${address}` as OperationId,
    workbookId,
    destructiveLevel: "format",
    reason: "Write style",
    target: { workbookId, sheetName: "Sheet1", address },
    style: { fillColor: "#E8EEF7" },
    preserveValues: true
  };
}

export function writeFormulaOperation(workbookId: WorkbookId): ExcelOperation {
  return {
    kind: "range.write_formulas",
    operationId: "op_formula" as OperationId,
    workbookId,
    destructiveLevel: "values",
    reason: "Write dependent formula",
    target: { workbookId, sheetName: "Sheet1", address: "B1" },
    formulas: [["=SUM(A1:A10)"]],
    preserveFormats: true
  };
}
