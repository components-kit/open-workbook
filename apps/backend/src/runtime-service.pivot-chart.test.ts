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

describe("RuntimeService chart template copy", () => {
  it("records a backup and transaction for deterministic chart copy", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_chart_copy" as WorkbookId;
    runtime.allowDestructiveActions(true);
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.get_map") {
          return {
            sheets: [
              { name: "Template", usedRange: { address: "A1:D10" } },
              { name: "Report", usedRange: { address: "A1:D10" } }
            ]
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "chart_copy_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["snapshot"]],
              fingerprint: {
                range,
                hash: "chart_copy_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "chart.copy_from_template") {
          return {
            ok: true,
            copied: ["chartType", "style", "title", "position"],
            source: { workbookId, sheetName: params.templateSheetName, chartName: params.templateChartName },
            target: { workbookId, sheetName: params.sheetName, chartName: params.chartName }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.copyChartFromTemplate({
      workbookId,
      sheetName: "Report",
      chartName: "Revenue",
      templateSheetName: "Template",
      templateChartName: "TemplateRevenue"
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
  });
});

describe("RuntimeService PivotTable template copy", () => {
  it("records a backup and transaction for deterministic PivotTable copy", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_copy" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: params.pivotTableName === "TemplatePivot" ? "Template" : "Report",
              source: "Transactions",
              sourceType: "Table"
            }
          };
        }
        if (method === "workbook.get_map") {
          return {
            sheets: [
              { name: "Template", usedRange: { address: "A1:H20" } },
              { name: "Report", usedRange: { address: "A1:H20" } }
            ]
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "pivot_copy_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["snapshot"]],
              fingerprint: {
                range,
                hash: "pivot_copy_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "pivot.copy_from_template") {
          return {
            ok: true,
            copied: ["layout", "rowHierarchyPositions", "dataHierarchySettings"],
            source: { workbookId, pivotTableName: params.templatePivotTableName },
            target: { workbookId, pivotTableName: params.pivotTableName }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.copyPivotFromTemplate({
      workbookId,
      pivotTableName: "ReportPivot",
      templatePivotTableName: "TemplatePivot"
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect((result as { capabilityStatus?: { capabilities?: Array<{ capability: string; status: string }> } }).capabilityStatus?.capabilities).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ capability: "template_copy", status: "partial" }),
        expect.objectContaining({ capability: "source_reassignment", status: "unsupported" })
      ])
    );
    expect((result as { warnings?: Array<{ code: string }> }).warnings).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: "PIVOT_TEMPLATE_COPY_PARTIAL" })])
    );
    expect((result as { result?: { capabilityStatus?: unknown; warnings?: Array<{ code: string }> } }).result?.capabilityStatus).toBeDefined();
    expect((result as { result?: { warnings?: Array<{ code: string }> } }).result?.warnings).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: "PIVOT_TEMPLATE_COPY_PARTIAL" })])
    );
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.warnings.some((warning) => warning.code === "PIVOT_TEMPLATE_COPY_PARTIAL"))).toBe(true);
  });

  it("blocks PivotTable template copy when target source fields are incompatible", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_copy_incompatible" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: string[] = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push(method);
        if (method === "pivot.get_info") {
          if (params.pivotTableName === "TemplatePivot") {
            return {
              ok: true,
              info: {
                workbookId,
                pivotTableName: "TemplatePivot",
                sheetName: "Template",
                source: "TemplateTransactions",
                sourceType: "Table",
                hierarchies: [{ name: "Region" }, { name: "Amount" }],
                rowHierarchies: [{ name: "Region" }],
                columnHierarchies: [],
                filterHierarchies: [],
                dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" } }]
              }
            };
          }
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "ReportPivot",
              sheetName: "Report",
              source: "ReportTransactions",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }],
              rowHierarchies: [],
              columnHierarchies: [],
              filterHierarchies: [],
              dataHierarchies: []
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.copyPivotFromTemplate({
      workbookId,
      pivotTableName: "ReportPivot",
      templatePivotTableName: "TemplatePivot"
    });

    expect((result as { ok?: boolean }).ok).toBe(false);
    expect((result as { error?: { code?: string } }).error?.code).toBe("TEMPLATE_MISMATCH");
    expect((result as { issues?: Array<{ code: string; details?: Record<string, unknown> }> }).issues?.some((issue) => issue.code === "PIVOT_TEMPLATE_SOURCE_FIELD_MISSING" && issue.details?.field === "Amount")).toBe(true);
    expect(calls).toEqual(["pivot.get_info", "pivot.get_info"]);
    expect(runtime.transactions.list(workbookId)).toHaveLength(0);
  });
});

describe("RuntimeService PivotTable validation", () => {
  it("records a scoped backup and transaction before deleting a PivotTable", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_delete" as WorkbookId;
    runtime.allowDestructiveActions(true);
    const session = runtime.sessions.createSession();
    const calls: Array<{ method: string; params: any }> = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push({ method, params });
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: "Report",
              range: { address: "Report!B4:F20", rowCount: 17, columnCount: 5 }
            }
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "pivot_delete_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["pivot"]],
              fingerprint: {
                range,
                hash: "pivot_delete_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "pivot.delete") {
          return { ok: true, deleted: true };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.deletePivotTable({
      workbookId,
      pivotTableName: "ReportPivot"
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect(calls.some((call) => call.method === "workbook.snapshot_ranges" && call.params.ranges[0].address === "B4:F20")).toBe(true);
    expect(calls.some((call) => call.method === "pivot.delete")).toBe(true);
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
  });

  it("reports useful PivotTable metadata issues", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "EmptyPivot",
              sheetName: "Report",
              range: { address: "Report!A3:C10" },
              sourceType: "Table",
              dataHierarchies: []
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "EmptyPivot"
    });

    expect(result.ok).toBe(true);
    expect(result.summary.hasOutputRange).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_SOURCE_UNAVAILABLE")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_HAS_NO_DATA_FIELDS")).toBe(true);
  });

  it("validates expected PivotTable source and layout fields", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate_fields" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "SalesPivot",
              sheetName: "Report",
              range: { address: "Report!A3:E20" },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }, { name: "Month" }, { name: "Amount" }],
              rowHierarchies: [{ name: "Region" }],
              columnHierarchies: [{ name: "Month" }],
              filterHierarchies: [],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" }, summarizeBy: "sum", numberFormat: "$#,##0" }],
              layout: { showRowGrandTotals: true }
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      expectedFields: ["Region", "Month", "Amount"],
      expectedRowFields: ["Region"],
      expectedColumnFields: ["Month"],
      expectedDataFields: ["Amount"],
      expectedDataFieldSettings: [{ sourceFieldName: "Amount", summarizeBy: "sum", numberFormat: "$#,##0" }],
      expectedLayout: { showRowGrandTotals: true }
    });

    expect(result.ok).toBe(true);
    expect(result.summary.sourceFieldCount).toBe(3);
    expect(result.summary.rowFields).toEqual(["Region"]);
    expect(result.summary.dataFields).toEqual(["Amount"]);
    expect(result.issues).toHaveLength(0);
  });

  it("reports missing and misplaced expected PivotTable fields", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate_bad_fields" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "SalesPivot",
              sheetName: "Report",
              range: { address: "Report!A3:E20" },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }, { name: "Amount" }],
              rowHierarchies: [{ name: "Region" }],
              columnHierarchies: [],
              filterHierarchies: [],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" } }]
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      expectedFields: ["Region", "Customer"],
      expectedColumnFields: ["Region"],
      expectedDataFields: ["Customer"]
    });

    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_FIELD_MISSING" && issue.details?.field === "Customer")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_LAYOUT_MISMATCH" && issue.details?.axis === "column")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_LAYOUT_MISMATCH" && issue.details?.axis === "data")).toBe(true);
  });

  it("reports PivotTable aggregation, number format, and layout mismatches", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate_settings" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "SalesPivot",
              sheetName: "Report",
              range: { address: "Report!A3:E20" },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Amount" }],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" }, summarizeBy: "sum", numberFormat: "$#,##0" }],
              layout: { showRowGrandTotals: true }
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      expectedDataFieldSettings: [{ sourceFieldName: "Amount", summarizeBy: "average", numberFormat: "0.00%" }],
      expectedLayout: { showRowGrandTotals: false }
    });

    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_AGGREGATION_MISMATCH")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_NUMBER_FORMAT_MISMATCH")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_LAYOUT_SETTING_MISMATCH")).toBe(true);
  });

  it("marks missing PivotTables as validation errors", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_missing" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return { ok: false };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "MissingPivot"
    });

    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.code === "PIVOT_NOT_FOUND" && issue.severity === "error")).toBe(true);
  });

  it("captures and diffs deterministic PivotTable fingerprints", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_fingerprint" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "pivot.get_info") {
          const isTemplate = params.pivotTableName === "TemplatePivot";
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: "Report",
              range: { address: "Report!A3:E20", rowCount: 18, columnCount: 5 },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }, { name: "Month" }, { name: "Amount" }],
              rowHierarchies: [{ name: "Region" }],
              columnHierarchies: isTemplate ? [{ name: "Month" }] : [],
              filterHierarchies: [],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" }, summarizeBy: "sum", numberFormat: "$#,##0" }],
              layout: { showRowGrandTotals: true }
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const fingerprint = await runtime.getPivotFingerprint({ workbookId, pivotTableName: "TemplatePivot" });
    const diff = await runtime.diffPivotTables({
      workbookId,
      pivotTableName: "TemplatePivot",
      targetPivotTableName: "ReportPivot"
    });

    expect((fingerprint as { ok?: boolean }).ok).toBe(true);
    expect((fingerprint as { fingerprint?: { hash?: string } }).fingerprint?.hash).toMatch(/^[a-f0-9]{16}$/);
    expect(diff.ok).toBe(false);
    expect(diff.changes.some((change) => change.path === "layout.columnFields")).toBe(true);
  });

  it("reports PivotTable source reassignment as an explicit capability limit", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_update_source" as WorkbookId;

    const result = runtime.updatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      sourceSheetName: "Data",
      sourceAddress: "A1:D100",
      destinationSheetName: "Report",
      destinationAddress: "B4"
    });

    expect(result.ok).toBe(false);
    expect(result.error.code).toBe("CAPABILITY_UNAVAILABLE");
    expect(result.capabilityStatus.fallback).toBe("excel.pivot.rebuild_with_source");
    expect(result.capabilityStatus.capabilities).toEqual(
      expect.arrayContaining([expect.objectContaining({ capability: "source_reassignment", status: "unsupported" })])
    );
    expect(result.warnings).toEqual(expect.arrayContaining([expect.objectContaining({ code: "PIVOT_SOURCE_REASSIGNMENT_UNSUPPORTED" })]));
  });

  it("rebuilds an existing PivotTable through explicit delete and create steps", async () => {
    const runtime = new RuntimeService({ persistState: false });
    runtime.allowDestructiveActions(true);
    const workbookId = "workbook_pivot_rebuild_replace" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: string[] = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push(method);
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: "Report",
              range: { address: "Report!B4:F20", rowCount: 17, columnCount: 5 }
            }
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "pivot_rebuild_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["pivot"]],
              fingerprint: { range, hash: "pivot_rebuild_range", cellCount: 1, capturedAt: new Date().toISOString() }
            }))
          };
        }
        if (method === "pivot.delete") {
          return { ok: true, deleted: true };
        }
        if (method === "pivot.create") {
          return { ok: true, info: { workbookId, pivotTableName: params.pivotTableName, sheetName: params.destinationSheetName } };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.rebuildPivotWithSource({
      workbookId,
      pivotTableName: "SalesPivot",
      sourceSheetName: "Data",
      sourceAddress: "A1:D100",
      destinationSheetName: "Report",
      destinationAddress: "B4",
      rowFields: ["Region"],
      dataFields: [{ sourceFieldName: "Amount", summarizeBy: "sum" }],
      replaceExisting: true
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect(calls).toContain("pivot.delete");
    expect(calls).toContain("pivot.create");
    expect(runtime.transactions.list(workbookId).filter((transaction) => transaction.status === "applied")).toHaveLength(2);
  });
});
