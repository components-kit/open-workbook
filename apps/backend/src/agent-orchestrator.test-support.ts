import type { BatchRequest, WorkbookId } from "@components-kit/open-workbook-protocol";
import { tryParseA1Address } from "@components-kit/open-workbook-excel-core";
import { createMetadataFingerprint, type WorkbookMetadata } from "./workbook-metadata-cache.js";

export const workbookId = "workbook_agent_unit" as WorkbookId;
export const activeWorkbook = { workbookId, name: "Agent Unit.xlsx", path: "/tmp/Agent Unit.xlsx" };
export const sheets = [
  { workbookId, worksheetId: "sheet_Data", name: "Data", usedRange: { address: "A1:D4", rowCount: 4, columnCount: 4 }, tables: [{ name: "Transactions" }] },
  { workbookId, worksheetId: "sheet_Report", name: "Report", usedRange: { address: "A1:B20", rowCount: 20, columnCount: 2 }, tables: [] },
  { workbookId, worksheetId: "sheet_CustomerMaster", name: "Customer Master", usedRange: { address: "A1:B4", rowCount: 4, columnCount: 2 }, tables: [] },
  { workbookId, worksheetId: "sheet_Operations", name: "Operations", usedRange: { address: "A1:J24", rowCount: 24, columnCount: 10 }, tables: [] },
  { workbookId, worksheetId: "sheet_Mar", name: "Mar 2026", usedRange: { address: "A1:AJ206", rowCount: 206, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_Apr", name: "Apr 2026", usedRange: { address: "A1:AJ244", rowCount: 244, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_May", name: "May 2026", usedRange: { address: "A1:AJ244", rowCount: 244, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_FinJun", name: "Financials - June 2026", usedRange: { address: "A1:C4", rowCount: 4, columnCount: 3 }, tables: [] },
  { workbookId, worksheetId: "sheet_FinMay", name: "Financials - May 2026", usedRange: { address: "A1:C4", rowCount: 4, columnCount: 3 }, tables: [] }
];

export class FakeAgentRuntime {
  readBatchCount = 0;
  writeBatchCount = 0;
  appendTableRowCount = 0;
  runtimeMethodCalls: Record<string, number> = {};
  tableMethodCalls: Array<{ method: string; request: any }> = [];
  returnDataOnly = false;
  omitOkOnWrite = false;
  batchResultOverride: unknown | undefined;
  snapshotRangesOverride: unknown | undefined;
  validationResult: any;
  lastBatchOperations: BatchRequest["operations"] = [];
  lastBatchRequest: BatchRequest | undefined;
  lastSnapshotRanges: Array<{ workbookId?: WorkbookId; sheetName: string; address: string }> = [];
  snapshotRangesHistory: Array<Array<{ workbookId?: WorkbookId; sheetName: string; address: string }>> = [];
  lastWriteOperations: Array<Extract<BatchRequest["operations"][number], { target: any }>> = [];
  selection: any;
  readiness: any;
  agentExecutionContext: any;
  collaborationStatus: any;
  failStyleCopyOnCall: number | undefined;
  workbookContentVersion = 0;

  sessions = { getActive: () => ({ activeWorkbook }) };

  currentAgentExecutionContext() {
    return this.agentExecutionContext;
  }

  runWithAgentExecutionContext<T>(context: any, work: () => T): T {
    const previous = this.agentExecutionContext;
    this.agentExecutionContext = context;
    try {
      return work();
    } finally {
      this.agentExecutionContext = previous;
    }
  }

  private recordRuntimeCall(method: string) {
    this.runtimeMethodCalls[method] = (this.runtimeMethodCalls[method] ?? 0) + 1;
  }

  getStatus() {
    this.recordRuntimeCall("runtime.get_status");
    return { activeAddinConnected: true, connectionState: "ready", activeWorkbookAvailable: true, activeWorkbook };
  }

  getCollaborationStatus(requestWorkbookId?: WorkbookId) {
    this.recordRuntimeCall("runtime.get_collaboration_status");
    if (this.collaborationStatus) return this.collaborationStatus;
    const agentId = this.agentExecutionContext?.agentId;
    return {
      ok: true,
      workbookId: requestWorkbookId ?? workbookId,
      agents: agentId ? [{
        agentId,
        agentName: this.agentExecutionContext?.agentName,
        clientType: this.agentExecutionContext?.clientType ?? "mcp",
        status: "active",
        lastSeenAt: "2026-06-18T00:00:00.000Z"
      }] : [],
      tasks: [],
      locks: [],
      transactions: [],
      conflicts: [],
      events: []
    };
  }

  async getConnectionReadiness() {
    this.recordRuntimeCall("runtime.get_connection_readiness");
    if (this.readiness) return this.readiness;
    return { ok: true, connectionState: "ready", activeWorkbook, status: this.getStatus() };
  }

  async getActiveContext() {
    this.recordRuntimeCall("runtime.get_active_context");
    return { ok: true, activeWorkbook };
  }

  getWorkbookContentVersion() {
    return this.workbookContentVersion;
  }

  getWorkbookChangeJournal() {
    return { ok: true, currentVersion: this.workbookContentVersion, overlapStatus: "changed", entries: [] };
  }

  async getSelection() {
    this.recordRuntimeCall("runtime.get_selection");
    return this.selection ? { workbook: activeWorkbook, selection: this.selection } : { ok: false };
  }

  async getWorkbookMap() {
    this.recordRuntimeCall("workbook.get_workbook_map");
    return { ok: true, map: { workbook: activeWorkbook, activeSheet: "Data", sheets } };
  }

  listOpenWorkbooks() {
    this.recordRuntimeCall("workbook.list_open_workbooks");
    return { ok: true, workbooks: [activeWorkbook] };
  }

  async getWorkbookInfo() {
    this.recordRuntimeCall("workbook.get_workbook_info");
    return { ok: true, info: { workbook: activeWorkbook, sheetCount: sheets.length } };
  }

  async closeWorkbook(requestWorkbookId: any, closeBehavior?: any) {
    this.recordRuntimeCall("workbook.close");
    return { ok: true, workbookId: requestWorkbookId, closeBehavior };
  }

  async listTables() {
    this.recordRuntimeCall("table.list");
    return { ok: true, tables: [{ name: "Transactions" }] };
  }

  async getTableInfo() {
    this.recordRuntimeCall("table.get_info");
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

  async readTable(request: any) {
    this.recordRuntimeCall("table.read");
    const rows = [
      ["2026-06-01", "A-100", 123, "Open"],
      ["2026-06-02", "A-101", 456, "Closed"],
      ["2026-06-03", "A-102", 789, "Open"]
    ];
    const headers = ["Date", "Account", "Amount", "Status"];
    const columnIndexes = Array.isArray(request.columns) && request.columns.length > 0
      ? request.columns.map((column: string | number) => typeof column === "number" ? column : headers.indexOf(column)).filter((index: number) => index >= 0)
      : headers.map((_header, index) => index);
    const offset = request.rowOffset ?? 0;
    const limit = request.rowLimit ?? rows.length;
    const values = rows.slice(offset, offset + limit).map((row) => columnIndexes.map((index: number) => row[index]));
    return {
      ok: true,
      table: {
        tableName: request.tableName,
        headers: columnIndexes.map((index: number) => headers[index]),
        values
      }
    };
  }

  async readRangeMetadata(method: string, request: any) {
    this.recordRuntimeCall(method);
    if (method === "range.read_data_validation") {
      return {
        ok: true,
        method,
        request,
        data: {
          address: request.address,
          rules: [{
            address: request.address,
            type: "list",
            source: ["Open", "Closed", "Pending"],
            allowBlank: true
          }]
        }
      };
    }
    return { ok: true, method, request, data: { address: request.address, count: method === "range.search" ? 1 : 0 } };
  }

  async listNames() {
    this.recordRuntimeCall("names.list");
    return { ok: true, names: [{ workbookId, name: "RevenueTotal", scope: "workbook", sheetName: "Report", address: "Report!B2" }] };
  }

  listRegions() {
    this.recordRuntimeCall("region.list");
    return { ok: true, regions: [{ name: "InputRegion", sheetName: "Report", address: "B1:B3" }] };
  }

  async getName(request: any) {
    this.recordRuntimeCall("names.get");
    return {
      ok: true,
      name: {
        workbookId,
        name: request.name,
        scope: request.sheetName ? "worksheet" : "workbook",
        ...(request.sheetName !== undefined ? { sheetName: request.sheetName } : { sheetName: "Report" }),
        address: request.name === "InputRegion" ? "Report!B1:B3" : "Report!B2",
        formula: request.name === "InputRegion" ? "=Report!$B$1:$B$3" : "=Report!$B$2"
      }
    };
  }

  async createName(request: any) {
    this.recordRuntimeCall("names.create");
    return { ok: true, result: { name: request.name, reference: request.reference, formula: request.formula } };
  }

  async updateName(request: any) {
    this.recordRuntimeCall("names.update");
    return { ok: true, result: { name: request.name, reference: request.reference, formula: request.formula } };
  }

  async deleteName(request: any) {
    this.recordRuntimeCall("names.delete");
    return { ok: true, result: { name: request.name } };
  }

  async getRegion(request: any) {
    this.recordRuntimeCall("region.get");
    return {
      ok: true,
      region: {
        workbookId,
        regionId: `region_${request.regionName}`,
        name: request.regionName,
        sheetName: "Report",
        address: request.regionName === "InputRegion" ? "B1:B3" : "B2",
        kind: "data",
        source: "manual",
        createdAt: "2026-06-18T00:00:00.000Z",
        updatedAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  async registerRegion(request: any) {
    this.recordRuntimeCall("region.register");
    return {
      ok: true,
      region: {
        workbookId,
        regionId: `region_${request.name}`,
        name: request.name,
        sheetName: request.sheetName,
        address: request.address,
        kind: request.kind ?? "data",
        source: request.createNamedRange ? "named-range" : "manual",
        createdAt: "2026-06-18T00:00:00.000Z",
        updatedAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  async clearRegionValues(request: any) {
    this.recordRuntimeCall("region.clear_values");
    return { ok: true, backups: [], warnings: [], telemetry: { regionName: request.regionName } };
  }

  async writeRegionValues(request: any) {
    this.recordRuntimeCall("region.write_values");
    return { ok: true, backups: [], warnings: [], telemetry: { rowsWritten: request.values?.length ?? 0 } };
  }

  async fillRegion(request: any) {
    this.recordRuntimeCall("region.fill");
    return { ok: true, backups: [], warnings: [], telemetry: { rowsWritten: request.values?.length ?? 0, clearFirst: request.clearFirst ?? false } };
  }

  async applyBatch(request: BatchRequest) {
    const effectiveRequest: BatchRequest = request.agentId === undefined && this.agentExecutionContext?.agentId
      ? { ...request, agentId: this.agentExecutionContext.agentId }
      : request;
    this.lastBatchRequest = effectiveRequest;
    this.lastBatchOperations = effectiveRequest.operations;
    if (effectiveRequest.operations.every((operation) => operation.kind === "range.read_full")) {
      this.readBatchCount += 1;
      const readData = effectiveRequest.operations.map((operation) => {
        const sheetName = operation.target.sheetName;
        const address = operation.target.address;
        return {
          operationId: operation.operationId,
          snapshot: {
            values: valuesFor(sheetName, address),
            formulas: formulasFor(sheetName, address),
            text: valuesFor(sheetName, address).map((row) => row.map((value) => value === null || value === undefined ? "" : String(value))),
            numberFormat: numberFormatsFor(sheetName, address),
            style: { fillColor: "#FFFFFF", fontName: "Calibri", fontSize: 11 }
          }
        };
      });
      return {
        ok: true,
        ...(this.returnDataOnly ? { data: readData } : { readData }),
        telemetry: { cellsRead: effectiveRequest.operations.reduce((total, operation) => total + valuesFor(operation.target.sheetName, operation.target.address).flat().length, 0) }
      };
    }
    this.writeBatchCount += 1;
    this.lastWriteOperations = effectiveRequest.operations.filter((operation): operation is Extract<BatchRequest["operations"][number], { target: any }> => "target" in operation);
    if (this.batchResultOverride !== undefined) {
      return this.batchResultOverride;
    }
    return {
      ...(this.omitOkOnWrite ? {} : { ok: true }),
      backups: [],
      warnings: [],
      telemetry: { cellsWritten: 1 }
    };
  }

  async snapshotRanges(requestWorkbookId: WorkbookId, ranges: Array<{ workbookId?: WorkbookId; sheetName: string; address: string }>) {
    this.recordRuntimeCall("workbook.snapshot_ranges");
    this.lastSnapshotRanges = ranges;
    this.snapshotRangesHistory.push(ranges);
    if (this.snapshotRangesOverride !== undefined) {
      return this.snapshotRangesOverride;
    }
    return {
      ok: true,
      workbookId: requestWorkbookId,
      rangeSnapshots: ranges.map((range) => {
        const values = valuesFor(range.sheetName, range.address);
        return {
          workbookId: range.workbookId ?? requestWorkbookId,
          sheetName: range.sheetName,
          address: range.address,
          rowCount: values.length,
          columnCount: values.reduce((max, row) => Math.max(max, row.length), 0),
          values,
          formulas: formulasFor(range.sheetName, range.address),
          text: values.map((row) => row.map((value) => value === null || value === undefined ? "" : String(value))),
          numberFormat: numberFormatsFor(range.sheetName, range.address)
        };
      })
    };
  }

  async appendTableRows(request: any) {
    this.appendTableRowCount += 1;
    this.tableMethodCalls.push({ method: "table.append_rows", request });
    return { ok: true, backups: [], warnings: [], telemetry: { rowsWritten: 1 } };
  }

  async updateTableRows(request: any) {
    this.tableMethodCalls.push({ method: "table.update_rows", request });
    return { ok: true, backups: [], warnings: [], telemetry: { rowsWritten: request.rows?.length ?? 0 } };
  }

  async createTable(request: any) {
    this.tableMethodCalls.push({ method: "table.create", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async resizeTable(request: any) {
    this.tableMethodCalls.push({ method: "table.resize", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async reorderTableColumns(request: any) {
    this.tableMethodCalls.push({ method: "table.reorder_columns", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async clearTableDataKeepFormulas(request: any) {
    this.tableMethodCalls.push({ method: "table.clear_data_keep_formulas", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async clearTableFilters(request: any) {
    this.tableMethodCalls.push({ method: "table.clear_filters", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async applyTableFilters(request: any) {
    this.tableMethodCalls.push({ method: "table.apply_filters", request });
    if (!Array.isArray(request.filters) || request.filters.some((filter: any) => !filter?.criteria)) {
      return { ok: false, error: { code: "INVALID_ARGUMENT", message: "Filter criteria are required." }, warnings: [], telemetry: {} };
    }
    return { ok: true, backups: [], warnings: [], telemetry: { filtersApplied: request.filters?.length ?? 0 } };
  }

  async sortTable(request: any) {
    this.tableMethodCalls.push({ method: "table.sort", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async applyTableView(request: any) {
    this.tableMethodCalls.push({ method: "table.apply_view", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async setTableTotalRow(request: any) {
    this.tableMethodCalls.push({ method: "table.set_total_row", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async setTableStyle(request: any) {
    this.tableMethodCalls.push({ method: "table.set_style", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async copyTableStructure(request: any) {
    this.tableMethodCalls.push({ method: "table.copy_structure", request });
    return { ok: true, backups: [], warnings: [], telemetry: {} };
  }

  async validateWorkbook() {
    this.recordRuntimeCall("validate.workbook");
    if (this.validationResult) return this.validationResult;
    return { ok: true, issues: [] };
  }

  async validateSheet(request: any) {
    this.recordRuntimeCall("validate.sheet");
    return validationReport("sheet", request);
  }

  async validateTemplateConsistency(request: any) {
    this.recordRuntimeCall("validate.template_consistency");
    return validationReport("template_consistency", request);
  }

  async validateFormulas(request: any) {
    this.recordRuntimeCall("validate.formulas");
    return validationReport("formulas", request);
  }

  getTemplate(requestTemplateId: any) {
    this.recordRuntimeCall("template.get");
    return {
      ok: true,
      template: {
        templateId: requestTemplateId,
        name: "Unit Template",
        scope: "workbook",
        version: 1,
        sourceSheetName: "Report",
        dataRegions: ["B2:B3"],
        fingerprint: {},
        fingerprintPayload: {},
        createdAt: "2026-06-18T00:00:00.000Z",
        updatedAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  listTemplates() {
    this.recordRuntimeCall("template.list");
    return [{
      templateId: "template_unit",
      name: "Unit Template",
      scope: "workbook",
      workbookId,
      sourceSheetName: "Report",
      dataRegions: ["B2:B3"],
      fingerprintPayload: {},
      createdAt: "2026-06-18T00:00:00.000Z",
      updatedAt: "2026-06-18T00:00:00.000Z"
    }];
  }

  async detectTemplates() {
    this.recordRuntimeCall("template.detect");
    return { ok: true, candidates: [{ workbookId, sheetName: "Report", usedRange: { address: "A1:B20" }, score: 1, reason: "Sheet has a used range." }] };
  }

  inferTemplateRegions(requestTemplateId: any) {
    this.recordRuntimeCall("template.infer_regions");
    return { ok: true, templateId: requestTemplateId, dataRegions: ["B2:B3"], inferredRegions: [{ address: "B2:B3", kind: "data-entry" }] };
  }

  async validateSheetAgainstTemplate(request: any) {
    this.recordRuntimeCall("template.validate_sheet");
    return { ok: true, sheetName: request.targetSheetName, templateId: request.templateId, issueCount: 0, issues: [], fingerprintPayload: {} };
  }

  async registerTemplate(request: any) {
    this.recordRuntimeCall("template.register");
    return { templateId: "template_registered", ...request, fingerprintPayload: {}, createdAt: "2026-06-18T00:00:00.000Z", updatedAt: "2026-06-18T00:00:00.000Z" };
  }

  unregisterTemplate(requestTemplateId: any) {
    this.recordRuntimeCall("template.unregister");
    return { ok: true, templateId: requestTemplateId };
  }

  async repairSheetFromTemplate(request: any) {
    this.recordRuntimeCall("template.repair_sheet");
    return { ok: true, result: { repaired: request.repair ?? ["styles", "formulas", "dataRegions"] }, validation: { ok: true, issueCount: 0, issues: [] } };
  }

  async compareFormulaPatterns(request: any) {
    this.recordRuntimeCall("formula.validate_against_template");
    return {
      ok: true,
      issueCount: 0,
      issues: [],
      sourcePatterns: { workbookId: request.workbookId, sheetName: request.sourceSheetName, address: request.sourceAddress ?? "A1:B20", formulas: [], patternMatrix: [], patterns: [], cells: [], warnings: [] },
      targetPatterns: { workbookId: request.workbookId, sheetName: request.targetSheetName, address: request.targetAddress ?? "A1:B20", formulas: [], patternMatrix: [], patterns: [], cells: [], warnings: [] }
    };
  }

  async validateStyles(request: any) {
    this.recordRuntimeCall("validate.styles");
    return validationReport("styles", request);
  }

  async getStyleFingerprint(request: any) {
    this.recordRuntimeCall("style.get_fingerprint");
    return { ok: true, fingerprint: { workbookId: request.workbookId, sheetName: request.sheetName, address: request.address ?? "A1:B20", dimensions: { fills: { hash: "fills" }, fonts: { hash: "fonts" } }, warnings: [] } };
  }

  async compareStyleFingerprints(request: any) {
    this.recordRuntimeCall("style.compare_fingerprint");
    return {
      ok: true,
      issueCount: 0,
      issues: [],
      sourceFingerprint: { workbookId: request.workbookId, sheetName: request.sourceSheetName, address: request.sourceAddress ?? "A1:B20", dimensions: {}, warnings: [] },
      targetFingerprint: { workbookId: request.workbookId, sheetName: request.targetSheetName, address: request.targetAddress ?? "A1:B20", dimensions: {}, warnings: [] }
    };
  }

  getTheme(requestWorkbookId: any) {
    this.recordRuntimeCall("style.get_theme");
    return { ok: false, workbookId: requestWorkbookId, operation: "get_theme", capabilityStatus: { capability: "excel.style.get_theme", status: "unsupported" }, warnings: [{ code: "THEME_READ_UNAVAILABLE", message: "Theme read unavailable." }] };
  }

  applyTheme(request: any) {
    this.recordRuntimeCall("style.apply_theme");
    return { ok: false, workbookId: request.workbookId, operation: "apply_theme", capabilityStatus: { capability: "excel.style.apply_theme", status: "unsupported" }, warnings: [{ code: "THEME_APPLY_UNAVAILABLE", message: "Theme apply unavailable." }] };
  }

  async copyStyleDimensions(request: any) {
    this.recordRuntimeCall("style.copy_dimensions");
    if (this.failStyleCopyOnCall === this.runtimeMethodCalls["style.copy_dimensions"]) {
      return { ok: false, backups: [], rollbackAvailable: false, warnings: ["Style copy failed"], error: { code: "STYLE_COPY_FAILED", message: "Style copy failed" }, telemetry: { styleCopyCount: 0 } };
    }
    return { ok: true, backups: [`backup_style_${this.runtimeMethodCalls["style.copy_dimensions"]}`], rollbackAvailable: true, result: { copied: request.dimensions }, validation: { ok: true, issueCount: 0, issues: [] }, telemetry: { styleCopyCount: 1 } };
  }

  async copyStyleDimensionsMany(request: any) {
    this.recordRuntimeCall("style.copy_dimensions_many");
    const copyCount = request.requests?.length ?? 0;
    if (this.failStyleCopyOnCall !== undefined) {
      const results = request.requests.map((_entry: unknown, index: number) => ({ ok: index + 1 !== this.failStyleCopyOnCall, warnings: index + 1 === this.failStyleCopyOnCall ? ["Style copy failed"] : [] }));
      return {
        ok: false,
        backups: ["backup_style_many"],
        rollbackAvailable: true,
        warnings: ["Style copy failed"],
        error: { code: "STYLE_COPY_FAILED", message: "Style copy failed" },
        result: {
          ok: false,
          copyCount,
          results
        },
        results,
        telemetry: { styleCopyCount: Math.max(0, Math.min(copyCount, this.failStyleCopyOnCall - 1)) }
      };
    }
    return {
      ok: true,
      backups: ["backup_style_many"],
      rollbackAvailable: true,
      result: { copied: request.requests.flatMap((entry: any) => entry.dimensions ?? []), copyCount },
      validation: { ok: true, issueCount: 0, issues: [] },
      telemetry: { styleCopyCount: copyCount }
    };
  }

  async repairStyleFromTemplate(request: any) {
    this.recordRuntimeCall("style.repair_consistency");
    return { ok: true, workbookId: request.workbookId, repair: "style_from_template", result: { templateId: request.templateId, targetSheetName: request.targetSheetName }, issueCount: 0, issues: [] };
  }

  async cleanDetectHeaderRow(request: any) {
    this.recordRuntimeCall("clean.detect_header_row");
    return cleaningReport("detect_header_row", request, 0, { headerRowIndex: 0, headers: ["Date", "Account", "Amount"] });
  }

  async cleanDetectOutliers(request: any) {
    this.recordRuntimeCall("clean.detect_outliers");
    return cleaningReport("detect_outliers", request, 0, { outliers: [] });
  }

  async cleanFuzzyMatch(request: any) {
    this.recordRuntimeCall("clean.fuzzy_match");
    return cleaningReport("fuzzy_match", request, 0, { matches: [] });
  }

  async cleanNormalizeHeaders(request: any) {
    this.recordRuntimeCall("clean.normalize_headers");
    return cleaningReport("normalize_headers", request, 3);
  }

  async cleanTrimWhitespace(request: any) {
    this.recordRuntimeCall("clean.trim_whitespace");
    return cleaningReport("trim_whitespace", request, 2);
  }

  async cleanRemoveDuplicates(request: any) {
    this.recordRuntimeCall("clean.remove_duplicates");
    return cleaningReport("remove_duplicates", request, 1);
  }

  async cleanParseDates(request: any) {
    this.recordRuntimeCall("clean.parse_dates");
    return cleaningReport("parse_dates", request, 2);
  }

  async cleanParseNumbers(request: any) {
    this.recordRuntimeCall("clean.parse_numbers");
    return cleaningReport("parse_numbers", request, 2);
  }

  async cleanStandardizeCurrency(request: any) {
    this.recordRuntimeCall("clean.standardize_currency");
    return cleaningReport("standardize_currency", request, 2);
  }

  async cleanFillMissingValues(request: any) {
    this.recordRuntimeCall("clean.fill_missing_values");
    return cleaningReport("fill_missing_values", request, 2);
  }

  async cleanSplitColumn(request: any) {
    this.recordRuntimeCall("clean.split_column");
    return cleaningReport("split_column", request, 4);
  }

  async cleanMergeColumns(request: any) {
    this.recordRuntimeCall("clean.merge_columns");
    return cleaningReport("merge_columns", request, 4);
  }

  async validateTables(request: any) {
    this.recordRuntimeCall("validate.tables");
    return validationReport("tables", request);
  }

  async validateTableAgainstTemplate(request: any) {
    this.recordRuntimeCall("table.validate_against_template");
    return { ok: true, table: { tableName: request.tableName }, templateTables: [] };
  }

  async validateFilters(request: any) {
    this.recordRuntimeCall("validate.filters");
    return validationReport("filters", request);
  }

  validatePrintLayout(request: any) {
    this.recordRuntimeCall("validate.print_layout");
    return validationReport("print_layout", request);
  }

  async validateNoBrokenReferences(request: any) {
    this.recordRuntimeCall("validate.no_broken_references");
    return validationReport("no_broken_references", request);
  }

  async validateNoFormulaErrors(request: any) {
    this.recordRuntimeCall("validate.no_formula_errors");
    return validationReport("no_formula_errors", request);
  }

  async validateNoUnintendedChanges(request: any) {
    this.recordRuntimeCall("validate.no_unintended_changes");
    return validationReport("no_unintended_changes", request);
  }

  async readFormulaPatterns(request: any) {
    this.recordRuntimeCall("formula.read_patterns");
    return {
      ok: true,
      patterns: {
        workbookId: request.workbookId,
        sheetName: request.sheetName,
        address: request.address,
        capturedAt: "2026-06-18T00:00:00.000Z",
        rowCount: 1,
        columnCount: 1,
        formulaCount: 1,
        formulas: [["=SUM(Data!C2:C4)"]],
        formulasR1C1: [["=SUM(Data!R2C3:R4C3)"]],
        patternMatrix: [["hash_sum_data"]],
        patterns: [{ patternHash: "hash_sum_data", formulaR1C1: "=SUM(Data!R2C3:R4C3)", count: 1, cells: [{ rowIndex: 0, columnIndex: 0 }] }],
        cells: [{ rowIndex: 0, columnIndex: 0, formula: "=SUM(Data!C2:C4)", formulaR1C1: "=SUM(Data!R2C3:R4C3)", patternHash: "hash_sum_data" }],
        warnings: []
      }
    };
  }

  async getFormulaDependencyGraph(request: any) {
    this.recordRuntimeCall("formula.get_dependency_graph");
    return {
      ok: true,
      graph: {
        workbookId: request.workbookId,
        sheetName: request.sheetName,
        address: request.address,
        capturedAt: "2026-06-18T00:00:00.000Z",
        nodes: [
          { id: "formula:Report!B2", workbookId: request.workbookId, kind: "range", sheetName: request.sheetName, address: request.address, formula: "=SUM(Data!C2:C4)" },
          { id: "range:Data!C2:C4", workbookId: request.workbookId, kind: "range", sheetName: "Data", address: "C2:C4" }
        ],
        edges: [{
          from: { id: "formula:Report!B2", workbookId: request.workbookId, kind: "range", sheetName: request.sheetName, address: request.address, formula: "=SUM(Data!C2:C4)" },
          to: { id: "range:Data!C2:C4", workbookId: request.workbookId, kind: "range", sheetName: "Data", address: "C2:C4" },
          kind: "precedent",
          confidence: "parsed"
        }],
        warnings: []
      }
    };
  }

  async traceFormulaPrecedents(request: any) {
    this.recordRuntimeCall("formula.trace_precedents");
    return {
      ok: true,
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      address: request.address,
      direction: "precedents",
      nodes: [
        { id: "formula:Report!B2", workbookId: request.workbookId, kind: "range", sheetName: request.sheetName, address: request.address },
        { id: "range:Data!C2:C4", workbookId: request.workbookId, kind: "range", sheetName: "Data", address: "C2:C4" }
      ],
      edges: [{
        from: { id: "formula:Report!B2", workbookId: request.workbookId, kind: "range", sheetName: request.sheetName, address: request.address },
        to: { id: "range:Data!C2:C4", workbookId: request.workbookId, kind: "range", sheetName: "Data", address: "C2:C4" },
        kind: "precedent",
        confidence: "parsed"
      }],
      warnings: []
    };
  }

  async traceFormulaDependents(request: any) {
    this.recordRuntimeCall("formula.trace_dependents");
    return {
      ok: true,
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      address: request.address,
      direction: "dependents",
      nodes: [
        { id: "range:Data!C2:C4", workbookId: request.workbookId, kind: "range", sheetName: request.sheetName, address: request.address },
        { id: "formula:Report!B2", workbookId: request.workbookId, kind: "range", sheetName: "Report", address: "B2" }
      ],
      edges: [{
        from: { id: "range:Data!C2:C4", workbookId: request.workbookId, kind: "range", sheetName: request.sheetName, address: request.address },
        to: { id: "formula:Report!B2", workbookId: request.workbookId, kind: "range", sheetName: "Report", address: "B2" },
        kind: "dependent",
        confidence: "parsed"
      }],
      warnings: []
    };
  }

  async copyFormulaPatterns(request: any) {
    this.recordRuntimeCall("formula.copy_patterns");
    return {
      ok: true,
      backup: { backupId: "backup_formula_copy", workbookId: request.workbookId },
      result: { ok: true, formulasChanged: 3, warnings: [] },
      validation: { ok: true, issueCount: 0, issues: [] }
    };
  }

  async fillFormulaPattern(request: any) {
    this.recordRuntimeCall(request.direction === "right" ? "formula.fill_right" : "formula.fill_down");
    return {
      ok: true,
      backup: { backupId: `backup_formula_fill_${request.direction}`, workbookId: request.workbookId },
      result: { ok: true, formulasChanged: 3, warnings: [] }
    };
  }

  async repairFormulasFromTemplate(request: any) {
    this.recordRuntimeCall("formula.repair_patterns");
    return {
      ok: true,
      backups: ["backup_formula_repair"],
      validation: { ok: true, issueCount: 0, issues: [] },
      result: { ok: true, formulasChanged: 3, warnings: [] },
      warnings: []
    };
  }

  repairFiltersFromTemplate(request: any) {
    this.recordRuntimeCall("repair.filters_from_template");
    return repairReport("filters_from_template", request, false);
  }

  repairPrintLayout(request: any) {
    this.recordRuntimeCall("repair.print_layout");
    return repairReport("print_layout", request, false);
  }

  repairNamedRanges(request: any) {
    this.recordRuntimeCall("repair.named_ranges");
    return repairReport("named_ranges", request, false);
  }

  repairFormulaErrors(request: any) {
    this.recordRuntimeCall("repair.formula_errors");
    return repairReport("formula_errors", request, false);
  }

  repairMergedCells(request: any) {
    this.recordRuntimeCall("repair.merged_cells");
    return repairReport("merged_cells", request, false);
  }

  async convertFormulasToValues(request: any) {
    this.recordRuntimeCall("formula.convert_to_values");
    return {
      ok: true,
      backup: { backupId: "backup_formula_convert", workbookId: request.workbookId },
      result: { ok: true, formulasChanged: 1, warnings: [] }
    };
  }

  async createWorkbookSnapshot(request: any) {
    this.recordRuntimeCall("workbook.snapshot");
    return {
      ok: true,
      snapshot: {
        snapshotId: "snapshot_agent_unit",
        workbookId: request.workbookId,
        reason: request.reason,
        affectedRanges: request.ranges ?? [{ workbookId: request.workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: { rangeSnapshots: [] },
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  async createWorkbookBackup(request: any) {
    this.recordRuntimeCall("workbook.create_backup");
    return {
      ok: true,
      backup: {
        backupId: "backup_agent_unit",
        workbookId: request.workbookId,
        reason: request.reason,
        affectedRanges: request.ranges ?? [{ workbookId: request.workbookId, sheetName: "Data", address: "A1:D4" }],
        kind: "workbook-copy",
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  async restoreBackup(requestBackupId: any) {
    this.recordRuntimeCall("workbook.restore_backup");
    return { ok: true, backups: [], warnings: [], telemetry: {}, backupId: requestBackupId };
  }

  async restoreWorkbookBackup(requestBackupId: any) {
    return this.restoreBackup(requestBackupId);
  }

  exportWorkbookLocalConfig(requestWorkbookId: any, options: any = {}) {
    this.recordRuntimeCall("workbook.export_local_config");
    return {
      ok: true,
      workbookId: requestWorkbookId,
      config: { version: 1, workbookId: requestWorkbookId, exportedAt: "2026-06-18T00:00:00.000Z", source: "open-workbook-local-config", templates: [], regions: [] },
      counts: { templates: 0, regions: 0, permissions: options.includePermissions ?? true }
    };
  }

  importWorkbookLocalConfig(request: any) {
    this.recordRuntimeCall("workbook.import_local_config");
    return { ok: true, workbookId: request.workbookId, imported: { templates: 0, regions: 0, permissions: false }, skipped: { templates: 0, regions: 0 } };
  }

  async embedWorkbookLocalConfig(requestWorkbookId: any, options: any = {}) {
    this.recordRuntimeCall("workbook.embed_local_config");
    return { ok: true, workbookId: requestWorkbookId, embedded: true, options };
  }

  async readWorkbookEmbeddedLocalConfig(requestWorkbookId: any) {
    this.recordRuntimeCall("workbook.read_embedded_local_config");
    return { ok: true, workbookId: requestWorkbookId, embedded: true, config: { version: 1, workbookId: requestWorkbookId, exportedAt: "2026-06-18T00:00:00.000Z", source: "open-workbook-local-config", templates: [], regions: [] } };
  }

  async importWorkbookEmbeddedLocalConfig(request: any) {
    this.recordRuntimeCall("workbook.import_embedded_local_config");
    return { ok: true, workbookId: request.workbookId, imported: { templates: 0, regions: 0, permissions: false }, skipped: { templates: 0, regions: 0 } };
  }

  async createFileBackup(request: any) {
    this.recordRuntimeCall("backup.create_file");
    return {
      ok: true,
      backup: { backupId: "backup_file_unit", workbookId: request.workbookId, kind: "file-copy" },
      manifest: { backupId: "backup_file_unit", workbookId: request.workbookId, filePath: request.targetPath ?? "/tmp/Agent Unit backup.xlsx", restoreStatus: "available", pinned: request.pin ?? false }
    };
  }

  async restoreFileBackup(request: any) {
    this.recordRuntimeCall("backup.restore_file");
    return {
      ok: true,
      workbookId: request.workbookId,
      backupId: request.backupId,
      mode: request.mode ?? "open-as-new",
      filePath: "/tmp/Agent Unit backup.xlsx"
    };
  }

  async pruneFileBackups(request: any) {
    this.recordRuntimeCall("backup.prune");
    return {
      ok: true,
      pruned: ["backup_old_unit"],
      candidates: [],
      skippedPinned: [],
      reclaimedBytes: 1024,
      request
    };
  }

  listSnapshots(workbookId: any) {
    this.recordRuntimeCall("snapshot.list");
    return {
      ok: true,
      snapshots: [{
        snapshotId: "snapshot_agent_unit",
        workbookId,
        reason: "Agent unit snapshot",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: {
          rangeSnapshots: [{
            fingerprint: {
              range: { workbookId, sheetName: "Data", address: "A1:D4" },
              hash: "hash_snapshot_agent_unit",
              cellCount: 16,
              capturedAt: "2026-06-18T00:00:00.000Z"
            }
          }]
        },
        createdAt: "2026-06-18T00:00:00.000Z"
      }]
    };
  }

  getSnapshot(snapshotId: any) {
    this.recordRuntimeCall("snapshot.get_compact");
    return {
      ok: true,
      snapshot: {
        snapshotId,
        workbookId,
        reason: "Agent unit snapshot",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: {
          rangeSnapshots: [{
            fingerprint: {
              range: { workbookId, sheetName: "Data", address: "A1:D4" },
              hash: "hash_snapshot_agent_unit",
              cellCount: 16,
              capturedAt: "2026-06-18T00:00:00.000Z"
            }
          }]
        },
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  getWorkbookSnapshot(snapshotId: any) {
    this.recordRuntimeCall("workbook.get_snapshot");
    return {
      ok: true,
      snapshot: {
        snapshotId,
        workbookId,
        reason: "Agent unit snapshot",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: { rangeSnapshots: [] },
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  compareSnapshots(leftSnapshotId: any, rightSnapshotId: any) {
    this.recordRuntimeCall("snapshot.compare_compact");
    return {
      ok: true,
      diff: {
        leftSnapshotId,
        rightSnapshotId,
        changedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        cellsChanged: 4,
        summary: "1 range changed"
      }
    };
  }

  async detectExternalChanges(request: any) {
    this.recordRuntimeCall("workbook.detect_external_changes");
    return { ok: true, diff: { leftSnapshotId: request.snapshotId, rightSnapshotId: "snapshot_current_unit", changedRanges: [], cellsChanged: 0, summary: "No external changes" } };
  }

  async refreshSnapshot(request: any) {
    this.recordRuntimeCall("snapshot.refresh");
    return {
      ok: true,
      snapshot: {
        snapshotId: "snapshot_refreshed_unit",
        workbookId,
        reason: request.reason,
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: { rangeSnapshots: [] },
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  async refreshWorkbookSnapshot(request: any) {
    this.recordRuntimeCall("workbook.refresh_snapshot");
    return {
      ok: true,
      snapshot: {
        snapshotId: "snapshot_refreshed_unit",
        workbookId,
        reason: request.reason,
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: { rangeSnapshots: [] },
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  invalidateSnapshot(snapshotId: any) {
    this.recordRuntimeCall("snapshot.invalidate");
    return {
      ok: true,
      snapshot: {
        snapshotId,
        workbookId,
        reason: "Agent unit snapshot",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        payload: { rangeSnapshots: [] },
        createdAt: "2026-06-18T00:00:00.000Z",
        invalidatedAt: "2026-06-18T00:01:00.000Z"
      }
    };
  }

  deleteSnapshot(snapshotId: any) {
    this.recordRuntimeCall("snapshot.delete");
    return { ok: true, snapshotId };
  }

  listFileBackups(workbookId: any) {
    this.recordRuntimeCall("backup.list");
    return {
      ok: true,
      backups: [{
        backup: {
          backupId: "backup_agent_unit",
          workbookId,
          kind: "workbook-copy",
          reason: "Agent unit backup",
          affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
          retention: "persistent",
          createdAt: "2026-06-18T00:00:00.000Z"
        },
        payload: {
          kind: "snapshot-json",
          path: "/tmp/open-workbook-backups/backup_agent_unit.json",
          pinned: false
        }
      }]
    };
  }

  getFileBackup(backupId: any) {
    this.recordRuntimeCall("backup.get");
    return {
      ok: true,
      backup: {
        backupId,
        workbookId,
        kind: "workbook-copy",
        reason: "Agent unit backup",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        retention: "persistent",
        createdAt: "2026-06-18T00:00:00.000Z"
      },
      payload: {
        kind: "snapshot-json",
        path: "/tmp/open-workbook-backups/backup_agent_unit.json",
        pinned: false
      }
    };
  }

  async verifyFileBackup(backupId: any) {
    this.recordRuntimeCall("backup.verify");
    return {
      ok: true,
      backup: {
        backupId,
        workbookId,
        kind: "file-copy",
        reason: "Agent unit backup",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        retention: "persistent",
        verifiedAt: "2026-06-18T00:02:00.000Z",
        restoreStatus: "available",
        createdAt: "2026-06-18T00:00:00.000Z"
      },
      manifest: {
        backupId,
        workbookId,
        filePath: "/tmp/open-workbook-backups/backup_agent_unit.xlsx",
        checksum: "checksum_agent_unit",
        restoreStatus: "available",
        verifiedAt: "2026-06-18T00:02:00.000Z"
      }
    };
  }

  pinFileBackup(backupId: any, pinned: boolean) {
    this.recordRuntimeCall(pinned ? "backup.pin" : "backup.unpin");
    return {
      ok: true,
      backup: {
        backupId,
        workbookId,
        kind: "workbook-copy",
        reason: "Agent unit backup",
        affectedRanges: [{ workbookId, sheetName: "Data", address: "A1:D4" }],
        retention: "persistent",
        pinned,
        createdAt: "2026-06-18T00:00:00.000Z"
      }
    };
  }

  async deleteFileBackup(backupId: any) {
    this.recordRuntimeCall("backup.delete");
    return { ok: true, backupId };
  }
}

export function selectionInfo(sheetName: string, address: string, position = { row: 2, column: 2 }) {
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
  if (sheetName === "Vendor Propose") {
    const rows = [
      ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "Reference Diesel Rate : 37.50 THB/liter"],
      ["", "YLTH_CTG_Zone BKK TT _Y2026", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "Item No.", "Transport Mode", "Route", "Orgin Name", "Orgin Area", "Destination Name", "Destination Area", "Truck Type", "Distance", "Est 20'ft vol", "Est 40'ft vol", "vol/month", "Est vol per year", "Truck Available\n(Truck/day)", "Vendor Propose\n(THB/trip)"],
      ["", 1, "Export", "Nongkhae - Klongtoey Port", "TOTO (Thai land)", "Nong Khae District, Saraburi", "Klongtoey Port", "Khlong Toei, Bangkok", "Trailer '20'40 HC", 194, 43, 5, 4, 48, "", ""],
      ["", 2, "Export", "Nongkhae - Ladkrabang Port", "TOTO (Thai land)", "Nong Khae District, Saraburi", "Ladkrabang Port", "Lat Krabang, Bangkok", "Trailer '20'40 HC", 196, 200, 30, 21, 230, "", ""],
      ["", 3, "Export", "Nongkhae - Sahathai Port", "TOTO (Thai land)", "Nong Khae District, Saraburi", "Sahathai Port", "Phra Pradaeng District, Samut Prakan", "Trailer '20'40 HC", 210, 43, 5, 4, 48, "", ""]
    ];
    if (address === "B3:D28") {
      return rows.slice(2).map((row) => row.slice(1, 4));
    }
    if (address === "O3:P5") {
      return rows.slice(2, 5).map((row) => row.slice(14, 16));
    }
    if (address === "O4:P6") {
      return rows.slice(3, 6).map((row) => row.slice(14, 16));
    }
    if (address === "A4:P6") {
      return rows.slice(3, 6);
    }
    if (address === "A1:P6") {
      return rows;
    }
    return rows;
  }
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
    if (address === "B2:B4") return [["A-100"], ["A-200"], ["A-300"]];
    if (address === "A2:D3") return [["2026-06-01", "A-100", 123, "Open"], ["2026-06-02", "A-101", 456, "Closed"]];
    if (address === "A3:D3") return [["2026-06-02", "A-101", 456, "Closed"]];
    if (address === "A2:B2") return [["20/6/26", "20/6/26"]];
    if (address === "B3") return [[""]];
    if (address === "C3") return [[200]];
    if (address === "C2") return [[100]];
    if (address === "C1:C4") return [["Amount"], [100], [200], [300]];
    if (address === "D1:D4") return [["Status"], ["Open"], ["Closed"], ["Open"]];
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
  if (sheetName === "Customer Master") {
    if (address === "A2:A4") return [["A-100"], ["A-200"], ["A-300"]];
    if (address === "B2:B4") return [["Gold"], ["Silver"], ["Bronze"]];
    return [["Account", "Tier"], ["A-100", "Gold"], ["A-200", "Silver"], ["A-300", "Bronze"]];
  }
  if (sheetName === "Mar 2026" || sheetName === "Apr 2026" || sheetName === "May 2026") {
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
  if (sheetName === "Report" && address === "B2") return [[600]];
  if (sheetName === "Report" && address === "B3") return [[""]];
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
  const summaryRows: unknown[][] = [
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
  ], summaryRows[0]!);
  const row2 = padToSummary([
    sheetName.startsWith("Mar") ? "2026-03-01" : "2026-04-01", "204", "71-4653", "Company gas top-up", "company_gas_topup", "Outflow", "2211.21", "2211.21", "0", "", "Bank", "proof.pdf", "text note", "",
    "INV-001", "204", sheetName.startsWith("Mar") ? "2026-03-01" : "2026-04-01", "ACME", "BK-001", "Customer A", "Job 204", "CONT-1", "20GP", "10000", "1000", "1000", "2000", "0", "12000", "360", "11640"
  ], summaryRows[1]!);
  if (sheetName.startsWith("May")) {
    const rows = Array.from({ length: 100 }, (_row, index) => padToSummary(index === 0 ? row1 : [], summaryRows[index % summaryRows.length] ?? []));
    rows[1] = padToSummary(["2026-05-01", "204", "71-4653", "Company gas top-up", "company_gas_topup", "Outflow", "2211.21", "2211.21", "0", "", "Bank", "proof.pdf", "text note"], summaryRows[1]!);
    for (const [rowNumber, amount] of [[25, 200], [27, 3400], [28, 60], [31, 1337.5], [33, 650], [37, 768], [38, 3600], [39, 1000], [40, 450], [44, 1500], [48, 665]] as Array<[number, number]>) {
      rows[rowNumber - 1] = padToSummary(["2026-05-18", "", "", "", "", "Outflow", amount, "", 0, "", "To X3488 MR. WITSARUT KONLA++"], summaryRows[(rowNumber - 1) % summaryRows.length] ?? []);
    }
    rows[32 - 1] = padToSummary(["2026-05-19", "", "", "", "", "Outflow", 500, "", 0, "", "To X0556 MR. PRAKRIT THARAS++"], summaryRows[31 % summaryRows.length] ?? []);
    rows[81 - 1] = padToSummary(["2026-05-30", "", "", "", "", "Outflow", 999, "", 0, "", "To X7010 OTHER PAYEE++"], summaryRows[80 % summaryRows.length] ?? []);
    return sliceMonthlyRows(rows, address);
  }
  const ownerFundRow = padToSummary([
    "2026-04-16", "", "", "เติมเงินเข้าบริษัท", "owner_fund_added", "Inflow", "10000", "10000", "0", "", "From X1183 MR. PRACH YOTHAPRA++", "fund-proof.pdf", "Owner adding fund"
  ], summaryRows[2]!);
  const rows = sheetName.startsWith("Apr")
    ? [row1, row2, ownerFundRow, ...summaryRows.slice(3).map((summary) => padToSummary([], summary))]
    : [row1, row2, ...summaryRows.slice(2).map((summary) => padToSummary([], summary))];
  return sliceMonthlyRows(rows, address);
}

function sliceMonthlyRows(rows: unknown[][], address: string) {
  if (address === "A2:AJ2") {
    return [rows[1] ?? []];
  }
  if (address === "A1:AJ1") {
    return [rows[0] ?? []];
  }
  if (address.startsWith("O")) {
    return rows.map((row) => row.slice(14, 31));
  }
  const parsed = tryParseA1Address(address);
  if (parsed) {
    const startRowIndex = Math.max(0, parsed.startRow - 1);
    const endRowIndex = Math.min(rows.length - 1, parsed.endRow - 1);
    const startColumnIndex = Math.max(0, parsed.startColumn - 1);
    const endColumnIndex = Math.min(35, parsed.endColumn - 1);
    return rows.slice(startRowIndex, endRowIndex + 1).map((row) => row.slice(startColumnIndex, endColumnIndex + 1));
  }
  return rows;
}

function padToSummary(left: unknown[], summary: unknown[]) {
  const row: unknown[] = Array.from({ length: 36 }, (_cell, index) => left[index] ?? "");
  summary.forEach((value, index) => {
    row[32 + index] = value;
  });
  return row;
}

function formulasFor(sheetName: string, address: string) {
  if (sheetName === "Report" && address === "B2") {
    return [["=SUM(Data!C2:C4)"]];
  }
  if (sheetName === "Apr 2026") {
    const parsed = tryParseA1Address(address);
    if (parsed && parsed.startColumn === 9 && parsed.endColumn === 9) {
      return Array.from({ length: parsed.endRow - parsed.startRow + 1 }, (_value, index) => {
        const row = parsed.startRow + index;
        return [`=H${row}-G${row}`];
      });
    }
    if (parsed && parsed.startColumn <= 9 && parsed.endColumn >= 9) {
      const values = valuesFor(sheetName, address);
      const formulaColumnIndex = 9 - parsed.startColumn;
      return values.map((row, index) => row.map((_value, columnIndex) => {
        const rowNumber = parsed.startRow + index;
        return columnIndex === formulaColumnIndex && rowNumber > 1 ? `=H${rowNumber}-G${rowNumber}` : null;
      }));
    }
  }
  const values = valuesFor(sheetName, address);
  return values.map((row) => row.map((value) => typeof value === "string" && value.startsWith("=") ? value : null));
}

function numberFormatsFor(sheetName: string, address: string) {
  const values = valuesFor(sheetName, address);
  if (sheetName === "Data" && address === "A2:B2") {
    return [["@", "@"]];
  }
  return values.map((row) => row.map(() => "General"));
}

function validationReport(scope: string, request: any) {
  return {
    ok: true,
    workbookId,
    scope,
    issues: [],
    data: { request }
  };
}

function cleaningReport(action: string, request: any, changedCells: number, data?: unknown) {
  return {
    ok: true,
    workbookId,
    target: { workbookId, sheetName: request.sheetName, address: request.address },
    action,
    changedCells,
    ...(data !== undefined ? { data } : {}),
    warnings: []
  };
}

function repairReport(repair: string, request: any, ok = true) {
  return {
    ok,
    workbookId,
    repair,
    repairedAt: "2026-06-18T00:00:00.000Z",
    issues: ok ? [] : [{ code: "CAPABILITY_UNAVAILABLE", severity: "warning", message: `${repair} repair is unavailable in fake runtime.` }],
    result: { request },
    warnings: []
  };
}

function monthlyColumns() {
  const names = [
    "Transaction Date", "Job ID", "Truck ID", "Description", "Transaction Type", "Direction", "Cash Amount", "Actual Amount", "Payment Variance", "Reconciliation Note", "Transfer From/To", "Proof File", "Detail Notes", "",
    "Invoice No", "Job ID", "Invoice Date", "Billed To", "Booking No", "Customer", "Job", "Container No", "Container Size", "Job Price", "Lifting On", "Lifting Off", "Total Lifting", "Other Fees", "Gross Billed", "W/H Tax", "Net Collect", "",
    "Metric", "Amount (THB)", "View", "Notes"
  ];
  return names.map((name, index) => ({
    name,
    normalizedName: name.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || `column_${index + 1}`,
    inferredType: /date/i.test(name) ? "date" as const : /amount|price|tax|collect|lifting|variance/i.test(name) ? "currency" as const : /type|direction|view|status/i.test(name) ? "status" as const : "text" as const,
    index,
    letter: columnName(index + 1)
  }));
}

function columnName(index: number): string {
  let value = index;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

export function createCachedMetadata(workbookContextId: string): WorkbookMetadata {
  const fingerprint = createMetadataFingerprint({ workbookId, workbook: activeWorkbook, sheets });
  return {
    workbookContextId,
    workbookKey: `id:${workbookId}`,
    detailLevel: "sampled",
    workbook: { workbookId, name: activeWorkbook.name, path: activeWorkbook.path, sheetCount: sheets.length, activeSheet: "Data" },
    sheets: [
      { id: "sheet:0", name: "Data", index: 0, usedRange: "A1:D4", rowCount: 4, columnCount: 4, kind: "transaction", headers: [], tableIds: ["table:Transactions"], sectionIds: [], summaryBlockIds: [], formulaRegionIds: [] },
      { id: "sheet:1", name: "Report", index: 1, usedRange: "A1:B20", rowCount: 20, columnCount: 2, kind: "summary", headers: [], tableIds: [], sectionIds: [], summaryBlockIds: [], formulaRegionIds: ["formula:manual"] },
      {
        id: "sheet:customer-master",
        name: "Customer Master",
        index: 2,
        usedRange: "A1:B4",
        rowCount: 4,
        columnCount: 2,
        kind: "lookup",
        headers: [{
          id: "header:Customer Master:1",
          sheetName: "Customer Master",
          row: 1,
          range: "A1:B1",
          confidence: 0.92,
          columns: [
            { name: "Account", normalizedName: "account", inferredType: "text", role: "account", importance: 0.9, index: 0, letter: "A" },
            { name: "Tier", normalizedName: "tier", inferredType: "text", role: "category", importance: 0.8, index: 1, letter: "B" }
          ]
        }],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      },
      {
        id: "sheet:2",
        name: "Mar 2026",
        index: 3,
        usedRange: "A1:AJ206",
        rowCount: 206,
        columnCount: 36,
        kind: "transaction",
        headers: [{
          id: "header:Mar 2026:1",
          sheetName: "Mar 2026",
          row: 1,
          range: "A1:AJ1",
          confidence: 0.9,
          columns: monthlyColumns()
        }],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      },
      {
        id: "sheet:3",
        name: "Apr 2026",
        index: 4,
        usedRange: "A1:AJ244",
        rowCount: 244,
        columnCount: 36,
        kind: "transaction",
        headers: [{
          id: "header:Apr 2026:1",
          sheetName: "Apr 2026",
          row: 1,
          range: "A1:AJ1",
          confidence: 0.9,
          columns: monthlyColumns()
        }],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      },
      {
        id: "sheet:4",
        name: "May 2026",
        index: 5,
        usedRange: "A1:AJ244",
        rowCount: 244,
        columnCount: 36,
        kind: "transaction",
        headers: [{
          id: "header:May 2026:1",
          sheetName: "May 2026",
          row: 1,
          range: "A1:AJ1",
          confidence: 0.9,
          columns: monthlyColumns()
        }],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      }
    ],
    tables: [{
      id: "table:Transactions",
      sheetName: "Data",
      name: "Transactions",
      range: "A1:D4",
      dataRange: "A2:D4",
      columns: [
        { name: "Date", normalizedName: "date", inferredType: "date", role: "date", importance: 0.97, index: 0, letter: "A" },
        { name: "Account", normalizedName: "account", inferredType: "text", role: "account", importance: 0.82, index: 1, letter: "B" },
        { name: "Amount", normalizedName: "amount", inferredType: "currency", role: "amount", importance: 0.99, index: 2, letter: "C" },
        { name: "Status", normalizedName: "status", inferredType: "text", role: "status", importance: 0.9, index: 3, letter: "D" }
      ]
    }],
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
