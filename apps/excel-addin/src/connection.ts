import type { JsonRpcMessage, JsonRpcRequest } from "@open-workbook/protocol";
import {
  calculateWorkbook,
  appendTableRows,
  applyTableFilters,
  captureStyleFingerprint,
  captureSheetFingerprint,
  captureTemplate,
  clearTableDataKeepFormulas,
  clearTableFilters,
  clearTableSort,
  closeWorkbook,
  convertFormulasToValues,
  copyChartFromTemplate,
  copyFormulaPatterns,
  copyPivotTableFromTemplate,
  copyStyleDimensions,
  copyTableStructure,
  createChart,
  createName,
  createPivotTable,
  createTable,
  deleteChart,
  deleteName,
  deletePivotTable,
  executeBatch,
  getActiveWorkbookContext,
  embedWorkbookLocalConfig,
  exportWorkbookFile,
  getChartInfo,
  getName,
  getPivotTableInfo,
  getRuntimeCapabilities,
  getSelection,
  getTableInfo,
  getWorkbookInfo,
  getWorkbookMap,
  findBlankCells,
  findFormulaErrors,
  listCharts,
  listNames,
  listPivotTables,
  listTables,
  repairTemplateConsistency,
  readTable,
  readRangeComments,
  readRangeConditionalFormatting,
  readRangeDataValidation,
  readRangeHyperlinks,
  readRangeMergedCells,
  readRangeNotes,
  readFormulaPatterns,
  readWorkbookEmbeddedLocalConfig,
  resizeTable,
  saveWorkbook,
  setActiveSheet,
  fillFormulaPattern,
  setTableStyle,
  setTableTotalRow,
  snapshotRanges,
  sortTable,
  searchRange,
  refreshAllPivotTables,
  refreshChart,
  refreshPivotTable,
  updateName,
  updateChartDataSource,
  updateTableRows
} from "./excel-executor.js";

export interface AddinConnectionOptions {
  backendUrl: string;
  heartbeatMs: number;
  reconnectMs?: number;
  onStatus?: (status: string) => void;
}

export class AddinConnection {
  private socket?: WebSocket;
  private heartbeat: number | undefined;
  private reconnect: number | undefined;
  private closedByUser = false;

  constructor(private readonly options: AddinConnectionOptions) {}

  connect(): void {
    this.closedByUser = false;
    this.stopReconnect();
    this.options.onStatus?.(`Connecting to ${this.options.backendUrl}...`);
    this.socket = new WebSocket(this.options.backendUrl);
    this.socket.addEventListener("open", () => {
      this.options.onStatus?.("Connected to local Open Workbook runtime.");
      this.sendNotification("addin.hello", {
        host: "excel",
        runtime: "office-js",
        capabilities: getRuntimeCapabilities(),
        connectedAt: new Date().toISOString()
      });
      this.startHeartbeat();
    });
    this.socket.addEventListener("message", (event) => this.handleMessage(JSON.parse(String(event.data))));
    this.socket.addEventListener("error", () => {
      this.options.onStatus?.(`Could not connect to ${this.options.backendUrl}. Retrying...`);
    });
    this.socket.addEventListener("close", () => {
      this.stopHeartbeat();
      if (!this.closedByUser) {
        this.options.onStatus?.("Disconnected from local runtime. Retrying...");
        this.scheduleReconnect();
      }
    });
  }

  disconnect(): void {
    this.closedByUser = true;
    this.stopHeartbeat();
    this.stopReconnect();
    this.socket?.close();
  }

  private scheduleReconnect(): void {
    this.stopReconnect();
    this.reconnect = window.setTimeout(() => this.connect(), this.options.reconnectMs ?? 2_000);
  }

  private stopReconnect(): void {
    if (this.reconnect !== undefined) {
      window.clearTimeout(this.reconnect);
      this.reconnect = undefined;
    }
  }

  private startHeartbeat(): void {
    this.stopHeartbeat();
    this.heartbeat = window.setInterval(() => {
      this.sendNotification("addin.heartbeat", { at: new Date().toISOString() });
    }, this.options.heartbeatMs);
  }

  private stopHeartbeat(): void {
    if (this.heartbeat !== undefined) {
      window.clearInterval(this.heartbeat);
      this.heartbeat = undefined;
    }
  }

  private handleMessage(message: JsonRpcMessage): void {
    if ("method" in message && "id" in message) {
      void this.handleRequest(message);
    }
  }

  private async handleRequest(request: JsonRpcRequest): Promise<void> {
    try {
      switch (request.method) {
        case "runtime.ping":
          this.sendSuccess(request.id, { ok: true, at: new Date().toISOString(), echo: request.params });
          break;
        case "runtime.disconnect":
          this.sendSuccess(request.id, { ok: true });
          this.disconnect();
          break;
        case "runtime.get_active_context":
          this.sendSuccess(request.id, await getActiveWorkbookContext());
          break;
        case "runtime.get_selection":
          this.sendSuccess(request.id, await getSelection());
          break;
        case "runtime.set_active_sheet": {
          const params = request.params as { sheetName: string };
          this.sendSuccess(request.id, await setActiveSheet(params.sheetName));
          break;
        }
        case "workbook.get_info":
          this.sendSuccess(request.id, await getWorkbookInfo());
          break;
        case "workbook.get_map":
          this.sendSuccess(request.id, await getWorkbookMap());
          break;
        case "workbook.calculate": {
          const params = request.params as { calculationType?: "full" | "recalculate" };
          this.sendSuccess(request.id, await calculateWorkbook(params.calculationType));
          break;
        }
        case "workbook.save":
          this.sendSuccess(request.id, await saveWorkbook());
          break;
        case "workbook.get_file": {
          const params = request.params as { workbookId: string; sliceSize?: number };
          this.sendSuccess(request.id, await exportWorkbookFile(params.workbookId, params.sliceSize));
          break;
        }
        case "workbook.close": {
          const params = request.params as { closeBehavior?: "Save" | "SkipSave" };
          this.sendSuccess(request.id, await closeWorkbook(params.closeBehavior));
          break;
        }
        case "workbook.snapshot_ranges": {
          const params = request.params as { workbookId: string; ranges: Parameters<typeof snapshotRanges>[1] };
          this.sendSuccess(request.id, await snapshotRanges(params.workbookId, params.ranges));
          break;
        }
        case "workbook.embed_local_config":
          this.sendSuccess(request.id, await embedWorkbookLocalConfig(request.params as Parameters<typeof embedWorkbookLocalConfig>[0]));
          break;
        case "workbook.read_embedded_local_config":
          this.sendSuccess(request.id, await readWorkbookEmbeddedLocalConfig((request.params as { workbookId: string }).workbookId));
          break;
        case "names.list":
          this.sendSuccess(request.id, await listNames((request.params as { workbookId: string }).workbookId));
          break;
        case "names.get":
          this.sendSuccess(request.id, await getName(request.params as Parameters<typeof getName>[0]));
          break;
        case "names.create":
          this.sendSuccess(request.id, await createName(request.params as Parameters<typeof createName>[0]));
          break;
        case "names.update":
          this.sendSuccess(request.id, await updateName(request.params as Parameters<typeof updateName>[0]));
          break;
        case "names.delete":
          this.sendSuccess(request.id, await deleteName(request.params as Parameters<typeof deleteName>[0]));
          break;
        case "pivot.list":
          this.sendSuccess(request.id, await listPivotTables((request.params as { workbookId: string }).workbookId));
          break;
        case "pivot.get_info":
          this.sendSuccess(request.id, await getPivotTableInfo(request.params as Parameters<typeof getPivotTableInfo>[0]));
          break;
        case "pivot.create":
          this.sendSuccess(request.id, await createPivotTable(request.params as Parameters<typeof createPivotTable>[0]));
          break;
        case "pivot.refresh":
          this.sendSuccess(request.id, await refreshPivotTable(request.params as Parameters<typeof refreshPivotTable>[0]));
          break;
        case "pivot.refresh_all":
          this.sendSuccess(request.id, await refreshAllPivotTables((request.params as { workbookId: string }).workbookId));
          break;
        case "pivot.copy_from_template":
          this.sendSuccess(request.id, await copyPivotTableFromTemplate(request.params as Parameters<typeof copyPivotTableFromTemplate>[0]));
          break;
        case "pivot.delete":
          this.sendSuccess(request.id, await deletePivotTable(request.params as Parameters<typeof deletePivotTable>[0]));
          break;
        case "chart.list":
          this.sendSuccess(request.id, await listCharts((request.params as { workbookId: string }).workbookId));
          break;
        case "chart.get_info":
          this.sendSuccess(request.id, await getChartInfo(request.params as Parameters<typeof getChartInfo>[0]));
          break;
        case "chart.create":
          this.sendSuccess(request.id, await createChart(request.params as Parameters<typeof createChart>[0]));
          break;
        case "chart.update_data_source":
          this.sendSuccess(request.id, await updateChartDataSource(request.params as Parameters<typeof updateChartDataSource>[0]));
          break;
        case "chart.copy_from_template":
          this.sendSuccess(request.id, await copyChartFromTemplate(request.params as Parameters<typeof copyChartFromTemplate>[0]));
          break;
        case "chart.refresh":
          this.sendSuccess(request.id, await refreshChart(request.params as Parameters<typeof refreshChart>[0]));
          break;
        case "chart.delete":
          this.sendSuccess(request.id, await deleteChart(request.params as Parameters<typeof deleteChart>[0]));
          break;
        case "range.read_hyperlinks":
          this.sendSuccess(request.id, await readRangeHyperlinks(request.params as Parameters<typeof readRangeHyperlinks>[0]));
          break;
        case "range.read_comments":
          this.sendSuccess(request.id, await readRangeComments(request.params as Parameters<typeof readRangeComments>[0]));
          break;
        case "range.read_notes":
          this.sendSuccess(request.id, await readRangeNotes(request.params as Parameters<typeof readRangeNotes>[0]));
          break;
        case "range.read_merged_cells":
          this.sendSuccess(request.id, await readRangeMergedCells(request.params as Parameters<typeof readRangeMergedCells>[0]));
          break;
        case "range.read_data_validation":
          this.sendSuccess(request.id, await readRangeDataValidation(request.params as Parameters<typeof readRangeDataValidation>[0]));
          break;
        case "range.read_conditional_formatting":
          this.sendSuccess(request.id, await readRangeConditionalFormatting(request.params as Parameters<typeof readRangeConditionalFormatting>[0]));
          break;
        case "range.search":
          this.sendSuccess(request.id, await searchRange(request.params as Parameters<typeof searchRange>[0]));
          break;
        case "range.find_blank_cells":
          this.sendSuccess(request.id, await findBlankCells(request.params as Parameters<typeof findBlankCells>[0]));
          break;
        case "range.find_errors":
          this.sendSuccess(request.id, await findFormulaErrors(request.params as Parameters<typeof findFormulaErrors>[0]));
          break;
        case "formula.read_patterns":
          this.sendSuccess(request.id, await readFormulaPatterns(request.params as Parameters<typeof readFormulaPatterns>[0]));
          break;
        case "formula.copy_patterns":
          this.sendSuccess(request.id, await copyFormulaPatterns(request.params as Parameters<typeof copyFormulaPatterns>[0]));
          break;
        case "formula.fill_pattern":
          this.sendSuccess(request.id, await fillFormulaPattern(request.params as Parameters<typeof fillFormulaPattern>[0]));
          break;
        case "formula.convert_to_values":
          this.sendSuccess(request.id, await convertFormulasToValues(request.params as Parameters<typeof convertFormulasToValues>[0]));
          break;
        case "table.list":
          this.sendSuccess(request.id, await listTables((request.params as { workbookId: string }).workbookId));
          break;
        case "table.get_info":
          this.sendSuccess(request.id, await getTableInfo(request.params as Parameters<typeof getTableInfo>[0]));
          break;
        case "table.read":
          this.sendSuccess(request.id, await readTable(request.params as Parameters<typeof readTable>[0]));
          break;
        case "table.create":
          this.sendSuccess(request.id, await createTable(request.params as Parameters<typeof createTable>[0]));
          break;
        case "table.resize":
          this.sendSuccess(request.id, await resizeTable(request.params as Parameters<typeof resizeTable>[0]));
          break;
        case "table.append_rows":
          this.sendSuccess(request.id, await appendTableRows(request.params as Parameters<typeof appendTableRows>[0]));
          break;
        case "table.update_rows":
          this.sendSuccess(request.id, await updateTableRows(request.params as Parameters<typeof updateTableRows>[0]));
          break;
        case "table.clear_data_keep_formulas":
          this.sendSuccess(request.id, await clearTableDataKeepFormulas(request.params as Parameters<typeof clearTableDataKeepFormulas>[0]));
          break;
        case "table.clear_filters":
          this.sendSuccess(request.id, await clearTableFilters(request.params as Parameters<typeof clearTableFilters>[0]));
          break;
        case "table.apply_filters":
          this.sendSuccess(request.id, await applyTableFilters(request.params as Parameters<typeof applyTableFilters>[0]));
          break;
        case "table.sort":
          this.sendSuccess(request.id, await sortTable(request.params as Parameters<typeof sortTable>[0]));
          break;
        case "table.clear_sort":
          this.sendSuccess(request.id, await clearTableSort(request.params as Parameters<typeof clearTableSort>[0]));
          break;
        case "table.set_total_row":
          this.sendSuccess(request.id, await setTableTotalRow(request.params as Parameters<typeof setTableTotalRow>[0]));
          break;
        case "table.set_style":
          this.sendSuccess(request.id, await setTableStyle(request.params as Parameters<typeof setTableStyle>[0]));
          break;
        case "table.copy_structure":
          this.sendSuccess(request.id, await copyTableStructure(request.params as Parameters<typeof copyTableStructure>[0]));
          break;
        case "operation.execute_batch":
          this.sendSuccess(request.id, await executeBatch(request.params as Parameters<typeof executeBatch>[0]));
          break;
        case "template.capture":
          this.sendSuccess(request.id, await captureTemplate(request.params as Parameters<typeof captureTemplate>[0]));
          break;
        case "template.capture_sheet":
          this.sendSuccess(request.id, await captureSheetFingerprint(request.params as Parameters<typeof captureSheetFingerprint>[0]));
          break;
        case "style.capture_fingerprint":
          this.sendSuccess(request.id, await captureStyleFingerprint(request.params as Parameters<typeof captureStyleFingerprint>[0]));
          break;
        case "style.copy_dimensions":
          this.sendSuccess(request.id, await copyStyleDimensions(request.params as Parameters<typeof copyStyleDimensions>[0]));
          break;
        case "template.repair":
          this.sendSuccess(request.id, await repairTemplateConsistency(request.params as Parameters<typeof repairTemplateConsistency>[0]));
          break;
        default:
          this.send({
            jsonrpc: "2.0",
            id: request.id,
            error: {
              code: -32601,
              message: `Method not implemented in add-in: ${request.method}`
            }
          });
      }
    } catch (error) {
      this.send({
        jsonrpc: "2.0",
        id: request.id,
        error: {
          code: -32000,
          message: error instanceof Error ? error.message : String(error)
        }
      });
    }
  }

  private sendNotification(method: string, params: unknown): void {
    this.send({ jsonrpc: "2.0", method, params });
  }

  private sendSuccess(id: string | number, result: unknown): void {
    this.send({ jsonrpc: "2.0", id, result });
  }

  private send(message: JsonRpcMessage): void {
    if (this.socket?.readyState === WebSocket.OPEN) {
      this.socket.send(JSON.stringify(message));
    }
  }
}
