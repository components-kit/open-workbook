import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { BackupManager, BatchCompiler, DefaultPermissionPolicy, hashStable, parseA1Address, PlanManager, SnapshotManager, TemplateRegistry } from "@open-workbook/excel-core";
import type { BackupRecord, PermissionPolicy } from "@open-workbook/excel-core";
import type {
  AddinTemplateRepairRequest,
  AddinExecuteBatchRequest,
  A1Range,
  BackupId,
  BatchRequest,
  CellMatrix,
  CellValue,
  ChartCreateRequest,
  ChartSelector,
  ChartUpdateDataSourceRequest,
  CleaningReport,
  ConnectionId,
  ExcelOperation,
  FormulaCompareResponse,
  FormulaCopyPatternsRequest,
  FormulaFillRequest,
  FormulaMutationResponse,
  FormulaPatternRequest,
  FormulaPatternResponse,
  NameCreateRequest,
  NameInfo,
  NameSelector,
  NameUpdateRequest,
  OperationResult,
  OperationId,
  OperationWarning,
  PermissionState,
  PivotCreateRequest,
  PivotSelector,
  PlanCreateRequest,
  PlanId,
  RangeFingerprint,
  RangeMetadataResponse,
  RangeAreasSummary,
  RangeSnapshot,
  RangeMetadataRequest,
  RangeSearchRequest,
  RepairReport,
  RegionRegisterRequest,
  RegionSelector,
  SnapshotId,
  TableAppendRowsRequest,
  TableApplyFiltersRequest,
  TableCopyStructureRequest,
  TableCreateRequest,
  TableInfo,
  TableResizeRequest,
  TableSelector,
  TableSetStyleRequest,
  TableSetTotalRowRequest,
  TableSortRequest,
  TableUpdateRowsRequest,
  TemplateId,
  TemplateExecutionSource,
  TemplateCaptureRequest,
  TemplateCaptureResponse,
  SheetTemplateFingerprintResponse,
  TemplateValidationIssue,
  TemplateValidationResponse,
  StyleCompareResponse,
  StyleCopyRequest,
  StyleCopyResponse,
  StyleDimension,
  StyleFingerprintRequest,
  StyleFingerprintResponse,
  ValidationIssue,
  ValidationReport,
  WorkbookRegion,
  WorkbookId,
  WorkbookRef,
  WorkbookSnapshotResponse
} from "@open-workbook/protocol";
import { getToolCatalogSummary, PromptCatalog, ResourceCatalog, makeId, runtimeError } from "@open-workbook/protocol";
import { SessionRegistry } from "./session-registry.js";
import type { AddinRpcClient } from "./addin-rpc-client.js";

export class RuntimeService {
  readonly sessions = new SessionRegistry();
  readonly backups = new BackupManager();
  readonly snapshots = new SnapshotManager();
  readonly templates = new TemplateRegistry();
  readonly compiler = new BatchCompiler();
  readonly plans = new PlanManager(this.compiler, this.backups);
  private readonly addinClients = new Map<ConnectionId, AddinRpcClient>();
  private readonly regions = new Map<string, WorkbookRegion>();
  private permissionState: PermissionState = {
    ...DefaultPermissionPolicy,
    requireConfirmationFor: [],
    allowMacroExecution: false,
    scope: {},
    lockedRegions: []
  };
  private readonly events: Array<{
    eventId: string;
    connectionId: ConnectionId;
    method: string;
    params?: unknown;
    receivedAt: string;
  }> = [];
  private eventSubscriptionEnabled = true;
  private eventDebounceMs = 250;

  attachAddinClient(connectionId: ConnectionId, client: AddinRpcClient): void {
    this.addinClients.set(connectionId, client);
  }

  detachAddinClient(connectionId: ConnectionId): void {
    this.addinClients.delete(connectionId);
  }

  recordAddinEvent(connectionId: ConnectionId, method: string, params?: unknown): void {
    if (!this.eventSubscriptionEnabled) {
      return;
    }
    this.events.push({
      eventId: makeId<string>("event"),
      connectionId,
      method,
      params,
      receivedAt: new Date().toISOString()
    });
    if (this.events.length > 250) {
      this.events.splice(0, this.events.length - 250);
    }
  }

  getStatus() {
    const activeSession = this.sessions.getActive();
    return {
      ok: true,
      activeAddinConnected: Boolean(activeSession),
      sessions: this.sessions.list(),
      activeWorkbook: activeSession?.activeWorkbook
    };
  }

  getCapabilities(options: { includePreview?: boolean } = {}) {
    const catalogOptions = options.includePreview === undefined ? {} : { includePreview: options.includePreview };
    return {
      runtime: this.getStatus(),
      catalog: getToolCatalogSummary(catalogOptions),
      resources: ResourceCatalog,
      prompts: PromptCatalog
    };
  }

  subscribeEvents() {
    this.eventSubscriptionEnabled = true;
    return { ok: true, subscribed: true, debounceMs: this.eventDebounceMs };
  }

  unsubscribeEvents() {
    this.eventSubscriptionEnabled = false;
    return { ok: true, subscribed: false };
  }

  getRecentEvents(limit = 50) {
    return {
      ok: true,
      subscribed: this.eventSubscriptionEnabled,
      debounceMs: this.eventDebounceMs,
      events: this.events.slice(-limit).reverse()
    };
  }

  clearEvents() {
    this.events.splice(0, this.events.length);
    return { ok: true };
  }

  setEventDebounce(debounceMs: number) {
    this.eventDebounceMs = Math.max(0, Math.min(60_000, debounceMs));
    return { ok: true, debounceMs: this.eventDebounceMs };
  }

  getPermissions() {
    return { ok: true, permissions: this.permissionState };
  }

  setPermissions(update: Partial<PermissionState>) {
    this.permissionState = mergePermissionState(this.permissionState, update);
    return this.getPermissions();
  }

  requireConfirmation(levels: PermissionState["requireConfirmationFor"]) {
    this.permissionState = {
      ...this.permissionState,
      requireConfirmationFor: [...new Set(levels)]
    };
    return this.getPermissions();
  }

  setPermissionScope(scope: PermissionState["scope"]) {
    this.permissionState = {
      ...this.permissionState,
      scope: { ...scope }
    };
    return this.getPermissions();
  }

  allowDestructiveActions(allow: boolean) {
    this.permissionState = {
      ...this.permissionState,
      allowDestructiveActions: allow
    };
    return this.getPermissions();
  }

  allowMacroExecution(allow: boolean) {
    this.permissionState = {
      ...this.permissionState,
      allowMacroExecution: allow
    };
    return this.getPermissions();
  }

  async lockRegions(input: { workbookId: WorkbookId; regions: Array<{ regionName: string; reason?: string }> }) {
    const locked: PermissionState["lockedRegions"] = [];
    for (const item of input.regions) {
      const resolved = await this.getRegion({ workbookId: input.workbookId, regionName: item.regionName });
      const region = (resolved as { region?: WorkbookRegion }).region;
      if (!region) {
        return resolved;
      }
      const lockedRegion: PermissionState["lockedRegions"][number] = {
        workbookId: input.workbookId,
        regionName: item.regionName,
        sheetName: region.sheetName,
        address: region.address,
        lockedAt: new Date().toISOString()
      };
      if (item.reason !== undefined) {
        lockedRegion.reason = item.reason;
      }
      locked.push(lockedRegion);
    }
    const existing = this.permissionState.lockedRegions.filter(
      (region) => region.workbookId !== input.workbookId || !locked.some((candidate) => candidate.regionName === region.regionName)
    );
    this.permissionState = {
      ...this.permissionState,
      lockedRegions: [...existing, ...locked]
    };
    return this.getPermissions();
  }

  unlockRegions(input: { workbookId: WorkbookId; regionNames?: string[] }) {
    this.permissionState = {
      ...this.permissionState,
      lockedRegions: this.permissionState.lockedRegions.filter((region) => {
        if (region.workbookId !== input.workbookId) {
          return true;
        }
        return input.regionNames !== undefined && !input.regionNames.includes(region.regionName);
      })
    };
    return this.getPermissions();
  }

  connectAddinInfo() {
    return {
      ok: true,
      backendUrl: `ws://${process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1"}:${process.env.OPEN_WORKBOOK_PORT ?? 37845}${
        process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin"
      }`,
      activeAddinConnected: Boolean(this.sessions.getActive())
    };
  }

  disconnectActiveAddin() {
    const activeSession = this.sessions.getActive();
    if (!activeSession) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    this.addinClients.get(activeSession.connectionId)?.close();
    this.detachAddinClient(activeSession.connectionId);
    this.sessions.remove(activeSession.connectionId);
    return { ok: true };
  }

  async pingAddin() {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("runtime.ping", { at: new Date().toISOString() });
  }

  createPlan(request: PlanCreateRequest) {
    return this.plans.createPlan(request);
  }

  async previewPlan(planId: PlanId) {
    const preview = this.plans.previewPlan(planId);
    const client = this.getActiveAddinClient();
    if (!client || preview.diffSummary.changedRanges.length === 0) {
      return preview;
    }

    const snapshot = await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
      workbookId: preview.workbookId,
      ranges: preview.diffSummary.changedRanges
    });

    return this.plans.replacePreviewFingerprints(planId, {
      beforeWorkbookFingerprint: snapshot.workbookFingerprint,
      targetFingerprints: snapshot.rangeSnapshots.map((rangeSnapshot) => rangeSnapshot.fingerprint)
    });
  }

  async getActiveContext() {
    const activeSession = this.sessions.getActive();
    const client = this.getActiveAddinClient();
    if (!activeSession || !client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }

    const activeWorkbook = await client.request<WorkbookRef | undefined>("runtime.get_active_context");
    if (activeWorkbook) {
      this.sessions.update(activeSession.connectionId, { activeWorkbook });
    }
    return {
      ok: true,
      activeWorkbook
    };
  }

  async getSelection() {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("runtime.get_selection");
  }

  async getWorkbookInfo() {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return {
      ok: true,
      info: await client.request("workbook.get_info")
    };
  }

  async getWorkbookMap() {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return {
      ok: true,
      map: await client.request("workbook.get_map")
    };
  }

  async createWorkbookSnapshot(input: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] }) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }

    const ranges = input.ranges?.length ? input.ranges : await this.getUsedRangesForSnapshot(input.workbookId);
    const payload = await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
      workbookId: input.workbookId,
      ranges
    });
    const snapshot = this.snapshots.createSnapshot({
      workbookId: input.workbookId,
      reason: input.reason ?? "Manual workbook snapshot",
      affectedRanges: ranges,
      payload
    });
    return { ok: true, snapshot };
  }

  getSnapshot(snapshotId: SnapshotId) {
    const snapshot = this.snapshots.getSnapshot(snapshotId);
    if (!snapshot) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Snapshot not found: ${snapshotId}`, { retryable: false })
      };
    }
    return { ok: true, snapshot };
  }

  listSnapshots(workbookId: WorkbookId) {
    return {
      ok: true,
      snapshots: this.snapshots.listSnapshots(workbookId)
    };
  }

  invalidateSnapshot(snapshotId: SnapshotId) {
    const snapshot = this.snapshots.invalidate(snapshotId);
    if (!snapshot) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Snapshot not found: ${snapshotId}`, { retryable: false })
      };
    }
    return { ok: true, snapshot };
  }

  deleteSnapshot(snapshotId: SnapshotId) {
    return {
      ok: this.snapshots.deleteSnapshot(snapshotId)
    };
  }

  compareSnapshots(leftSnapshotId: SnapshotId, rightSnapshotId: SnapshotId) {
    const diff = this.snapshots.compare(leftSnapshotId, rightSnapshotId);
    if (!diff) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", "One or both snapshots were not found.", { retryable: false })
      };
    }
    return { ok: true, diff };
  }

  async detectExternalChanges(input: { workbookId: WorkbookId; snapshotId: SnapshotId }) {
    const base = this.snapshots.getSnapshot(input.snapshotId);
    if (!base) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Snapshot not found: ${input.snapshotId}`, { retryable: false })
      };
    }
    const current = await this.createWorkbookSnapshot({
      workbookId: input.workbookId,
      reason: `External change check against ${input.snapshotId}`,
      ranges: base.affectedRanges
    });
    if (!current.ok || !("snapshot" in current)) {
      return current;
    }
    return this.compareSnapshots(input.snapshotId, current.snapshot.snapshotId);
  }

  async calculateWorkbook(workbookId: WorkbookId, calculationType?: "full" | "recalculate") {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("workbook.calculate", { workbookId, calculationType });
  }

  async saveWorkbook(workbookId: WorkbookId) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("workbook.save", { workbookId });
  }

  saveWorkbookAs(workbookId: WorkbookId, targetPath?: string) {
    return {
      ok: false,
      workbookId,
      targetPath,
      error: runtimeError(
        "CAPABILITY_UNAVAILABLE",
        "Office.js does not expose a local Save As file path API. Use Excel UI or a future native host bridge for true save_as.",
        { retryable: false }
      )
    };
  }

  async exportWorkbookCopy(input: { workbookId: WorkbookId; reason?: string; targetPath?: string; ranges?: A1Range[] }) {
    const backupRequest: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] } = {
      workbookId: input.workbookId,
      reason: input.reason ?? "Export workbook copy requested"
    };
    if (input.ranges !== undefined) {
      backupRequest.ranges = input.ranges;
    }
    const backup = await this.createWorkbookBackup(backupRequest);
    return {
      ok: false,
      workbookId: input.workbookId,
      targetPath: input.targetPath,
      backup,
      error: runtimeError(
        "CAPABILITY_UNAVAILABLE",
        "Office.js cannot export a local .xlsx copy from the add-in. A persistent snapshot backup was created instead.",
        { retryable: false }
      )
    };
  }

  async closeWorkbook(workbookId: WorkbookId, closeBehavior?: "Save" | "SkipSave") {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("workbook.close", { workbookId, closeBehavior });
  }

  async readRangeMetadata(method: string, request: RangeMetadataRequest | RangeSearchRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request(method, request);
  }

  async listNames(workbookId: WorkbookId) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("names.list", { workbookId });
  }

  async getName(request: NameSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("names.get", request);
  }

  async createName(request: NameCreateRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    const backup = request.reference && request.sheetName ? await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason: `Before creating named range ${request.name}`,
      ranges: [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.reference }]
    }) : undefined;
    const result = await client.request("names.create", request);
    return backup ? { ok: true, backup, result } : result;
  }

  async updateName(request: NameUpdateRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    const backup = request.reference && request.sheetName ? await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason: `Before updating named range ${request.name}`,
      ranges: [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.reference }]
    }) : undefined;
    const result = await client.request("names.update", request);
    return backup ? { ok: true, backup, result } : result;
  }

  async deleteName(request: NameSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("names.delete", request);
  }

  async listPivotTables(workbookId: WorkbookId) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("pivot.list", { workbookId });
  }

  async getPivotTableInfo(request: PivotSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("pivot.get_info", request);
  }

  async createPivotTable(request: PivotCreateRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    if (!request.sourceTableName && !request.sourceAddress) {
      return {
        ok: false,
        error: runtimeError("RANGE_INVALID", "PivotTable creation requires sourceTableName or sourceAddress.", { retryable: false })
      };
    }
    const ranges = pivotCreateRanges(request);
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("PivotTable creation is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const backup = await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason: `Before creating PivotTable ${request.pivotTableName}`,
      ranges
    });
    const result = await client.request("pivot.create", request);
    return { ok: true, backup, result };
  }

  async refreshPivotTable(request: PivotSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("pivot.refresh", request);
  }

  async refreshAllPivotTables(workbookId: WorkbookId) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("pivot.refresh_all", { workbookId });
  }

  updatePivotSource(request: PivotCreateRequest) {
    return {
      ok: false,
      request,
      error: runtimeError("CAPABILITY_UNAVAILABLE", "Office.js does not expose safe in-place PivotTable source reassignment in this runtime. Create a new PivotTable from the desired source.", {
        retryable: false
      })
    };
  }

  copyPivotFromTemplate(request: PivotSelector & { templateId?: TemplateId }) {
    return {
      ok: false,
      request,
      error: runtimeError("CAPABILITY_UNAVAILABLE", "PivotTable template copy is not implemented yet. Use create plus template/style validation once field layout copy is added.", {
        retryable: false
      })
    };
  }

  async validatePivotSource(request: PivotSelector) {
    const info = await this.getPivotTableInfo(request);
    const source = (info as { info?: { source?: string } }).info?.source;
    return {
      ok: Boolean(source),
      info,
      issues: source ? [] : [{ code: "PIVOT_SOURCE_UNAVAILABLE", severity: "warning", message: "Pivot source string is unavailable from Office.js." }]
    };
  }

  async listCharts(workbookId: WorkbookId) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("chart.list", { workbookId });
  }

  async getChartInfo(request: ChartSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("chart.get_info", request);
  }

  async createChart(request: ChartCreateRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    const ranges = [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.sourceAddress }];
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Chart creation is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const backup = await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason: `Before creating chart ${request.chartName ?? request.chartType}`,
      ranges
    });
    const result = await client.request("chart.create", request);
    return { ok: true, backup, result };
  }

  async updateChartDataSource(request: ChartUpdateDataSourceRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    const ranges = [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.sourceAddress }];
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Chart data-source update is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const backup = await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason: `Before updating chart ${request.chartName} data source`,
      ranges
    });
    const result = await client.request("chart.update_data_source", request);
    return { ok: true, backup, result };
  }

  async copyChartFromTemplate(request: ChartSelector & { templateChartName: string; templateSheetName: string }) {
    const source = await this.getChartInfo({
      workbookId: request.workbookId,
      sheetName: request.templateSheetName,
      chartName: request.templateChartName
    });
    const target = await this.getChartInfo(request);
    return {
      ok: Boolean((source as { ok?: boolean }).ok && (target as { ok?: boolean }).ok),
      source,
      target,
      note: "Chart template copy currently reports source/target metadata. Style replay will be added after chart style fingerprints are captured."
    };
  }

  async refreshChart(request: ChartSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    return client.request("chart.refresh", request);
  }

  async deleteChart(request: ChartSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    const permissionWarnings = this.validateDirectMutation(request.workbookId, [], "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Chart deletion is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    return client.request("chart.delete", request);
  }

  async validateChartAgainstTemplate(request: ChartSelector & { templateChartName?: string; templateSheetName?: string }) {
    const target = await this.getChartInfo(request);
    const template =
      request.templateChartName && request.templateSheetName
        ? await this.getChartInfo({ workbookId: request.workbookId, sheetName: request.templateSheetName, chartName: request.templateChartName })
        : undefined;
    return {
      ok: Boolean((target as { ok?: boolean }).ok && (template === undefined || (template as { ok?: boolean }).ok)),
      target,
      template,
      note: "Chart validation currently verifies chart existence and metadata availability. Deep chart fingerprints are planned."
    };
  }

  async detectRegions(workbookId: WorkbookId) {
    const map = await this.getWorkbookMap();
    const names = await this.listNames(workbookId);
    const registered = this.listRegions(workbookId).regions;
    const nameCandidates = ((names as { names?: NameInfo[] }).names ?? [])
      .filter((name) => name.address && name.sheetName)
      .map((name) => ({
        workbookId,
        name: name.name,
        sheetName: name.sheetName,
        address: stripSheetName(name.address ?? ""),
        kind: "named-range",
        source: "named-range"
      }));
    const mapSheets = (map as { map?: { sheets?: Array<{ name: string; usedRange?: { address: string }; tables?: Array<{ name: string }> }> } }).map?.sheets ?? [];
    const usedRangeCandidates = mapSheets
      .filter((sheet) => sheet.usedRange?.address)
      .map((sheet) => ({
        workbookId,
        name: `${sheet.name}_UsedRange`,
        sheetName: sheet.name,
        address: stripSheetName(sheet.usedRange!.address),
        kind: "data",
        source: "detected"
      }));
    return {
      ok: true,
      registered,
      candidates: [...nameCandidates, ...usedRangeCandidates],
      sources: { map, names }
    };
  }

  async registerRegion(request: RegionRegisterRequest) {
    let namedItem: string | undefined;
    if (request.createNamedRange) {
      const createNameRequest: NameCreateRequest = {
        workbookId: request.workbookId,
        name: request.name,
        sheetName: request.sheetName,
        reference: request.address
      };
      if (request.description !== undefined) {
        createNameRequest.comment = request.description;
      }
      const createResult = await this.createName(createNameRequest);
      const created = (createResult as { result?: { name?: NameInfo }; name?: NameInfo }).result?.name ?? (createResult as { name?: NameInfo }).name;
      namedItem = created?.name ?? request.name;
    }
    const now = new Date().toISOString();
    const region: WorkbookRegion = {
      workbookId: request.workbookId,
      regionId: makeId<string>("region"),
      name: request.name,
      sheetName: request.sheetName,
      address: request.address,
      kind: request.kind ?? "data",
      source: request.createNamedRange ? "named-range" : "manual",
      createdAt: now,
      updatedAt: now
    };
    if (request.description !== undefined) {
      region.description = request.description;
    }
    if (request.templateId !== undefined) {
      region.templateId = request.templateId;
    }
    if (namedItem !== undefined) {
      region.namedItem = namedItem;
    }
    this.regions.set(regionKey(request.workbookId, request.name), region);
    return { ok: true, region };
  }

  listRegions(workbookId: WorkbookId) {
    return {
      ok: true,
      regions: [...this.regions.values()].filter((region) => region.workbookId === workbookId)
    };
  }

  async getRegion(request: RegionSelector) {
    const registered = this.regions.get(regionKey(request.workbookId, request.regionName));
    if (registered) {
      return { ok: true, region: registered };
    }
    const nameResult = await this.getName({ workbookId: request.workbookId, name: request.regionName });
    const name = (nameResult as { name?: NameInfo }).name;
    if (name?.sheetName && name.address) {
      const now = new Date().toISOString();
      const region: WorkbookRegion = {
        workbookId: request.workbookId,
        regionId: makeId<string>("region"),
        name: request.regionName,
        sheetName: name.sheetName,
        address: stripSheetName(name.address),
        kind: "named-range",
        source: "named-range",
        namedItem: name.name,
        createdAt: now,
        updatedAt: now
      };
      return { ok: true, region };
    }
    return {
      ok: false,
      error: runtimeError("WORKBOOK_NOT_FOUND", `Region not found: ${request.regionName}`, { retryable: false })
    };
  }

  async clearRegionValues(request: RegionSelector) {
    const region = await this.resolveRegion(request);
    if (!region.ok) {
      return region;
    }
    return this.applyBatch({
      workbookId: request.workbookId,
      mode: "apply",
      operations: [regionOperation("range.clear_values_keep_format", request.workbookId, region.region, `Clear region ${request.regionName}`)]
    });
  }

  async writeRegionValues(request: RegionSelector & { values: unknown[][] }) {
    const region = await this.resolveRegion(request);
    if (!region.ok) {
      return region;
    }
    return this.applyBatch({
      workbookId: request.workbookId,
      mode: "apply",
      operations: [
        {
          ...regionOperation("range.write_values", request.workbookId, region.region, `Write region ${request.regionName}`),
          values: request.values,
          preserveFormats: true
        } as ExcelOperation
      ]
    });
  }

  async fillRegion(request: RegionSelector & { values: unknown[][]; clearFirst?: boolean }) {
    const region = await this.resolveRegion(request);
    if (!region.ok) {
      return region;
    }
    const operations: ExcelOperation[] = [];
    if (request.clearFirst) {
      operations.push(regionOperation("range.clear_values_keep_format", request.workbookId, region.region, `Clear region ${request.regionName}`));
    }
    operations.push({
      ...regionOperation("range.write_values", request.workbookId, region.region, `Fill region ${request.regionName}`),
      values: request.values,
      preserveFormats: true
    } as ExcelOperation);
    return this.applyBatch({ workbookId: request.workbookId, mode: "apply", operations });
  }

  async cleanDetectHeaderRow(input: CleanRangeInput & { maxRows?: number }): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const values = await this.readRangeValues(target);
    if (!values.ok) {
      return cleaningError(input.workbookId, "detect_header_row", target, values.error);
    }
    const candidates = detectHeaderCandidates(values.values, input.maxRows ?? 10);
    return cleaningReport(input.workbookId, "detect_header_row", target, 0, {
      candidates,
      headerRowIndex: candidates[0]?.rowIndex ?? 0,
      headers: candidates[0] ? values.values[candidates[0].rowIndex] : []
    });
  }

  async cleanNormalizeHeaders(input: CleanRangeInput & { headerRowIndex?: number }): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const read = await this.readRangeValues(target);
    if (!read.ok) {
      return cleaningError(input.workbookId, "normalize_headers", target, read.error);
    }
    const headerRowIndex = input.headerRowIndex ?? detectHeaderCandidates(read.values, 10)[0]?.rowIndex ?? 0;
    const values = cloneMatrix(read.values);
    const before = values[headerRowIndex] ?? [];
    const normalized = dedupeHeaders(before.map((value) => normalizeHeader(String(value ?? ""))));
    values[headerRowIndex] = normalized;
    const result = await this.writeCleanValues(target, values, "Normalize headers");
    return cleaningReport(input.workbookId, "normalize_headers", target, changedCellCount([before], [normalized]), { headerRowIndex, headers: normalized }, result);
  }

  async cleanTrimWhitespace(input: CleanRangeInput): Promise<CleaningReport> {
    return this.cleanTransform(input, "trim_whitespace", (value) => (typeof value === "string" ? value.trim().replace(/\s+/g, " ") : value));
  }

  async cleanRemoveDuplicates(input: CleanRangeInput & { hasHeader?: boolean; keyColumns?: number[] }): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const read = await this.readRangeValues(target);
    if (!read.ok) {
      return cleaningError(input.workbookId, "remove_duplicates", target, read.error);
    }
    const hasHeader = input.hasHeader ?? true;
    const header = hasHeader ? [read.values[0] ?? []] : [];
    const body = hasHeader ? read.values.slice(1) : read.values;
    const seen = new Set<string>();
    const unique: CellMatrix = [];
    const keyColumns = input.keyColumns?.length ? input.keyColumns : undefined;
    for (const row of body) {
      const keyValues = keyColumns ? keyColumns.map((index) => row[index]) : row;
      const key = JSON.stringify(keyValues.map((value) => normalizeComparable(value)));
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      unique.push(row);
    }
    const compact = [...header, ...unique];
    const values = padMatrixRows(compact, read.values.length, read.values[0]?.length ?? 0);
    const result = await this.writeCleanValues(target, values, "Remove duplicate rows");
    return cleaningReport(input.workbookId, "remove_duplicates", target, read.values.length - compact.length, {
      removedRows: read.values.length - compact.length,
      remainingRows: compact.length
    }, result);
  }

  async cleanParseDates(input: CleanRangeInput): Promise<CleaningReport> {
    return this.cleanTransform(input, "parse_dates", (value) => parseDateValue(value));
  }

  async cleanParseNumbers(input: CleanRangeInput): Promise<CleaningReport> {
    return this.cleanTransform(input, "parse_numbers", (value) => parseNumberValue(value));
  }

  async cleanStandardizeCurrency(input: CleanRangeInput): Promise<CleaningReport> {
    return this.cleanTransform(input, "standardize_currency", (value) => parseCurrencyValue(value));
  }

  async cleanFillMissingValues(input: CleanRangeInput & { strategy?: "value" | "zero" | "previous" | "next"; value?: unknown }): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const read = await this.readRangeValues(target);
    if (!read.ok) {
      return cleaningError(input.workbookId, "fill_missing_values", target, read.error);
    }
    const strategy = input.strategy ?? "value";
    const values = cloneMatrix(read.values);
    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < (values[rowIndex]?.length ?? 0); columnIndex += 1) {
        if (!isMissing(values[rowIndex]![columnIndex])) {
          continue;
        }
        values[rowIndex]![columnIndex] = (
          strategy === "zero"
            ? 0
            : strategy === "previous"
              ? previousNonMissing(values, rowIndex, columnIndex)
              : strategy === "next"
                ? nextNonMissing(values, rowIndex, columnIndex)
                : input.value
        ) as CellValue;
      }
    }
    const result = await this.writeCleanValues(target, values, "Fill missing values");
    return cleaningReport(input.workbookId, "fill_missing_values", target, changedCellCount(read.values, values), { strategy }, result);
  }

  async cleanSplitColumn(input: CleanRangeInput & { columnIndex: number; delimiter?: string; targetAddress: string }): Promise<CleaningReport> {
    const source = targetFromCleanInput(input);
    const target = { ...source, address: input.targetAddress };
    const read = await this.readRangeValues(source);
    if (!read.ok) {
      return cleaningError(input.workbookId, "split_column", source, read.error);
    }
    const delimiter = input.delimiter ?? ",";
    const values = read.values.map((row) => String(row[input.columnIndex] ?? "").split(delimiter).map((part) => part.trim()));
    const result = await this.writeCleanValues(target, rectangularize(values), "Split column");
    return cleaningReport(input.workbookId, "split_column", target, values.length * (values[0]?.length ?? 0), { source, delimiter }, result);
  }

  async cleanMergeColumns(input: CleanRangeInput & { columnIndexes: number[]; separator?: string; targetAddress: string }): Promise<CleaningReport> {
    const source = targetFromCleanInput(input);
    const target = { ...source, address: input.targetAddress };
    const read = await this.readRangeValues(source);
    if (!read.ok) {
      return cleaningError(input.workbookId, "merge_columns", source, read.error);
    }
    const separator = input.separator ?? " ";
    const values = read.values.map((row) => [input.columnIndexes.map((index) => row[index]).filter((value) => !isMissing(value)).join(separator)]);
    const result = await this.writeCleanValues(target, values, "Merge columns");
    return cleaningReport(input.workbookId, "merge_columns", target, values.length, { source, columnIndexes: input.columnIndexes }, result);
  }

  async cleanDetectOutliers(input: CleanRangeInput & { columnIndex?: number; threshold?: number }): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const read = await this.readRangeValues(target);
    if (!read.ok) {
      return cleaningError(input.workbookId, "detect_outliers", target, read.error);
    }
    const columnIndex = input.columnIndex ?? 0;
    const threshold = input.threshold ?? 3;
    const numbers = read.values.map((row, rowIndex) => ({ rowIndex, value: typeof row[columnIndex] === "number" ? row[columnIndex] as number : parseCurrencyValue(row[columnIndex]) }));
    const numeric = numbers.filter((item): item is { rowIndex: number; value: number } => typeof item.value === "number" && Number.isFinite(item.value));
    const mean = numeric.reduce((sum, item) => sum + item.value, 0) / Math.max(1, numeric.length);
    const stddev = Math.sqrt(numeric.reduce((sum, item) => sum + (item.value - mean) ** 2, 0) / Math.max(1, numeric.length));
    const outliers = numeric
      .map((item) => ({ ...item, zScore: stddev === 0 ? 0 : (item.value - mean) / stddev }))
      .filter((item) => Math.abs(item.zScore) >= threshold);
    return cleaningReport(input.workbookId, "detect_outliers", target, 0, { columnIndex, threshold, mean, stddev, outliers });
  }

  async cleanFuzzyMatch(input: CleanRangeInput & { lookupValues: string[]; threshold?: number }): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const read = await this.readRangeValues(target);
    if (!read.ok) {
      return cleaningError(input.workbookId, "fuzzy_match", target, read.error);
    }
    const threshold = input.threshold ?? 0.75;
    const matches = read.values.flatMap((row, rowIndex) =>
      row.map((value, columnIndex) => {
        const text = String(value ?? "");
        const best = bestFuzzyMatch(text, input.lookupValues);
        return { rowIndex, columnIndex, value: text, match: best.value, score: best.score, accepted: best.score >= threshold };
      })
    );
    return cleaningReport(input.workbookId, "fuzzy_match", target, 0, { threshold, matches });
  }

  async validateWorkbook(input: { workbookId: WorkbookId }): Promise<ValidationReport> {
    const mapResult = await this.getWorkbookMap();
    const issues: ValidationIssue[] = [];
    if (!mapResult.ok || !("map" in mapResult)) {
      return makeValidationReport(input.workbookId, "workbook", [
        {
          code: "WORKBOOK_MAP_UNAVAILABLE",
          severity: "error",
          category: "workbook",
          message: "Workbook map could not be read from the connected Excel add-in.",
          details: { result: mapResult }
        }
      ]);
    }

    const map = mapResult.map as { sheets?: Array<{ name: string; usedRange?: { address: string; rowCount?: number; columnCount?: number } }> };
    if (!map.sheets?.length) {
      issues.push({
        code: "WORKBOOK_HAS_NO_SHEETS",
        severity: "error",
        category: "workbook",
        message: "Workbook has no visible sheets in the workbook map."
      });
    }
    for (const sheet of map.sheets ?? []) {
      if (!sheet.usedRange?.address) {
        issues.push({
          code: "SHEET_EMPTY",
          severity: "info",
          category: "sheet",
          message: `Sheet ${sheet.name} has no used range.`,
          target: { workbookId: input.workbookId, sheetName: sheet.name, address: "" }
        });
      }
    }

    const formulaReport = await this.validateFormulas(input);
    issues.push(...formulaReport.issues);
    return makeValidationReport(input.workbookId, "workbook", issues, { map });
  }

  async validateSheet(input: { workbookId: WorkbookId; sheetName: string }): Promise<ValidationReport> {
    const ranges = await this.getValidationRanges(input.workbookId, input.sheetName);
    const issues: ValidationIssue[] = [];
    if (ranges.length === 0) {
      issues.push({
        code: "SHEET_USED_RANGE_EMPTY",
        severity: "info",
        category: "sheet",
        message: `Sheet ${input.sheetName} has no used range.`,
        target: { workbookId: input.workbookId, sheetName: input.sheetName, address: "" }
      });
    }
    const formulaReport = await this.validateFormulas(input);
    issues.push(...formulaReport.issues);
    return makeValidationReport(input.workbookId, `sheet:${input.sheetName}`, issues, { ranges });
  }

  async validateTemplateConsistency(input: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string }): Promise<ValidationReport> {
    const result = await this.validateSheetAgainstTemplate(input);
    if (!result.ok && "error" in result) {
      return makeValidationReport(input.workbookId, `template:${input.templateId}:${input.targetSheetName}`, [
        {
          code: result.error.code,
          severity: "error",
          category: "template",
          message: result.error.message,
          details: { error: result.error }
        }
      ]);
    }
    return makeValidationReport(
      input.workbookId,
      `template:${input.templateId}:${input.targetSheetName}`,
      templateIssuesToValidationIssues(input.workbookId, result.issues),
      { templateValidation: result }
    );
  }

  async validateFormulas(input: { workbookId: WorkbookId; sheetName?: string; address?: string }): Promise<ValidationReport> {
    const ranges = await this.getValidationRanges(input.workbookId, input.sheetName, input.address);
    const issues: ValidationIssue[] = [];
    for (const range of ranges) {
      const result = (await this.readRangeMetadata("range.find_errors", range)) as RangeMetadataResponse;
      issues.push(...rangeMetadataWarningsToIssues("formula", result));
      if (result.ok && rangeAreasHasCells(result.data)) {
        issues.push({
          code: "FORMULA_ERRORS_FOUND",
          severity: "error",
          category: "formula",
          message: `Formula errors were found in ${range.sheetName}!${range.address}.`,
          target: range,
          details: { errors: result.data }
        });
      }
    }
    return makeValidationReport(input.workbookId, validationScope("formulas", input.sheetName, input.address), issues, { ranges });
  }

  async validateStyles(input: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string; sheetName?: string }): Promise<ValidationReport> {
    if (input.templateId && input.targetSheetName) {
      const report = await this.validateTemplateConsistency({
        workbookId: input.workbookId,
        templateId: input.templateId,
        targetSheetName: input.targetSheetName
      });
      return makeValidationReport(
        input.workbookId,
        `styles:${input.targetSheetName}`,
        report.issues.filter((issue) => issue.category === "style" || issue.category === "template"),
        report.data
      );
    }
    if (input.sheetName) {
      const fingerprint = await this.getSheetTemplateFingerprint({ workbookId: input.workbookId, sheetName: input.sheetName });
      return makeValidationReport(input.workbookId, `styles:${input.sheetName}`, [], { fingerprint });
    }
    return makeValidationReport(input.workbookId, "styles", [
      {
        code: "STYLE_VALIDATION_SCOPE_REQUIRED",
        severity: "warning",
        category: "style",
        message: "Provide sheetName for a style fingerprint or templateId with targetSheetName for consistency validation."
      }
    ]);
  }

  async validateTables(input: { workbookId: WorkbookId; tableName?: string; templateId?: TemplateId }): Promise<ValidationReport> {
    const issues: ValidationIssue[] = [];
    const data: Record<string, unknown> = {};
    if (input.tableName) {
      const info = await this.getTableInfo({ workbookId: input.workbookId, tableName: input.tableName });
      data.table = info;
      if (!(info as { ok?: boolean }).ok) {
        issues.push({
          code: "TABLE_INFO_UNAVAILABLE",
          severity: "error",
          category: "table",
          message: `Table ${input.tableName} could not be read.`,
          details: { result: info }
        });
      }
      if (input.templateId) {
        const templateResult = await this.validateTableAgainstTemplate({
          workbookId: input.workbookId,
          tableName: input.tableName,
          templateId: input.templateId
        });
        data.templateValidation = templateResult;
      }
    } else {
      const tables = await this.listTables(input.workbookId);
      data.tables = tables;
      const tableList = (tables as { tables?: unknown[] }).tables;
      if (Array.isArray(tableList) && tableList.length === 0) {
        issues.push({
          code: "NO_TABLES_FOUND",
          severity: "info",
          category: "table",
          message: "No structured tables were found in the workbook."
        });
      }
    }
    return makeValidationReport(input.workbookId, input.tableName ? `table:${input.tableName}` : "tables", issues, data);
  }

  async validateFilters(input: { workbookId: WorkbookId; tableName?: string }): Promise<ValidationReport> {
    const issues: ValidationIssue[] = [];
    const data: Record<string, unknown> = {};
    if (input.tableName) {
      const info = await this.getTableInfo({ workbookId: input.workbookId, tableName: input.tableName });
      data.table = info;
    } else {
      data.tables = await this.listTables(input.workbookId);
    }
    return makeValidationReport(input.workbookId, input.tableName ? `filters:${input.tableName}` : "filters", issues, data);
  }

  validatePrintLayout(input: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string }): ValidationReport {
    const issues: ValidationIssue[] = [];
    if (!input.templateId || !input.targetSheetName) {
      issues.push({
        code: "PRINT_LAYOUT_DEEP_VALIDATION_UNAVAILABLE",
        severity: "warning",
        category: "printLayout",
        message: "Office.js print layout coverage is limited; provide templateId and targetSheetName for template fingerprint comparison."
      });
    }
    return makeValidationReport(input.workbookId, "print_layout", issues, {
      templateId: input.templateId,
      targetSheetName: input.targetSheetName
    });
  }

  async validateNoBrokenReferences(input: { workbookId: WorkbookId; sheetName?: string; address?: string }): Promise<ValidationReport> {
    const ranges = await this.getValidationRanges(input.workbookId, input.sheetName, input.address);
    const issues: ValidationIssue[] = [];
    for (const range of ranges) {
      const result = (await this.readRangeMetadata("range.search", { ...range, text: "#REF!" })) as { ok: boolean; matches?: RangeAreasSummary };
      if (result.ok && rangeAreasHasCells(result.matches)) {
        issues.push({
          code: "BROKEN_REFERENCES_FOUND",
          severity: "error",
          category: "reference",
          message: `Broken #REF! references were found in ${range.sheetName}!${range.address}.`,
          target: range,
          details: { matches: result.matches }
        });
      }
    }
    return makeValidationReport(input.workbookId, validationScope("broken_references", input.sheetName, input.address), issues, { ranges });
  }

  async validateNoFormulaErrors(input: { workbookId: WorkbookId; sheetName?: string; address?: string }): Promise<ValidationReport> {
    const report = await this.validateFormulas(input);
    return makeValidationReport(input.workbookId, validationScope("no_formula_errors", input.sheetName, input.address), report.issues, report.data);
  }

  async validateNoUnintendedChanges(input: {
    workbookId: WorkbookId;
    snapshotId?: SnapshotId;
    leftSnapshotId?: SnapshotId;
    rightSnapshotId?: SnapshotId;
  }): Promise<ValidationReport> {
    const issues: ValidationIssue[] = [];
    let diffResult: unknown;
    if (input.leftSnapshotId && input.rightSnapshotId) {
      diffResult = this.compareSnapshots(input.leftSnapshotId, input.rightSnapshotId);
    } else if (input.snapshotId) {
      diffResult = await this.detectExternalChanges({ workbookId: input.workbookId, snapshotId: input.snapshotId });
    } else {
      return makeValidationReport(input.workbookId, "unintended_changes", [
        {
          code: "SNAPSHOT_REQUIRED",
          severity: "error",
          category: "change",
          message: "Provide snapshotId or both leftSnapshotId and rightSnapshotId to validate unintended changes."
        }
      ]);
    }
    const diff = (diffResult as { diff?: { cellsChanged?: number; formulasChanged?: number; stylesChanged?: number; tablesChanged?: number; sheetsChanged?: number } }).diff;
    const changed =
      (diff?.cellsChanged ?? 0) + (diff?.formulasChanged ?? 0) + (diff?.stylesChanged ?? 0) + (diff?.tablesChanged ?? 0) + (diff?.sheetsChanged ?? 0);
    if (changed > 0) {
      issues.push({
        code: "UNINTENDED_CHANGES_FOUND",
        severity: "error",
        category: "change",
        message: "Snapshot comparison detected workbook changes.",
        details: { diff }
      });
    }
    return makeValidationReport(input.workbookId, "unintended_changes", issues, { diffResult });
  }

  async repairStyleFromTemplate(input: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string }): Promise<RepairReport> {
    return this.templateRepairReport("style_from_template", input, ["styles"]);
  }

  async repairFormulasFromTemplate(input: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string }): Promise<RepairReport> {
    return this.templateRepairReport("formulas_from_template", input, ["formulas"]);
  }

  repairFiltersFromTemplate(input: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string }): RepairReport {
    return unsupportedRepairReport(
      input.workbookId,
      "filters_from_template",
      "FILTER_REPAIR_UNAVAILABLE",
      "Office.js does not expose enough filter fingerprint detail yet to safely replay registered template filters."
    );
  }

  async repairTableStructure(input: TableCopyStructureRequest): Promise<RepairReport> {
    const result = await this.copyTableStructure(input);
    return {
      ok: Boolean((result as { ok?: boolean }).ok),
      workbookId: input.workbookId,
      repair: "table_structure",
      repairedAt: new Date().toISOString(),
      backups: extractBackupIds(result),
      result,
      warnings: []
    };
  }

  repairPrintLayout(input: { workbookId: WorkbookId }): RepairReport {
    return unsupportedRepairReport(
      input.workbookId,
      "print_layout",
      "PRINT_LAYOUT_REPAIR_UNAVAILABLE",
      "Office.js does not expose enough page setup and print layout APIs here for safe repair."
    );
  }

  repairNamedRanges(input: { workbookId: WorkbookId }): RepairReport {
    return unsupportedRepairReport(
      input.workbookId,
      "named_ranges",
      "NAMED_RANGE_REPAIR_UNAVAILABLE",
      "Named range repair requires a template-aware names implementation that is not enabled yet."
    );
  }

  repairFormulaErrors(input: { workbookId: WorkbookId }): RepairReport {
    return unsupportedRepairReport(
      input.workbookId,
      "formula_errors",
      "FORMULA_ERROR_AUTO_REPAIR_UNAVAILABLE",
      "Formula errors are validated and reported, but automatic formula repair requires a template or explicit formula operation."
    );
  }

  repairMergedCells(input: { workbookId: WorkbookId }): RepairReport {
    return unsupportedRepairReport(
      input.workbookId,
      "merged_cells",
      "MERGED_CELL_REPAIR_UNAVAILABLE",
      "Merged-cell repair is not safe without an explicit template or target range policy."
    );
  }

  async listTables(workbookId: WorkbookId) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("table.list", { workbookId });
  }

  async getTableInfo(request: TableSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("table.get_info", request);
  }

  async readTable(request: TableSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("table.read", request);
  }

  async createTable(request: TableCreateRequest) {
    return this.mutateTable("table.create", request, `Before creating table ${request.tableName ?? request.address}`, [
      { workbookId: request.workbookId, sheetName: request.sheetName, address: request.address }
    ]);
  }

  async resizeTable(request: TableResizeRequest) {
    return this.mutateTable("table.resize", request, `Before resizing table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async appendTableRows(request: TableAppendRowsRequest) {
    return this.mutateTable("table.append_rows", request, `Before appending rows to table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async updateTableRows(request: TableUpdateRowsRequest) {
    return this.mutateTable("table.update_rows", request, `Before updating rows in table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async clearTableDataKeepFormulas(request: TableSelector) {
    return this.mutateTable(
      "table.clear_data_keep_formulas",
      request,
      `Before clearing table data for ${request.tableName}`,
      await this.getTableBackupRanges(request)
    );
  }

  async clearTableFilters(request: TableSelector) {
    return this.mutateTable("table.clear_filters", request, `Before clearing filters for table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async applyTableFilters(request: TableApplyFiltersRequest) {
    return this.mutateTable("table.apply_filters", request, `Before applying filters to table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async sortTable(request: TableSortRequest) {
    return this.mutateTable("table.sort", request, `Before sorting table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async clearTableSort(request: TableSelector) {
    return this.mutateTable("table.clear_sort", request, `Before clearing sort for table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async setTableTotalRow(request: TableSetTotalRowRequest) {
    return this.mutateTable("table.set_total_row", request, `Before setting total row for table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async setTableStyle(request: TableSetStyleRequest) {
    return this.mutateTable("table.set_style", request, `Before setting style for table ${request.tableName}`, await this.getTableBackupRanges(request));
  }

  async copyTableStructure(request: TableCopyStructureRequest) {
    return this.mutateTable("table.copy_structure", request, `Before copying table structure from ${request.tableName}`, [
      { workbookId: request.workbookId, sheetName: request.targetSheetName, address: request.targetAddress }
    ]);
  }

  async validateTableAgainstTemplate(request: TableSelector & { templateId: TemplateId }) {
    const tableInfo = await this.getTableInfo(request);
    const template = this.templates.get(request.templateId);
    if (!template) {
      return {
        ok: false,
        error: runtimeError("TEMPLATE_MISMATCH", `Template not found: ${request.templateId}`, { retryable: false })
      };
    }
    return {
      ok: true,
      table: tableInfo,
      templateTables: template.fingerprintPayload.tables,
      note: "Table validation compares current table metadata with registered template table fingerprints."
    };
  }

  async createWorkbookBackup(input: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] }) {
    const snapshotRequest: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] } = {
      workbookId: input.workbookId,
      reason: input.reason ?? "Manual backup snapshot"
    };
    if (input.ranges !== undefined) {
      snapshotRequest.ranges = input.ranges;
    }
    const snapshotResult = await this.createWorkbookSnapshot(snapshotRequest);
    if (!snapshotResult.ok || !("snapshot" in snapshotResult)) {
      return snapshotResult;
    }
    const backup = this.backups.createBackup({
      workbookId: input.workbookId,
      kind: "workbook-copy",
      reason: input.reason ?? "Manual workbook backup",
      affectedRanges: snapshotResult.snapshot.affectedRanges,
      payloadRef: snapshotResult.snapshot.snapshotId
    });
    backup.payload = snapshotResult.snapshot.payload;
    backup.payloadRef = await this.persistBackupPayload(backup.backupId, snapshotResult.snapshot.payload);
    return { ok: true, backup };
  }

  async restoreBackup(backupId: BackupId, confirmationToken?: string): Promise<OperationResult> {
    const backup = this.backups.getBackup(backupId);
    const snapshot = backup ? await this.loadBackupPayload(backup) : undefined;
    if (!backup || !snapshot?.rangeSnapshots) {
      return {
        ok: false,
        rollbackAvailable: false,
        backups: [],
        warnings: [],
        telemetry: {},
        error: runtimeError("BACKUP_UNAVAILABLE", `Backup is unavailable or has no restorable snapshot: ${backupId}`, {
          retryable: false
        })
      };
    }

    const operations: ExcelOperation[] = snapshot.rangeSnapshots.map((rangeSnapshot) => ({
      kind: "range.restore_snapshot",
      operationId: makeId<OperationId>("op"),
      workbookId: backup.workbookId,
      destructiveLevel: "format",
      reason: `Restore backup ${backupId}`,
      target: rangeSnapshot.fingerprint.range,
      snapshot: rangeSnapshot as RangeSnapshot
    }));
    const request: BatchRequest = {
      workbookId: backup.workbookId,
      mode: "apply",
      operations
    };
    if (confirmationToken !== undefined) {
      request.confirmationToken = confirmationToken;
    }
    return this.applyBatch(request);
  }

  private async mutateTable(method: string, request: { workbookId: WorkbookId }, reason: string, ranges: A1Range[]) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "values");
    if (permissionWarnings.length > 0) {
      return {
        ok: false,
        error: runtimeError("PERMISSION_DENIED", "Table mutation is blocked by the current Open Workbook permission policy.", {
          retryable: false,
          details: { permissionWarnings }
        }),
        warnings: permissionWarnings
      };
    }

    const backup = await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason,
      ranges
    });
    const result = await client.request(method, request);
    return { ok: true, backup, result };
  }

  private async mutateFormulas(
    method: string,
    request: FormulaCopyPatternsRequest | FormulaFillRequest | FormulaPatternRequest,
    reason: string,
    validate?: () => Promise<FormulaCompareResponse | { ok: false; error: ReturnType<typeof runtimeError> }>
  ) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    const ranges = await this.getFormulaMutationRanges(request);
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "values");
    if (permissionWarnings.length > 0) {
      return {
        ok: false,
        error: runtimeError("PERMISSION_DENIED", "Formula mutation is blocked by the current Open Workbook permission policy.", {
          retryable: false,
          details: { permissionWarnings }
        }),
        warnings: permissionWarnings
      };
    }

    const backup = await this.createWorkbookBackup({
      workbookId: request.workbookId,
      reason,
      ranges
    });
    const result = await client.request<FormulaMutationResponse>(method, request);
    const validation = validate ? await validate() : undefined;
    return { ok: result.ok, backup, result, validation };
  }

  private async getTableBackupRanges(request: TableSelector): Promise<A1Range[]> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return [];
    }
    const response = await client.request<{ ok: boolean; info: TableInfo }>("table.get_info", request);
    if (!response.info.sheetName || !response.info.address) {
      return [];
    }
    return [
      {
        workbookId: request.workbookId,
        sheetName: response.info.sheetName,
        address: response.info.address
      }
    ];
  }

  setActiveWorkbook(workbookIdOrName: string) {
    const session = this.sessions.setActiveWorkbook(workbookIdOrName);
    if (!session) {
      return {
        ok: false,
        error: runtimeError("WORKBOOK_NOT_FOUND", `No connected workbook matched ${workbookIdOrName}.`, {
          retryable: true
        })
      };
    }
    return {
      ok: true,
      activeWorkbook: session.activeWorkbook
    };
  }

  async setActiveSheet(sheetName: string) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return client.request("runtime.set_active_sheet", { sheetName });
  }

  async registerTemplate(request: TemplateCaptureRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      throw runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true });
    }
    const captured = await client.request<TemplateCaptureResponse>("template.capture", request);
    const input = {
      name: request.name,
      scope: request.scope,
      sourceSheetName: captured.sourceSheetName,
      dataRegions: captured.dataRegions,
      fingerprintPayload: captured.fingerprintPayload
    };
    return this.templates.register(
      request.scope === "workbook"
        ? { ...input, workbookId: request.workbookId }
        : input
    );
  }

  listTemplates(workbookId?: WorkbookId) {
    return workbookId === undefined ? this.templates.list() : this.templates.list({ workbookId });
  }

  getTemplate(templateId: TemplateId) {
    const template = this.templates.get(templateId);
    if (!template) {
      return {
        ok: false,
        error: runtimeError("TEMPLATE_MISMATCH", `Template not found: ${templateId}`, { retryable: false })
      };
    }
    return { ok: true, template };
  }

  unregisterTemplate(templateId: TemplateId) {
    return { ok: this.templates.unregister(templateId) };
  }

  async detectTemplates(workbookId: WorkbookId) {
    const mapResult = await this.getWorkbookMap();
    if (!mapResult.ok || !("map" in mapResult)) {
      return mapResult;
    }
    const map = mapResult.map as { sheets?: Array<{ name: string; usedRange?: { address: string; rowCount: number; columnCount: number } }> };
    return {
      ok: true,
      candidates:
        map.sheets?.map((sheet) => ({
          workbookId,
          sheetName: sheet.name,
          usedRange: sheet.usedRange,
          score: sheet.usedRange ? Math.min(1, (sheet.usedRange.rowCount * sheet.usedRange.columnCount) / 100) : 0,
          reason: sheet.usedRange ? "Sheet has a used range that can be registered as a template." : "Sheet is empty."
        })) ?? []
    };
  }

  inferTemplateRegions(templateId: TemplateId) {
    const template = this.templates.get(templateId);
    if (!template) {
      return {
        ok: false,
        error: runtimeError("TEMPLATE_MISMATCH", `Template not found: ${templateId}`, { retryable: false })
      };
    }
    return {
      ok: true,
      templateId,
      dataRegions: template.dataRegions,
      inferredRegions: template.dataRegions.map((address) => ({ address, kind: "data-entry" }))
    };
  }

  async validateSheetAgainstTemplate(input: {
    workbookId: WorkbookId;
    templateId: TemplateId;
    targetSheetName: string;
  }): Promise<TemplateValidationResponse | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const template = this.templates.get(input.templateId);
    const client = this.getActiveAddinClient();
    if (!template) {
      return {
        ok: false,
        error: runtimeError("TEMPLATE_MISMATCH", `Template not found: ${input.templateId}`, { retryable: false })
      };
    }
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }

    const captured = await client.request<SheetTemplateFingerprintResponse>("template.capture_sheet", {
      workbookId: input.workbookId,
      sheetName: input.targetSheetName,
      dataRegions: template.dataRegions
    });
    const issues = compareTemplatePayload(input.templateId, input.targetSheetName, template.fingerprintPayload, captured.fingerprintPayload);
    return {
      ok: issues.every((issue) => issue.severity !== "error"),
      sheetName: input.targetSheetName,
      templateId: input.templateId,
      issueCount: issues.length,
      issues,
      fingerprintPayload: captured.fingerprintPayload
    };
  }

  async getSheetTemplateFingerprint(input: { workbookId: WorkbookId; sheetName: string; dataRegions?: string[] }) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    const request: { workbookId: WorkbookId; sheetName: string; dataRegions?: string[] } = {
      workbookId: input.workbookId,
      sheetName: input.sheetName
    };
    if (input.dataRegions !== undefined) {
      request.dataRegions = input.dataRegions;
    }
    return {
      ok: true,
      fingerprint: await client.request<SheetTemplateFingerprintResponse>("template.capture_sheet", request)
    };
  }

  async getStyleFingerprint(
    input: StyleFingerprintRequest
  ): Promise<{ ok: true; fingerprint: StyleFingerprintResponse } | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return {
      ok: true,
      fingerprint: await client.request<StyleFingerprintResponse>("style.capture_fingerprint", input)
    };
  }

  async compareStyleFingerprints(input: {
    workbookId: WorkbookId;
    sourceSheetName: string;
    targetSheetName: string;
    sourceAddress?: string;
    targetAddress?: string;
    dimensions?: StyleDimension[];
    maxCellSamples?: number;
  }): Promise<StyleCompareResponse | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const sourceRequest: StyleFingerprintRequest = {
      workbookId: input.workbookId,
      sheetName: input.sourceSheetName
    };
    if (input.sourceAddress !== undefined) {
      sourceRequest.address = input.sourceAddress;
    }
    if (input.maxCellSamples !== undefined) {
      sourceRequest.maxCellSamples = input.maxCellSamples;
    }
    const targetRequest: StyleFingerprintRequest = {
      workbookId: input.workbookId,
      sheetName: input.targetSheetName
    };
    if (input.targetAddress !== undefined) {
      targetRequest.address = input.targetAddress;
    }
    if (input.maxCellSamples !== undefined) {
      targetRequest.maxCellSamples = input.maxCellSamples;
    }

    const source = await this.getStyleFingerprint(sourceRequest);
    const target = await this.getStyleFingerprint(targetRequest);
    if (!source.ok) {
      return source;
    }
    if (!target.ok) {
      return target;
    }

    const issues = compareStylePayloads(source.fingerprint, target.fingerprint, input.dimensions);
    return {
      ok: issues.length === 0,
      issueCount: issues.length,
      issues,
      sourceFingerprint: source.fingerprint,
      targetFingerprint: target.fingerprint
    };
  }

  async copyStyleDimensions(input: StyleCopyRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }

    const targetRanges =
      input.targetAddress !== undefined
        ? [
            {
              workbookId: input.workbookId,
              sheetName: input.targetSheetName,
              address: input.targetAddress
            }
          ]
        : await this.getSheetUsedRange(input.workbookId, input.targetSheetName);
    const permissionWarnings = this.validateDirectMutation(input.workbookId, targetRanges, "format");
    if (permissionWarnings.length > 0) {
      return {
        ok: false,
        error: runtimeError("PERMISSION_DENIED", "Style copy is blocked by the current Open Workbook permission policy.", {
          retryable: false,
          details: { permissionWarnings }
        }),
        warnings: permissionWarnings
      };
    }

    const backup = await this.createWorkbookBackup({
      workbookId: input.workbookId,
      reason: `Before copying style dimensions to ${input.targetSheetName}`,
      ranges: targetRanges
    });
    const result = await client.request<StyleCopyResponse>("style.copy_dimensions", input);
    const validation = await this.compareStyleFingerprints({
      workbookId: input.workbookId,
      sourceSheetName: input.sourceSheetName,
      targetSheetName: input.targetSheetName,
      ...(input.sourceAddress !== undefined ? { sourceAddress: input.sourceAddress } : {}),
      ...(input.targetAddress !== undefined ? { targetAddress: input.targetAddress } : {}),
      dimensions: input.dimensions
    });
    return { ok: result.ok, backup, result, validation };
  }

  async readFormulaPatterns(
    input: FormulaPatternRequest
  ): Promise<{ ok: true; patterns: FormulaPatternResponse } | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return {
      ok: true,
      patterns: await client.request<FormulaPatternResponse>("formula.read_patterns", input)
    };
  }

  async compareFormulaPatterns(input: {
    workbookId: WorkbookId;
    sourceSheetName: string;
    targetSheetName: string;
    sourceAddress?: string;
    targetAddress?: string;
  }): Promise<FormulaCompareResponse | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const sourceRequest: FormulaPatternRequest = {
      workbookId: input.workbookId,
      sheetName: input.sourceSheetName,
      ...(input.sourceAddress !== undefined ? { address: input.sourceAddress } : {})
    };
    const targetRequest: FormulaPatternRequest = {
      workbookId: input.workbookId,
      sheetName: input.targetSheetName,
      ...(input.targetAddress !== undefined ? { address: input.targetAddress } : {})
    };
    const source = await this.readFormulaPatterns(sourceRequest);
    const target = await this.readFormulaPatterns(targetRequest);
    if (!source.ok) {
      return source;
    }
    if (!target.ok) {
      return target;
    }
    const issues = compareFormulaPatternPayloads(source.patterns, target.patterns);
    return {
      ok: issues.length === 0,
      issueCount: issues.length,
      issues,
      sourcePatterns: source.patterns,
      targetPatterns: target.patterns
    };
  }

  async copyFormulaPatterns(input: FormulaCopyPatternsRequest) {
    return this.mutateFormulas("formula.copy_patterns", input, `Before copying formula patterns to ${input.targetSheetName}`, async () =>
      this.compareFormulaPatterns({
        workbookId: input.workbookId,
        sourceSheetName: input.sourceSheetName,
        targetSheetName: input.targetSheetName,
        ...(input.sourceAddress !== undefined ? { sourceAddress: input.sourceAddress } : {}),
        ...(input.targetAddress !== undefined ? { targetAddress: input.targetAddress } : {})
      })
    );
  }

  async fillFormulaPattern(input: FormulaFillRequest) {
    return this.mutateFormulas("formula.fill_pattern", input, `Before filling formulas in ${input.sheetName}!${input.targetAddress}`);
  }

  async convertFormulasToValues(input: FormulaPatternRequest) {
    return this.mutateFormulas("formula.convert_to_values", input, `Before converting formulas to values in ${input.sheetName}`);
  }

  async repairSheetFromTemplate(input: {
    workbookId: WorkbookId;
    templateId: TemplateId;
    targetSheetName: string;
    repair?: AddinTemplateRepairRequest["repair"];
  }) {
    const template = this.templates.get(input.templateId);
    const client = this.getActiveAddinClient();
    if (!template) {
      return {
        ok: false,
        error: runtimeError("TEMPLATE_MISMATCH", `Template not found: ${input.templateId}`, { retryable: false })
      };
    }
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }

    const ranges = await this.getSheetUsedRange(input.workbookId, input.targetSheetName);
    const backup = await this.createWorkbookBackup({
      workbookId: input.workbookId,
      reason: `Before repairing ${input.targetSheetName} from template ${input.templateId}`,
      ranges
    });
    const repairRequest: AddinTemplateRepairRequest = {
      workbookId: input.workbookId,
      templateId: input.templateId,
      sourceSheetName: template.sourceSheetName,
      targetSheetName: input.targetSheetName,
      dataRegions: template.dataRegions,
      repair: input.repair ?? ["styles", "formulas", "dataRegions"]
    };
    const result = await client.request("template.repair", repairRequest);
    const validation = await this.validateSheetAgainstTemplate({
      workbookId: input.workbookId,
      templateId: input.templateId,
      targetSheetName: input.targetSheetName
    });
    return { ok: true, backup, result, validation };
  }

  async applyPlan(planId: PlanId, confirmationToken?: string): Promise<OperationResult> {
    const batch = this.plans.createBatchRequest(planId, confirmationToken);
    const result = await this.applyBatch(batch);
    this.plans.markApplyResult(planId, result);
    return { ...result, planId };
  }

  async rollbackPlan(planId: PlanId, confirmationToken?: string): Promise<OperationResult> {
    const plan = this.plans.getPlan(planId);
    if (!plan?.preview) {
      return {
        ok: false,
        planId,
        rollbackAvailable: false,
        backups: [],
        warnings: [],
        telemetry: {},
        error: runtimeError("BACKUP_UNAVAILABLE", "Plan has no preview or rollback metadata.", { retryable: false })
      };
    }

    const operations = this.createRollbackOperations(planId);
    if (operations.length === 0) {
      return {
        ok: false,
        planId,
        rollbackAvailable: false,
        backups: [],
        warnings: [],
        telemetry: {},
        error: runtimeError("BACKUP_UNAVAILABLE", "No rollback operations could be created for this plan.", {
          retryable: false
        })
      };
    }

    const request: BatchRequest = {
      workbookId: plan.workbookId,
      mode: "apply",
      operations
    };
    if (confirmationToken !== undefined) {
      request.confirmationToken = confirmationToken;
    }
    const result = await this.applyBatch(request);
    if (result.ok) {
      this.plans.markRolledBack(planId);
    }
    const rollbackResult: OperationResult = {
      ...result,
      planId
    };
    if (result.diffSummary) {
      rollbackResult.diffSummary = {
        ...result.diffSummary,
        title: `Rollback for ${plan.goal}`
      };
    }
    return rollbackResult;
  }

  async applyBatch(request: BatchRequest): Promise<OperationResult> {
    const activeSession = this.sessions.getActive();
    const client = this.getActiveAddinClient();
    if (!activeSession || !client) {
      return {
        ok: false,
        rollbackAvailable: false,
        backups: [],
        warnings: [],
        telemetry: {},
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }

    const compiled = this.compiler.compile(request);
    const permissionWarnings = this.validateBatchPermissions(request, compiled);
    if (permissionWarnings.length > 0) {
      return {
        ok: false,
        rollbackAvailable: false,
        backups: [],
        warnings: permissionWarnings,
        telemetry: {},
        error: runtimeError("PERMISSION_DENIED", "Batch is blocked by the current Open Workbook permission policy.", {
          retryable: false,
          details: { permissionWarnings }
        })
      };
    }
    const beforeSnapshot =
      request.mode === "apply" && compiled.targetFingerprints.length > 0
        ? await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
            workbookId: request.workbookId,
            ranges: compiled.targetFingerprints.map((fingerprint) => fingerprint.range)
          })
        : undefined;
    const conflictWarnings =
      request.mode === "apply" && request.expectedTargetFingerprints?.length && beforeSnapshot
        ? detectFingerprintConflicts(request.expectedTargetFingerprints, beforeSnapshot.rangeSnapshots.map((snapshot) => snapshot.fingerprint))
        : [];
    if (conflictWarnings.length > 0) {
      return {
        ok: false,
        rollbackAvailable: compiled.requiredBackups.length > 0,
        backups: [],
        warnings: conflictWarnings,
        telemetry: {
          syncCount: 1,
          cellsRead: compiled.estimatedCellsTouched,
          rangeCount: compiled.targetFingerprints.length,
          warningCount: conflictWarnings.length
        },
        error: runtimeError("EXTERNAL_CHANGE_DETECTED", "Target ranges changed after preview. Refresh the plan before applying.", {
          retryable: true
        })
      };
    }

    const backups =
      request.mode === "apply"
        ? compiled.requiredBackups.map((kind) =>
            this.backups.createBackup({
              workbookId: request.workbookId,
              kind,
              reason: `Before ${request.operations.map((operation) => operation.kind).join(", ")}`,
              affectedRanges: compiled.targetFingerprints.map((fingerprint) => fingerprint.range),
              payload: kind === "region" ? beforeSnapshot : undefined
            })
          )
        : [];

    const executionRequest =
      request.expectedTargetFingerprints === undefined
        ? request
        : omitExpectedTargetFingerprints(request);
    const payload: AddinExecuteBatchRequest = {
      request: executionRequest,
      compiled,
      templateSources: this.resolveTemplateSources(request)
    };

    const result = await client.request<OperationResult>("operation.execute_batch", payload);
    return {
      ...result,
      backups: [...new Set([...result.backups, ...backups.map((backup) => backup.backupId)])],
      rollbackAvailable: result.rollbackAvailable || backups.length > 0
    };
  }

  private async cleanTransform(input: CleanRangeInput, action: string, transform: (value: unknown) => unknown): Promise<CleaningReport> {
    const target = targetFromCleanInput(input);
    const read = await this.readRangeValues(target);
    if (!read.ok) {
      return cleaningError(input.workbookId, action, target, read.error);
    }
    const values = read.values.map((row) => row.map((value) => transform(value) as CellValue));
    const changedCells = changedCellCount(read.values, values);
    const result = changedCells > 0 ? await this.writeCleanValues(target, values, action.replace(/_/g, " ")) : undefined;
    return cleaningReport(input.workbookId, action, target, changedCells, undefined, result);
  }

  private async readRangeValues(target: A1Range): Promise<{ ok: true; values: CellMatrix } | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    const snapshot = await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
      workbookId: target.workbookId,
      ranges: [target]
    });
    return {
      ok: true,
      values: snapshot.rangeSnapshots[0]?.values ?? []
    };
  }

  private async writeCleanValues(target: A1Range, values: CellMatrix, reason: string): Promise<OperationResult> {
    return this.applyBatch({
      workbookId: target.workbookId,
      mode: "apply",
      operations: [
        {
          kind: "range.write_values",
          operationId: makeId<OperationId>("op"),
          workbookId: target.workbookId,
          destructiveLevel: "values",
          reason,
          target,
          values,
          preserveFormats: true
        }
      ]
    });
  }

  private validateBatchPermissions(request: BatchRequest, compiled: ReturnType<BatchCompiler["compile"]>): OperationWarning[] {
    if (request.mode !== "apply") {
      return [];
    }
    const warnings: OperationWarning[] = [];
    const policy = this.permissionState;
    if (!policy.allowWrites && compiled.destructiveLevel !== "none") {
      warnings.push({ code: "WRITES_DISABLED", message: "Writes are disabled by permission policy." });
    }
    if (!policy.allowDestructiveActions && (compiled.destructiveLevel === "structure" || compiled.destructiveLevel === "workbook")) {
      warnings.push({ code: "DESTRUCTIVE_ACTION_BLOCKED", message: "Structure and workbook actions are disabled by permission policy." });
    }
    if (!policy.allowWorkbookActions && compiled.destructiveLevel === "workbook") {
      warnings.push({ code: "WORKBOOK_ACTION_BLOCKED", message: "Workbook-level actions are disabled by permission policy." });
    }
    if (policy.requireConfirmationFor.includes(compiled.destructiveLevel) && !request.confirmationToken) {
      warnings.push({ code: "CONFIRMATION_REQUIRED", message: `Confirmation token is required for ${compiled.destructiveLevel} operations.` });
    }
    warnings.push(...this.validatePermissionScope(request.workbookId, compiled.targetFingerprints.map((fingerprint) => fingerprint.range)));
    warnings.push(...this.validateLockedRegions(request.workbookId, compiled.targetFingerprints.map((fingerprint) => fingerprint.range)));
    return warnings;
  }

  private validatePermissionScope(workbookId: WorkbookId, ranges: A1Range[]): OperationWarning[] {
    const scope = this.permissionState.scope;
    const warnings: OperationWarning[] = [];
    if (scope.workbookId !== undefined && scope.workbookId !== workbookId) {
      warnings.push({ code: "WORKBOOK_SCOPE_BLOCKED", message: `Permission scope is restricted to workbook ${scope.workbookId}.` });
    }
    if (scope.sheetNames?.length) {
      for (const range of ranges) {
        if (!scope.sheetNames.includes(range.sheetName)) {
          warnings.push({ code: "SHEET_SCOPE_BLOCKED", message: `Sheet ${range.sheetName} is outside the permission scope.`, target: range });
        }
      }
    }
    if (scope.regionNames?.length) {
      const allowedRegions = scope.regionNames
        .map((regionName) => this.regions.get(regionKey(workbookId, regionName)))
        .filter((region): region is WorkbookRegion => region !== undefined);
      for (const range of ranges) {
        if (!allowedRegions.some((region) => rangesOverlap(range, region))) {
          warnings.push({ code: "REGION_SCOPE_BLOCKED", message: `${range.sheetName}!${range.address} is outside the allowed region scope.`, target: range });
        }
      }
    }
    return warnings;
  }

  private validateLockedRegions(workbookId: WorkbookId, ranges: A1Range[]): OperationWarning[] {
    const warnings: OperationWarning[] = [];
    const locked = this.permissionState.lockedRegions.filter((region) => region.workbookId === workbookId);
    for (const range of ranges) {
      for (const region of locked) {
        if (rangesOverlap(range, region)) {
          warnings.push({
            code: "LOCKED_REGION_BLOCKED",
            message: `${range.sheetName}!${range.address} overlaps locked region ${region.regionName}.`,
            target: range,
            details: { lockedRegion: region }
          });
        }
      }
    }
    return warnings;
  }

  private validateDirectMutation(workbookId: WorkbookId, ranges: A1Range[], destructiveLevel: PermissionPolicy["requireConfirmationFor"][number]): OperationWarning[] {
    const warnings: OperationWarning[] = [];
    if (!this.permissionState.allowWrites && destructiveLevel !== "none") {
      warnings.push({ code: "WRITES_DISABLED", message: "Writes are disabled by permission policy." });
    }
    if (!this.permissionState.allowDestructiveActions && (destructiveLevel === "structure" || destructiveLevel === "workbook")) {
      warnings.push({ code: "DESTRUCTIVE_ACTION_BLOCKED", message: "Structure and workbook actions are disabled by permission policy." });
    }
    if (!this.permissionState.allowWorkbookActions && destructiveLevel === "workbook") {
      warnings.push({ code: "WORKBOOK_ACTION_BLOCKED", message: "Workbook-level actions are disabled by permission policy." });
    }
    warnings.push(...this.validatePermissionScope(workbookId, ranges));
    warnings.push(...this.validateLockedRegions(workbookId, ranges));
    return warnings;
  }

  private async templateRepairReport(
    repair: string,
    input: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string },
    repairKinds: AddinTemplateRepairRequest["repair"]
  ): Promise<RepairReport> {
    const result = await this.repairSheetFromTemplate({ ...input, repair: repairKinds });
    const report: RepairReport = {
      ok: Boolean((result as { ok?: boolean }).ok),
      workbookId: input.workbookId,
      repair,
      repairedAt: new Date().toISOString(),
      backups: extractBackupIds(result),
      result,
      warnings: []
    };
    const validation = (result as { validation?: TemplateValidationResponse }).validation;
    if (validation !== undefined) {
      report.validation = validation;
    }
    return report;
  }

  private getActiveAddinClient(): AddinRpcClient | undefined {
    const activeSession = this.sessions.getActive();
    return activeSession ? this.addinClients.get(activeSession.connectionId) : undefined;
  }

  private resolveTemplateSources(request: BatchRequest): TemplateExecutionSource[] {
    return request.operations
      .filter((operation) => operation.kind === "template.create_sheet_from_template")
      .map((operation) => {
        const template = this.templates.get(operation.templateId);
        if (!template) {
          return undefined;
        }
        return {
          templateId: operation.templateId,
          sourceSheetName: template.sourceSheetName,
          dataRegions: template.dataRegions
        };
      })
      .filter((source): source is TemplateExecutionSource => source !== undefined);
  }

  private createRollbackOperations(planId: PlanId): ExcelOperation[] {
    const plan = this.plans.getPlan(planId);
    if (!plan?.preview) {
      return [];
    }

    const operations: ExcelOperation[] = [];
    const restoredRanges = new Set<string>();

    for (const original of [...plan.operations].reverse()) {
      if (original.kind === "template.create_sheet_from_template") {
        operations.push({
          kind: "sheet.delete",
          operationId: makeId<OperationId>("op"),
          workbookId: plan.workbookId,
          destructiveLevel: "structure",
          reason: `Rollback sheet created by plan ${planId}`,
          sheetName: original.newSheetName
        });
      }
    }

    for (const backupId of plan.preview.requiredBackups) {
      const backup = this.backups.getBackup(backupId);
      const snapshot = backup?.payload as WorkbookSnapshotResponse | undefined;
      if (!snapshot?.rangeSnapshots) {
        continue;
      }

      for (const rangeSnapshot of snapshot.rangeSnapshots) {
        const key = `${rangeSnapshot.fingerprint.range.sheetName}!${rangeSnapshot.fingerprint.range.address}`;
        if (restoredRanges.has(key)) {
          continue;
        }
        restoredRanges.add(key);
        operations.push({
          kind: "range.restore_snapshot",
          operationId: makeId<OperationId>("op"),
          workbookId: plan.workbookId as WorkbookId,
          destructiveLevel: "format",
          reason: `Rollback range snapshot from plan ${planId}`,
          target: rangeSnapshot.fingerprint.range,
          snapshot: rangeSnapshot as RangeSnapshot
        });
      }
    }

    return operations;
  }

  private async getUsedRangesForSnapshot(workbookId: WorkbookId): Promise<A1Range[]> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return [];
    }
    const workbookMap = await client.request<{
      sheets: Array<{ name: string; usedRange?: { address: string } }>;
    }>("workbook.get_map");
    return workbookMap.sheets
      .filter((sheet) => sheet.usedRange?.address)
      .map((sheet) => ({
        workbookId,
        sheetName: sheet.name,
        address: sheet.usedRange!.address
      }));
  }

  private async getSheetUsedRange(workbookId: WorkbookId, sheetName: string): Promise<A1Range[]> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return [];
    }
    const workbookMap = await client.request<{
      sheets: Array<{ name: string; usedRange?: { address: string } }>;
    }>("workbook.get_map");
    const sheet = workbookMap.sheets.find((candidate) => candidate.name === sheetName);
    if (!sheet?.usedRange?.address) {
      return [];
    }
    return [
      {
        workbookId,
        sheetName,
        address: sheet.usedRange.address
      }
    ];
  }

  private async getFormulaMutationRanges(request: FormulaCopyPatternsRequest | FormulaFillRequest | FormulaPatternRequest): Promise<A1Range[]> {
    if ("targetSheetName" in request) {
      if (request.targetAddress !== undefined) {
        return [{ workbookId: request.workbookId, sheetName: request.targetSheetName, address: request.targetAddress }];
      }
      return this.getSheetUsedRange(request.workbookId, request.targetSheetName);
    }
    if ("targetAddress" in request) {
      return [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.targetAddress }];
    }
    if (request.address !== undefined) {
      return [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.address }];
    }
    return this.getSheetUsedRange(request.workbookId, request.sheetName);
  }

  private async getValidationRanges(workbookId: WorkbookId, sheetName?: string, address?: string): Promise<A1Range[]> {
    if (sheetName && address) {
      return [{ workbookId, sheetName, address }];
    }
    if (sheetName) {
      return this.getSheetUsedRange(workbookId, sheetName);
    }
    return this.getUsedRangesForSnapshot(workbookId);
  }

  private async resolveRegion(request: RegionSelector): Promise<{ ok: true; region: WorkbookRegion } | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const result = await this.getRegion(request);
    if ((result as { ok?: boolean }).ok && (result as { region?: WorkbookRegion }).region) {
      return { ok: true, region: (result as { region: WorkbookRegion }).region };
    }
    return {
      ok: false,
      error:
        (result as { error?: ReturnType<typeof runtimeError> }).error ??
        runtimeError("WORKBOOK_NOT_FOUND", `Region not found: ${request.regionName}`, { retryable: false })
    };
  }

  private async persistBackupPayload(backupId: BackupId, payload: WorkbookSnapshotResponse): Promise<string> {
    const directory = this.getBackupDirectory();
    await mkdir(directory, { recursive: true });
    const filePath = path.join(directory, `${backupId}.json`);
    await writeFile(
      filePath,
      JSON.stringify(
        {
          backupId,
          persistedAt: new Date().toISOString(),
          payload
        },
        null,
        2
      ),
      "utf8"
    );
    return filePath;
  }

  private async loadBackupPayload(backup: BackupRecord): Promise<WorkbookSnapshotResponse | undefined> {
    if (backup.payload) {
      return backup.payload as WorkbookSnapshotResponse;
    }
    if (!backup.payloadRef || !backup.payloadRef.endsWith(".json")) {
      return undefined;
    }
    const raw = await readFile(backup.payloadRef, "utf8");
    const parsed = JSON.parse(raw) as { payload?: WorkbookSnapshotResponse };
    return parsed.payload;
  }

  private getBackupDirectory(): string {
    return process.env.OPEN_WORKBOOK_BACKUP_DIR ?? path.join(process.cwd(), ".open-workbook", "backups");
  }
}

function compareTemplatePayload(
  templateId: TemplateId,
  sheetName: string,
  expected: TemplateCaptureResponse["fingerprintPayload"],
  actual: TemplateCaptureResponse["fingerprintPayload"]
): TemplateValidationIssue[] {
  const issues: TemplateValidationIssue[] = [];
  for (const component of ["structure", "formulas", "styles", "filters", "tables", "printLayout"] as const) {
    const expectedHash = hashStable(normalizeTemplateComponent(component, expected[component]));
    const actualHash = hashStable(normalizeTemplateComponent(component, actual[component]));
    if (expectedHash === actualHash) {
      continue;
    }
    issues.push({
      code: `TEMPLATE_${component.toUpperCase()}_MISMATCH`,
      severity: component === "filters" || component === "printLayout" ? "warning" : "error",
      component,
      message: `${sheetName} differs from template ${templateId} for ${component}.`,
      expected: expectedHash,
      actual: actualHash
    });
  }
  return issues;
}

function compareStylePayloads(
  source: StyleFingerprintResponse,
  target: StyleFingerprintResponse,
  dimensions?: StyleDimension[]
): TemplateValidationIssue[] {
  const selectedDimensions = dimensions?.length ? dimensions : (Object.keys(source.dimensions) as StyleDimension[]);
  const issues: TemplateValidationIssue[] = [];
  for (const dimension of selectedDimensions) {
    const expectedHash = hashStable(source.dimensions[dimension] ?? null);
    const actualHash = hashStable(target.dimensions[dimension] ?? null);
    if (expectedHash === actualHash) {
      continue;
    }
    issues.push({
      code: `STYLE_${dimension.replace(/[A-Z]/g, (letter) => `_${letter}`).toUpperCase()}_MISMATCH`,
      severity: "error",
      component: "styles",
      message: `${target.sheetName} differs from ${source.sheetName} for ${dimension}.`,
      expected: expectedHash,
      actual: actualHash,
      target: {
        workbookId: target.workbookId,
        sheetName: target.sheetName,
        address: target.address
      }
    });
  }
  return issues;
}

function compareFormulaPatternPayloads(source: FormulaPatternResponse, target: FormulaPatternResponse): TemplateValidationIssue[] {
  const issues: TemplateValidationIssue[] = [];
  if (source.rowCount !== target.rowCount || source.columnCount !== target.columnCount) {
    issues.push({
      code: "FORMULA_RANGE_SHAPE_MISMATCH",
      severity: "error",
      component: "formulas",
      message: `${target.sheetName}!${target.address} formula range shape differs from ${source.sheetName}!${source.address}.`,
      expected: { rowCount: source.rowCount, columnCount: source.columnCount },
      actual: { rowCount: target.rowCount, columnCount: target.columnCount },
      target: { workbookId: target.workbookId, sheetName: target.sheetName, address: target.address }
    });
  }

  const rowCount = Math.min(source.patternMatrix.length, target.patternMatrix.length);
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const columnCount = Math.min(source.patternMatrix[rowIndex]?.length ?? 0, target.patternMatrix[rowIndex]?.length ?? 0);
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const expected = source.patternMatrix[rowIndex]?.[columnIndex] ?? null;
      const actual = target.patternMatrix[rowIndex]?.[columnIndex] ?? null;
      if (expected === actual) {
        continue;
      }
      issues.push({
        code: expected === null ? "FORMULA_UNEXPECTED_PATTERN" : actual === null ? "FORMULA_MISSING_PATTERN" : "FORMULA_PATTERN_MISMATCH",
        severity: "error",
        component: "formulas",
        message: `${target.sheetName}!${target.address} formula pattern differs at relative cell ${rowIndex},${columnIndex}.`,
        expected,
        actual,
        target: { workbookId: target.workbookId, sheetName: target.sheetName, address: target.address }
      });
      if (issues.length >= 100) {
        issues.push({
          code: "FORMULA_PATTERN_DIFF_TRUNCATED",
          severity: "warning",
          component: "formulas",
          message: "Formula pattern comparison stopped after 100 mismatches.",
          target: { workbookId: target.workbookId, sheetName: target.sheetName, address: target.address }
        });
        return issues;
      }
    }
  }
  return issues;
}

function normalizeTemplateComponent(component: string, payload: unknown): unknown {
  if (component !== "structure" || !payload || typeof payload !== "object" || Array.isArray(payload)) {
    return payload;
  }
  const { sheetName: _sheetName, ...rest } = payload as Record<string, unknown>;
  return rest;
}

function makeValidationReport(workbookId: WorkbookId, scope: string, issues: ValidationIssue[], data?: unknown): ValidationReport {
  const errorCount = issues.filter((issue) => issue.severity === "error").length;
  const warningCount = issues.filter((issue) => issue.severity === "warning").length;
  const infoCount = issues.filter((issue) => issue.severity === "info").length;
  const report: ValidationReport = {
    ok: errorCount === 0,
    workbookId,
    scope,
    checkedAt: new Date().toISOString(),
    issueCount: issues.length,
    summary: {
      errorCount,
      warningCount,
      infoCount
    },
    issues
  };
  if (data !== undefined) {
    report.data = data;
  }
  return report;
}

function templateIssuesToValidationIssues(workbookId: WorkbookId, issues: TemplateValidationIssue[]): ValidationIssue[] {
  return issues.map((issue) => {
    const mapped: ValidationIssue = {
      code: issue.code,
      severity: issue.severity,
      category: templateComponentToValidationCategory(issue.component),
      message: issue.message
    };
    if (issue.target !== undefined) {
      mapped.target = issue.target;
    } else {
      mapped.details = { component: issue.component };
    }
    if (issue.expected !== undefined || issue.actual !== undefined) {
      mapped.details = {
        ...(mapped.details ?? {}),
        expected: issue.expected,
        actual: issue.actual,
        workbookId
      };
    }
    return mapped;
  });
}

function templateComponentToValidationCategory(component: TemplateValidationIssue["component"]): ValidationIssue["category"] {
  if (component === "formulas") {
    return "formula";
  }
  if (component === "styles") {
    return "style";
  }
  if (component === "tables") {
    return "table";
  }
  if (component === "filters") {
    return "filter";
  }
  if (component === "printLayout") {
    return "printLayout";
  }
  return "template";
}

function rangeMetadataWarningsToIssues(category: ValidationIssue["category"], result: RangeMetadataResponse): ValidationIssue[] {
  return result.warnings.map((warning) => {
    const issue: ValidationIssue = {
      code: warning.code,
      severity: result.ok ? "warning" : "error",
      category,
      message: warning.message
    };
    if (warning.target !== undefined) {
      issue.target = warning.target;
    }
    if (warning.details !== undefined) {
      issue.details = warning.details;
    }
    return issue;
  });
}

function rangeAreasHasCells(data: unknown): boolean {
  if (!data || typeof data !== "object") {
    return false;
  }
  const summary = data as RangeAreasSummary;
  return summary.isNullObject !== true && (summary.cellCount ?? 0) > 0;
}

function validationScope(prefix: string, sheetName?: string, address?: string): string {
  if (sheetName && address) {
    return `${prefix}:${sheetName}!${address}`;
  }
  if (sheetName) {
    return `${prefix}:${sheetName}`;
  }
  return prefix;
}

function extractBackupIds(result: unknown): BackupId[] {
  const backup = (result as { backup?: { backup?: { backupId?: BackupId }; backupId?: BackupId } }).backup;
  if (backup?.backupId) {
    return [backup.backupId];
  }
  if (backup?.backup?.backupId) {
    return [backup.backup.backupId];
  }
  const backups = (result as { backups?: BackupId[] }).backups;
  return Array.isArray(backups) ? backups : [];
}

function detectFingerprintConflicts(expected: RangeFingerprint[], current: RangeFingerprint[]): OperationWarning[] {
  const currentByRange = new Map(current.map((fingerprint) => [rangeFingerprintKey(fingerprint), fingerprint]));
  const warnings: OperationWarning[] = [];
  for (const expectedFingerprint of expected) {
    const currentFingerprint = currentByRange.get(rangeFingerprintKey(expectedFingerprint));
    if (!currentFingerprint || currentFingerprint.hash !== expectedFingerprint.hash) {
      warnings.push({
        code: "TARGET_REGION_CHANGED",
        message: `Target changed after preview: ${expectedFingerprint.range.sheetName}!${expectedFingerprint.range.address}`,
        target: expectedFingerprint.range
      });
    }
  }
  return warnings;
}

function omitExpectedTargetFingerprints(request: BatchRequest): BatchRequest {
  const { expectedTargetFingerprints: _expectedTargetFingerprints, ...rest } = request;
  return rest;
}

function rangeFingerprintKey(fingerprint: RangeFingerprint): string {
  return `${fingerprint.range.workbookId}:${fingerprint.range.sheetName}!${fingerprint.range.address}`;
}

function regionKey(workbookId: WorkbookId, regionName: string): string {
  return `${workbookId}:${regionName.toLowerCase()}`;
}

function regionOperation(kind: "range.clear_values_keep_format" | "range.write_values", workbookId: WorkbookId, region: WorkbookRegion, reason: string): ExcelOperation {
  const operation: ExcelOperation = {
    kind,
    operationId: makeId<OperationId>("op"),
    workbookId,
    destructiveLevel: "values",
    reason,
    target: {
      workbookId,
      sheetName: region.sheetName,
      address: region.address
    }
  } as ExcelOperation;
  return operation;
}

function stripSheetName(address: string): string {
  const bangIndex = address.lastIndexOf("!");
  return bangIndex >= 0 ? address.slice(bangIndex + 1) : address;
}

function unsupportedRepairReport(workbookId: WorkbookId, repair: string, code: string, message: string): RepairReport {
  return {
    ok: false,
    workbookId,
    repair,
    repairedAt: new Date().toISOString(),
    backups: [],
    warnings: [],
    error: runtimeError("CAPABILITY_UNAVAILABLE", message, { retryable: false, details: { reasonCode: code } })
  };
}

function disconnectedError() {
  return {
    ok: false,
    error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
  };
}

function permissionDenied(message: string, permissionWarnings: OperationWarning[]) {
  return {
    ok: false,
    warnings: permissionWarnings,
    error: runtimeError("PERMISSION_DENIED", message, { retryable: false, details: { permissionWarnings } })
  };
}

function pivotCreateRanges(request: PivotCreateRequest): A1Range[] {
  const ranges: A1Range[] = [
    {
      workbookId: request.workbookId,
      sheetName: request.destinationSheetName,
      address: request.destinationAddress
    }
  ];
  if (request.sourceSheetName !== undefined && request.sourceAddress !== undefined) {
    ranges.push({
      workbookId: request.workbookId,
      sheetName: request.sourceSheetName,
      address: request.sourceAddress
    });
  }
  return ranges;
}

interface CleanRangeInput {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
}

function targetFromCleanInput(input: CleanRangeInput): A1Range {
  return {
    workbookId: input.workbookId,
    sheetName: input.sheetName,
    address: input.address
  };
}

function cleaningReport(
  workbookId: WorkbookId,
  action: string,
  target: A1Range,
  changedCells: number,
  data?: unknown,
  result?: OperationResult
): CleaningReport {
  const report: CleaningReport = {
    ok: result ? result.ok : true,
    workbookId,
    target,
    action,
    changedCells,
    warnings: result?.warnings ?? []
  };
  const affectedRows = target.address ? safeRowCount(target.address) : undefined;
  const affectedColumns = target.address ? safeColumnCount(target.address) : undefined;
  if (affectedRows !== undefined) {
    report.affectedRows = affectedRows;
  }
  if (affectedColumns !== undefined) {
    report.affectedColumns = affectedColumns;
  }
  if (data !== undefined) {
    report.data = data;
  }
  if (result !== undefined) {
    report.result = result;
  }
  if (result?.error !== undefined) {
    report.error = result.error;
  }
  return report;
}

function cleaningError(workbookId: WorkbookId, action: string, target: A1Range, error: ReturnType<typeof runtimeError>): CleaningReport {
  return {
    ok: false,
    workbookId,
    target,
    action,
    changedCells: 0,
    warnings: [],
    error
  };
}

function cloneMatrix(values: CellMatrix): CellMatrix {
  return values.map((row) => [...row]);
}

function changedCellCount(before: CellMatrix, after: CellMatrix): number {
  const rowCount = Math.max(before.length, after.length);
  let changed = 0;
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const columnCount = Math.max(before[rowIndex]?.length ?? 0, after[rowIndex]?.length ?? 0);
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      if (before[rowIndex]?.[columnIndex] !== after[rowIndex]?.[columnIndex]) {
        changed += 1;
      }
    }
  }
  return changed;
}

function detectHeaderCandidates(values: CellMatrix, maxRows: number): Array<{ rowIndex: number; score: number; nonEmptyCount: number; uniqueCount: number }> {
  return values
    .slice(0, Math.max(1, maxRows))
    .map((row, rowIndex) => {
      const nonEmpty = row.map((value) => String(value ?? "").trim()).filter(Boolean);
      const unique = new Set(nonEmpty.map((value) => normalizeHeader(value)));
      const textCount = nonEmpty.filter((value) => Number.isNaN(Number(value))).length;
      return {
        rowIndex,
        score: nonEmpty.length === 0 ? 0 : textCount / nonEmpty.length + unique.size / Math.max(1, nonEmpty.length),
        nonEmptyCount: nonEmpty.length,
        uniqueCount: unique.size
      };
    })
    .sort((left, right) => right.score - left.score);
}

function normalizeHeader(value: string): string {
  return value
    .trim()
    .replace(/([a-z])([A-Z])/g, "$1_$2")
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_+/g, "_")
    .toLowerCase();
}

function dedupeHeaders(headers: string[]): string[] {
  const seen = new Map<string, number>();
  return headers.map((header, index) => {
    const base = header || `column_${index + 1}`;
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function normalizeComparable(value: unknown): unknown {
  return typeof value === "string" ? value.trim().toLowerCase() : value;
}

function padMatrixRows(values: CellMatrix, rowCount: number, columnCount: number): CellMatrix {
  const padded = values.map((row) => [...row, ...Array(Math.max(0, columnCount - row.length)).fill("")]);
  while (padded.length < rowCount) {
    padded.push(Array(columnCount).fill(""));
  }
  return padded;
}

function rectangularize(values: CellMatrix): CellMatrix {
  const width = Math.max(0, ...values.map((row) => row.length));
  return values.map((row) => [...row, ...Array(width - row.length).fill("")]);
}

function parseDateValue(value: unknown): unknown {
  if (typeof value !== "string" || !value.trim()) {
    return value;
  }
  const timestamp = Date.parse(value);
  if (Number.isNaN(timestamp)) {
    return value;
  }
  return new Date(timestamp).toISOString().slice(0, 10);
}

function parseNumberValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }
  const normalized = value.replace(/,/g, "").trim();
  if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) {
    return value;
  }
  return Number(normalized);
}

function parseCurrencyValue(value: unknown): unknown {
  if (typeof value === "number") {
    return value;
  }
  if (typeof value !== "string") {
    return value;
  }
  const normalized = value.replace(/[,$€£¥\s]/g, "").replace(/^\((.*)\)$/, "-$1");
  if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) {
    return value;
  }
  return Number(normalized);
}

function isMissing(value: unknown): boolean {
  return value === null || value === undefined || (typeof value === "string" && value.trim() === "");
}

function previousNonMissing(values: CellMatrix, rowIndex: number, columnIndex: number): unknown {
  for (let index = rowIndex - 1; index >= 0; index -= 1) {
    if (!isMissing(values[index]?.[columnIndex])) {
      return values[index]?.[columnIndex];
    }
  }
  return "";
}

function nextNonMissing(values: CellMatrix, rowIndex: number, columnIndex: number): unknown {
  for (let index = rowIndex + 1; index < values.length; index += 1) {
    if (!isMissing(values[index]?.[columnIndex])) {
      return values[index]?.[columnIndex];
    }
  }
  return "";
}

function bestFuzzyMatch(value: string, candidates: string[]): { value: string; score: number } {
  return candidates.reduce(
    (best, candidate) => {
      const score = similarity(value, candidate);
      return score > best.score ? { value: candidate, score } : best;
    },
    { value: "", score: 0 }
  );
}

function similarity(left: string, right: string): number {
  const a = left.toLowerCase().trim();
  const b = right.toLowerCase().trim();
  if (a === b) {
    return 1;
  }
  if (!a || !b) {
    return 0;
  }
  const distance = levenshtein(a, b);
  return 1 - distance / Math.max(a.length, b.length);
}

function levenshtein(left: string, right: string): number {
  const previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let i = 0; i < left.length; i += 1) {
    let last = i;
    previous[0] = i + 1;
    for (let j = 0; j < right.length; j += 1) {
      const old = previous[j + 1]!;
      previous[j + 1] = Math.min(previous[j + 1]! + 1, previous[j]! + 1, last + (left[i] === right[j] ? 0 : 1));
      last = old;
    }
  }
  return previous[right.length]!;
}

function rangesOverlap(range: A1Range, region: { sheetName: string; address: string }): boolean {
  if (range.sheetName !== region.sheetName) {
    return false;
  }
  try {
    const left = parseA1Address(stripSheetName(range.address));
    const right = parseA1Address(stripSheetName(region.address));
    return left.startRow <= right.endRow && left.endRow >= right.startRow && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
  } catch {
    return false;
  }
}

function safeRowCount(address: string): number | undefined {
  try {
    const parsed = parseA1Address(stripSheetName(address));
    return parsed.endRow - parsed.startRow + 1;
  } catch {
    return undefined;
  }
}

function safeColumnCount(address: string): number | undefined {
  try {
    const parsed = parseA1Address(stripSheetName(address));
    return parsed.endColumn - parsed.startColumn + 1;
  } catch {
    return undefined;
  }
}

function mergePermissionState(current: PermissionState, update: Partial<PermissionState>): PermissionState {
  return {
    ...current,
    ...update,
    scope: update.scope ? { ...update.scope } : current.scope,
    lockedRegions: update.lockedRegions ? [...update.lockedRegions] : current.lockedRegions
  };
}
