import { BackupManager, BatchCompiler, PlanManager, TemplateRegistry } from "@open-workbook/excel-core";
import type {
  AddinExecuteBatchRequest,
  BatchRequest,
  ConnectionId,
  OperationResult,
  PlanCreateRequest,
  PlanId,
  TemplateExecutionSource,
  WorkbookRef,
  WorkbookSnapshotResponse
} from "@open-workbook/protocol";
import { runtimeError } from "@open-workbook/protocol";
import { SessionRegistry } from "./session-registry.js";
import type { AddinRpcClient } from "./addin-rpc-client.js";

export class RuntimeService {
  readonly sessions = new SessionRegistry();
  readonly backups = new BackupManager();
  readonly templates = new TemplateRegistry();
  readonly compiler = new BatchCompiler();
  readonly plans = new PlanManager(this.compiler, this.backups);
  private readonly addinClients = new Map<ConnectionId, AddinRpcClient>();

  attachAddinClient(connectionId: ConnectionId, client: AddinRpcClient): void {
    this.addinClients.set(connectionId, client);
  }

  detachAddinClient(connectionId: ConnectionId): void {
    this.addinClients.delete(connectionId);
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

  async applyPlan(planId: PlanId, confirmationToken?: string): Promise<OperationResult> {
    const batch = this.plans.createBatchRequest(planId, confirmationToken);
    const result = await this.applyBatch(batch);
    this.plans.markApplyResult(planId, result);
    return { ...result, planId };
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
    const beforeSnapshot =
      request.mode === "apply" && compiled.targetFingerprints.length > 0
        ? await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
            workbookId: request.workbookId,
            ranges: compiled.targetFingerprints.map((fingerprint) => fingerprint.range)
          })
        : undefined;

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

    const payload: AddinExecuteBatchRequest = {
      request,
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
}
