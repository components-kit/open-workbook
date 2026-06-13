import {
  type BackupId,
  type AgentId,
  type BatchRequest,
  type ExcelOperation,
  type OperationResult,
  type PlanCreateRequest,
  type PlanId,
  type PlanPreview,
  type SnapshotId,
  type TaskId,
  type WorkbookId,
  makeId,
  runtimeError
} from "@component-kit/open-workbook-protocol";
import { BackupManager } from "./backup-manager.js";
import { BatchCompiler } from "./batch-compiler.js";
import { createWorkbookFingerprint } from "./fingerprint.js";

export interface PlanRecord {
  planId: PlanId;
  workbookId: WorkbookId;
  goal: string;
  operations: ExcelOperation[];
  agentId?: AgentId | undefined;
  agentName?: string | undefined;
  taskId?: TaskId | undefined;
  role?: string | undefined;
  status: "draft" | "previewed" | "applying" | "applied" | "rolled_back" | "failed" | "cancelled";
  baseSnapshotId: SnapshotId;
  preview?: PlanPreview;
  createdAt: string;
  updatedAt: string;
}

export class PlanManager {
  private readonly plans = new Map<PlanId, PlanRecord>();

  constructor(
    private readonly batchCompiler = new BatchCompiler(),
    private readonly backupManager = new BackupManager()
  ) {}

  createPlan(request: PlanCreateRequest): PlanRecord {
    const planId = makeId<PlanId>("plan");
    const now = new Date().toISOString();
    const plan: PlanRecord = {
      planId,
      workbookId: request.workbookId,
      goal: request.goal,
      operations: request.operations,
      status: "draft",
      baseSnapshotId: request.baseSnapshotId ?? makeId<SnapshotId>("snapshot"),
      createdAt: now,
      updatedAt: now
    };
    if (request.agentId !== undefined) {
      plan.agentId = request.agentId;
    }
    if (request.agentName !== undefined) {
      plan.agentName = request.agentName;
    }
    if (request.taskId !== undefined) {
      plan.taskId = request.taskId;
    }
    if (request.role !== undefined) {
      plan.role = request.role;
    }
    this.plans.set(planId, plan);
    return plan;
  }

  previewPlan(planId: PlanId): PlanPreview {
    const plan = this.requirePlan(planId);
    const compiled = this.batchCompiler.compile({
      workbookId: plan.workbookId,
      operations: plan.operations,
      mode: "dry_run"
    });

    const backups = compiled.requiredBackups.map((kind) =>
      this.backupManager.createBackup({
        workbookId: plan.workbookId,
        kind,
        reason: `Plan preview: ${plan.goal}`,
        affectedRanges: compiled.targetFingerprints.map((fingerprint) => fingerprint.range)
      })
    );

    const preview: PlanPreview = {
      planId,
      workbookId: plan.workbookId,
      baseSnapshotId: plan.baseSnapshotId,
      requiredBackups: backups.map((backup) => backup.backupId),
      beforeWorkbookFingerprint: createWorkbookFingerprint(
        plan.workbookId,
        { planId, operationCount: plan.operations.length },
        { operations: plan.operations.map((operation) => operation.kind) }
      ),
      targetFingerprints: compiled.targetFingerprints,
      diffSummary: {
        title: `Preview for ${plan.goal}`,
        changedRanges: compiled.targetFingerprints.map((fingerprint) => fingerprint.range),
        cellsChanged: compiled.estimatedCellsTouched,
        formulasChanged: plan.operations.filter((operation) => operation.kind === "range.write_formulas").length,
        stylesChanged: 0,
        tablesChanged: 0,
        sheetsChanged: plan.operations.filter((operation) => operation.kind === "template.create_sheet_from_template").length,
        destructiveLevel: compiled.destructiveLevel
      },
      warnings: []
    };

    plan.preview = preview;
    plan.status = "previewed";
    plan.updatedAt = new Date().toISOString();
    return preview;
  }

  markApplyResult(planId: PlanId, result: OperationResult): PlanRecord {
    const plan = this.requirePlan(planId);
    if (result.ok && plan.preview) {
      plan.preview = {
        ...plan.preview,
        requiredBackups: [...new Set([...plan.preview.requiredBackups, ...result.backups])]
      };
    }
    plan.status = result.ok ? "applied" : "failed";
    plan.updatedAt = new Date().toISOString();
    return plan;
  }

  markRolledBack(planId: PlanId): PlanRecord {
    const plan = this.requirePlan(planId);
    plan.status = "rolled_back";
    plan.updatedAt = new Date().toISOString();
    return plan;
  }

  createBatchRequest(planId: PlanId, confirmationToken?: string): BatchRequest {
    const plan = this.requirePlan(planId);
    if (!plan.preview) {
      throw runtimeError("CONFIRMATION_REQUIRED", "Plan must be previewed before apply.", {
        retryable: true
      });
    }
    const request: BatchRequest = {
      workbookId: plan.workbookId,
      operations: plan.operations,
      mode: "apply" as const,
      planId,
      baseSnapshotId: plan.baseSnapshotId,
      expectedTargetFingerprints: plan.preview.targetFingerprints
    };
    if (plan.agentId !== undefined) {
      request.agentId = plan.agentId;
    }
    if (plan.agentName !== undefined) {
      request.agentName = plan.agentName;
    }
    if (plan.taskId !== undefined) {
      request.taskId = plan.taskId;
    }
    if (plan.role !== undefined) {
      request.role = plan.role;
    }
    return confirmationToken === undefined ? request : { ...request, confirmationToken };
  }

  replacePreviewFingerprints(
    planId: PlanId,
    replacement: Pick<PlanPreview, "beforeWorkbookFingerprint" | "targetFingerprints">
  ): PlanPreview {
    const plan = this.requirePlan(planId);
    if (!plan.preview) {
      throw new Error(`Plan has no preview: ${planId}`);
    }
    plan.preview = {
      ...plan.preview,
      beforeWorkbookFingerprint: replacement.beforeWorkbookFingerprint,
      targetFingerprints: replacement.targetFingerprints
    };
    plan.updatedAt = new Date().toISOString();
    return plan.preview;
  }

  getPlan(planId: PlanId): PlanRecord | undefined {
    return this.plans.get(planId);
  }

  listRecent(workbookId: WorkbookId): PlanRecord[] {
    return [...this.plans.values()]
      .filter((plan) => plan.workbookId === workbookId)
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  }

  load(records: PlanRecord[]): void {
    this.plans.clear();
    for (const record of records) {
      this.plans.set(record.planId, { ...record });
    }
  }

  dump(): PlanRecord[] {
    return [...this.plans.values()].map((plan) => ({ ...plan }));
  }

  requireBackups(planId: PlanId): BackupId[] {
    const plan = this.requirePlan(planId);
    return plan.preview?.requiredBackups ?? [];
  }

  private requirePlan(planId: PlanId): PlanRecord {
    const plan = this.plans.get(planId);
    if (!plan) {
      throw new Error(`Plan not found: ${planId}`);
    }
    return plan;
  }
}
