import { mkdir, readFile, stat, unlink, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { AsyncLocalStorage } from "node:async_hooks";
import path from "node:path";
import {
  BackupManager,
  BatchCompiler,
  cellCount,
  cloneMatrix,
  DefaultPermissionPolicy,
  buildFormulaDependencyGraph,
  attachConflictGuidance,
  extractFormulaReferences,
  hashStable,
  LockManager,
  formatA1Address,
  padMatrixRows,
  parseA1Address,
  PlanManager,
  rangesOverlap as rangesOverlapAddresses,
  rectangularize,
  SnapshotManager,
  stripSheetName,
  TaskRegistry,
  TemplateRegistry,
  TransactionManager,
  traceDependents,
  tracePrecedents
} from "@components-kit/open-workbook-excel-core";
import { makeRollbackConflict } from "@components-kit/open-workbook-excel-core";
import type { BackupRecord, PermissionPolicy } from "@components-kit/open-workbook-excel-core";
import type { TemplateRecord } from "@components-kit/open-workbook-excel-core";
import type {
  AddinTemplateRepairRequest,
  AddinExecuteBatchRequest,
  A1Range,
  AgentId,
  AgentRunExecutionContext,
  AgentRunInput,
  AgentRunOutput,
  AgentRecord,
  BackupId,
  BatchChunkPlan,
  BatchPreflightResult,
  BatchRequest,
  CellMatrix,
  CellValue,
  ChartCreateRequest,
  ChartSelector,
  ChartUpdateDataSourceRequest,
  CleaningReport,
  CollaborationEvent,
  ConflictRecord,
  ConflictTelemetryRecord,
  ConflictTelemetrySummary,
  ConnectionId,
  DestructiveLevel,
  DiffSummary,
  ExcelOperation,
  FormulaCompareResponse,
  FormulaDependencyGraph,
  FormulaTraceResponse,
  FormulaCopyPatternsRequest,
  ConflictGuidanceResponse,
  JobId,
  JobRecord,
  JobStatus,
  LockAcquireResponse,
  LockId,
  LockLeasePolicy,
  LockMode,
  LockRenewResponse,
  LockReleaseResponse,
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
  PivotCopyFromTemplateResponse,
  PivotCopyFromTemplateRequest,
  PivotCompareFingerprintRequest,
  PivotCreateRequest,
  PivotDiff,
  PivotFingerprint,
  PivotLayoutInfo,
  PivotOperationCapabilityStatus,
  PivotRebuildWithSourceRequest,
  PivotRepairFromTemplateRequest,
  PivotSelector,
  PivotTableInfo,
  PivotValidateSourceRequest,
  PlanCreateRequest,
  PlanId,
  PlanRefreshResult,
  RangeFingerprint,
  RangeMetadataResponse,
  RangeAreasSummary,
  RangeSnapshot,
  RangeMetadataRequest,
  RangeSearchRequest,
  RepairReport,
  RegionRegisterRequest,
  RegionSelector,
  RuntimeCapabilities,
  RuntimeSelectionResponse,
  SnapshotId,
  TaskId,
  TaskBlocker,
  TaskScheduleResponse,
  TaskRecord,
  TableAppendRowsRequest,
  TableApplyFiltersRequest,
  TableApplyViewRequest,
  TableCopyStructureRequest,
  TableCreateRequest,
  TableInfo,
  TableReadRequest,
  TableReorderColumnsRequest,
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
  TransactionId,
  TransactionRecord,
  TransactionRollbackPreview,
  TransactionRollbackConflict,
  TransactionRollbackChainPreview,
  SheetTemplateFingerprintResponse,
  TemplateValidationIssue,
  TemplateValidationResponse,
  StyleCompareResponse,
  StyleCopyManyResponse,
  StyleCopyRequest,
  StyleCopyResponse,
  StyleDimension,
  StyleFingerprintRequest,
  StyleFingerprintResponse,
  ValidationIssue,
  ValidationReport,
  WorkbookBackupRetentionRequest,
  WorkbookCreateFileBackupRequest,
  WorkbookFileContent,
  WorkbookFileBackupManifest,
  WorkbookRestoreFileBackupRequest,
  WorkbookScope,
  WorkbookRegion,
  WorkbookId,
  WorkbookEmbeddedLocalConfigResponse,
  WorkbookLocalConfig,
  WorkbookLocalConfigImportRequest,
  WorkbookLocalConfigImportResponse,
  WorkbookRef,
  WorkbookSnapshotResponse
} from "@components-kit/open-workbook-protocol";
import { getInternalCapabilityCatalogSummary, getPublicAgentToolCatalogSummary, PromptCatalog, ResourceCatalog, makeId, runtimeError } from "@components-kit/open-workbook-protocol";
import { SessionRegistry } from "./session-registry.js";
import type { AddinRpcClient } from "./addin-rpc-client.js";
import { NativeFileBridge } from "./native-file-bridge.js";
import { RuntimeStateStore } from "./state-store.js";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { addinHealthTimeoutMs, addinStaleTtlMs, defaultLockLeasePolicy, runtimeVersion } from "./runtime/config.js";

type AddinConnectionState = "disconnected" | "connected_no_workbook" | "ready" | "stale";

export interface RuntimeServiceOptions {
  stateDir?: string;
  persistState?: boolean;
  fileBridge?: NativeFileBridge;
}

export class RuntimeService {
  readonly sessions = new SessionRegistry();
  readonly backups = new BackupManager();
  readonly snapshots = new SnapshotManager();
  readonly templates = new TemplateRegistry();
  readonly compiler = new BatchCompiler();
  readonly plans = new PlanManager(this.compiler, this.backups);
  readonly tasks = new TaskRegistry();
  readonly locks = new LockManager();
  readonly transactions = new TransactionManager();
  readonly agent: AgentOrchestrator;
  private readonly agentExecutionContext = new AsyncLocalStorage<AgentRunExecutionContext>();
  private readonly addinClients = new Map<ConnectionId, AddinRpcClient>();
  private readonly regions = new Map<string, WorkbookRegion>();
  private readonly agents = new Map<AgentId, AgentRecord>();
  private readonly jobs = new Map<JobId, JobRecord>();
  private readonly directMutationResults = new Map<string, unknown>();
  private readonly workbookContentVersions = new Map<string, number>();
  private readonly collabEvents: CollaborationEvent[] = [];
  private readonly conflicts: ConflictRecord[] = [];
  private readonly conflictTelemetry: ConflictTelemetryRecord[] = [];
  private transactionQueue: Promise<void> = Promise.resolve();
  private runtimeMutationActive = false;
  private runtimeMutationQueuedCount = 0;
  private readonly cancelledQueuedTransactions = new Set<TransactionId>();
  private readonly defaultAgentId: AgentId = "agent_daemon" as AgentId;
  private readonly stateStore: RuntimeStateStore | undefined;
  private readonly fileBridge: NativeFileBridge;
  private lockLeasePolicy: LockLeasePolicy = defaultLockLeasePolicy();
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

  constructor(options: RuntimeServiceOptions = {}) {
    this.stateStore = options.persistState === false ? undefined : new RuntimeStateStore(options.stateDir);
    this.fileBridge = options.fileBridge ?? new NativeFileBridge();
    this.agent = new AgentOrchestrator(this, { onOperationsChanged: () => this.persistState() });
    this.restoreState();
    this.recoverRuntimeState();
    void this.applyDefaultBackupRetention("startup");
    this.registerAgent({
      agentId: this.defaultAgentId,
      agentName: process.env.OPEN_WORKBOOK_AGENT_NAME ?? "local-agent",
      clientType: "daemon",
      pid: process.pid
    });
  }

  registerAgent(input: { agentId?: AgentId | undefined; agentName?: string | undefined; clientType?: AgentRecord["clientType"] | undefined; pid?: number | undefined } = {}) {
    const now = new Date().toISOString();
    const agentId = input.agentId ?? makeId<AgentId>("agent");
    const existing = this.agents.get(agentId);
    const agent: AgentRecord = existing
      ? {
          ...existing,
          agentName: input.agentName ?? existing.agentName,
          clientType: input.clientType ?? existing.clientType,
          pid: input.pid ?? existing.pid,
          status: "active",
          lastSeenAt: now
        }
      : {
          agentId,
          agentName: input.agentName,
          clientType: input.clientType ?? "mcp",
          pid: input.pid,
          status: "active",
          connectedAt: now,
          lastSeenAt: now
        };
    this.agents.set(agentId, agent);
    this.recordCollabEvent({
      type: existing ? "agent.heartbeat" : "agent.registered",
      agentId,
      message: existing ? `Agent ${agent.agentName ?? agentId} heartbeat.` : `Agent ${agent.agentName ?? agentId} registered.`
    });
    return { ok: true, agent };
  }

  createTask(input: {
    workbookId: WorkbookId;
    goal: string;
    role?: string | undefined;
    priority?: TaskRecord["priority"] | undefined;
    assignedAgentId?: AgentId | undefined;
    allowedScopes?: WorkbookScope[] | undefined;
    dependencies?: TaskId[] | undefined;
  }) {
    const task = this.tasks.create(input);
    this.recordCollabEvent({
      type: "task.created",
      workbookId: input.workbookId,
      agentId: task.assignedAgentId,
      taskId: task.taskId,
      message: `Task created: ${task.goal}`
    });
    return { ok: true, task };
  }

  claimTask(taskId: TaskId, agentId: AgentId) {
    const task = this.tasks.claim(taskId, agentId);
    this.recordCollabEvent({
      type: "task.updated",
      workbookId: task.workbookId,
      agentId,
      taskId,
      message: `Task claimed: ${task.goal}`
    });
    return { ok: true, task };
  }

  updateTask(
    taskId: TaskId,
    patch: Partial<
      Pick<
        TaskRecord,
        "goal" | "role" | "priority" | "status" | "progress" | "currentStep" | "blockers" | "assignedAgentId" | "allowedScopes" | "dependencies" | "errorMessage"
      >
    >
  ) {
    const task = this.tasks.update(taskId, patch);
    this.recordCollabEvent({
      type: patch.status === "completed" ? "task.completed" : patch.status === "failed" ? "task.failed" : "task.updated",
      workbookId: task.workbookId,
      agentId: task.assignedAgentId,
      taskId,
      message: `Task ${task.status}: ${task.goal}`
    });
    return { ok: true, task };
  }

  completeTask(taskId: TaskId) {
    return this.updateTask(taskId, { status: "completed", progress: 100, currentStep: "Completed" });
  }

  failTask(taskId: TaskId, errorMessage?: string | undefined) {
    return this.updateTask(taskId, { status: "failed", errorMessage: errorMessage ?? "Task failed." });
  }

  cancelTask(taskId: TaskId) {
    return this.updateTask(taskId, { status: "cancelled", currentStep: "Cancelled" });
  }

  setTaskProgress(taskId: TaskId, progress: number, currentStep?: string | undefined) {
    const task = this.tasks.setProgress(taskId, progress, currentStep);
    this.recordCollabEvent({
      type: "task.updated",
      workbookId: task.workbookId,
      agentId: task.assignedAgentId,
      taskId,
      message: `Task progress ${task.progress}%: ${task.goal}`,
      details: { progress: task.progress, currentStep: task.currentStep }
    });
    return { ok: true, task };
  }

  addTaskBlocker(taskId: TaskId, input: Pick<TaskBlocker, "message" | "severity"> & { scope?: WorkbookScope | undefined }) {
    const task = this.tasks.addBlocker(taskId, input);
    const blocker = task.blockers.at(-1);
    this.recordCollabEvent({
      type: "task.updated",
      workbookId: task.workbookId,
      agentId: task.assignedAgentId,
      taskId,
      message: `${input.severity === "blocked" ? "Task blocked" : "Task note added"}: ${task.goal}`,
      details: { blocker }
    });
    return { ok: true, task, blocker };
  }

  resolveTaskBlocker(taskId: TaskId, blockerId: string) {
    const task = this.tasks.resolveBlocker(taskId, blockerId);
    this.recordCollabEvent({
      type: "task.updated",
      workbookId: task.workbookId,
      agentId: task.assignedAgentId,
      taskId,
      message: `Task blocker resolved: ${task.goal}`,
      details: { blockerId }
    });
    return { ok: true, task };
  }

  evaluateTaskSchedule(input: { workbookId?: WorkbookId | undefined; apply?: boolean | undefined; lockMode?: LockMode | undefined } = {}): TaskScheduleResponse {
    const tasks = this.tasks.list(input.workbookId);
    const decisions = tasks.map((task) => {
      const activeBlockers = task.blockers.filter((blocker) => blocker.status === "open" && blocker.severity === "blocked");
      const waitingForTaskIds = task.dependencies.filter((dependencyId) => this.tasks.get(dependencyId)?.status !== "completed");
      const lockConflicts =
        task.allowedScopes.length > 0
          ? this.locks.findConflicts(task.workbookId, task.allowedScopes, input.lockMode ?? "write_values")
          : [];
      const done = task.status === "completed" || task.status === "failed" || task.status === "cancelled";
      const state: "ready" | "waiting_dependencies" | "waiting_locks" | "blocked" | "done" = done
        ? "done"
        : activeBlockers.length > 0
          ? "blocked"
          : waitingForTaskIds.length > 0
            ? "waiting_dependencies"
            : lockConflicts.length > 0
              ? "waiting_locks"
              : "ready";
      return {
        taskId: task.taskId,
        workbookId: task.workbookId,
        ready: state === "ready",
        state,
        waitingForTaskIds,
        lockConflicts,
        blockers: activeBlockers,
        nextRetryAt: nextRetryAt(lockConflicts),
        suggestedAction: taskScheduleAction(task.status, state),
        message: taskScheduleMessage(task.goal, state, waitingForTaskIds.length, lockConflicts.length, activeBlockers.length)
      };
    });

    const updatedTasks: TaskRecord[] = [];
    if (input.apply) {
      for (const decision of decisions) {
        const task = this.tasks.get(decision.taskId);
        if (!task || decision.state === "done") {
          continue;
        }
        if (decision.ready && task.status === "blocked") {
          updatedTasks.push(this.tasks.update(task.taskId, {
            status: task.assignedAgentId ? "claimed" : "open",
            currentStep: "Ready to resume"
          }));
          continue;
        }
        if (!decision.ready && task.status !== "blocked") {
          updatedTasks.push(this.tasks.update(task.taskId, {
            status: "blocked",
            currentStep: decision.message
          }));
        }
      }
      if (updatedTasks.length > 0) {
        this.recordCollabEvent({
          type: "task.updated",
          workbookId: input.workbookId,
          message: `Task schedule evaluated; ${updatedTasks.length} task(s) updated.`,
          details: { decisions }
        });
      } else {
        this.persistState();
      }
    }

    return {
      ok: true,
      workbookId: input.workbookId,
      applied: Boolean(input.apply),
      decisions,
      updatedTasks
    };
  }

  resumeReadyTasks(workbookId?: WorkbookId | undefined): TaskScheduleResponse {
    return this.evaluateTaskSchedule({ workbookId, apply: true });
  }

  getTask(taskId: TaskId) {
    const task = this.tasks.get(taskId);
    return task ? { ok: true, task } : { ok: false, error: runtimeError("NOT_FOUND", `Task not found: ${taskId}`, { retryable: false }) };
  }

  listTasks(workbookId?: WorkbookId) {
    return { ok: true, tasks: this.tasks.list(workbookId) };
  }

  getCollaborationStatus(workbookId?: WorkbookId) {
    return {
      ok: true,
      agents: this.listAgents().agents,
      tasks: this.tasks.list(workbookId),
      locks: this.locks.list(workbookId),
      transactions: this.transactions.list(workbookId),
      conflicts: this.conflicts.filter((conflict) => workbookId === undefined || conflict.workbookId === workbookId).slice(-50).reverse().map(attachConflictGuidance),
      events: this.collabEvents.filter((event) => workbookId === undefined || event.workbookId === workbookId).slice(-50).reverse()
    };
  }

  listAgents() {
    return { ok: true, agents: [...this.agents.values()].sort((a, b) => b.lastSeenAt.localeCompare(a.lastSeenAt)) };
  }

  listLocks(workbookId?: WorkbookId) {
    return { ok: true, locks: this.locks.list(workbookId) };
  }

  getLockPolicy() {
    return { ok: true, policy: this.lockLeasePolicy };
  }

  setLockPolicy(update: Partial<LockLeasePolicy>) {
    this.lockLeasePolicy = normalizeLockLeasePolicy({ ...this.lockLeasePolicy, ...update });
    this.recordCollabEvent({
      type: "lock.policy_updated",
      message: "Lock lease policy updated.",
      details: { policy: this.lockLeasePolicy }
    });
    return { ok: true, policy: this.lockLeasePolicy };
  }

  acquireLocks(input: {
    workbookId: WorkbookId;
    scopes: WorkbookScope[];
    mode: LockMode;
    reason: string;
    ownerAgentId?: AgentId | undefined;
    taskId?: TaskId | undefined;
    ttlMs?: number | undefined;
  }): LockAcquireResponse {
    if (!this.lockLeasePolicy.allowManualLocks) {
      return {
        ok: false,
        locks: [],
        conflicts: [
          {
            conflictId: makeId<string>("conflict"),
            code: "MANUAL_LOCKS_DISABLED",
            message: "Manual lock acquisition is disabled by runtime lock policy.",
            workbookId: input.workbookId,
            scopes: input.scopes,
            retryable: false,
            suggestedAction: "manual_review",
            createdAt: new Date().toISOString()
          }
        ],
        policy: this.lockLeasePolicy
      };
    }
    const lockResult = this.locks.acquire({
      workbookId: input.workbookId,
      ownerAgentId: input.ownerAgentId,
      taskId: input.taskId,
      scopes: input.scopes,
      mode: input.mode,
      ttlMs: lockTtl(input.ttlMs ?? this.lockLeasePolicy.defaultTtlMs, this.lockLeasePolicy),
      reason: input.reason
    });
    if (!lockResult.ok) {
      this.conflicts.push(...lockResult.conflicts);
      for (const conflict of lockResult.conflicts) {
        this.recordConflictTelemetry(conflict);
        this.recordCollabEvent({
          type: "conflict.detected",
          workbookId: input.workbookId,
          agentId: input.ownerAgentId,
          taskId: input.taskId,
          message: conflict.message,
          details: { conflict }
        });
      }
      return { ok: false, locks: [], conflicts: lockResult.conflicts, policy: this.lockLeasePolicy };
    }
    for (const lock of lockResult.locks) {
      this.recordCollabEvent({
        type: "lock.acquired",
        workbookId: input.workbookId,
        agentId: input.ownerAgentId,
        taskId: input.taskId,
        lockId: lock.lockId,
        message: `Lock acquired: ${input.reason}`,
        details: { lock }
      });
    }
    return { ok: true, locks: lockResult.locks, conflicts: [], policy: this.lockLeasePolicy };
  }

  renewLocks(lockIds: LockId[], ttlMs?: number | undefined): LockRenewResponse {
    const result = this.locks.renewWithMissing(lockIds, lockTtl(ttlMs ?? this.lockLeasePolicy.defaultTtlMs, this.lockLeasePolicy));
    for (const lock of result.renewed) {
      this.recordCollabEvent({
        type: "lock.acquired",
        workbookId: lock.workbookId,
        agentId: lock.ownerAgentId,
        taskId: lock.taskId,
        lockId: lock.lockId,
        message: `Lock renewed: ${lock.reason}`,
        details: { lock }
      });
    }
    if (result.renewed.length === 0) {
      this.persistState();
    }
    return { ok: result.missingLockIds.length === 0, renewed: result.renewed, missingLockIds: result.missingLockIds, policy: this.lockLeasePolicy };
  }

  releaseLocks(lockIds: LockId[]): LockReleaseResponse {
    const result = this.locks.releaseWithMissing(lockIds);
    this.markConflictTelemetryClearedByLock(result.released.map((lock) => lock.lockId));
    for (const lock of result.released) {
      this.recordCollabEvent({
        type: "lock.released",
        workbookId: lock.workbookId,
        agentId: lock.ownerAgentId,
        taskId: lock.taskId,
        lockId: lock.lockId,
        message: `Lock released: ${lock.reason}`,
        details: { lock }
      });
    }
    if (result.released.length === 0) {
      this.persistState();
    }
    return { ok: result.missingLockIds.length === 0, released: result.released, missingLockIds: result.missingLockIds };
  }

  listTransactions(workbookId?: WorkbookId) {
    return {
      ok: true,
      transactions: this.transactions.list(workbookId).map((transaction) => this.transactions.withQueueMetadata(transaction))
    };
  }

  getTransaction(transactionId: TransactionId) {
    const transaction = this.transactions.get(transactionId);
    return transaction
      ? { ok: true, transaction: this.transactions.withQueueMetadata(transaction) }
      : { ok: false, error: runtimeError("NOT_FOUND", `Transaction not found: ${transactionId}`, { retryable: false }) };
  }

  async waitTransaction(transactionId: TransactionId, timeoutMs = 30_000, pollMs = 250) {
    const started = Date.now();
    while (true) {
      const transaction = this.transactions.get(transactionId);
      if (!transaction) {
        return { ok: false, completed: false, error: runtimeError("NOT_FOUND", `Transaction not found: ${transactionId}`, { retryable: false }) };
      }
      const withMetadata = this.transactions.withQueueMetadata(transaction);
      if (isTerminalTransactionStatus(withMetadata.status)) {
        return { ok: true, completed: true, transaction: withMetadata };
      }
      if (Date.now() - started >= timeoutMs) {
        return {
          ok: true,
          completed: false,
          transaction: withMetadata,
          error: runtimeError("TIMEOUT", `Timed out waiting for transaction ${transactionId}.`, { retryable: true })
        };
      }
      await sleep(Math.max(25, pollMs));
    }
  }

  cancelTransaction(transactionId: TransactionId) {
    const transaction = this.transactions.get(transactionId);
    if (!transaction) {
      return { ok: false, error: runtimeError("NOT_FOUND", `Transaction not found: ${transactionId}`, { retryable: false }) };
    }
    if (transaction.status !== "queued") {
      return {
        ok: false,
        transaction: this.transactions.withQueueMetadata(transaction),
        error: runtimeError("OPERATION_FAILED", `Only queued transactions can be cancelled. Current status: ${transaction.status}.`, { retryable: false })
      };
    }
    this.cancelledQueuedTransactions.add(transactionId);
    const cancelled = this.transactions.markCancelled(transactionId);
    this.recordCollabEvent({
      type: "transaction.cancelled",
      workbookId: cancelled.workbookId,
      agentId: cancelled.agentId,
      taskId: cancelled.taskId,
      transactionId,
      message: cancelled.progressMessage ?? "Queued transaction cancelled.",
      details: { transaction: cancelled }
    });
    if (cancelled.taskId !== undefined) {
      this.updateTask(cancelled.taskId, { status: "cancelled", errorMessage: cancelled.errorMessage });
    } else {
      this.persistState();
    }
    return { ok: true, transaction: this.transactions.withQueueMetadata(cancelled) };
  }

  previewTransactionRollback(transactionId: TransactionId): TransactionRollbackPreview {
    const transaction = this.transactions.get(transactionId);
    if (!transaction) {
      return {
        ok: false,
        transactionId,
        rollbackAvailable: false,
        scopes: [],
        laterTransactions: [],
        conflicts: [
          {
            code: "ROLLBACK_UNAVAILABLE",
            message: `Transaction not found: ${transactionId}`,
            transactionId,
            scopes: [],
            suggestedAction: "manual_review"
          }
        ],
        warnings: []
      };
    }

    const laterTransactions = this.transactions.getLaterApplied(transaction);
    const conflicts: TransactionRollbackConflict[] = [];
    for (const later of laterTransactions) {
      for (const scope of transaction.scopes) {
        for (const laterScope of later.scopes) {
          const conflict = makeRollbackConflict({
            transactionId,
            conflictingTransactionId: later.transactionId,
            left: scope,
            right: laterScope
          });
          if (conflict) {
            conflicts.push(conflict);
          }
        }
      }
    }

    if (transaction.status !== "applied" && transaction.status !== "rolled_back") {
      conflicts.push({
        code: "ROLLBACK_UNAVAILABLE",
        message: `Only applied transactions can be rolled back. Current status is ${transaction.status}.`,
        transactionId,
        scopes: transaction.scopes,
        suggestedAction: "manual_review"
      });
    }

    if (transaction.status === "rolled_back") {
      conflicts.push({
        code: "ROLLBACK_UNAVAILABLE",
        message: "Transaction has already been rolled back.",
        transactionId,
        scopes: transaction.scopes,
        suggestedAction: "manual_review"
      });
    }

    if (!transaction.planId) {
      conflicts.push({
        code: "ROLLBACK_UNAVAILABLE",
        message: "Transaction has no plan rollback metadata. Use backup repair or workbook restore instead.",
        transactionId,
        scopes: transaction.scopes,
        suggestedAction: "repair_from_backup"
      });
    }

    const preview: TransactionRollbackPreview = {
      ok: conflicts.length === 0,
      transactionId,
      workbookId: transaction.workbookId,
      planId: transaction.planId,
      taskId: transaction.taskId,
      rollbackAvailable: conflicts.length === 0 && transaction.planId !== undefined,
      rollbackMethod: transaction.planId !== undefined ? "plan" : undefined,
      scopes: transaction.scopes,
      laterTransactions,
      conflicts,
      warnings: transaction.warnings
    };
    this.recordCollabEvent({
      type: "transaction.rollback_previewed",
      workbookId: transaction.workbookId,
      agentId: transaction.agentId,
      taskId: transaction.taskId,
      transactionId,
      message: preview.ok ? `Rollback preview ready for ${transactionId}.` : `Rollback blocked for ${transactionId}.`,
      details: { preview }
    });
    return preview;
  }

  async rollbackTransaction(transactionId: TransactionId, confirmationToken?: string): Promise<OperationResult> {
    const preview = this.previewTransactionRollback(transactionId);
    const transaction = this.transactions.get(transactionId);
    if (transaction && !transaction.planId && transaction.backups.length > 0) {
      const failedWarnings: OperationWarning[] = [];
      let lastResult: OperationResult | undefined;
      for (const backupId of transaction.backups) {
        const result = await this.restoreBackup(backupId, confirmationToken);
        lastResult = result;
        if (!result.ok) {
          failedWarnings.push({
            code: "BACKUP_RESTORE_SKIPPED",
            message: result.error?.message ?? `Backup restore failed for ${backupId}.`,
            details: { backupId, error: result.error }
          });
          continue;
        }
        this.transactions.markRolledBack(transactionId);
        this.recordCollabEvent({
          type: "transaction.rolled_back",
          workbookId: transaction.workbookId,
          agentId: transaction.agentId,
          taskId: transaction.taskId,
          transactionId,
          message: `Transaction rolled back from backup: ${transaction.goal}`
        });
        this.persistState();
        return {
          ...result,
          transactionId,
          taskId: transaction.taskId,
          agentId: transaction.agentId,
          warnings: [...(result.warnings ?? []), ...failedWarnings]
        };
      }
      return {
        ...(lastResult ?? {
          ok: false,
          rollbackAvailable: false,
          backups: [],
          warnings: [],
          telemetry: {},
          error: {
            code: "BACKUP_UNAVAILABLE",
            message: "No transaction backups were restorable.",
            severity: "error",
            retryable: false
          }
        }),
        transactionId,
        taskId: transaction.taskId,
        agentId: transaction.agentId,
        warnings: [...(lastResult?.warnings ?? []), ...failedWarnings]
      };
    }
    if (!preview.ok || !transaction?.planId) {
      return {
        ok: false,
        transactionId,
        planId: transaction?.planId,
        taskId: transaction?.taskId,
        agentId: transaction?.agentId,
        rollbackAvailable: false,
        backups: [],
        warnings: preview.conflicts.map((conflict) => ({
          code: conflict.code,
          message: conflict.message,
          details: { conflict }
        })),
        telemetry: { warningCount: preview.conflicts.length },
        error: runtimeError("BACKUP_UNAVAILABLE", "Transaction rollback is blocked. Review rollback preview conflicts.", {
          retryable: false,
          details: { preview }
        })
      };
    }

    const result = await this.rollbackPlan(transaction.planId, confirmationToken);
    if (result.ok) {
      this.transactions.markRolledBack(transactionId);
      this.recordCollabEvent({
        type: "transaction.rolled_back",
        workbookId: transaction.workbookId,
        agentId: transaction.agentId,
        taskId: transaction.taskId,
        transactionId,
        message: `Transaction rolled back: ${transaction.goal}`
      });
      this.persistState();
    }
    return {
      ...result,
      transactionId,
      taskId: transaction.taskId,
      agentId: transaction.agentId
    };
  }

  previewTransactionRollbackChain(transactionId: TransactionId): TransactionRollbackChainPreview {
    const transaction = this.transactions.get(transactionId);
    if (!transaction) {
      return {
        ok: false,
        rootTransactionId: transactionId,
        rollbackAvailable: false,
        rollbackOrder: [],
        affectedTransactions: [],
        conflicts: [
          {
            code: "ROLLBACK_UNAVAILABLE",
            message: `Transaction not found: ${transactionId}`,
            transactionId,
            scopes: [],
            suggestedAction: "manual_review"
          }
        ],
        warnings: [],
        requiresConfirmation: false
      };
    }

    const affected = this.collectRollbackChain(transaction);
    const rollbackOrder = [...affected].reverse();
    const conflicts: TransactionRollbackConflict[] = [];
    for (const affectedTransaction of affected) {
      if (affectedTransaction.status !== "applied") {
        conflicts.push({
          code: "ROLLBACK_UNAVAILABLE",
          message: `Only applied transactions can be included in a rollback chain. ${affectedTransaction.transactionId} is ${affectedTransaction.status}.`,
          transactionId: affectedTransaction.transactionId,
          scopes: affectedTransaction.scopes,
          suggestedAction: "manual_review"
        });
      }
      if (!affectedTransaction.planId) {
        conflicts.push({
          code: "ROLLBACK_UNAVAILABLE",
          message: `Transaction ${affectedTransaction.transactionId} has no plan rollback metadata.`,
          transactionId: affectedTransaction.transactionId,
          scopes: affectedTransaction.scopes,
          suggestedAction: "repair_from_backup"
        });
      }
    }

    const warnings: OperationWarning[] =
      affected.length > 1
        ? [
            {
              code: "ROLLBACK_CHAIN_REQUIRED",
              message: `Rollback requires ${affected.length} related transactions to be rolled back newest-first.`,
              details: { transactionIds: rollbackOrder.map((candidate) => candidate.transactionId) }
            }
          ]
        : [];
    const confirmationToken = affected.length > 1 ? this.rollbackChainConfirmationToken(transactionId, rollbackOrder) : undefined;
    const preview: TransactionRollbackChainPreview = {
      ok: conflicts.length === 0,
      rootTransactionId: transactionId,
      workbookId: transaction.workbookId,
      rollbackAvailable: conflicts.length === 0,
      rollbackOrder,
      affectedTransactions: affected,
      conflicts,
      warnings,
      requiresConfirmation: affected.length > 1
    };
    if (confirmationToken !== undefined) {
      preview.confirmationToken = confirmationToken;
    }
    this.recordCollabEvent({
      type: "transaction.rollback_previewed",
      workbookId: transaction.workbookId,
      agentId: transaction.agentId,
      taskId: transaction.taskId,
      transactionId,
      message: preview.ok ? `Rollback chain preview ready for ${transactionId}.` : `Rollback chain blocked for ${transactionId}.`,
      details: { preview }
    });
    return preview;
  }

  async rollbackTransactionChain(transactionId: TransactionId, confirmationToken?: string): Promise<OperationResult> {
    const preview = this.previewTransactionRollbackChain(transactionId);
    if (!preview.ok) {
      return {
        ok: false,
        transactionId,
        rollbackAvailable: false,
        backups: [],
        warnings: preview.conflicts.map((conflict) => ({
          code: conflict.code,
          message: conflict.message,
          details: { conflict }
        })),
        telemetry: { warningCount: preview.conflicts.length },
        error: runtimeError("BACKUP_UNAVAILABLE", "Transaction rollback chain is blocked. Review rollback preview conflicts.", {
          retryable: false,
          details: { preview }
        })
      };
    }
    if (preview.requiresConfirmation && confirmationToken !== preview.confirmationToken) {
      return {
        ok: false,
        transactionId,
        rollbackAvailable: true,
        backups: [],
        warnings: [
          {
            code: "CONFIRMATION_REQUIRED",
            message: "Rollback chain requires the exact confirmation token from preview.",
            details: { confirmationToken: preview.confirmationToken }
          }
        ],
        telemetry: { warningCount: 1 },
        error: runtimeError("CONFIRMATION_REQUIRED", "Rollback chain requires explicit confirmation.", {
          retryable: true,
          details: { preview }
        })
      };
    }

    const results: OperationResult[] = [];
    for (const transaction of preview.rollbackOrder) {
      const result = await this.rollbackTransaction(transaction.transactionId, confirmationToken);
      results.push(result);
      if (!result.ok) {
        return {
          ...result,
          transactionId,
          warnings: [
            ...result.warnings,
            {
              code: "ROLLBACK_CHAIN_PARTIAL",
              message: `Rollback chain stopped at ${transaction.transactionId}.`,
              details: { results }
            }
          ]
        };
      }
    }
    return {
      ok: true,
      transactionId,
      rollbackAvailable: false,
      backups: results.flatMap((result) => result.backups),
      warnings: preview.warnings,
      telemetry: {
        warningCount: preview.warnings.length,
        chunkCount: results.length
      },
      data: {
        rollbackChain: preview.rollbackOrder.map((transaction) => transaction.transactionId),
        results
      }
    };
  }

  private collectRollbackChain(root: TransactionRecord): TransactionRecord[] {
    const affected: TransactionRecord[] = [root];
    const laterTransactions = this.transactions.getLaterApplied(root);
    let changed = true;
    while (changed) {
      changed = false;
      for (const later of laterTransactions) {
        if (affected.some((candidate) => candidate.transactionId === later.transactionId)) {
          continue;
        }
        const conflictsAffected = affected.some((candidate) =>
          candidate.scopes.some((scope) =>
            later.scopes.some((laterScope) =>
              makeRollbackConflict({
                transactionId: candidate.transactionId,
                conflictingTransactionId: later.transactionId,
                left: scope,
                right: laterScope
              })
            )
          )
        );
        if (conflictsAffected) {
          affected.push(later);
          changed = true;
        }
      }
    }
    return affected;
  }

  private rollbackChainConfirmationToken(rootTransactionId: TransactionId, rollbackOrder: TransactionRecord[]): string {
    const chainHash = hashStable(rollbackOrder.map((transaction) => transaction.transactionId));
    return `rollback-chain:${rootTransactionId}:${chainHash}`;
  }

  listConflicts(workbookId?: WorkbookId) {
    return {
      ok: true,
      conflicts: this.conflicts.filter((conflict) => workbookId === undefined || conflict.workbookId === workbookId).slice(-100).reverse().map(attachConflictGuidance)
    };
  }

  getConflictGuidance(workbookId?: WorkbookId): ConflictGuidanceResponse {
    const conflicts = this.conflicts.filter((conflict) => workbookId === undefined || conflict.workbookId === workbookId).slice(-100).reverse();
    return {
      ok: true,
      workbookId,
      guidance: conflicts.map((conflict) => attachConflictGuidance(conflict).guidance).filter((guidance): guidance is NonNullable<ConflictRecord["guidance"]> => guidance !== undefined)
    };
  }

  explainConflict(conflict: ConflictRecord) {
    return {
      ok: true,
      conflict: attachConflictGuidance(conflict),
      guidance: attachConflictGuidance(conflict).guidance
    };
  }

  getConflictTelemetry(workbookId?: WorkbookId, windowSize = 250): ConflictTelemetrySummary {
    const recent = this.conflictTelemetry
      .filter((record) => workbookId === undefined || record.workbookId === workbookId)
      .slice(-Math.max(1, Math.min(1_000, Math.round(windowSize))));
    return {
      ok: true,
      workbookId,
      windowSize: recent.length,
      totalCount: recent.length,
      openCount: recent.filter((record) => record.status === "open").length,
      clearedCount: recent.filter((record) => record.status === "cleared").length,
      byCode: telemetryBuckets(recent, (record) => [record.code]),
      byPrimaryAction: telemetryBuckets(recent, (record) => [record.primaryAction]),
      hotScopes: telemetryBuckets(recent, (record) => record.scopeKeys),
      hotTasks: telemetryBuckets(recent, (record) => record.taskId ? [record.taskId] : []),
      hotAgents: telemetryBuckets(recent, (record) => record.ownerAgentId ? [record.ownerAgentId] : []),
      recent: [...recent].reverse()
    };
  }

  clearConflictTelemetry(workbookId?: WorkbookId) {
    const before = this.conflictTelemetry.length;
    if (workbookId === undefined) {
      this.conflictTelemetry.splice(0, this.conflictTelemetry.length);
    } else {
      for (let index = this.conflictTelemetry.length - 1; index >= 0; index -= 1) {
        if (this.conflictTelemetry[index]?.workbookId === workbookId) {
          this.conflictTelemetry.splice(index, 1);
        }
      }
    }
    const cleared = before - this.conflictTelemetry.length;
    this.persistState();
    return { ok: true, workbookId, cleared };
  }

  private restoreState(): void {
    const snapshot = this.stateStore?.load();
    if (!snapshot) {
      return;
    }
    this.agents.clear();
    for (const agent of snapshot.agents) {
      this.agents.set(agent.agentId, { ...agent });
    }
    this.tasks.load(snapshot.tasks);
    this.locks.load(snapshot.locks);
    if (snapshot.lockLeasePolicy !== undefined) {
      this.lockLeasePolicy = normalizeLockLeasePolicy(snapshot.lockLeasePolicy);
    }
    this.transactions.load(snapshot.transactions);
    this.jobs.clear();
    for (const job of snapshot.jobs ?? []) {
      this.jobs.set(job.jobId, { ...job, transactionIds: [...job.transactionIds], warnings: [...job.warnings] });
    }
    this.conflicts.splice(0, this.conflicts.length, ...snapshot.conflicts.slice(-250));
    this.conflictTelemetry.splice(0, this.conflictTelemetry.length, ...(snapshot.conflictTelemetry ?? []).slice(-1_000));
    this.collabEvents.splice(0, this.collabEvents.length, ...snapshot.collaborationEvents.slice(-1_000));
    this.templates.load(snapshot.templates ?? []);
    this.regions.clear();
    for (const region of snapshot.regions ?? []) {
      this.regions.set(regionKey(region.workbookId, region.name), { ...region });
    }
    if (snapshot.permissions !== undefined) {
      this.permissionState = mergePermissionState(this.permissionState, snapshot.permissions);
    }
    this.plans.load(snapshot.plans ?? []);
    this.backups.load(snapshot.backups ?? []);
    this.agent.loadOperations(snapshot.agentOperations ?? []);
  }

  private recoverRuntimeState(): void {
    const now = new Date().toISOString();
    let recovered = false;
    for (const agent of this.agents.values()) {
      if (agent.status === "active" || agent.status === "idle") {
        agent.status = "disconnected";
        agent.lastSeenAt = now;
        recovered = true;
      }
    }
    const expiredLocks = this.locks.expireActive("Daemon restarted before the lock was released.");
    if (expiredLocks.length > 0) {
      recovered = true;
    }
    const interruptedTransactions = this.transactions.markInterrupted();
    if (interruptedTransactions.length > 0) {
      recovered = true;
      for (const transaction of interruptedTransactions) {
        if (transaction.taskId !== undefined) {
          const task = this.tasks.get(transaction.taskId);
          if (task && (task.status === "queued" || task.status === "applying")) {
            this.tasks.update(transaction.taskId, {
              status: "failed",
              errorMessage: "Daemon restarted before the task transaction finished."
            });
          }
        }
      }
    }
    if (recovered) {
      this.recordCollabEvent({
        type: "transaction.failed",
        message: "Recovered daemon state after restart.",
        details: {
          expiredLocks: expiredLocks.length,
          interruptedTransactions: interruptedTransactions.length
        }
      });
      return;
    }
    this.persistState();
  }

  private persistState(): void {
    this.stateStore?.save({
      version: 1,
      savedAt: new Date().toISOString(),
      agents: [...this.agents.values()],
      tasks: this.tasks.dump(),
      locks: this.locks.dump(),
      lockLeasePolicy: this.lockLeasePolicy,
      transactions: this.transactions.dump(),
      jobs: [...this.jobs.values()].map((job) => ({ ...job, transactionIds: [...job.transactionIds], warnings: [...job.warnings] })),
      conflicts: this.conflicts.slice(-250),
      conflictTelemetry: this.conflictTelemetry.slice(-1_000),
      collaborationEvents: this.collabEvents.slice(-1_000),
      templates: this.templates.dump(),
      regions: [...this.regions.values()].map((region) => ({ ...region })),
      permissions: this.permissionState,
      plans: this.plans.dump(),
      backups: this.backups.dump(),
      agentOperations: this.agent.dumpOperations()
    });
  }

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
    const workbookId = eventWorkbookId(params) ?? this.sessions.getActive()?.activeWorkbook?.workbookId;
    if (workbookId && /change|changed|mutation|write|clear|format|style|table|sheet|range/i.test(method)) {
      this.bumpWorkbookContentVersion(workbookId);
    }
  }

  getWorkbookContentVersion(workbookId: WorkbookId | string): number {
    return this.workbookContentVersions.get(String(workbookId)) ?? 0;
  }

  private bumpWorkbookContentVersion(workbookId: WorkbookId | string): number {
    const key = String(workbookId);
    const next = (this.workbookContentVersions.get(key) ?? 0) + 1;
    this.workbookContentVersions.set(key, next);
    return next;
  }

  private recordCollabEvent(input: Omit<CollaborationEvent, "eventId" | "createdAt">): void {
    const event: CollaborationEvent = {
      ...input,
      eventId: makeId<string>("collab_event"),
      createdAt: new Date().toISOString()
    };
    this.collabEvents.push(event);
    if (this.collabEvents.length > 1_000) {
      this.collabEvents.splice(0, this.collabEvents.length - 1_000);
    }
    this.persistState();
  }

  private recordConflictTelemetry(conflict: ConflictRecord): void {
    const guided = attachConflictGuidance(conflict);
    const telemetry: ConflictTelemetryRecord = {
      telemetryId: makeId<string>("conflict_telemetry"),
      conflictId: conflict.conflictId,
      code: conflict.code,
      workbookId: conflict.workbookId,
      scopes: conflict.scopes,
      scopeKeys: conflict.scopes.map(scopeTelemetryKey),
      primaryAction: guided.guidance?.primaryAction ?? conflict.suggestedAction,
      retryable: conflict.retryable,
      status: "open",
      createdAt: conflict.createdAt
    };
    if (conflict.ownerAgentId !== undefined) {
      telemetry.ownerAgentId = conflict.ownerAgentId;
    }
    if (conflict.taskId !== undefined) {
      telemetry.taskId = conflict.taskId;
    }
    if (conflict.transactionId !== undefined) {
      telemetry.transactionId = conflict.transactionId;
    }
    if (conflict.lockId !== undefined) {
      telemetry.lockId = conflict.lockId;
    }
    this.conflictTelemetry.push(telemetry);
    if (this.conflictTelemetry.length > 1_000) {
      this.conflictTelemetry.splice(0, this.conflictTelemetry.length - 1_000);
    }
  }

  private markConflictTelemetryClearedByLock(lockIds: LockId[]): void {
    if (lockIds.length === 0) {
      return;
    }
    const lockIdSet = new Set(lockIds);
    const clearedAt = new Date().toISOString();
    let changed = false;
    for (const record of this.conflictTelemetry) {
      if (record.status === "open" && record.lockId !== undefined && lockIdSet.has(record.lockId)) {
        record.status = "cleared";
        record.clearedAt = clearedAt;
        record.clearReason = "lock_released";
        changed = true;
      }
    }
    if (changed) {
      this.persistState();
    }
  }

  getStatus() {
    const activeSession = this.sessions.getActive();
    const connectionState = this.connectionStateFor(activeSession);
    const readySession = connectionState === "ready" || connectionState === "connected_no_workbook" ? activeSession : undefined;
    return {
      ok: true,
      runtime: {
        service: "open-workbook-backend",
        packageName: "@components-kit/open-workbook-backend",
        version: runtimeVersion(),
        pid: process.pid
      },
      connectionState,
      activeAddinConnected: Boolean(readySession),
      activeAddinReachable: Boolean(readySession),
      activeWorkbookAvailable: Boolean(readySession?.activeWorkbook),
      fileBridge: this.fileBridge.getStatus(),
      sessions: this.sessions.list().map((session) => ({
        ...session,
        ageMs: this.sessionAgeMs(session),
        stale: this.isSessionStale(session)
      })),
      activeWorkbook: readySession?.activeWorkbook
    };
  }

  async getConnectionReadiness() {
    const activeSession = this.sessions.getActive();
    if (!activeSession) {
      return {
        ok: false,
        connectionState: "disconnected" as const,
        status: this.getStatus(),
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    if (this.isSessionStale(activeSession)) {
      return {
        ok: false,
        connectionState: "stale" as const,
        status: this.getStatus(),
        error: runtimeError("ADDIN_STALE", "The Excel add-in session is stale. Reload the OpenWorkbook Local taskpane, then retry.", { retryable: true })
      };
    }
    const client = this.addinClients.get(activeSession.connectionId);
    if (!client) {
      this.sessions.remove(activeSession.connectionId);
      return {
        ok: false,
        connectionState: "disconnected" as const,
        status: this.getStatus(),
        error: runtimeError("ADDIN_DISCONNECTED", "No active Excel add-in client is available.", { retryable: true })
      };
    }
    try {
      const activeWorkbook = await client.request<WorkbookRef | undefined>("runtime.get_active_context", undefined, { timeoutMs: addinHealthTimeoutMs() });
      if (activeWorkbook) {
        this.sessions.update(activeSession.connectionId, { activeWorkbook });
        return { ok: true, connectionState: "ready" as const, activeWorkbook, status: this.getStatus() };
      }
      return {
        ok: false,
        connectionState: "connected_no_workbook" as const,
        status: this.getStatus(),
        error: runtimeError("NO_ACTIVE_WORKBOOK", "Open Workbook is connected to Excel, but there is no active workbook.", { retryable: true })
      };
    } catch (error) {
      this.addinClients.get(activeSession.connectionId)?.close();
      this.detachAddinClient(activeSession.connectionId);
      this.sessions.remove(activeSession.connectionId);
      return {
        ok: false,
        connectionState: "stale" as const,
        status: this.getStatus(),
        error: runtimeError("ADDIN_STALE", `The Excel add-in session stopped responding: ${error instanceof Error ? error.message : String(error)}`, { retryable: true })
      };
    }
  }

  async runAgent(input: AgentRunInput, context?: AgentRunExecutionContext): Promise<AgentRunOutput> {
    const normalized = normalizeAgentExecutionContext(context);
    if (!normalized) {
      return this.agent.run(input);
    }
    this.registerAgent({
      agentId: normalized.agentId as AgentId,
      ...(normalized.agentName !== undefined ? { agentName: normalized.agentName } : {}),
      clientType: normalized.clientType
    });
    return this.agentExecutionContext.run(normalized, () => this.agent.run(input, normalized));
  }

  currentAgentExecutionContext(): AgentRunExecutionContext | undefined {
    return this.agentExecutionContext.getStore();
  }

  runWithAgentExecutionContext<T>(context: AgentRunExecutionContext | undefined, work: () => T): T {
    const normalized = normalizeAgentExecutionContext(context);
    return normalized ? this.agentExecutionContext.run(normalized, work) : work();
  }

  private currentAgentId(): AgentId {
    const agentId = this.agentExecutionContext.getStore()?.agentId;
    return typeof agentId === "string" && agentId.length > 0 ? agentId as AgentId : this.defaultAgentId;
  }

  getAgentContextResource(workbookContextId: string) {
    return this.agent.getContextResource(workbookContextId);
  }

  getAgentSemanticIndexResource(workbookContextId: string) {
    return this.agent.getSemanticIndexResource(workbookContextId);
  }

  getAgentOperationResource(operationId: string) {
    return this.agent.getOperationResource(operationId);
  }

  getAgentResultResource(resultId: string, options?: { view?: "summary" | "full"; maxBytes?: number }) {
    return this.agent.getResultResource(resultId, options);
  }

  getCompactResource(resourceId: string, options?: { view?: "summary" | "full"; maxBytes?: number }) {
    return this.agent.getCompactResource(resourceId, options);
  }

  async getStatusWithFileBridgeProbe() {
    const status = this.getStatus();
    return {
      ...status,
      fileBridge: await this.fileBridge.probeStatus()
    };
  }

  getCapabilities(options: { includePreview?: boolean } = {}) {
    const catalogOptions = options.includePreview === undefined ? {} : { includePreview: options.includePreview };
    const sessions = this.sessions.list();
    const activeSession = this.sessions.getActive();
    const readyActiveSession = activeSession && !this.isSessionStale(activeSession) ? activeSession : undefined;
    return {
      runtime: this.getStatus(),
      activeHostCapabilities: readyActiveSession?.capabilities ?? disconnectedRuntimeCapabilities(),
      connectedHostCapabilities: sessions
        .filter((session) => !this.isSessionStale(session))
        .filter((session) => session.capabilities !== undefined)
        .map((session) => ({
          connectionId: session.connectionId,
          connectedAt: session.connectedAt,
          lastSeenAt: session.lastSeenAt,
          activeWorkbook: session.activeWorkbook,
          capabilities: session.capabilities
        })),
      catalog: getPublicAgentToolCatalogSummary(catalogOptions),
      internalCapabilities: getInternalCapabilityCatalogSummary(catalogOptions),
      fileBridge: this.fileBridge.getStatus(),
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
    this.persistState();
    return this.getPermissions();
  }

  requireConfirmation(levels: PermissionState["requireConfirmationFor"]) {
    this.permissionState = {
      ...this.permissionState,
      requireConfirmationFor: [...new Set(levels)]
    };
    this.persistState();
    return this.getPermissions();
  }

  setPermissionScope(scope: PermissionState["scope"]) {
    this.permissionState = {
      ...this.permissionState,
      scope: { ...scope }
    };
    this.persistState();
    return this.getPermissions();
  }

  allowDestructiveActions(allow: boolean) {
    this.permissionState = {
      ...this.permissionState,
      allowDestructiveActions: allow
    };
    this.persistState();
    return this.getPermissions();
  }

  allowMacroExecution(allow: boolean) {
    this.permissionState = {
      ...this.permissionState,
      allowMacroExecution: allow
    };
    this.persistState();
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
    this.persistState();
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
    this.persistState();
    return this.getPermissions();
  }

  connectAddinInfo() {
    const status = this.getStatus();
    return {
      ok: true,
      backendUrl: `ws://${process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1"}:${process.env.OPEN_WORKBOOK_PORT ?? 37845}${
        process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin"
      }`,
      activeAddinConnected: status.activeAddinConnected,
      connectionState: status.connectionState,
      activeWorkbookAvailable: status.activeWorkbookAvailable
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
    const agentId = request.agentId ?? this.currentAgentId();
    this.registerAgent({ agentId, agentName: request.agentName, clientType: "mcp" });
    const plan = this.plans.createPlan({ ...request, agentId });
    if (request.taskId !== undefined) {
      this.tasks.attachPlan(request.taskId, plan.planId);
      this.updateTask(request.taskId, { status: "planning" });
    }
    this.persistState();
    return plan;
  }

  compileBatch(request: BatchRequest) {
    return this.compiler.compile(request);
  }

  preflightBatch(request: BatchRequest): BatchPreflightResult {
    const compiled = this.compiler.compile(request);
    const estimatedPayloadBytes = Buffer.byteLength(JSON.stringify(request.operations), "utf8");
    const chunkPlan = planBatchChunks(request.operations);
    const warnings: OperationWarning[] = [];
    const needsQueue =
      request.operations.length > batchDirectOperationThreshold() ||
      estimatedPayloadBytes > batchDirectPayloadThresholdBytes() ||
      compiled.estimatedCellsTouched > batchDirectCellThreshold() ||
      (chunkPlan.safeToAutoChunk && chunkPlan.chunksTotal > 1);
    let recommendedExecutionMode: BatchPreflightResult["recommendedExecutionMode"] = needsQueue ? "submit" : "apply";
    if (needsQueue && chunkPlan.safeToAutoChunk && chunkPlan.chunksTotal > 1) {
      recommendedExecutionMode = "chunked_submit";
      warnings.push({
        code: "LARGE_BATCH_WILL_BE_CHUNKED",
        message: `Batch is large and can be safely split into ${chunkPlan.chunksTotal} queued chunks.`,
        details: { strategy: chunkPlan.strategy, chunkSize: chunkPlan.chunkSize }
      });
    } else if (needsQueue) {
      warnings.push({
        code: "LARGE_BATCH_SHOULD_BE_QUEUED",
        message: "Batch is large enough that it should be submitted to the serialized queue instead of applied synchronously.",
        details: {
          operationThreshold: batchDirectOperationThreshold(),
          payloadThresholdBytes: batchDirectPayloadThresholdBytes(),
          cellThreshold: batchDirectCellThreshold()
        }
      });
    }
    return {
      ok: true,
      workbookId: request.workbookId,
      operationCount: request.operations.length,
      estimatedCellsTouched: compiled.estimatedCellsTouched,
      estimatedPayloadBytes,
      destructiveLevel: compiled.destructiveLevel,
      recommendedExecutionMode,
      safeToAutoChunk: chunkPlan.safeToAutoChunk,
      chunkPlan: chunkPlan.strategy === "none" ? undefined : chunkPlan,
      warnings
    };
  }

  submitChunkedBatch(request: BatchRequest, input: { goal?: string | undefined; retryStrategy?: string | undefined } = {}) {
    if (request.mode !== "apply") {
      return {
        ok: false,
        error: runtimeError("INVALID_ARGUMENT", "Only apply-mode batches can be submitted as chunked jobs.", { retryable: false })
      };
    }
    const preflight = this.preflightBatch(request);
    if (!preflight.safeToAutoChunk || preflight.chunkPlan === undefined || preflight.chunkPlan.chunksTotal <= 1) {
      const submitted = this.submitBatch(request);
      return {
        ok: submitted.ok,
        status: "queued",
        progressMessage: "Batch was not safely chunkable, so it was submitted as one queued transaction.",
        preflight,
        transactionIds: submitted.transactionId ? [submitted.transactionId] : [],
        transactions: [submitted]
      };
    }
    const chunks = chunkBatchOperations(request.operations);
    const now = new Date().toISOString();
    const compiled = this.compiler.compile(request);
    const agentId = request.agentId ?? this.currentAgentId();
    const job: JobRecord = {
      jobId: makeId<JobId>("job"),
      workbookId: request.workbookId,
      agentId,
      kind: preflight.chunkPlan.strategy === "split_style_entries" ? "style_chunked" : preflight.chunkPlan.strategy === "split_matrix_rows" ? "matrix_chunked" : "batch_chunked",
      status: "queued",
      goal: input.goal ?? request.progressMessage ?? "Apply chunked Excel batch",
      transactionIds: [],
      chunksTotal: chunks.length,
      chunksCompleted: 0,
      progressMessage: `Queued ${chunks.length} workbook update chunks.`,
      retryStrategy: input.retryStrategy ?? preflight.chunkPlan.strategy,
      destructiveLevel: compiled.destructiveLevel,
      warnings: preflight.warnings,
      queuedAt: now
    };
    if (request.taskId !== undefined) {
      job.taskId = request.taskId;
    }
    if (request.planId !== undefined) {
      job.planId = request.planId;
    }
    this.jobs.set(job.jobId, job);
    const transactions = chunks.map((chunk, index) =>
      this.submitBatch({
        ...request,
        operations: chunk,
        retryStrategy: job.retryStrategy,
        chunksTotal: chunks.length,
        chunksCompleted: index,
        progressMessage: `Queued workbook update chunk ${index + 1} of ${chunks.length}.`
      })
    );
    job.transactionIds = transactions.map((transaction: any) => transaction.transactionId).filter(Boolean);
    this.refreshJob(job.jobId);
    this.persistState();
    return {
      ok: true,
      status: "queued",
      job: this.getJobRecord(job.jobId),
      jobId: job.jobId,
      preflight,
      retryStrategy: job.retryStrategy,
      chunksTotal: chunks.length,
      chunksCompleted: 0,
      transactionIds: job.transactionIds,
      transactions,
      progressMessage: `Batch is large, so Open Workbook queued it as job ${job.jobId} with ${chunks.length} chunks. Use excel.job.wait or excel.job.get for progress.`
    };
  }

  listJobs(workbookId?: WorkbookId) {
    return {
      ok: true,
      jobs: [...this.jobs.values()]
        .filter((job) => workbookId === undefined || job.workbookId === workbookId)
        .map((job) => this.refreshJob(job.jobId))
        .sort((a, b) => (b.finishedAt ?? b.startedAt ?? b.queuedAt).localeCompare(a.finishedAt ?? a.startedAt ?? a.queuedAt))
    };
  }

  getJob(jobId: JobId) {
    const job = this.getJobRecord(jobId);
    return job
      ? { ok: true, job }
      : { ok: false, error: runtimeError("NOT_FOUND", `Job not found: ${jobId}`, { retryable: false }) };
  }

  async waitJob(jobId: JobId, timeoutMs = 30_000, pollMs = 250) {
    const started = Date.now();
    while (true) {
      const job = this.getJobRecord(jobId);
      if (!job) {
        return { ok: false, completed: false, error: runtimeError("NOT_FOUND", `Job not found: ${jobId}`, { retryable: false }) };
      }
      if (isTerminalJobStatus(job.status)) {
        return { ok: true, completed: true, job };
      }
      if (Date.now() - started >= timeoutMs) {
        return {
          ok: true,
          completed: false,
          job,
          error: runtimeError("TIMEOUT", `Timed out waiting for job ${jobId}.`, { retryable: true })
        };
      }
      await sleep(Math.max(25, pollMs));
    }
  }

  cancelJob(jobId: JobId) {
    const job = this.getJobRecord(jobId);
    if (!job) {
      return { ok: false, error: runtimeError("NOT_FOUND", `Job not found: ${jobId}`, { retryable: false }) };
    }
    const cancelled: TransactionRecord[] = [];
    const skipped: TransactionRecord[] = [];
    for (const transactionId of job.transactionIds) {
      const transaction = this.transactions.get(transactionId);
      if (!transaction) {
        continue;
      }
      if (transaction.status === "queued") {
        const result = this.cancelTransaction(transactionId);
        if (result.ok && result.transaction) {
          cancelled.push(result.transaction);
        }
      } else if (!isTerminalTransactionStatus(transaction.status)) {
        skipped.push(this.transactions.withQueueMetadata(transaction));
      }
    }
    const refreshed = this.refreshJob(jobId);
    this.persistState();
    return {
      ok: skipped.length === 0,
      job: refreshed,
      cancelledTransactions: cancelled,
      skippedTransactions: skipped,
      progressMessage:
        skipped.length === 0
          ? `Cancelled queued work for job ${jobId}.`
          : `Cancelled queued chunks for job ${jobId}; ${skipped.length} chunk(s) were already applying and could not be cancelled.`
    };
  }

  private getJobRecord(jobId: JobId): JobRecord | undefined {
    const job = this.jobs.get(jobId);
    return job ? { ...this.refreshJob(jobId), transactionIds: [...job.transactionIds], warnings: [...job.warnings] } : undefined;
  }

  private refreshJob(jobId: JobId): JobRecord {
    const job = this.jobs.get(jobId);
    if (!job) {
      throw new Error(`Job not found: ${jobId}`);
    }
    const transactions = job.transactionIds.map((transactionId) => this.transactions.get(transactionId)).filter((transaction): transaction is TransactionRecord => transaction !== undefined);
    const appliedCount = transactions.filter((transaction) => transaction.status === "applied").length;
    const terminalTransactions = transactions.filter((transaction) => isTerminalTransactionStatus(transaction.status));
    const failed = transactions.find((transaction) => transaction.status === "failed" || transaction.status === "blocked");
    const applying = transactions.some((transaction) => transaction.status === "applying");
    const queued = transactions.some((transaction) => transaction.status === "queued");
    const cancelledCount = transactions.filter((transaction) => transaction.status === "cancelled").length;
    let status: JobStatus = job.status;
    if (transactions.length === 0) {
      status = "queued";
    } else if (appliedCount === transactions.length) {
      status = "applied";
    } else if (cancelledCount === transactions.length) {
      status = "cancelled";
    } else if (terminalTransactions.length === transactions.length && appliedCount > 0) {
      status = "partially_applied";
    } else if (terminalTransactions.length === transactions.length) {
      status = failed !== undefined ? "failed" : "cancelled";
    } else if (failed !== undefined) {
      status = appliedCount > 0 ? "partially_applied" : "failed";
    } else if (appliedCount > 0 && terminalTransactions.length < transactions.length) {
      status = "partially_applied";
    } else if (applying) {
      status = "applying";
    } else if (queued) {
      status = "queued";
    }
    const now = new Date().toISOString();
    job.status = status;
    job.chunksCompleted = appliedCount;
    job.startedAt = job.startedAt ?? transactions.find((transaction) => transaction.startedAt !== undefined)?.startedAt;
    if (isTerminalJobStatus(status)) {
      job.finishedAt = job.finishedAt ?? now;
    }
    if (failed !== undefined) {
      job.errorCode = failed.errorCode;
      job.errorMessage = failed.errorMessage;
    }
    job.progressMessage = jobProgressMessage(job, transactions.length);
    return job;
  }

  async previewPlan(planId: PlanId) {
    const preview = this.plans.previewPlan(planId);
    const client = this.getActiveAddinClient();
    if (!client || preview.diffSummary.changedRanges.length === 0) {
      this.persistState();
      return preview;
    }

    const snapshot = await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
      workbookId: preview.workbookId,
      ranges: preview.diffSummary.changedRanges
    });

    const refreshed = this.plans.replacePreviewFingerprints(planId, {
      beforeWorkbookFingerprint: snapshot.workbookFingerprint,
      targetFingerprints: snapshot.rangeSnapshots.map((rangeSnapshot) => rangeSnapshot.fingerprint)
    });
    this.persistState();
    return refreshed;
  }

  async refreshPlanPreview(planId: PlanId): Promise<PlanRefreshResult> {
    const plan = this.plans.getPlan(planId);
    if (!plan) {
      return {
        ok: false,
        planId,
        refreshed: false,
        conflicts: [{ code: "PLAN_NOT_FOUND", message: `Plan not found: ${planId}` }],
        warnings: []
      };
    }
    const existingPreview = plan.preview ?? await this.previewPlan(planId);
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        planId,
        refreshed: false,
        preview: existingPreview,
        conflicts: [{ code: "ADDIN_DISCONNECTED", message: "No Excel add-in session is connected." }],
        warnings: []
      };
    }
    if (existingPreview.targetFingerprints.length === 0) {
      return { ok: true, planId, refreshed: false, preview: existingPreview, conflicts: [], warnings: [] };
    }

    const snapshot = await client.request<WorkbookSnapshotResponse>("workbook.snapshot_ranges", {
      workbookId: existingPreview.workbookId,
      ranges: existingPreview.targetFingerprints.map((fingerprint) => fingerprint.range)
    });
    const currentFingerprints = snapshot.rangeSnapshots.map((rangeSnapshot) => rangeSnapshot.fingerprint);
    const conflicts = detectFingerprintConflicts(existingPreview.targetFingerprints, currentFingerprints);
    if (conflicts.length > 0) {
      return {
        ok: false,
        planId,
        refreshed: false,
        preview: existingPreview,
        conflicts,
        warnings: [
          {
            code: "PLAN_REBASE_BLOCKED",
            message: "Plan target ranges changed after preview. Refresh/rebase is blocked to avoid overwriting newer work."
          }
        ]
      };
    }

    const refreshed = this.plans.replacePreviewFingerprints(planId, {
      beforeWorkbookFingerprint: snapshot.workbookFingerprint,
      targetFingerprints: currentFingerprints
    });
    this.recordCollabEvent({
      type: "task.updated",
      workbookId: plan.workbookId,
      agentId: plan.agentId,
      taskId: plan.taskId,
      message: `Plan preview refreshed: ${plan.goal}`,
      details: { planId }
    });
    return { ok: true, planId, refreshed: true, preview: refreshed, conflicts: [], warnings: [] };
  }

  async rebasePlan(planId: PlanId): Promise<PlanRefreshResult> {
    return this.refreshPlanPreview(planId);
  }

  getPlanDiffResource(workbookId: WorkbookId, planId: PlanId) {
    const plan = this.plans.getPlan(planId);
    if (!plan || plan.workbookId !== workbookId) {
      return {
        ok: false,
        workbookId,
        planId,
        error: runtimeError("NOT_FOUND", `Plan not found: ${planId}`, { retryable: false })
      };
    }
    return {
      ok: true,
      workbookId,
      planId,
      diffSummary: plan.preview?.diffSummary,
      preview: plan.preview,
      plan
    };
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
    return client.request<RuntimeSelectionResponse>("runtime.get_selection");
  }

  listOpenWorkbooks() {
    const workbooks = this.sessions
      .list()
      .map((session) => session.activeWorkbook)
      .filter((workbook): workbook is WorkbookRef => workbook !== undefined);
    const byId = new Map(workbooks.map((workbook) => [workbook.workbookId, workbook]));
    return { ok: true, workbooks: [...byId.values()] };
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

  async refreshSnapshot(input: { snapshotId: SnapshotId; reason?: string }) {
    const base = this.snapshots.getSnapshot(input.snapshotId);
    if (!base) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Snapshot not found: ${input.snapshotId}`, { retryable: false })
      };
    }
    return this.createWorkbookSnapshot({
      workbookId: base.workbookId,
      reason: input.reason ?? `Refresh snapshot ${input.snapshotId}`,
      ranges: base.affectedRanges
    });
  }

  refreshWorkbookSnapshot(input: { snapshotId: SnapshotId; reason?: string }) {
    return this.refreshSnapshot(input);
  }

  getWorkbookSnapshot(snapshotId: SnapshotId) {
    return this.getSnapshot(snapshotId);
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

  async saveWorkbookAs(workbookId: WorkbookId, targetPath?: string) {
    const bridgeStatus = this.fileBridge.getStatus();
    if (bridgeStatus.available) {
      const bridgeRequest: Parameters<NativeFileBridge["request"]>[0] = {
        operation: "workbook.save_as",
        workbookId
      };
      if (targetPath !== undefined) {
        bridgeRequest.targetPath = targetPath;
      }
      const bridge = await this.fileBridge.request(bridgeRequest);
      if (bridge.ok) {
        return { ok: true, workbookId, targetPath: bridge.filePath ?? bridge.targetPath ?? targetPath, bridge };
      }
      return {
        ok: false,
        workbookId,
        targetPath,
        bridge,
        error: runtimeError(
          "OPERATION_FAILED",
          `Native file bridge failed to save workbook as${targetPath ? ` ${targetPath}` : ""}: ${bridge.error ?? "unknown error"}`,
          { retryable: true, details: { bridgeStatus } }
        )
      };
    }
    return {
      ok: false,
      workbookId,
      targetPath,
      bridgeStatus,
      error: runtimeError(
        "CAPABILITY_UNAVAILABLE",
        "Office.js does not expose a local Save As file path API. Configure OPEN_WORKBOOK_FILE_BRIDGE_URL for true save_as.",
        { retryable: false, details: { bridgeStatus } }
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
    if (!("backup" in backup)) {
      return {
        ok: false,
        workbookId: input.workbookId,
        targetPath: input.targetPath,
        backup,
        error: runtimeError("BACKUP_UNAVAILABLE", "Workbook export requires a persistent safety backup before writing a file copy.", { retryable: true })
      };
    }
    const bridgeStatus = this.fileBridge.getStatus();
    let bridge: Awaited<ReturnType<NativeFileBridge["request"]>> | undefined;
    if (bridgeStatus.available) {
      const bridgeRequest: Parameters<NativeFileBridge["request"]>[0] = {
        operation: "workbook.export_copy",
        workbookId: input.workbookId,
        sourceBackupId: backup.backup.backupId
      };
      if (input.targetPath !== undefined) {
        bridgeRequest.targetPath = input.targetPath;
      }
      if (input.ranges !== undefined) {
        bridgeRequest.ranges = input.ranges;
      }
      if (input.reason !== undefined) {
        bridgeRequest.reason = input.reason;
      }
      bridge = await this.fileBridge.request(bridgeRequest);
      if (bridge.ok) {
        const exportedPath = bridge.filePath ?? bridge.targetPath ?? input.targetPath;
        return {
          ok: true,
          workbookId: input.workbookId,
          targetPath: exportedPath,
          backup,
          bridge
        };
      }
    }

    const client = this.getActiveAddinClient();
    if (client) {
      try {
        const exported = await client.request<WorkbookFileContent | { ok: false; error?: ReturnType<typeof runtimeError> }>("workbook.get_file", {
          workbookId: input.workbookId
        });
        if (exported.ok) {
          const targetPath = input.targetPath ?? this.defaultWorkbookExportPath(input.workbookId);
          const writtenPath = await this.writeWorkbookFileContent(targetPath, exported);
          return {
            ok: true,
            workbookId: input.workbookId,
            targetPath: writtenPath,
            backup,
            bridgeStatus,
            ...(bridge !== undefined ? { bridge } : {}),
            file: {
              path: writtenPath,
              size: exported.size,
              sliceCount: exported.sliceCount,
              capturedAt: exported.capturedAt,
              method: "office-js-compressed-file"
            }
          };
        }
        return {
          ok: false,
          workbookId: input.workbookId,
          targetPath: input.targetPath,
          backup,
          bridgeStatus,
          ...(bridge !== undefined ? { bridge } : {}),
          error:
            exported.error ??
            runtimeError("CAPABILITY_UNAVAILABLE", "The connected Excel add-in could not export a compressed workbook file.", {
              retryable: false,
              details: { bridgeStatus }
            })
        };
      } catch (error) {
        return {
          ok: false,
          workbookId: input.workbookId,
          targetPath: input.targetPath,
          backup,
          bridgeStatus,
          ...(bridge !== undefined ? { bridge } : {}),
          error: runtimeError(
            bridge !== undefined ? "OPERATION_FAILED" : "CAPABILITY_UNAVAILABLE",
            `Workbook export copy failed: ${error instanceof Error ? error.message : String(error)}`,
            { retryable: true, details: { bridgeStatus } }
          )
        };
      }
    }

    return {
      ok: false,
      workbookId: input.workbookId,
      targetPath: input.targetPath,
      backup,
      bridgeStatus,
      ...(bridge !== undefined ? { bridge } : {}),
      error: runtimeError(
        "CAPABILITY_UNAVAILABLE",
        "No native file bridge or Excel add-in file export path is available. A persistent snapshot backup was created instead.",
        { retryable: false, details: { bridgeStatus } }
      )
    };
  }

  async createFileBackup(input: WorkbookCreateFileBackupRequest) {
    const result = await this.exportWorkbookCopy({
      workbookId: input.workbookId,
      reason: input.reason ?? "Create full workbook file backup",
      targetPath: input.targetPath ?? this.defaultWorkbookFileBackupPath(input.workbookId)
    });
    if (!(result as { ok?: boolean }).ok || !(result as { targetPath?: string }).targetPath) {
      return result;
    }
    const filePath = (result as { targetPath: string }).targetPath;
    const fileStat = await stat(filePath);
    const checksum = await this.hashFile(filePath);
    const sourceSnapshotBackupId = extractBackupIds(result)[0];
    const manifest: WorkbookFileBackupManifest = {
      backupId: makeId<BackupId>("backup"),
      workbookId: input.workbookId,
      createdAt: new Date().toISOString(),
      reason: input.reason ?? "Create full workbook file backup",
      filePath,
      mode: input.mode ?? "export-copy",
      size: fileStat.size,
      checksum,
      pinned: input.pin ?? false,
      verifiedAt: new Date().toISOString(),
      restoreStatus: "available",
      metadata: {
        exportResult: result
      }
    };
    if (sourceSnapshotBackupId !== undefined) {
      manifest.sourceSnapshotBackupId = sourceSnapshotBackupId;
    }
    const bridge = (result as { bridge?: WorkbookFileBackupManifest["bridge"] }).bridge;
    if (bridge !== undefined) {
      manifest.bridge = bridge;
    }
    const backup = this.backups.createBackup({
      workbookId: input.workbookId,
      kind: "file-copy",
      reason: manifest.reason,
      affectedRanges: [],
      payloadRef: filePath,
      payload: manifest
    });
    manifest.backupId = backup.backupId;
    backup.payload = manifest;
    backup.pinned = manifest.pinned ?? false;
    backup.verifiedAt = manifest.verifiedAt ?? new Date().toISOString();
    backup.restoreStatus = manifest.restoreStatus ?? "available";
    this.recordCollabEvent({
      type: "backup.created",
      workbookId: input.workbookId,
      message: `File backup created: ${backup.backupId}`,
      details: { backupId: backup.backupId, filePath, checksum, size: fileStat.size, pinned: backup.pinned }
    });
    this.persistState();
    void this.applyDefaultBackupRetention("create_file_backup");
    return { ok: true, backup, manifest, export: result };
  }

  listFileBackups(workbookId?: WorkbookId) {
    const backups = this.backups
      .dump()
      .filter((backup) => isPersistedBackup(backup))
      .filter((backup) => workbookId === undefined || backup.workbookId === workbookId)
      .map((backup) => this.describePersistedBackup(backup));
    return { ok: true, backups };
  }

  getFileBackup(backupId: BackupId) {
    const backup = this.backups.getBackup(backupId);
    if (!backup || !isPersistedBackup(backup)) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Persisted backup not found: ${backupId}`, { retryable: false })
      };
    }
    return { ok: true, ...this.describePersistedBackup(backup) };
  }

  async verifyFileBackup(backupId: BackupId) {
    const backup = this.backups.getBackup(backupId);
    const manifest = backup?.payload as WorkbookFileBackupManifest | undefined;
    if (!backup || backup.kind !== "file-copy" || !manifest?.filePath) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `File backup cannot be verified: ${backupId}`, { retryable: false })
      };
    }
    try {
      const fileStat = await stat(manifest.filePath);
      const checksum = await this.hashFile(manifest.filePath);
      const restoreStatus = manifest.checksum !== undefined && checksum !== manifest.checksum ? "checksum_mismatch" : "available";
      const updatedManifest: WorkbookFileBackupManifest = {
        ...manifest,
        size: fileStat.size,
        checksum: manifest.checksum ?? checksum,
        verifiedAt: new Date().toISOString(),
        restoreStatus
      };
      const updated = this.backups.updateBackup(backupId, {
        payload: updatedManifest,
        verifiedAt: updatedManifest.verifiedAt ?? new Date().toISOString(),
        restoreStatus
      });
      this.recordCollabEvent({
        type: "backup.verified",
        workbookId: backup.workbookId,
        message: `File backup verified: ${backupId}`,
        details: { backupId, ok: restoreStatus === "available", restoreStatus, checksum }
      });
      this.persistState();
      return { ok: restoreStatus === "available", backup: updated, manifest: updatedManifest };
    } catch (error) {
      const updatedManifest: WorkbookFileBackupManifest = {
        ...manifest,
        verifiedAt: new Date().toISOString(),
        restoreStatus: "missing",
        metadata: { ...(manifest.metadata ?? {}), verifyError: error instanceof Error ? error.message : String(error) }
      };
      const updated = this.backups.updateBackup(backupId, {
        payload: updatedManifest,
        verifiedAt: updatedManifest.verifiedAt ?? new Date().toISOString(),
        restoreStatus: "missing"
      });
      this.recordCollabEvent({
        type: "backup.verified",
        workbookId: backup.workbookId,
        message: `File backup verification failed: ${backupId}`,
        details: { backupId, ok: false, restoreStatus: "missing", error: error instanceof Error ? error.message : String(error) }
      });
      this.persistState();
      return { ok: false, backup: updated, manifest: updatedManifest };
    }
  }

  async restoreFileBackup(input: WorkbookRestoreFileBackupRequest) {
    const verified = await this.verifyFileBackup(input.backupId);
    const manifest = (verified as { manifest?: WorkbookFileBackupManifest }).manifest;
    if (!(verified as { ok?: boolean }).ok || !manifest?.filePath) {
      return verified;
    }
    const restorableManifest: WorkbookFileBackupManifest & { filePath: string } = { ...manifest, filePath: manifest.filePath };
    const mode = input.mode ?? "open-as-new";
    if (mode === "open-as-new") {
      this.recordCollabEvent({
        type: "backup.restored",
        workbookId: input.workbookId,
        message: `File backup verified for open-as-new recovery: ${input.backupId}`,
        details: { backupId: input.backupId, mode, filePath: manifest.filePath }
      });
      return {
        ok: true,
        workbookId: input.workbookId,
        backupId: input.backupId,
        mode,
        filePath: manifest.filePath,
        note: "The file backup is verified and ready to open as a separate workbook. Open Workbook does not auto-open desktop Excel in headless MCP mode."
      };
    }
    if (input.confirmationToken === undefined && input.force !== true) {
      return {
        ok: false,
        workbookId: input.workbookId,
        backupId: input.backupId,
        mode,
        error: runtimeError("CONFIRMATION_REQUIRED", "Replacing or restoring into an open workbook requires explicit confirmation.", {
          retryable: false,
          details: { requiredConfirmation: "restore_file_backup" }
        })
      };
    }
    const permissionWarnings = this.validateDirectMutation(input.workbookId, [], "workbook");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Full-file restore is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const bridgeStatus = this.fileBridge.getStatus();
    if (bridgeStatus.available) {
      return this.applyDirectTransaction(
        {
          workbookId: input.workbookId,
          goal: `Restore file backup ${input.backupId}`,
          scopes: [{ type: "workbook", workbookId: input.workbookId }],
          destructiveLevel: "workbook"
        },
        async () => this.restoreFileBackupThroughBridge(input, restorableManifest, mode, bridgeStatus)
      );
    }
    return {
      ok: false,
      workbookId: input.workbookId,
      backupId: input.backupId,
      mode,
      bridgeStatus,
      error: runtimeError("CAPABILITY_UNAVAILABLE", "Full workbook file restore requires a native bridge with close/reopen support. Use open-as-new mode for verified recovery.", {
        retryable: false,
        details: { bridgeStatus }
      })
    };
  }

  private async restoreFileBackupThroughBridge(
    input: WorkbookRestoreFileBackupRequest,
    manifest: WorkbookFileBackupManifest & { filePath: string },
    mode: NonNullable<WorkbookRestoreFileBackupRequest["mode"]>,
    bridgeStatus: ReturnType<NativeFileBridge["getStatus"]>
  ) {
      const emergencyBackup = await this.createFileBackup({
        workbookId: input.workbookId,
        reason: `Emergency file backup before restoring ${input.backupId}`,
        pin: true
      });
      if (!(emergencyBackup as { ok?: boolean }).ok) {
        return {
          ok: false,
          workbookId: input.workbookId,
          backupId: input.backupId,
          mode,
          emergencyBackup,
          error: runtimeError("BACKUP_UNAVAILABLE", "Destructive file restore requires an emergency file backup before replacing the workbook.", {
            retryable: true,
            details: { bridgeStatus }
          })
        };
      }
      const bridge = await this.fileBridge.request({
        operation: "workbook.restore_file_backup",
        workbookId: input.workbookId,
        targetPath: manifest.filePath,
        backupPath: manifest.filePath,
        ...(input.restoreTargetPath !== undefined ? { restoreTargetPath: input.restoreTargetPath } : {}),
        restoreMode: mode,
        sourceBackupId: input.backupId,
        reason: `Restore file backup ${input.backupId}`
      });
      if (bridge.ok) {
        const emergencyBackupId = (emergencyBackup as { manifest?: { backupId?: BackupId } }).manifest?.backupId;
        const restoredManifest: WorkbookFileBackupManifest = {
          ...manifest,
          metadata: {
            ...(manifest.metadata ?? {}),
            lastRestoredAt: new Date().toISOString(),
            lastRestoreMode: mode,
            lastRestoreBridge: bridge,
            emergencyBackupId
          }
        };
        this.backups.updateBackup(input.backupId, { payload: restoredManifest });
        this.recordCollabEvent({
          type: "backup.restored",
          workbookId: input.workbookId,
          message: `File backup restored: ${input.backupId}`,
          details: {
            backupId: input.backupId,
            mode,
            bridge,
            emergencyBackupId
          }
        });
        this.persistState();
        return {
          ok: true,
          workbookId: input.workbookId,
          backupId: input.backupId,
          mode,
          bridge,
          emergencyBackup,
          ...((emergencyBackup as { backup?: unknown }).backup !== undefined ? { backup: (emergencyBackup as { backup: unknown }).backup } : {})
        };
      }
      return {
        ok: false,
        workbookId: input.workbookId,
        backupId: input.backupId,
        mode,
        bridge,
        emergencyBackup,
        error: runtimeError("OPERATION_FAILED", "The native file bridge failed to restore the workbook file backup.", {
          retryable: false,
          details: { bridgeStatus }
        })
      };
  }

  async deleteFileBackup(backupId: BackupId) {
    const backup = this.backups.getBackup(backupId);
    if (!backup || !isPersistedBackup(backup)) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Persisted backup not found: ${backupId}`, { retryable: false })
      };
    }
    if (backup.pinned) {
      return {
        ok: false,
        backup,
        error: runtimeError("PERMISSION_DENIED", `Persisted backup is pinned and cannot be deleted: ${backupId}`, { retryable: false })
      };
    }
    const payload = await this.unlinkBackupPayloadIfSafe(backup);
    const deleted = this.backups.deleteBackup(backupId);
    this.recordCollabEvent({
      type: "backup.deleted",
      workbookId: backup.workbookId,
      message: `Persisted backup deleted: ${backupId}`,
      details: { backupId, kind: backup.kind, payload }
    });
    this.persistState();
    return { ok: deleted, backupId, backup, payload };
  }

  async pruneFileBackups(input: WorkbookBackupRetentionRequest) {
    const now = Date.now();
    const maxAgeMs = input.maxAgeDays !== undefined ? input.maxAgeDays * 24 * 60 * 60 * 1000 : undefined;
    const byWorkbook = new Map<string, BackupRecord[]>();
    for (const backup of this.backups.dump().filter((item) => isPersistedBackup(item) && backupMatchesRetentionKind(item, input.kind ?? "all"))) {
      if (input.workbookId !== undefined && backup.workbookId !== input.workbookId) {
        continue;
      }
      const list = byWorkbook.get(backup.workbookId) ?? [];
      list.push(backup);
      byWorkbook.set(backup.workbookId, list);
    }
    const candidates = new Map<BackupId, PruneCandidate>();
    for (const backups of byWorkbook.values()) {
      for (const backup of backups) {
        if (!backup.pinned && maxAgeMs !== undefined && now - Date.parse(backup.createdAt) > maxAgeMs) {
          addPruneCandidate(candidates, backup, "age");
        }
      }
      if (input.maxBackupsPerWorkbook !== undefined) {
        const overflow = backups
          .filter((backup) => !backup.pinned)
          .sort((a, b) => Date.parse(b.createdAt) - Date.parse(a.createdAt))
          .slice(input.maxBackupsPerWorkbook);
        for (const backup of overflow) {
          addPruneCandidate(candidates, backup, "count");
        }
      }
    }
    const groupedBackups = [...byWorkbook.values()].flat();
    if (input.maxTotalBytes !== undefined) {
      const sized = await Promise.all(groupedBackups.map(async (backup) => ({ backup, payload: await this.getBackupPayloadInfo(backup) })));
      let totalBytes = sized.reduce((total, item) => total + (item.payload.bytes ?? 0), 0);
      const oldestFirst = sized
        .filter((item) => !item.backup.pinned)
        .sort((a, b) => Date.parse(a.backup.createdAt) - Date.parse(b.backup.createdAt));
      for (const item of oldestFirst) {
        if (totalBytes <= input.maxTotalBytes) {
          break;
        }
        addPruneCandidate(candidates, item.backup, "size", item.payload);
        totalBytes -= item.payload.bytes ?? 0;
      }
    }
    for (const candidate of candidates.values()) {
      if (candidate.bytes === undefined && candidate.payloadPath === undefined) {
        const payload = await this.getBackupPayloadInfo(candidate.backup);
        candidate.bytes = payload.bytes;
        candidate.payloadPath = payload.path;
        candidate.missingPayload = payload.missing;
      }
    }
    const candidateList = [...candidates.values()].map((candidate) => ({
      backup: candidate.backup,
      reasons: candidate.reasons,
      bytes: candidate.bytes ?? 0,
      payloadPath: candidate.payloadPath,
      missingPayload: candidate.missingPayload ?? false
    }));
    const skippedPinned = groupedBackups.filter((backup) => backup.pinned).map((backup) => backup.backupId);
    const reclaimedBytes = candidateList.reduce((total, candidate) => total + candidate.bytes, 0);
    const missingPayloads = candidateList.filter((candidate) => candidate.missingPayload).map((candidate) => candidate.backup.backupId);
    if (input.dryRun) {
      return { ok: true, dryRun: true, candidates: candidateList, skippedPinned, reclaimedBytes, missingPayloads };
    }
    if (candidateList.length === 0) {
      return { ok: true, pruned: [], candidates: candidateList, skippedPinned, reclaimedBytes, missingPayloads };
    }
    const pruned: BackupId[] = [];
    for (const candidate of candidateList) {
      await this.unlinkBackupPayloadIfSafe(candidate.backup);
      this.backups.deleteBackup(candidate.backup.backupId);
      pruned.push(candidate.backup.backupId);
    }
    this.recordCollabEvent({
      type: "backup.pruned",
      ...(input.workbookId !== undefined ? { workbookId: input.workbookId } : {}),
      message: `Persisted backups pruned: ${pruned.length}`,
      details: { prunedBackupIds: pruned, criteria: input, reclaimedBytes, missingPayloads }
    });
    this.persistState();
    return { ok: true, pruned, candidates: candidateList, skippedPinned, reclaimedBytes, missingPayloads };
  }

  pinFileBackup(backupId: BackupId, pinned: boolean) {
    const backup = this.backups.getBackup(backupId);
    if (!backup || !isPersistedBackup(backup)) {
      return {
        ok: false,
        error: runtimeError("BACKUP_UNAVAILABLE", `Persisted backup not found: ${backupId}`, { retryable: false })
      };
    }
    const manifest = backup.payload as WorkbookFileBackupManifest | undefined;
    const updatedManifest = manifest ? { ...manifest, pinned } : undefined;
    const updated = this.backups.updateBackup(backupId, {
      pinned,
      ...(updatedManifest ? { payload: updatedManifest } : {})
    });
    this.recordCollabEvent({
      type: "backup.updated",
      workbookId: backup.workbookId,
      message: `Persisted backup ${pinned ? "pinned" : "unpinned"}: ${backupId}`,
      details: { backupId, pinned }
    });
    this.persistState();
    return { ok: true, backup: updated, manifest: updatedManifest };
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
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before creating named range ${request.name}`,
        scopes: nameMutationScopes(request),
        destructiveLevel: "structure"
      },
      async () => {
        const backup = request.reference && request.sheetName ? await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before creating named range ${request.name}`,
          ranges: [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.reference }]
        }) : undefined;
        const result = await client.request("names.create", request);
        return backup ? { ok: true, backup, result } : { ok: true, result };
      }
    );
  }

  async updateName(request: NameUpdateRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before updating named range ${request.name}`,
        scopes: nameMutationScopes(request),
        destructiveLevel: "structure"
      },
      async () => {
        const backup = request.reference && request.sheetName ? await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before updating named range ${request.name}`,
          ranges: [{ workbookId: request.workbookId, sheetName: request.sheetName, address: request.reference }]
        }) : undefined;
        const result = await client.request("names.update", request);
        return backup ? { ok: true, backup, result } : { ok: true, result };
      }
    );
  }

  async deleteName(request: NameSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before deleting named range ${request.name}`,
        scopes: nameMutationScopes(request),
        destructiveLevel: "structure"
      },
      async () => {
        const result = await client.request("names.delete", request);
        return { ok: true, result };
      }
    );
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
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before creating PivotTable ${request.pivotTableName}`,
        scopes: pivotMutationScopes(request, ranges),
        destructiveLevel: "structure"
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before creating PivotTable ${request.pivotTableName}`,
          ranges
        });
        const result = await client.request("pivot.create", request);
        return { ok: true, backup, result };
      }
    );
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
    const capabilityStatus = pivotOperationCapabilityStatus("pivot.update_source");
    return {
      ok: false,
      request,
      warnings: capabilityStatus.warnings,
      capabilityStatus,
      error: runtimeError("CAPABILITY_UNAVAILABLE", "Office.js does not expose safe in-place PivotTable source reassignment in this runtime. Create a new PivotTable from the desired source.", {
        retryable: false
      })
    };
  }

  async copyPivotFromTemplate(request: PivotCopyFromTemplateRequest) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    if (!request.templatePivotTableName) {
      return {
        ok: false,
        request,
        error: runtimeError("INVALID_ARGUMENT", "PivotTable template copy requires templatePivotTableName.", { retryable: false })
      };
    }
    const targetInfo = await this.getPivotTableInfo(request);
    const sourceInfo = await this.getPivotTableInfo({
      workbookId: request.workbookId,
      pivotTableName: request.templatePivotTableName
    });
    const targetPivotInfo = (targetInfo as { ok?: boolean; info?: PivotTableInfo }).info;
    const sourcePivotInfo = (sourceInfo as { ok?: boolean; info?: PivotTableInfo }).info;
    if (!(targetInfo as { ok?: boolean }).ok || !targetPivotInfo) {
      return {
        ok: false,
        target: targetInfo,
        error: runtimeError("NOT_FOUND", `Target PivotTable not found or unavailable: ${request.pivotTableName}`, { retryable: false })
      };
    }
    if (!(sourceInfo as { ok?: boolean }).ok || !sourcePivotInfo) {
      return {
        ok: false,
        source: sourceInfo,
        error: runtimeError("NOT_FOUND", `Template PivotTable not found or unavailable: ${request.templatePivotTableName}`, { retryable: false })
      };
    }
    const compatibilityIssues = validatePivotTemplateCompatibility(sourcePivotInfo, targetPivotInfo);
    if (compatibilityIssues.some((issue) => issue.severity === "error")) {
      return {
        ok: false,
        source: sourceInfo,
        target: targetInfo,
        issues: compatibilityIssues,
        error: runtimeError("TEMPLATE_MISMATCH", "PivotTable template copy is blocked because the target pivot source is missing fields required by the template.", {
          retryable: false,
          details: { issues: compatibilityIssues }
        })
      };
    }
    const targetSheetName = targetPivotInfo?.sheetName;
    const sourceSheetName = sourcePivotInfo?.sheetName;
    const ranges = await this.getPivotTemplateCopyRanges(request.workbookId, targetSheetName, targetPivotInfo?.range?.address);
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "format");
    if (permissionWarnings.length > 0) {
      return permissionDenied("PivotTable template copy is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const capabilityStatus = pivotOperationCapabilityStatus("pivot.copy_from_template", request.dimensions);
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before copying PivotTable template ${request.templatePivotTableName} to ${request.pivotTableName}`,
        scopes: pivotTemplateCopyScopes(request, targetSheetName, sourceSheetName, ranges),
        destructiveLevel: "format"
      },
      async () => {
        const backupResult = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before copying PivotTable template ${request.templatePivotTableName} to ${request.pivotTableName}`,
          ranges
        });
        if (!("backup" in backupResult)) {
          return backupResult;
        }
        const result = await client.request<PivotCopyFromTemplateResponse>("pivot.copy_from_template", request);
        const warnings = mergeOperationWarnings(capabilityStatus.warnings, result.warnings ?? []);
        return {
          ok: true,
          backup: backupResult.backup,
          result: { ...result, warnings, capabilityStatus },
          warnings,
          capabilityStatus
        };
      }
    );
  }

  async deletePivotTable(request: PivotSelector) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    const targetInfo = await this.getPivotTableInfo(request);
    const pivotInfo = (targetInfo as { ok?: boolean; info?: { sheetName?: string; range?: { address?: string } } }).info;
    if (!(targetInfo as { ok?: boolean }).ok || !pivotInfo) {
      return {
        ok: false,
        target: targetInfo,
        error: runtimeError("NOT_FOUND", `PivotTable not found or unavailable: ${request.pivotTableName}`, { retryable: false })
      };
    }
    const ranges = pivotInfo.sheetName && pivotInfo.range?.address
      ? [{ workbookId: request.workbookId, sheetName: pivotInfo.sheetName, address: stripSheetName(pivotInfo.range.address) }]
      : [];
    const permissionWarnings = this.validateDirectMutation(request.workbookId, ranges, "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("PivotTable deletion is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before deleting PivotTable ${request.pivotTableName}`,
        scopes: pivotDeleteScopes(request, pivotInfo.sheetName, ranges),
        destructiveLevel: "structure"
      },
      async () => {
        const backupResult = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before deleting PivotTable ${request.pivotTableName}`,
          ...(ranges.length > 0 ? { ranges } : {})
        });
        if (!("backup" in backupResult)) {
          return backupResult;
        }
        const result = await client.request("pivot.delete", request);
        return { ok: true, backup: backupResult.backup, result };
      }
    );
  }

  async validatePivotSource(request: PivotValidateSourceRequest) {
    const info = await this.getPivotTableInfo(request);
    const pivotInfo = (info as {
      ok?: boolean;
      info?: {
        source?: string;
        sourceType?: string;
        sheetName?: string;
        range?: { address?: string };
        hierarchies?: Array<{ name: string }>;
        rowHierarchies?: Array<{ name: string }>;
        columnHierarchies?: Array<{ name: string }>;
        filterHierarchies?: Array<{ name: string }>;
        dataHierarchies?: Array<{ name: string; field?: { name: string } }>;
        layout?: PivotLayoutInfo;
      };
    }).info;
    const issues: Array<{ code: string; severity: "info" | "warning" | "error"; message: string; details?: Record<string, unknown> }> = [];
    if (!pivotInfo) {
      issues.push({
        code: "PIVOT_NOT_FOUND",
        severity: "error",
        message: `PivotTable not found or unavailable: ${request.pivotTableName}`
      });
    }
    if (pivotInfo && !pivotInfo.source) {
      issues.push({
        code: "PIVOT_SOURCE_UNAVAILABLE",
        severity: "warning",
        message: "Pivot source string is unavailable from Office.js."
      });
    }
    if (pivotInfo && !pivotInfo.sourceType) {
      issues.push({
        code: "PIVOT_SOURCE_TYPE_UNAVAILABLE",
        severity: "warning",
        message: "Pivot source type is unavailable from Office.js."
      });
    }
    if (pivotInfo && (!pivotInfo.sheetName || !pivotInfo.range?.address)) {
      issues.push({
        code: "PIVOT_OUTPUT_RANGE_UNAVAILABLE",
        severity: "warning",
        message: "PivotTable output sheet or range is unavailable from Office.js."
      });
    }
    if (pivotInfo && !pivotInfo.dataHierarchies?.length) {
      issues.push({
        code: "PIVOT_HAS_NO_DATA_FIELDS",
        severity: "info",
        message: "PivotTable has no data fields. It may be blank or only partially configured."
      });
    }
    const availableSourceFields = uniqueDefined((pivotInfo?.hierarchies ?? []).map((hierarchy) => hierarchy.name));
    const rowFields = uniqueDefined((pivotInfo?.rowHierarchies ?? []).map((hierarchy) => hierarchy.name));
    const columnFields = uniqueDefined((pivotInfo?.columnHierarchies ?? []).map((hierarchy) => hierarchy.name));
    const filterFields = uniqueDefined((pivotInfo?.filterHierarchies ?? []).map((hierarchy) => hierarchy.name));
    const dataFields = uniqueDefined((pivotInfo?.dataHierarchies ?? []).map((hierarchy) => hierarchy.field?.name ?? hierarchy.name));
    const expectedSourceFields = uniqueDefined([
      ...(request.expectedFields ?? []),
      ...(request.expectedRowFields ?? []),
      ...(request.expectedColumnFields ?? []),
      ...(request.expectedFilterFields ?? []),
      ...(request.expectedDataFields ?? [])
    ]);
    if (pivotInfo && expectedSourceFields.length > 0 && availableSourceFields.length === 0) {
      issues.push({
        code: "PIVOT_SOURCE_FIELDS_UNAVAILABLE",
        severity: "warning",
        message: "Pivot source field metadata is unavailable from Office.js, so expected fields cannot be verified.",
        details: { expectedFields: expectedSourceFields }
      });
    }
    if (availableSourceFields.length > 0) {
      const missingSourceFields = expectedSourceFields.filter((field) => !availableSourceFields.includes(field));
      for (const field of missingSourceFields) {
        issues.push({
          code: "PIVOT_EXPECTED_FIELD_MISSING",
          severity: "error",
          message: `Expected PivotTable source field is missing: ${field}`,
          details: { field, availableSourceFields }
        });
      }
    }
    addExpectedPivotAxisIssues(issues, "row", request.expectedRowFields, rowFields);
    addExpectedPivotAxisIssues(issues, "column", request.expectedColumnFields, columnFields);
    addExpectedPivotAxisIssues(issues, "filter", request.expectedFilterFields, filterFields);
    addExpectedPivotAxisIssues(issues, "data", request.expectedDataFields, dataFields);
    addExpectedPivotDataSettingIssues(issues, request.expectedDataFieldSettings, pivotInfo?.dataHierarchies ?? []);
    addExpectedPivotLayoutIssues(issues, request.expectedLayout, pivotInfo?.layout);
    const hierarchyCount =
      (pivotInfo?.rowHierarchies?.length ?? 0) +
      (pivotInfo?.columnHierarchies?.length ?? 0) +
      (pivotInfo?.filterHierarchies?.length ?? 0) +
      (pivotInfo?.dataHierarchies?.length ?? 0);
    return {
      ok: issues.every((issue) => issue.severity !== "error"),
      info,
      summary: {
        hasSource: Boolean(pivotInfo?.source),
        sourceType: pivotInfo?.sourceType,
        hasOutputRange: Boolean(pivotInfo?.sheetName && pivotInfo.range?.address),
        hierarchyCount,
        sourceFieldCount: availableSourceFields.length,
        sourceFields: availableSourceFields,
        rowFields,
        columnFields,
        filterFields,
        dataFields
      },
      issues
    };
  }

  getPivotCapabilityMatrix(workbookId?: WorkbookId) {
    return {
      ok: true,
      matrix: {
        workbookId,
        hostPlatform: this.sessions.getActive()?.activeWorkbook?.platform,
        capabilities: [
          { capability: "create", status: "supported" },
          { capability: "read_source_metadata", status: "partial", reason: "Office.js exposes source metadata inconsistently across hosts and pivot source types." },
          { capability: "read_axis_fields", status: "partial", reason: "Available when the host exposes PivotTable hierarchy metadata." },
          { capability: "write_axis_fields", status: "supported" },
          { capability: "write_data_fields", status: "supported" },
          { capability: "aggregation", status: "supported" },
          { capability: "number_format", status: "supported" },
          { capability: "layout_flags", status: "partial", reason: "Only deterministic Office.js layout fields are replayed." },
          { capability: "refresh", status: "supported" },
          { capability: "delete", status: "supported" },
          { capability: "template_copy", status: "partial", reason: "Template copy replays deterministic fields and returns warnings for unavailable dimensions." },
          { capability: "fingerprint", status: "partial", reason: "Fingerprint includes metadata Office.js exposes and marks missing dimensions as warnings." },
          { capability: "diff", status: "partial", reason: "Diff is deterministic for captured fingerprint dimensions only." },
          { capability: "rebuild_with_source", status: "partial", reason: "Rebuild creates a new PivotTable from the desired source; in-place replacement remains guarded." },
          { capability: "source_reassignment", status: "unsupported", reason: "Office.js does not expose safe in-place source reassignment in this runtime." },
          { capability: "pivot_chart", status: "partial", reason: "PivotChart-specific controls remain host/API limited." }
        ]
      }
    };
  }

  async getPivotFingerprint(request: PivotSelector) {
    const infoResult = await this.getPivotTableInfo(request);
    const info = (infoResult as { ok?: boolean; info?: PivotTableInfo }).info;
    if (!(infoResult as { ok?: boolean }).ok || !info) {
      return {
        ok: false,
        info: infoResult,
        error: runtimeError("NOT_FOUND", `PivotTable not found or unavailable: ${request.pivotTableName}`, { retryable: false })
      };
    }
    const fingerprint = makePivotFingerprint(info);
    return { ok: true, fingerprint, info };
  }

  async comparePivotFingerprint(request: PivotCompareFingerprintRequest) {
    const source = await this.getPivotFingerprint(request);
    const target = await this.getPivotFingerprint({
      workbookId: request.workbookId,
      pivotTableName: request.targetPivotTableName
    });
    if (!(source as { ok?: boolean }).ok || !(target as { ok?: boolean }).ok) {
      return {
        ok: false,
        source,
        target,
        error: runtimeError("NOT_FOUND", "One or both PivotTables are unavailable for fingerprint comparison.", { retryable: false })
      };
    }
    const sourceFingerprint = (source as { fingerprint: PivotFingerprint }).fingerprint;
    const targetFingerprint = (target as { fingerprint: PivotFingerprint }).fingerprint;
    const diff = diffPivotFingerprints(sourceFingerprint, targetFingerprint, request.targetPivotTableName);
    return { ok: diff.changes.length === 0, source: sourceFingerprint, target: targetFingerprint, diff };
  }

  async diffPivotTables(request: PivotCompareFingerprintRequest): Promise<PivotDiff> {
    const comparison = await this.comparePivotFingerprint(request);
    if (!(comparison as { diff?: PivotDiff }).diff) {
      return {
        ok: false,
        workbookId: request.workbookId,
        sourcePivotTableName: request.pivotTableName,
        targetPivotTableName: request.targetPivotTableName,
        changes: [],
        warnings: [{ code: "PIVOT_DIFF_UNAVAILABLE", message: "One or both PivotTable fingerprints are unavailable." }]
      };
    }
    return (comparison as { diff: PivotDiff }).diff;
  }

  async repairPivotFromTemplate(request: PivotRepairFromTemplateRequest) {
    const capabilityStatus = pivotOperationCapabilityStatus("pivot.repair_from_template", request.dimensions);
    const before = await this.comparePivotFingerprint({
      workbookId: request.workbookId,
      pivotTableName: request.templatePivotTableName,
      targetPivotTableName: request.pivotTableName
    });
    if (request.strict && (before as { diff?: PivotDiff }).diff?.warnings.some((warning) => warning.code.includes("UNAVAILABLE"))) {
      return {
        ok: false,
        before,
        warnings: capabilityStatus.warnings,
        capabilityStatus,
        error: runtimeError("CAPABILITY_UNAVAILABLE", "Strict PivotTable repair cannot proceed because one or more fingerprint dimensions are unavailable.", {
          retryable: false
        })
      };
    }
    const copy = await this.copyPivotFromTemplate(request);
    if (!(copy as { ok?: boolean }).ok) {
      return copy;
    }
    const after = await this.comparePivotFingerprint({
      workbookId: request.workbookId,
      pivotTableName: request.templatePivotTableName,
      targetPivotTableName: request.pivotTableName
    });
    const warnings = mergeOperationWarnings(capabilityStatus.warnings, (copy as { warnings?: OperationWarning[] }).warnings ?? []);
    return { ok: true, before, copy, after, warnings, capabilityStatus };
  }

  async rebuildPivotWithSource(request: PivotRebuildWithSourceRequest) {
    const capabilityStatus = pivotOperationCapabilityStatus("pivot.rebuild_with_source");
    if (request.replaceExisting) {
      if (request.templatePivotTableName === request.pivotTableName) {
        return {
          ok: false,
          request,
          error: runtimeError("INVALID_ARGUMENT", "replaceExisting cannot use the same PivotTable as its template because the target must be deleted first.", {
            retryable: false
          })
        };
      }
      const deleted = await this.deletePivotTable({
        workbookId: request.workbookId,
        pivotTableName: request.pivotTableName
      });
      if (!(deleted as { ok?: boolean }).ok) {
        return { ok: false, request, deleted };
      }
      const created = await this.createPivotTable(request);
      if (!(created as { ok?: boolean }).ok || !request.templatePivotTableName) {
        return {
          ok: (created as { ok?: boolean }).ok,
          deleted,
          created,
          warnings: capabilityStatus.warnings,
          capabilityStatus
        };
      }
      const repairRequest: PivotRepairFromTemplateRequest = {
        workbookId: request.workbookId,
        pivotTableName: request.pivotTableName,
        templatePivotTableName: request.templatePivotTableName
      };
      if (request.strict !== undefined) {
        repairRequest.strict = request.strict;
      }
      const repaired = await this.repairPivotFromTemplate(repairRequest);
      const warnings = mergeOperationWarnings(capabilityStatus.warnings, (repaired as { warnings?: OperationWarning[] }).warnings ?? []);
      return { ok: (repaired as { ok?: boolean }).ok, deleted, created, repaired, warnings, capabilityStatus };
    }
    const created = await this.createPivotTable(request);
    if (!(created as { ok?: boolean }).ok || !request.templatePivotTableName) {
      return {
        ...(created as object),
        warnings: mergeOperationWarnings(capabilityStatus.warnings, (created as { warnings?: OperationWarning[] }).warnings ?? []),
        capabilityStatus
      };
    }
    const repairRequest: PivotRepairFromTemplateRequest = {
      workbookId: request.workbookId,
      pivotTableName: request.pivotTableName,
      templatePivotTableName: request.templatePivotTableName
    };
    if (request.strict !== undefined) {
      repairRequest.strict = request.strict;
    }
    const repaired = await this.repairPivotFromTemplate(repairRequest);
    const warnings = mergeOperationWarnings(capabilityStatus.warnings, (repaired as { warnings?: OperationWarning[] }).warnings ?? []);
    return { ok: (repaired as { ok?: boolean }).ok, created, repaired, warnings, capabilityStatus };
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
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before creating chart ${request.chartName ?? request.chartType}`,
        scopes: chartMutationScopes(request, ranges),
        destructiveLevel: "structure"
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before creating chart ${request.chartName ?? request.chartType}`,
          ranges
        });
        const result = await client.request("chart.create", request);
        return { ok: true, backup, result };
      }
    );
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
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before updating chart ${request.chartName} data source`,
        scopes: chartMutationScopes(request, ranges),
        destructiveLevel: "structure"
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before updating chart ${request.chartName} data source`,
          ranges
        });
        const result = await client.request("chart.update_data_source", request);
        return { ok: true, backup, result };
      }
    );
  }

  async copyChartFromTemplate(request: ChartSelector & { templateChartName: string; templateSheetName: string }) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    const permissionWarnings = this.validateDirectMutation(request.workbookId, [], "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Chart template copy is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before copying chart template ${request.templateChartName} to ${request.chartName}`,
        scopes: dedupeScopes([
          { type: "chart", workbookId: request.workbookId, sheetName: request.templateSheetName, chartName: request.templateChartName },
          { type: "chart", workbookId: request.workbookId, sheetName: request.sheetName, chartName: request.chartName }
        ]),
        destructiveLevel: "structure"
      },
      async () => {
        const backupResult = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before copying chart template ${request.templateChartName} to ${request.chartName}`
        });
        if (!("backup" in backupResult)) {
          return backupResult;
        }
        const result = await client.request("chart.copy_from_template", request);
        return { ok: true, backup: backupResult.backup, result };
      }
    );
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
    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: `Before deleting chart ${request.chartName}`,
        scopes: [{ type: "chart", workbookId: request.workbookId, sheetName: request.sheetName, chartName: request.chartName }],
        destructiveLevel: "structure"
      },
      async () => {
        const backupResult = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason: `Before deleting chart ${request.chartName}`
        });
        if (!("backup" in backupResult)) {
          return backupResult;
        }
        const result = await client.request("chart.delete", request);
        return { ok: (result as { ok?: boolean }).ok !== false, backup: backupResult.backup, result };
      }
    );
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
    this.persistState();
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
    const before = read.values[headerRowIndex] ?? [];
    const normalized = dedupeHeaders(before.map((value) => normalizeHeader(String(value ?? ""))));
    const result = await this.writeChangedCleanValues(headerRowTarget(target, headerRowIndex), [before], [normalized], "Normalize headers");
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
    const values = padMatrixRows(compact, read.values.length, read.values[0]?.length ?? 0, "");
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
    const result = await this.writeChangedCleanValues(target, read.values, values, "Fill missing values");
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
    const result = await this.writeCleanValues(target, rectangularize(values, ""), "Split column");
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

  async readTable(request: TableReadRequest) {
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

  async reorderTableColumns(request: TableReorderColumnsRequest) {
    const infoResult = await this.getTableInfo(request);
    const info = (infoResult as { ok?: boolean; info?: TableInfo }).info;
    if (!info) {
      return {
        ok: false,
        error: runtimeError("NOT_FOUND", `Table ${request.tableName} could not be read before column reorder.`, { retryable: false }),
        table: infoResult
      };
    }
    const validation = validateTableColumnOrder(info, request.columnOrder);
    if (!validation.ok) {
      return validation;
    }
    return this.mutateTable("table.reorder_columns", request, `Before reordering columns in table ${request.tableName}`, await this.getTableBackupRanges(request));
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

  async applyTableView(request: TableApplyViewRequest) {
    return this.mutateTable("table.apply_view", request, `Before applying table view to ${request.tableName}`, await this.getTableBackupRanges(request));
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
    backup.payloadRef = await this.persistBackupPayload(backup.backupId, snapshotResult.snapshot.payload);
    this.persistState();
    void this.applyDefaultBackupRetention("create_backup");
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

  restoreWorkbookBackup(backupId: BackupId, confirmationToken?: string): Promise<OperationResult> {
    return this.restoreBackup(backupId, confirmationToken);
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

    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: reason,
        scopes: tableMutationScopes(request, ranges),
        destructiveLevel: method.includes("resize") || method.includes("copy_structure") || method.includes("reorder_columns") ? "structure" : "values"
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason,
          ranges
        });
        if (!("backup" in backup)) {
          return backup;
        }
        const result = await client.request(method, request);
        if ((result as { ok?: boolean }).ok === false) {
          return {
            ok: false,
            backup: backup.backup,
            warnings: (result as { warnings?: OperationWarning[] }).warnings ?? [],
            error: (result as { error?: ReturnType<typeof runtimeError> }).error ?? runtimeError("OPERATION_FAILED", `Table operation ${method} failed.`, { retryable: false }),
            result
          };
        }
        return { ok: true, backup: backup.backup, result };
      }
    );
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

    return this.applyDirectTransaction(
      {
        workbookId: request.workbookId,
        goal: reason,
        scopes: formulaMutationScopes(request, ranges),
        destructiveLevel: "values"
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: request.workbookId,
          reason,
          ranges
        });
        const result = await client.request<FormulaMutationResponse>(method, request);
        const validation = validate ? await validate() : undefined;
        return { ok: result.ok, backup, result, validation };
      }
    );
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
    const template = this.templates.register(
      request.scope === "workbook"
        ? { ...input, workbookId: request.workbookId }
        : input
    );
    this.persistState();
    return template;
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
    const ok = this.templates.unregister(templateId);
    this.persistState();
    return { ok };
  }

  exportWorkbookLocalConfig(workbookId: WorkbookId, options: { includePermissions?: boolean } = {}) {
    const regions = [...this.regions.values()]
      .filter((region) => region.workbookId === workbookId)
      .map((region) => ({ ...region }));
    const templates = this.templates
      .dump()
      .filter((template) => template.scope === "local" || template.workbookId === workbookId)
      .map((template) => ({ ...template }));
    const config: WorkbookLocalConfig = {
      version: 1,
      workbookId,
      exportedAt: new Date().toISOString(),
      source: "open-workbook-local-config",
      templates: templates as unknown as Array<Record<string, unknown>>,
      regions
    };
    if (options.includePermissions ?? true) {
      config.permissions = clonePermissionStateForWorkbook(this.permissionState, workbookId);
    }
    return {
      ok: true,
      workbookId,
      config,
      counts: {
        templates: templates.length,
        regions: regions.length,
        permissions: config.permissions !== undefined
      }
    };
  }

  importWorkbookLocalConfig(request: WorkbookLocalConfigImportRequest): WorkbookLocalConfigImportResponse {
    const includeTemplates = request.includeTemplates ?? true;
    const includeRegions = request.includeRegions ?? true;
    const includePermissions = request.includePermissions ?? true;
    const overwrite = request.overwrite ?? false;
    if (request.config.version !== 1 || request.config.source !== "open-workbook-local-config") {
      return {
        ok: false,
        workbookId: request.workbookId,
        imported: { templates: 0, regions: 0, permissions: false },
        skipped: { templates: 0, regions: 0 },
        error: runtimeError("INVALID_ARGUMENT", "Unsupported workbook local config format.", { retryable: false })
      };
    }
    if (request.config.workbookId !== request.workbookId) {
      return {
        ok: false,
        workbookId: request.workbookId,
        imported: { templates: 0, regions: 0, permissions: false },
        skipped: { templates: 0, regions: 0 },
        error: runtimeError("INVALID_ARGUMENT", "Config workbookId does not match the import target.", {
          retryable: false,
          details: { configWorkbookId: request.config.workbookId }
        })
      };
    }

    let importedTemplates = 0;
    let skippedTemplates = 0;
    if (includeTemplates) {
      const existingTemplates = this.templates.dump();
      const byId = new Map(existingTemplates.map((template) => [template.templateId, template]));
      for (const rawTemplate of request.config.templates) {
        const template = normalizeImportedTemplate(rawTemplate, request.workbookId);
        if (!template) {
          skippedTemplates += 1;
          continue;
        }
        if (!overwrite && byId.has(template.templateId)) {
          skippedTemplates += 1;
          continue;
        }
        byId.set(template.templateId, template);
        importedTemplates += 1;
      }
      this.templates.load([...byId.values()]);
    }

    let importedRegions = 0;
    let skippedRegions = 0;
    if (includeRegions) {
      for (const rawRegion of request.config.regions) {
        const region = normalizeImportedRegion(rawRegion, request.workbookId);
        if (!region) {
          skippedRegions += 1;
          continue;
        }
        const key = regionKey(region.workbookId, region.name);
        if (!overwrite && this.regions.has(key)) {
          skippedRegions += 1;
          continue;
        }
        this.regions.set(key, region);
        importedRegions += 1;
      }
    }

    let importedPermissions = false;
    if (includePermissions && request.config.permissions !== undefined) {
      this.permissionState = mergePermissionState(
        this.permissionState,
        clonePermissionStateForWorkbook(request.config.permissions, request.workbookId)
      );
      importedPermissions = true;
    }

    this.persistState();
    return {
      ok: true,
      workbookId: request.workbookId,
      imported: {
        templates: importedTemplates,
        regions: importedRegions,
        permissions: importedPermissions
      },
      skipped: {
        templates: skippedTemplates,
        regions: skippedRegions
      }
    };
  }

  async embedWorkbookLocalConfig(workbookId: WorkbookId, options: { includePermissions?: boolean } = {}) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError();
    }
    const permissionWarnings = this.validateDirectMutation(workbookId, [], "workbook");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Embedding workbook local config is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const exported = this.exportWorkbookLocalConfig(workbookId, options);
    return this.applyDirectTransaction(
      {
        workbookId,
        goal: "Embed Open Workbook local config into workbook custom XML",
        scopes: [{ type: "workbook", workbookId }],
        destructiveLevel: "workbook"
      },
      async () => {
        const result = await client.request<WorkbookEmbeddedLocalConfigResponse>("workbook.embed_local_config", {
          workbookId,
          config: exported.config
        });
        return {
          ...result,
          config: exported.config
        };
      }
    );
  }

  async readWorkbookEmbeddedLocalConfig(workbookId: WorkbookId): Promise<WorkbookEmbeddedLocalConfigResponse> {
    const client = this.getActiveAddinClient();
    if (!client) {
      return disconnectedError() as WorkbookEmbeddedLocalConfigResponse;
    }
    const result = await client.request<WorkbookEmbeddedLocalConfigResponse>("workbook.read_embedded_local_config", { workbookId });
    if (result.ok && result.config !== undefined && result.config.workbookId !== workbookId) {
      const mismatch: WorkbookEmbeddedLocalConfigResponse = {
        ok: false,
        workbookId,
        embedded: true,
        error: runtimeError("INVALID_ARGUMENT", "Embedded Open Workbook config belongs to a different workbook id.", {
          retryable: false,
          details: { embeddedWorkbookId: result.config.workbookId }
        })
      };
      if (result.partCount !== undefined) {
        mismatch.partCount = result.partCount;
      }
      return mismatch;
    }
    return result;
  }

  async importWorkbookEmbeddedLocalConfig(input: {
    workbookId: WorkbookId;
    includeTemplates?: boolean;
    includeRegions?: boolean;
    includePermissions?: boolean;
    overwrite?: boolean;
  }): Promise<WorkbookLocalConfigImportResponse | WorkbookEmbeddedLocalConfigResponse> {
    const embedded = await this.readWorkbookEmbeddedLocalConfig(input.workbookId);
    if (!embedded.ok || embedded.config === undefined) {
      return embedded;
    }
    const request: WorkbookLocalConfigImportRequest = {
      workbookId: input.workbookId,
      config: embedded.config
    };
    if (input.includeTemplates !== undefined) {
      request.includeTemplates = input.includeTemplates;
    }
    if (input.includeRegions !== undefined) {
      request.includeRegions = input.includeRegions;
    }
    if (input.includePermissions !== undefined) {
      request.includePermissions = input.includePermissions;
    }
    if (input.overwrite !== undefined) {
      request.overwrite = input.overwrite;
    }
    return this.importWorkbookLocalConfig(request);
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

  getTheme(workbookId: WorkbookId) {
    return unsupportedThemeReport(
      workbookId,
      "get_theme",
      "THEME_READ_UNAVAILABLE",
      "Office.js does not expose a deterministic workbook theme read path in this runtime."
    );
  }

  applyTheme(input: { workbookId: WorkbookId; theme?: unknown }) {
    return unsupportedThemeReport(
      input.workbookId,
      "apply_theme",
      "THEME_APPLY_UNAVAILABLE",
      "Office.js does not expose a deterministic workbook theme apply path in this runtime."
    );
  }

  async copyStyleDimensions(input: StyleCopyRequest, options: { idempotencyKey?: string } = {}) {
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

    return this.applyDirectTransaction(
      {
        workbookId: input.workbookId,
        goal: `Before copying style dimensions to ${input.targetSheetName}`,
        scopes: targetRanges.map(rangeScope),
        destructiveLevel: "format",
        idempotencyKey: options.idempotencyKey
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: input.workbookId,
          reason: `Before copying style dimensions to ${input.targetSheetName}`,
          ranges: targetRanges
        });
        if (!("backup" in backup)) {
          return backup;
        }
        const result = await client.request<StyleCopyResponse>("style.copy_dimensions", input);
        const validation = await this.compareStyleFingerprints({
          workbookId: input.workbookId,
          sourceSheetName: input.sourceSheetName,
          targetSheetName: input.targetSheetName,
          ...(input.sourceAddress !== undefined ? { sourceAddress: input.sourceAddress } : {}),
          ...(input.targetAddress !== undefined ? { targetAddress: input.targetAddress } : {}),
          dimensions: input.dimensions
        });
        const backupIds = backup.backup?.backupId ? [backup.backup.backupId] : [];
        return {
          ok: result.ok,
          backup: backup.backup,
          backups: backupIds,
          rollbackAvailable: backupIds.length > 0,
          warnings: result.warnings ?? [],
          telemetry: { styleCopyCount: 1, validationIssueCount: "issueCount" in validation ? validation.issueCount : 0 },
          result,
          validation
        };
      }
    );
  }

  async copyStyleDimensionsMany(input: { workbookId: WorkbookId; requests: StyleCopyRequest[] }, options: { idempotencyKey?: string } = {}) {
    const client = this.getActiveAddinClient();
    if (!client) {
      return {
        ok: false,
        error: runtimeError("ADDIN_DISCONNECTED", "No Excel add-in session is connected.", { retryable: true })
      };
    }
    if (input.requests.length === 0) {
      return { ok: true, copied: [], copyCount: 0, results: [], warnings: [], telemetry: { styleCopyCount: 0 } };
    }

    const targetRanges = (await Promise.all(input.requests.map((request) =>
      request.targetAddress !== undefined
        ? Promise.resolve([{ workbookId: request.workbookId, sheetName: request.targetSheetName, address: request.targetAddress }])
        : this.getSheetUsedRange(request.workbookId, request.targetSheetName)
    ))).flat();
    const permissionWarnings = this.validateDirectMutation(input.workbookId, targetRanges, "format");
    if (permissionWarnings.length > 0) {
      return {
        ok: false,
        error: runtimeError("PERMISSION_DENIED", "Grouped style copy is blocked by the current Open Workbook permission policy.", {
          retryable: false,
          details: { permissionWarnings }
        }),
        warnings: permissionWarnings
      };
    }

    return this.applyDirectTransaction(
      {
        workbookId: input.workbookId,
        goal: `Before copying ${input.requests.length} grouped style range(s)`,
        scopes: targetRanges.map(rangeScope),
        destructiveLevel: "format",
        idempotencyKey: options.idempotencyKey
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: input.workbookId,
          reason: `Before copying ${input.requests.length} grouped style range(s)`,
          ranges: targetRanges
        });
        if (!("backup" in backup)) {
          return backup;
        }
        const result = await client.request<StyleCopyManyResponse>("style.copy_dimensions_many", input);
        const backupIds = backup.backup?.backupId ? [backup.backup.backupId] : [];
        return {
          ok: result.ok,
          backup: backup.backup,
          backups: backupIds,
          rollbackAvailable: backupIds.length > 0,
          warnings: result.warnings ?? [],
          telemetry: { styleCopyCount: result.copyCount },
          result,
          results: result.results
        };
      }
    );
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

  async getFormulaDependencyGraph(input: FormulaPatternRequest): Promise<{ ok: true; graph: FormulaDependencyGraph } | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const patterns = await this.readFormulaPatterns(input);
    if (!patterns.ok) {
      return patterns;
    }
    const tables = await this.getFormulaGraphTables(input.workbookId);
    return {
      ok: true,
      graph: buildFormulaDependencyGraph(patterns.patterns, { tables })
    };
  }

  async traceFormulaPrecedents(input: FormulaPatternRequest): Promise<FormulaTraceResponse | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const graphResult = await this.getFormulaDependencyGraph(input);
    if (!graphResult.ok) {
      return graphResult;
    }
    const traced = tracePrecedents(graphResult.graph, input.sheetName, input.address ?? graphResult.graph.address);
    return {
      ok: true,
      workbookId: input.workbookId,
      sheetName: input.sheetName,
      address: input.address ?? graphResult.graph.address,
      direction: "precedents",
      nodes: traced.nodes,
      edges: traced.edges,
      warnings: graphResult.graph.warnings
    };
  }

  async traceFormulaDependents(input: FormulaPatternRequest): Promise<FormulaTraceResponse | { ok: false; error: ReturnType<typeof runtimeError> }> {
    const graphResult = await this.getFormulaDependencyGraph({
      workbookId: input.workbookId,
      sheetName: input.sheetName
    });
    if (!graphResult.ok) {
      return graphResult;
    }
    const traced = traceDependents(graphResult.graph, input.sheetName, input.address ?? graphResult.graph.address);
    return {
      ok: true,
      workbookId: input.workbookId,
      sheetName: input.sheetName,
      address: input.address ?? graphResult.graph.address,
      direction: "dependents",
      nodes: traced.nodes,
      edges: traced.edges,
      warnings: graphResult.graph.warnings
    };
  }

  private async getFormulaGraphTables(workbookId: WorkbookId): Promise<TableInfo[]> {
    const listed = await this.listTables(workbookId);
    const tables = (listed as { ok?: boolean; tables?: TableInfo[] }).tables;
    if (!Array.isArray(tables)) {
      return [];
    }
    return tables.filter((table) => table.workbookId === workbookId);
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
    const permissionWarnings = this.validateDirectMutation(input.workbookId, ranges, "structure");
    if (permissionWarnings.length > 0) {
      return permissionDenied("Template repair is blocked by the current Open Workbook permission policy.", permissionWarnings);
    }
    const repairRequest: AddinTemplateRepairRequest = {
      workbookId: input.workbookId,
      templateId: input.templateId,
      sourceSheetName: template.sourceSheetName,
      targetSheetName: input.targetSheetName,
      dataRegions: template.dataRegions,
      repair: input.repair ?? ["styles", "formulas", "dataRegions"]
    };
    return this.applyDirectTransaction(
      {
        workbookId: input.workbookId,
        goal: `Before repairing ${input.targetSheetName} from template ${input.templateId}`,
        scopes: ranges.map(rangeScope),
        destructiveLevel: "structure"
      },
      async () => {
        const backup = await this.createWorkbookBackup({
          workbookId: input.workbookId,
          reason: `Before repairing ${input.targetSheetName} from template ${input.templateId}`,
          ranges
        });
        if (!("backup" in backup)) {
          return backup;
        }
        const result = await client.request("template.repair", repairRequest);
        const validation = await this.validateSheetAgainstTemplate({
          workbookId: input.workbookId,
          templateId: input.templateId,
          targetSheetName: input.targetSheetName
        });
        return {
          ok: (result as { ok?: boolean }).ok !== false,
          backup: backup.backup,
          warnings: (result as { warnings?: OperationWarning[] }).warnings ?? [],
          result,
          validation
        };
      }
    );
  }

  async applyPlan(planId: PlanId, confirmationToken?: string): Promise<OperationResult> {
    const batch = this.plans.createBatchRequest(planId, confirmationToken);
    const result = await this.applyBatch(batch);
    this.plans.markApplyResult(planId, result);
    return { ...result, planId };
  }

  async previewRiskyEdit(input: {
    workbookId: WorkbookId;
    operations: ExcelOperation[];
    reason?: string;
    goal?: string;
    ranges?: A1Range[];
    apply?: boolean;
    allowSparseOverwrite?: boolean;
    confirmationToken?: string;
    agentId?: AgentId;
    agentName?: string;
    taskId?: TaskId;
    role?: string;
  }) {
    const goal = input.goal ?? input.reason ?? "Scoped risky edit";
    if (input.operations.length === 0) {
      return {
        ok: false,
        workflow: "excel.workflow.preview_risky_edit",
        applied: false,
        completedSteps: [],
        errorStep: "operations",
        error: runtimeError("INVALID_ARGUMENT", "previewRiskyEdit requires at least one scoped operation.", {
          retryable: false
        })
      };
    }
    const sparseWarnings = detectSparseOverwriteWarnings(input.operations);
    if (sparseWarnings.length > 0 && input.allowSparseOverwrite !== true) {
      return {
        ok: false,
        workflow: "excel.workflow.preview_risky_edit",
        applied: false,
        completedSteps: [],
        errorStep: "sparse_write_guard",
        warnings: sparseWarnings,
        error: runtimeError("INVALID_ARGUMENT", "Risky workflow blocked a sparse/null-padded range write. Use the smallest changed range, use clear_values_keep_format for explicit clearing, or pass allowSparseOverwrite when this broad overwrite is intentional.", {
          retryable: false,
          details: { warnings: sparseWarnings }
        })
      };
    }
    const ranges = input.ranges?.length ? input.ranges : snapshotRangesFromOperations(input.workbookId, input.operations);
    const before = await this.createWorkbookSnapshot({
      workbookId: input.workbookId,
      reason: `Before ${goal}`,
      ...(ranges.length > 0 ? { ranges } : {})
    });
    if (!before.ok || !("snapshot" in before)) {
      return {
        ok: false,
        workflow: "excel.workflow.preview_risky_edit",
        applied: false,
        completedSteps: [],
        errorStep: "before_snapshot",
        beforeSnapshotResult: before
      };
    }

    const planRequest: PlanCreateRequest = {
      workbookId: input.workbookId,
      goal,
      operations: input.operations,
      baseSnapshotId: before.snapshot.snapshotId
    };
    if (input.agentId !== undefined) {
      planRequest.agentId = input.agentId;
    }
    if (input.agentName !== undefined) {
      planRequest.agentName = input.agentName;
    }
    if (input.taskId !== undefined) {
      planRequest.taskId = input.taskId;
    }
    if (input.role !== undefined) {
      planRequest.role = input.role;
    }
    const plan = this.createPlan(planRequest);
    const planPreview = await this.previewPlan(plan.planId);
    const completedSteps = ["before_snapshot", "plan_create", "plan_preview"];

    if (input.apply === false) {
      return {
        ok: true,
        workflow: "excel.workflow.preview_risky_edit",
        applied: false,
        completedSteps,
        planId: plan.planId,
        beforeSnapshot: before.snapshot,
        planPreview,
        recovery: {
          beforeSnapshotId: before.snapshot.snapshotId,
          rollbackAvailable: false
        },
        nextSteps: ["Apply the previewed plan with excel.plan.apply, then capture an after snapshot, diff, and rollback preview."]
      };
    }

    const applyResult = await this.applyPlan(plan.planId, input.confirmationToken);
    completedSteps.push("plan_apply");
    if (!applyResult.ok) {
      return {
        ok: false,
        workflow: "excel.workflow.preview_risky_edit",
        applied: false,
        completedSteps,
        errorStep: "plan_apply",
        planId: plan.planId,
        transactionId: applyResult.transactionId,
        beforeSnapshot: before.snapshot,
        planPreview,
        applyResult,
        recovery: {
          beforeSnapshotId: before.snapshot.snapshotId,
          transactionId: applyResult.transactionId,
          rollbackAvailable: false
        }
      };
    }

    const after = await this.createWorkbookSnapshot({
      workbookId: input.workbookId,
      reason: `After ${goal}`,
      ranges: before.snapshot.affectedRanges
    });
    if (after.ok && "snapshot" in after) {
      completedSteps.push("after_snapshot");
    }
    const diff = after.ok && "snapshot" in after
      ? this.compareSnapshots(before.snapshot.snapshotId, after.snapshot.snapshotId)
      : undefined;
    if (diff?.ok) {
      completedSteps.push("snapshot_diff");
    }
    const rollbackPreview = applyResult.transactionId !== undefined
      ? this.previewTransactionRollback(applyResult.transactionId)
      : undefined;
    if (rollbackPreview !== undefined) {
      completedSteps.push("rollback_preview");
    }

    return {
      ok: Boolean(after.ok && diff?.ok),
      workflow: "excel.workflow.preview_risky_edit",
      applied: true,
      completedSteps,
      summary: riskyEditSummary(applyResult, diff, rollbackPreview),
      planId: plan.planId,
      transactionId: applyResult.transactionId,
      beforeSnapshot: before.snapshot,
      afterSnapshot: after.ok && "snapshot" in after ? after.snapshot : undefined,
      planPreview,
      applyResult,
      diff,
      rollbackPreview,
      recovery: {
        beforeSnapshotId: before.snapshot.snapshotId,
        afterSnapshotId: after.ok && "snapshot" in after ? after.snapshot.snapshotId : undefined,
        transactionId: applyResult.transactionId,
        rollbackAvailable: rollbackPreview?.rollbackAvailable ?? false
      }
    };
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

    const operations = await this.createRollbackOperations(planId);
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
      planId,
      mode: "apply",
      operations
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
    if (request.mode !== "apply") {
      return this.applyBatchDirect(request);
    }
    const returnQueuedProgress = this.isRuntimeMutationBusy();
    const scheduled = this.scheduleBatch(request);
    void scheduled.promise.then((result) => {
      if (result.ok) {
        this.bumpWorkbookContentVersion(request.workbookId);
        if (batchInvalidatesWorkbookMetadata(request)) {
          this.agent.invalidateWorkbook(request.workbookId);
        }
      }
    }, () => undefined);
    if (returnQueuedProgress) {
      void scheduled.promise.catch(() => undefined);
      return queuedOperationResult(this.transactions.withQueueMetadata(scheduled.transaction));
    }
    return scheduled.promise;
  }

  submitBatch(request: BatchRequest) {
    if (request.mode !== "apply") {
      return {
        ok: false,
        error: runtimeError("INVALID_ARGUMENT", "Only apply-mode batches can be submitted to the mutation queue.", { retryable: false })
      };
    }
    const scheduled = this.scheduleBatch(request);
    void scheduled.promise.catch(() => undefined);
    const transaction = this.transactions.withQueueMetadata(scheduled.transaction);
    return {
      ok: true,
      transactionId: transaction.transactionId,
      status: transaction.status,
      queuePosition: transaction.queuePosition,
      progressMessage: transaction.progressMessage,
      transaction
    };
  }

  private scheduleBatch(request: BatchRequest): { transaction: TransactionRecord; promise: Promise<OperationResult> } {
    const compiled = this.compiler.compile(request);
    const agentId = request.agentId ?? this.currentAgentId();
    const scopes = scopesFromBatch(request.workbookId, request.operations);
    const transaction = this.transactions.create({
      workbookId: request.workbookId,
      agentId,
      taskId: request.taskId,
      planId: request.planId,
      goal: request.operations.map((operation) => operation.reason || operation.kind).join("; ") || "Apply Excel batch",
      scopes,
      baseFingerprints: request.expectedTargetFingerprints ?? [],
      destructiveLevel: compiled.destructiveLevel,
      progressMessage: request.progressMessage,
      retryStrategy: request.retryStrategy,
      chunksTotal: request.chunksTotal,
      chunksCompleted: request.chunksCompleted
    });
    if (request.taskId !== undefined) {
      this.tasks.attachTransaction(request.taskId, transaction.transactionId);
      this.updateTask(request.taskId, { status: "queued" });
    }
    this.recordCollabEvent({
      type: "transaction.queued",
      workbookId: request.workbookId,
      agentId,
      taskId: request.taskId,
      transactionId: transaction.transactionId,
      message: `Transaction queued: ${transaction.goal}`
    });

    const promise = this.enqueueTransaction(transaction.transactionId, async () => {
      const lockResult = this.locks.acquire({
        workbookId: request.workbookId,
        ownerAgentId: agentId,
        taskId: request.taskId,
        transactionId: transaction.transactionId,
        scopes,
        mode: lockModeForDestructiveLevel(compiled.destructiveLevel),
        ttlMs: lockTtl(this.lockLeasePolicy.transactionTtlMs, this.lockLeasePolicy),
        reason: transaction.goal
      });
      if (!lockResult.ok) {
        this.conflicts.push(...lockResult.conflicts);
        for (const conflict of lockResult.conflicts) {
          this.recordConflictTelemetry(conflict);
          this.recordCollabEvent({
            type: "conflict.detected",
            workbookId: request.workbookId,
            agentId,
            taskId: request.taskId,
            transactionId: transaction.transactionId,
            message: conflict.message,
            details: { conflict }
          });
        }
        this.transactions.markBlocked(transaction.transactionId, "LOCK_CONFLICT", "Transaction conflicts with an active lock.");
        if (request.taskId !== undefined) {
          this.updateTask(request.taskId, { status: "blocked", currentStep: "Waiting for conflicting workbook lock" });
        }
        this.persistState();
        return {
          ok: false,
          transactionId: transaction.transactionId,
          transactionStatus: "blocked",
          progressMessage: "Workbook mutation is blocked by an active lock.",
          taskId: request.taskId,
          agentId,
          rollbackAvailable: compiled.requiredBackups.length > 0,
          backups: [],
          warnings: lockResult.conflicts.map((conflict) => ({
            code: conflict.code,
            message: conflict.message,
            details: { conflict }
          })),
          telemetry: { warningCount: lockResult.conflicts.length },
          error: runtimeError("LOCK_CONFLICT", "Transaction conflicts with active workbook work. Wait for the lock or refresh the plan.", {
            retryable: true
          })
        };
      }

      this.transactions.markApplying(transaction.transactionId, lockResult.locks.map((lock) => lock.lockId));
      for (const lock of lockResult.locks) {
        this.recordCollabEvent({
          type: "lock.acquired",
          workbookId: request.workbookId,
          agentId,
          taskId: request.taskId,
          transactionId: transaction.transactionId,
          lockId: lock.lockId,
          message: `Lock acquired for transaction ${transaction.transactionId}.`,
          details: { lock }
        });
      }
      if (request.taskId !== undefined) {
        this.updateTask(request.taskId, { status: "applying" });
      }
      this.recordCollabEvent({
        type: "transaction.applying",
        workbookId: request.workbookId,
        agentId,
        taskId: request.taskId,
        transactionId: transaction.transactionId,
        message: `Transaction applying: ${transaction.goal}`
      });

      try {
        const result = await this.applyBatchDirect({ ...request, agentId });
        const enriched: OperationResult = {
          ...result,
          transactionId: transaction.transactionId,
          transactionStatus: result.ok ? "applied" : "failed",
          progressMessage: result.ok ? "Workbook mutation applied successfully." : (result.error?.message ?? "Workbook mutation failed."),
          agentId
        };
        if (request.taskId !== undefined) {
          enriched.taskId = request.taskId;
        }
        if (result.ok) {
          this.transactions.markApplied(transaction.transactionId, {
            backups: enriched.backups,
            warnings: enriched.warnings,
            diffSummary: enriched.diffSummary,
            telemetry: enriched.telemetry
          });
          if (request.taskId !== undefined) {
            this.tasks.attachBackups(request.taskId, enriched.backups);
            this.updateTask(request.taskId, { status: "completed" });
          }
          this.recordCollabEvent({
            type: "transaction.applied",
            workbookId: request.workbookId,
            agentId,
            taskId: request.taskId,
            transactionId: transaction.transactionId,
            message: `Transaction applied: ${transaction.goal}`
          });
        } else {
          this.transactions.markFailed(
            transaction.transactionId,
            result.error?.code ?? "TRANSACTION_FAILED",
            result.error?.message ?? "Transaction failed.",
            result.warnings
          );
          if (request.taskId !== undefined) {
            this.updateTask(request.taskId, { status: "failed", errorMessage: result.error?.message ?? "Transaction failed." });
          }
          this.recordCollabEvent({
            type: "transaction.failed",
            workbookId: request.workbookId,
            agentId,
            taskId: request.taskId,
            transactionId: transaction.transactionId,
            message: result.error?.message ?? `Transaction failed: ${transaction.goal}`
          });
        }
        return enriched;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        const code = /timed out|timeout/i.test(message) ? "TIMEOUT" : "TRANSACTION_FAILED";
        if (code === "TIMEOUT" && shouldRetryStyleBatch(request)) {
          const chunks = chunkOperationsForRetry(request.operations);
          const warnings: OperationWarning[] = [
            {
              code: "RETRYING_SMALLER_BATCH",
              message: `Style batch timed out and was resubmitted as ${chunks.length} smaller queued chunks.`
            }
          ];
          this.transactions.markFailed(transaction.transactionId, code, message, warnings);
          if (request.taskId !== undefined) {
            this.updateTask(request.taskId, { status: "queued", currentStep: "Retrying style update in smaller chunks" });
          }
          this.recordCollabEvent({
            type: "transaction.failed",
            workbookId: request.workbookId,
            agentId,
            taskId: request.taskId,
            transactionId: transaction.transactionId,
            message: `${message}; retrying style update in smaller chunks.`
          });
          const retryTransactions = chunks.map((chunk, index) =>
            this.submitBatch({
              ...request,
              operations: chunk,
              retryStrategy: "retry_timeout_split_style_entries",
              chunksTotal: chunks.length,
              chunksCompleted: index,
              progressMessage: `Retrying style update chunk ${index + 1} of ${chunks.length}.`
            })
          );
          return {
            ok: true,
            transactionId: transaction.transactionId,
            transactionStatus: "failed",
            progressMessage: `Style batch timed out, so Open Workbook queued ${chunks.length} smaller retry chunks.`,
            taskId: request.taskId,
            agentId,
            rollbackAvailable: false,
            backups: [],
            warnings,
            telemetry: { warningCount: warnings.length },
            data: {
              retryStrategy: "retry_timeout_split_style_entries",
              chunksTotal: chunks.length,
              retryTransactionIds: retryTransactions.map((retry: any) => retry.transactionId).filter(Boolean),
              retryTransactions
            }
          };
        }
        this.transactions.markFailed(transaction.transactionId, code, message);
        if (request.taskId !== undefined) {
          this.updateTask(request.taskId, { status: "failed", errorMessage: message });
        }
        this.recordCollabEvent({
          type: "transaction.failed",
          workbookId: request.workbookId,
          agentId,
          taskId: request.taskId,
          transactionId: transaction.transactionId,
          message
        });
        return {
          ok: false,
          transactionId: transaction.transactionId,
          transactionStatus: "failed",
          progressMessage: message,
          taskId: request.taskId,
          agentId,
          rollbackAvailable: compiled.requiredBackups.length > 0,
          backups: [],
          warnings: [],
          telemetry: { warningCount: 0 },
          error: runtimeError(code === "TIMEOUT" ? "TIMEOUT" : "OPERATION_FAILED", message, { retryable: code === "TIMEOUT" })
        };
      } finally {
        const releasedLocks = this.locks.release(lockResult.locks.map((lock) => lock.lockId));
        this.markConflictTelemetryClearedByLock(releasedLocks.map((lock) => lock.lockId));
        for (const lock of releasedLocks) {
          this.recordCollabEvent({
            type: "lock.released",
            workbookId: request.workbookId,
            agentId,
            taskId: request.taskId,
            transactionId: transaction.transactionId,
            lockId: lock.lockId,
            message: `Lock released for transaction ${transaction.transactionId}.`,
            details: { lock }
          });
        }
        if (releasedLocks.length === 0) {
          this.persistState();
        }
      }
    });
    return { transaction, promise };
  }

  private enqueueTransaction(transactionId: TransactionId, work: () => Promise<OperationResult>): Promise<OperationResult> {
    return this.enqueueRuntimeMutation(async () => {
      const transaction = this.transactions.get(transactionId);
      if (this.cancelledQueuedTransactions.delete(transactionId) || transaction?.status === "cancelled") {
        return cancelledOperationResult(transactionId);
      }
      return work();
    });
  }

  private enqueueRuntimeMutation<T>(work: () => Promise<T>): Promise<T> {
    this.runtimeMutationQueuedCount += 1;
    const run = this.transactionQueue.then(
      () => this.runQueuedRuntimeMutation(work),
      () => this.runQueuedRuntimeMutation(work)
    );
    this.transactionQueue = run.then(
      () => undefined,
      () => undefined
    );
    return run;
  }

  private async runQueuedRuntimeMutation<T>(work: () => Promise<T>): Promise<T> {
    this.runtimeMutationQueuedCount = Math.max(0, this.runtimeMutationQueuedCount - 1);
    this.runtimeMutationActive = true;
    try {
      return await work();
    } finally {
      this.runtimeMutationActive = false;
    }
  }

  private isRuntimeMutationBusy(): boolean {
    return this.runtimeMutationActive || this.runtimeMutationQueuedCount > 0;
  }

  private async applyDirectTransaction<T>(
    input: {
      workbookId: WorkbookId;
      goal: string;
      scopes: WorkbookScope[];
      destructiveLevel: DestructiveLevel;
      taskId?: TaskId | undefined;
      agentId?: AgentId | undefined;
      idempotencyKey?: string | undefined;
    },
    work: () => Promise<T>
  ): Promise<T | { ok: false; transactionId: TransactionId; rollbackAvailable: false; backups: []; warnings: OperationWarning[]; telemetry: { warningCount: number }; error: ReturnType<typeof runtimeError> }> {
    if (input.idempotencyKey !== undefined && this.directMutationResults.has(input.idempotencyKey)) {
      return this.directMutationResults.get(input.idempotencyKey) as T;
    }
    const agentId = input.agentId ?? this.currentAgentId();
    const transaction = this.transactions.create({
      workbookId: input.workbookId,
      agentId,
      taskId: input.taskId,
      goal: input.goal,
      scopes: input.scopes,
      destructiveLevel: input.destructiveLevel
    });
    this.recordCollabEvent({
      type: "transaction.queued",
      workbookId: input.workbookId,
      agentId,
      taskId: input.taskId,
      transactionId: transaction.transactionId,
      message: `Transaction queued: ${transaction.goal}`
    });
    if (input.taskId !== undefined) {
      this.tasks.attachTransaction(input.taskId, transaction.transactionId);
      this.updateTask(input.taskId, { status: "queued" });
    }

    return this.enqueueRuntimeMutation(async () => {
      const lockResult = this.locks.acquire({
        workbookId: input.workbookId,
        ownerAgentId: agentId,
        taskId: input.taskId,
        transactionId: transaction.transactionId,
        scopes: input.scopes,
        mode: lockModeForDestructiveLevel(input.destructiveLevel),
        ttlMs: lockTtl(this.lockLeasePolicy.transactionTtlMs, this.lockLeasePolicy),
        reason: input.goal
      });
      if (!lockResult.ok) {
        this.conflicts.push(...lockResult.conflicts);
        for (const conflict of lockResult.conflicts) {
          this.recordConflictTelemetry(conflict);
          this.recordCollabEvent({
            type: "conflict.detected",
            workbookId: input.workbookId,
            agentId,
            taskId: input.taskId,
            transactionId: transaction.transactionId,
            message: conflict.message,
            details: { conflict }
          });
        }
        this.transactions.markBlocked(transaction.transactionId, "LOCK_CONFLICT", "Direct transaction conflicts with an active lock.", lockResult.conflicts.map((conflict) => ({
          code: conflict.code,
          message: conflict.message,
          details: { conflict }
        })));
        if (input.taskId !== undefined) {
          this.updateTask(input.taskId, { status: "blocked", currentStep: "Waiting for conflicting workbook lock" });
        } else {
          this.persistState();
        }
        const blockedResult = {
          ok: false as const,
          transactionId: transaction.transactionId,
          rollbackAvailable: false as const,
          backups: [] as [],
          warnings: lockResult.conflicts.map((conflict) => ({
            code: conflict.code,
            message: conflict.message,
            details: { conflict }
          })),
          telemetry: { warningCount: lockResult.conflicts.length },
          error: runtimeError("LOCK_CONFLICT", "Direct transaction conflicts with active workbook work. Wait for the lock or split the task scope.", {
            retryable: true
          })
        };
        return this.cacheDirectMutationResult(input.idempotencyKey, blockedResult);
      }

      this.transactions.markApplying(transaction.transactionId, lockResult.locks.map((lock) => lock.lockId));
      this.recordCollabEvent({
        type: "transaction.applying",
        workbookId: input.workbookId,
        agentId,
        taskId: input.taskId,
        transactionId: transaction.transactionId,
        message: `Transaction applying: ${transaction.goal}`
      });
      if (input.taskId !== undefined) {
        this.updateTask(input.taskId, { status: "applying" });
      }
      try {
        const result = await work();
        const resultRecord = result as { ok?: boolean; backup?: { backupId?: BackupId }; warnings?: OperationWarning[]; error?: { code?: string; message?: string } };
        const backups = resultRecord.backup?.backupId ? [resultRecord.backup.backupId] : [];
        const warnings = resultRecord.warnings ?? [];
        if (resultRecord.ok === false) {
          this.transactions.markFailed(transaction.transactionId, resultRecord.error?.code ?? "TRANSACTION_FAILED", resultRecord.error?.message ?? "Direct transaction failed.", warnings);
          if (input.taskId !== undefined) {
            this.updateTask(input.taskId, { status: "failed", errorMessage: resultRecord.error?.message ?? "Direct transaction failed." });
          }
          this.recordCollabEvent({
            type: "transaction.failed",
            workbookId: input.workbookId,
            agentId,
            taskId: input.taskId,
            transactionId: transaction.transactionId,
            message: resultRecord.error?.message ?? `Transaction failed: ${transaction.goal}`
          });
        } else {
          this.transactions.markApplied(transaction.transactionId, { backups, warnings });
          if (input.taskId !== undefined) {
            this.tasks.attachBackups(input.taskId, backups);
            this.updateTask(input.taskId, { status: "completed" });
          }
          this.recordCollabEvent({
            type: "transaction.applied",
            workbookId: input.workbookId,
            agentId,
            taskId: input.taskId,
            transactionId: transaction.transactionId,
            message: `Transaction applied: ${transaction.goal}`
          });
        }
        if (resultRecord.ok !== false) {
          this.bumpWorkbookContentVersion(input.workbookId);
          if (transactionInvalidatesWorkbookMetadata(input)) {
            this.agent.invalidateWorkbook(input.workbookId);
          }
        }
        if (typeof result === "object" && result !== null && !Array.isArray(result)) {
          return this.cacheDirectMutationResult(input.idempotencyKey, {
            ...result,
            transactionId: transaction.transactionId,
            backups: Array.isArray((result as { backups?: unknown }).backups) ? ((result as unknown) as { backups: unknown[] }).backups : backups,
            rollbackAvailable: Boolean((result as { rollbackAvailable?: unknown }).rollbackAvailable) || backups.length > 0,
            warnings
          } as T);
        }
        return this.cacheDirectMutationResult(input.idempotencyKey, result);
      } finally {
        const releasedLocks = this.locks.release(lockResult.locks.map((lock) => lock.lockId));
        this.markConflictTelemetryClearedByLock(releasedLocks.map((lock) => lock.lockId));
        for (const lock of releasedLocks) {
          this.recordCollabEvent({
            type: "lock.released",
            workbookId: input.workbookId,
            agentId,
            taskId: input.taskId,
            transactionId: transaction.transactionId,
            lockId: lock.lockId,
            message: `Lock released for transaction ${transaction.transactionId}.`,
            details: { lock }
          });
        }
      }
    });
  }

  private cacheDirectMutationResult<T>(idempotencyKey: string | undefined, result: T): T {
    if (idempotencyKey !== undefined) {
      this.directMutationResults.set(idempotencyKey, result);
    }
    return result;
  }

  private async applyBatchDirect(request: BatchRequest): Promise<OperationResult> {
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
        ? await Promise.all(compiled.requiredBackups.map(async (kind) => {
            const backup = this.backups.createBackup({
              workbookId: request.workbookId,
              kind,
              reason: `Before ${request.operations.map((operation) => operation.kind).join(", ")}`,
              affectedRanges: compiled.targetFingerprints.map((fingerprint) => fingerprint.range)
            });
            if (kind === "region" && beforeSnapshot !== undefined) {
              backup.payloadRef = await this.persistBackupPayload(backup.backupId, beforeSnapshot);
            }
            return backup;
          }))
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
    if (backups.length > 0) {
      void this.applyDefaultBackupRetention("apply_batch");
    }
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
    const result = changedCells > 0 ? await this.writeChangedCleanValues(target, read.values, values, action.replace(/_/g, " ")) : undefined;
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

  private async writeChangedCleanValues(target: A1Range, before: CellMatrix, after: CellMatrix, reason: string): Promise<OperationResult | undefined> {
    const operations = changedValueRunOperations(target, before, after, reason);
    if (operations.length === 0) {
      return undefined;
    }
    return this.applyBatch({
      workbookId: target.workbookId,
      mode: "apply",
      operations
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
    if (!activeSession || this.isSessionStale(activeSession)) {
      return undefined;
    }
    return this.addinClients.get(activeSession.connectionId);
  }

  private sessionAgeMs(session: { lastSeenAt: string }): number {
    const parsed = Date.parse(session.lastSeenAt);
    return Number.isFinite(parsed) ? Math.max(0, Date.now() - parsed) : Number.POSITIVE_INFINITY;
  }

  private isSessionStale(session: { lastSeenAt: string }): boolean {
    return this.sessionAgeMs(session) > addinStaleTtlMs();
  }

  private connectionStateFor(session: { lastSeenAt: string; activeWorkbook?: WorkbookRef } | undefined): AddinConnectionState {
    if (!session) {
      return "disconnected";
    }
    if (this.isSessionStale(session)) {
      return "stale";
    }
    return session.activeWorkbook ? "ready" : "connected_no_workbook";
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

  private async createRollbackOperations(planId: PlanId): Promise<ExcelOperation[]> {
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
      const snapshot = backup ? await this.loadBackupPayload(backup) : undefined;
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

  private async getPivotTemplateCopyRanges(workbookId: WorkbookId, targetSheetName: string | undefined, targetAddress: string | undefined): Promise<A1Range[]> {
    if (targetSheetName !== undefined && targetAddress !== undefined) {
      return [{ workbookId, sheetName: targetSheetName, address: stripSheetName(targetAddress) }];
    }
    if (targetSheetName !== undefined) {
      const ranges = await this.getSheetUsedRange(workbookId, targetSheetName);
      if (ranges.length > 0) {
        return ranges;
      }
    }
    return this.getUsedRangesForSnapshot(workbookId);
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

  private getExportDirectory(): string {
    return process.env.OPEN_WORKBOOK_EXPORT_DIR ?? path.join(process.cwd(), ".open-workbook", "exports");
  }

  private defaultWorkbookExportPath(workbookId: WorkbookId): string {
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    return path.join(this.getExportDirectory(), `${sanitizeFileName(workbookId)}-${timestamp}.xlsx`);
  }

  private defaultWorkbookFileBackupPath(workbookId: WorkbookId): string {
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    return path.join(this.getBackupDirectory(), "files", `${sanitizeFileName(workbookId)}-${timestamp}.xlsx`);
  }

  private async applyDefaultBackupRetention(reason: string): Promise<void> {
    const request = defaultBackupRetentionRequest();
    if (!request) {
      return;
    }
    try {
      await this.pruneFileBackups(request);
    } catch (error) {
      this.recordCollabEvent({
        type: "backup.pruned",
        message: "Automatic backup retention failed.",
        details: { reason, error: error instanceof Error ? error.message : String(error) }
      });
    }
  }

  private describePersistedBackup(backup: BackupRecord) {
    const manifest = backup.kind === "file-copy" ? backup.payload as WorkbookFileBackupManifest | undefined : undefined;
    const payloadPath = this.backupPayloadPath(backup);
    return {
      backup,
      ...(manifest !== undefined ? { manifest } : {}),
      payload: {
        kind: backup.kind === "file-copy" ? "file-copy" : "snapshot-json",
        path: payloadPath,
        pinned: backup.pinned === true
      }
    };
  }

  private async getBackupPayloadInfo(backup: BackupRecord): Promise<BackupPayloadInfo> {
    const payloadPath = this.backupPayloadPath(backup);
    if (!payloadPath) {
      return {};
    }
    try {
      const fileStat = await stat(payloadPath);
      return { path: payloadPath, bytes: fileStat.size, missing: false };
    } catch {
      return { path: payloadPath, missing: true };
    }
  }

  private backupPayloadPath(backup: BackupRecord): string | undefined {
    if (backup.kind === "file-copy") {
      const manifest = backup.payload as WorkbookFileBackupManifest | undefined;
      return manifest?.filePath ?? backup.payloadRef;
    }
    return backup.payloadRef?.endsWith(".json") ? backup.payloadRef : undefined;
  }

  private async writeWorkbookFileContent(targetPath: string, content: WorkbookFileContent): Promise<string> {
    const resolved = path.resolve(targetPath);
    await mkdir(path.dirname(resolved), { recursive: true });
    await writeFile(resolved, Buffer.from(content.base64, "base64"));
    return resolved;
  }

  private async hashFile(filePath: string): Promise<string> {
    const hash = createHash("sha256");
    hash.update(await readFile(filePath));
    return `sha256:${hash.digest("hex")}`;
  }

  private async unlinkBackupPayloadIfSafe(backup: BackupRecord): Promise<BackupPayloadInfo> {
    const payload = await this.getBackupPayloadInfo(backup);
    if (!payload.path) {
      return payload;
    }
    const resolvedPayload = path.resolve(payload.path);
    const resolvedBackupDir = path.resolve(this.getBackupDirectory());
    const canDelete = backup.kind === "file-copy" || (backup.payloadRef?.endsWith(".json") === true && isPathInside(resolvedPayload, resolvedBackupDir));
    if (!canDelete) {
      return { ...payload, skipped: true };
    }
    try {
      await unlink(resolvedPayload);
      return { ...payload, deleted: true };
    } catch {
      // Missing backup files are reported by verify; deletion should remain idempotent.
      return { ...payload, missing: true };
    }
  }
}

function sanitizeFileName(value: string): string {
  return value.replace(/[^a-zA-Z0-9._-]+/g, "_").replace(/^_+|_+$/g, "") || "workbook";
}

interface BackupPayloadInfo {
  path?: string | undefined;
  bytes?: number | undefined;
  missing?: boolean | undefined;
  deleted?: boolean | undefined;
  skipped?: boolean | undefined;
}

interface PruneCandidate {
  backup: BackupRecord;
  reasons: string[];
  bytes?: number | undefined;
  payloadPath?: string | undefined;
  missingPayload?: boolean | undefined;
}

function isPersistedBackup(backup: BackupRecord): boolean {
  return backup.kind === "file-copy" || backup.payloadRef?.endsWith(".json") === true;
}

function backupMatchesRetentionKind(backup: BackupRecord, kind: NonNullable<WorkbookBackupRetentionRequest["kind"]>): boolean {
  if (kind === "all") {
    return true;
  }
  if (kind === "file-copy") {
    return backup.kind === "file-copy";
  }
  return backup.kind !== "file-copy" && backup.payloadRef?.endsWith(".json") === true;
}

function addPruneCandidate(candidates: Map<BackupId, PruneCandidate>, backup: BackupRecord, reason: string, payload?: BackupPayloadInfo): void {
  const existing = candidates.get(backup.backupId);
  if (existing) {
    if (!existing.reasons.includes(reason)) {
      existing.reasons.push(reason);
    }
    if (payload?.bytes !== undefined) existing.bytes = payload.bytes;
    if (payload?.path !== undefined) existing.payloadPath = payload.path;
    if (payload?.missing !== undefined) existing.missingPayload = payload.missing;
    return;
  }
  candidates.set(backup.backupId, {
    backup,
    reasons: [reason],
    bytes: payload?.bytes,
    payloadPath: payload?.path,
    missingPayload: payload?.missing
  });
}

function defaultBackupRetentionRequest(): WorkbookBackupRetentionRequest | undefined {
  if (process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED === "1") {
    return undefined;
  }
  return {
    kind: "all",
    maxAgeDays: positiveEnvInteger("OPEN_WORKBOOK_BACKUP_RETENTION_DAYS", 30),
    maxBackupsPerWorkbook: positiveEnvInteger("OPEN_WORKBOOK_BACKUP_RETENTION_COUNT", 20),
    maxTotalBytes: positiveEnvInteger("OPEN_WORKBOOK_BACKUP_RETENTION_BYTES", 1024 * 1024 * 1024)
  };
}

function positiveEnvInteger(name: string, fallback: number): number {
  const raw = process.env[name];
  if (raw === undefined) {
    return fallback;
  }
  const parsed = Number(raw);
  return Number.isFinite(parsed) && parsed >= 0 ? Math.floor(parsed) : fallback;
}

function isPathInside(child: string, parent: string): boolean {
  const relative = path.relative(parent, child);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function pivotOperationCapabilityStatus(
  operation: PivotOperationCapabilityStatus["operation"],
  dimensions?: PivotCopyFromTemplateRequest["dimensions"]
): PivotOperationCapabilityStatus {
  if (operation === "pivot.update_source") {
    const warnings: OperationWarning[] = [
      {
        code: "PIVOT_SOURCE_REASSIGNMENT_UNSUPPORTED",
        message: "Office.js does not expose safe in-place PivotTable source reassignment in this runtime.",
        details: {
          safeAlternative: "Use excel.pivot.rebuild_with_source to create a PivotTable from the desired source and optionally replay template settings."
        }
      }
    ];
    return {
      operation,
      capabilities: [
        {
          capability: "source_reassignment",
          status: "unsupported",
          reason: "No deterministic cross-platform Office.js API is available for in-place PivotTable source reassignment."
        },
        {
          capability: "rebuild_with_source",
          status: "partial",
          reason: "Safe source changes are implemented as explicit create or delete/create transactions instead of mutating the existing source pointer."
        }
      ],
      warnings,
      fallback: "excel.pivot.rebuild_with_source"
    };
  }

  if (operation === "pivot.rebuild_with_source") {
    return {
      operation,
      capabilities: [
        {
          capability: "rebuild_with_source",
          status: "partial",
          reason: "Rebuild changes the source by creating a new PivotTable, with optional delete/create replacement and template replay."
        },
        {
          capability: "source_reassignment",
          status: "unsupported",
          reason: "The original PivotTable source is not reassigned in place."
        }
      ],
      warnings: [
        {
          code: "PIVOT_REBUILD_NOT_IN_PLACE",
          message: "PivotTable source changes are applied through rebuild/create semantics, not in-place source reassignment.",
          details: { safeAlternativeForInPlaceReassignment: "excel.pivot.rebuild_with_source" }
        }
      ]
    };
  }

  const requestedDimensions = dimensions?.length ? dimensions : ["metadata", "layout", "fields", "dataFields", "numberFormats", "filters", "refresh"];
  const warnings: OperationWarning[] = [
    {
      code: "PIVOT_TEMPLATE_COPY_PARTIAL",
      message: "PivotTable template copy replays deterministic Office.js dimensions only.",
      details: {
        requestedDimensions,
        replayedWhenRequested: [
          "metadata",
          "layout flags",
          "row hierarchy membership and order",
          "column hierarchy membership and order",
          "filter hierarchy membership and order",
          "data hierarchy membership and order",
          "data aggregation",
          "data number formats",
          "basic field settings",
          "refresh"
        ],
        notReplayed: [
          "source reassignment",
          "PivotChart-specific settings",
          "slicers and timelines",
          "item-level manual filters and sorts not exposed by Office.js",
          "grouping details not exposed by Office.js",
          "calculated fields/items when not exposed by Office.js",
          "host-specific PivotTable settings without deterministic Office.js setters"
        ]
      }
    }
  ];
  return {
    operation,
    capabilities: [
      {
        capability: "template_copy",
        status: "partial",
        reason: "Only deterministic PivotTable metadata and layout dimensions exposed by Office.js are replayed."
      },
      {
        capability: "source_reassignment",
        status: "unsupported",
        reason: "Template copy does not change the target PivotTable source."
      },
      {
        capability: "pivot_chart",
        status: "partial",
        reason: "PivotChart-specific settings are not part of PivotTable template replay."
      }
    ],
    warnings,
    fallback: "Use excel.pivot.rebuild_with_source when the target source must change."
  };
}

function mergeOperationWarnings(...groups: OperationWarning[][]): OperationWarning[] {
  const warnings = groups.flat();
  const seen = new Set<string>();
  return warnings.filter((warning) => {
    const key = `${warning.code}:${warning.message}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function makePivotFingerprint(info: PivotTableInfo): PivotFingerprint {
  const warnings: PivotFingerprint["warnings"] = [];
  if (!info.source) {
    warnings.push({ code: "PIVOT_SOURCE_UNAVAILABLE", message: "Pivot source is unavailable from Office.js." });
  }
  if (!info.hierarchies?.length) {
    warnings.push({ code: "PIVOT_SOURCE_FIELDS_UNAVAILABLE", message: "Pivot source fields are unavailable from Office.js." });
  }
  const source: NonNullable<PivotFingerprint["source"]> = {
    fields: uniqueDefined((info.hierarchies ?? []).map((hierarchy) => hierarchy.name)).sort()
  };
  if (info.sourceType !== undefined) {
    source.type = info.sourceType;
  }
  if (info.source !== undefined) {
    source.value = info.source;
  }
  const dataFields: PivotFingerprint["layout"]["dataFields"] = (info.dataHierarchies ?? []).map((hierarchy) => {
    const field: PivotFingerprint["layout"]["dataFields"][number] = { name: hierarchy.name };
    if (hierarchy.field?.name !== undefined) {
      field.sourceFieldName = hierarchy.field.name;
    }
    if (hierarchy.summarizeBy !== undefined) {
      field.summarizeBy = hierarchy.summarizeBy;
    }
    if (hierarchy.numberFormat !== undefined) {
      field.numberFormat = hierarchy.numberFormat;
    }
    return field;
  });
  const layout: PivotFingerprint["layout"] = {
    rowFields: uniqueDefined((info.rowHierarchies ?? []).map((hierarchy) => hierarchy.name)),
    columnFields: uniqueDefined((info.columnHierarchies ?? []).map((hierarchy) => hierarchy.name)),
    filterFields: uniqueDefined((info.filterHierarchies ?? []).map((hierarchy) => hierarchy.name)),
    dataFields
  };
  if (info.layout !== undefined) {
    layout.flags = info.layout;
  }
  const fingerprint: Omit<PivotFingerprint, "hash"> = {
    workbookId: info.workbookId,
    pivotTableName: info.pivotTableName,
    capturedAt: new Date().toISOString(),
    source,
    layout,
    warnings
  };
  if (info.range !== undefined) {
    fingerprint.output = { ...info.range };
    if (info.sheetName !== undefined) {
      fingerprint.output.sheetName = info.sheetName;
    }
  }
  return {
    ...fingerprint,
    hash: hashStable(normalizePivotFingerprintForHash(fingerprint))
  };
}

function normalizePivotFingerprintForHash(fingerprint: Omit<PivotFingerprint, "hash">): unknown {
  return {
    source: fingerprint.source,
    layout: fingerprint.layout,
    output: fingerprint.output
      ? {
          sheetName: fingerprint.output.sheetName,
          rowCount: fingerprint.output.rowCount,
          columnCount: fingerprint.output.columnCount
        }
      : undefined
  };
}

function diffPivotFingerprints(source: PivotFingerprint, target: PivotFingerprint, targetPivotTableName: string): PivotDiff {
  const changes: PivotDiff["changes"] = [];
  addPivotDiff(changes, "source.type", source.source?.type, target.source?.type);
  addPivotDiff(changes, "source.value", source.source?.value, target.source?.value);
  addPivotDiff(changes, "source.fields", source.source?.fields ?? [], target.source?.fields ?? []);
  addPivotDiff(changes, "layout.rowFields", source.layout.rowFields, target.layout.rowFields);
  addPivotDiff(changes, "layout.columnFields", source.layout.columnFields, target.layout.columnFields);
  addPivotDiff(changes, "layout.filterFields", source.layout.filterFields, target.layout.filterFields);
  addPivotDiff(changes, "layout.dataFields", source.layout.dataFields, target.layout.dataFields);
  addPivotDiff(changes, "layout.flags", source.layout.flags ?? {}, target.layout.flags ?? {});
  addPivotDiff(changes, "output.shape", source.output ? { rowCount: source.output.rowCount, columnCount: source.output.columnCount } : undefined, target.output ? { rowCount: target.output.rowCount, columnCount: target.output.columnCount } : undefined);
  return {
    ok: changes.length === 0,
    workbookId: source.workbookId,
    sourcePivotTableName: source.pivotTableName,
    targetPivotTableName,
    source,
    target,
    changes,
    warnings: [...source.warnings, ...target.warnings]
  };
}

function addPivotDiff(changes: PivotDiff["changes"], pathName: string, before: unknown, after: unknown): void {
  const beforeHash = hashStable(before);
  const afterHash = hashStable(after);
  if (beforeHash === afterHash) {
    return;
  }
  changes.push({
    path: pathName,
    kind: before === undefined ? "added" : after === undefined ? "removed" : "changed",
    before,
    after
  });
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

function unsupportedThemeReport(workbookId: WorkbookId, operation: "get_theme" | "apply_theme", code: string, message: string) {
  return {
    ok: false,
    workbookId,
    operation,
    capabilityStatus: {
      capability: `excel.style.${operation}`,
      status: "unsupported",
      reasonCode: code
    },
    warnings: [{ code, message }],
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

function scopesFromBatch(workbookId: WorkbookId, operations: ExcelOperation[]): WorkbookScope[] {
  const scopes: WorkbookScope[] = [];
  for (const operation of operations) {
    scopes.push(...scopesFromOperation(workbookId, operation));
  }
  return dedupeScopes(scopes.length > 0 ? scopes : [{ type: "workbook", workbookId }]);
}

function snapshotRangesFromOperations(workbookId: WorkbookId, operations: ExcelOperation[]): A1Range[] {
  const ranges: A1Range[] = [];
  for (const scope of scopesFromBatch(workbookId, operations)) {
    if (scope.type === "range" && scope.address !== undefined) {
      ranges.push({
        workbookId: scope.workbookId,
        sheetName: scope.sheetName,
        address: scope.address
      });
    }
  }
  const seen = new Set<string>();
  return ranges.filter((range) => {
    const key = `${range.workbookId}:${range.sheetName}:${range.address}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function detectSparseOverwriteWarnings(operations: ExcelOperation[]): OperationWarning[] {
  const warnings: OperationWarning[] = [];
  for (const operation of operations) {
    if (operation.kind !== "range.write_values") {
      continue;
    }
    const values = (operation as { values?: unknown[][] }).values;
    if (!Array.isArray(values) || values.length === 0) {
      continue;
    }
    const matrixCells = values.reduce((sum, row) => sum + (Array.isArray(row) ? row.length : 0), 0);
    const nonEmptyCells = values.reduce(
      (sum, row) => sum + (Array.isArray(row) ? row.filter((value) => value !== null && value !== undefined && value !== "").length : 0),
      0
    );
    const targetCells = cellCount(operation.target.address);
    const touchedCells = Math.max(matrixCells, targetCells);
    if (touchedCells < 8 || nonEmptyCells === 0) {
      continue;
    }
    const nonEmptyRatio = nonEmptyCells / touchedCells;
    if (nonEmptyRatio <= 0.25 && touchedCells - nonEmptyCells >= 4) {
      warnings.push({
        code: "SPARSE_RANGE_WRITE_RISK",
        message: `Sparse range write to ${operation.target.sheetName}!${operation.target.address} has ${nonEmptyCells}/${touchedCells} non-empty cell(s). Use a smaller range or an explicit clear operation.`,
        target: operation.target,
        details: {
          operationId: operation.operationId,
          nonEmptyCells,
          touchedCells,
          nonEmptyRatio
        }
      });
    }
  }
  return warnings;
}

function riskyEditSummary(
  applyResult: OperationResult,
  diffResult: unknown,
  rollbackPreview: { rollbackAvailable?: boolean; ok?: boolean; conflicts?: unknown[] } | undefined
) {
  const diff = (diffResult as { diff?: DiffSummary } | undefined)?.diff;
  return {
    transactionId: applyResult.transactionId,
    backupIds: applyResult.backups,
    warningCount: applyResult.warnings.length,
    diff: diff
      ? {
          changedRanges: diff.changedRanges,
          cellsChanged: diff.cellsChanged,
          formulasChanged: diff.formulasChanged,
          stylesChanged: diff.stylesChanged,
          tablesChanged: diff.tablesChanged,
          sheetsChanged: diff.sheetsChanged,
          destructiveLevel: diff.destructiveLevel
        }
      : undefined,
    rollback: {
      previewed: rollbackPreview !== undefined,
      available: rollbackPreview?.rollbackAvailable ?? false,
      conflictCount: Array.isArray(rollbackPreview?.conflicts) ? rollbackPreview.conflicts.length : 0
    }
  };
}

function tableMutationScopes(request: { workbookId: WorkbookId; tableName?: string; sheetName?: string; address?: string }, ranges: A1Range[]): WorkbookScope[] {
  const scopes: WorkbookScope[] = ranges.map(rangeScope);
  if (request.tableName !== undefined) {
    scopes.push({ type: "table", workbookId: request.workbookId, sheetName: request.sheetName, tableName: request.tableName });
  } else if (request.sheetName !== undefined && request.address !== undefined) {
    scopes.push({ type: "table", workbookId: request.workbookId, sheetName: request.sheetName, tableName: `${request.sheetName}!${request.address}` });
  }
  return dedupeScopes(scopes);
}

function formulaMutationScopes(
  request: FormulaCopyPatternsRequest | FormulaFillRequest | FormulaPatternRequest,
  ranges: A1Range[]
): WorkbookScope[] {
  const scopes: WorkbookScope[] = ranges.flatMap((range) => [
    rangeScope(range),
    { type: "formula" as const, workbookId: range.workbookId, sheetName: range.sheetName, address: range.address }
  ]);
  if ("sourceSheetName" in request) {
    scopes.push({
      type: "formula",
      workbookId: request.workbookId,
      sheetName: request.sourceSheetName,
      address: request.sourceAddress
    });
  }
  return dedupeScopes(scopes);
}

function chartMutationScopes(request: { workbookId: WorkbookId; sheetName: string; sourceAddress: string; chartName?: string; chartType?: string }, ranges: A1Range[]): WorkbookScope[] {
  return dedupeScopes([
    ...ranges.map(rangeScope),
    {
      type: "chart",
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      chartName: request.chartName ?? request.chartType ?? `${request.sheetName}!${request.sourceAddress}`
    }
  ]);
}

function pivotMutationScopes(request: PivotCreateRequest, ranges: A1Range[]): WorkbookScope[] {
  const scopes: WorkbookScope[] = [
    ...ranges.map(rangeScope),
    {
      type: "pivot",
      workbookId: request.workbookId,
      sheetName: request.destinationSheetName,
      pivotName: request.pivotTableName
    }
  ];
  if (request.sourceTableName !== undefined) {
    scopes.push({ type: "table", workbookId: request.workbookId, tableName: request.sourceTableName });
  }
  return dedupeScopes(scopes);
}

function pivotTemplateCopyScopes(
  request: PivotCopyFromTemplateRequest,
  targetSheetName: string | undefined,
  sourceSheetName: string | undefined,
  ranges: A1Range[]
): WorkbookScope[] {
  return dedupeScopes([
    ...ranges.map(rangeScope),
    {
      type: "pivot",
      workbookId: request.workbookId,
      sheetName: targetSheetName,
      pivotName: request.pivotTableName
    },
    {
      type: "pivot",
      workbookId: request.workbookId,
      sheetName: sourceSheetName,
      pivotName: request.templatePivotTableName
    }
  ]);
}

function pivotDeleteScopes(request: PivotSelector, sheetName: string | undefined, ranges: A1Range[]): WorkbookScope[] {
  return dedupeScopes([
    ...ranges.map(rangeScope),
    {
      type: "pivot",
      workbookId: request.workbookId,
      sheetName,
      pivotName: request.pivotTableName
    }
  ]);
}

function validatePivotTemplateCompatibility(source: PivotTableInfo, target: PivotTableInfo): ValidationIssue[] {
  const requiredFields = pivotTemplateRequiredSourceFields(source);
  const targetSourceFields = uniqueDefined((target.hierarchies ?? []).map((hierarchy) => hierarchy.name));
  if (requiredFields.length === 0 || targetSourceFields.length === 0) {
    return [];
  }
  return requiredFields
    .filter((field) => !targetSourceFields.includes(field))
    .map((field) => ({
      code: "PIVOT_TEMPLATE_SOURCE_FIELD_MISSING",
      severity: "error",
      category: "template",
      message: `Target PivotTable source is missing template field: ${field}`,
      details: {
        field,
        templatePivotTableName: source.pivotTableName,
        targetPivotTableName: target.pivotTableName,
        requiredFields,
        targetSourceFields
      }
    }));
}

function pivotTemplateRequiredSourceFields(info: PivotTableInfo): string[] {
  return uniqueDefined([
    ...(info.rowHierarchies ?? []).map((hierarchy) => hierarchy.name),
    ...(info.columnHierarchies ?? []).map((hierarchy) => hierarchy.name),
    ...(info.filterHierarchies ?? []).map((hierarchy) => hierarchy.name),
    ...(info.dataHierarchies ?? []).map((hierarchy) => hierarchy.field?.name ?? hierarchy.name)
  ]);
}

function uniqueDefined(values: Array<string | undefined>): string[] {
  return [...new Set(values.filter((value): value is string => typeof value === "string" && value.length > 0))];
}

function validateTableColumnOrder(info: TableInfo, columnOrder: Array<string | number>) {
  const issues: OperationWarning[] = [];
  if (columnOrder.length !== info.columns.length) {
    issues.push({
      code: "TABLE_COLUMN_ORDER_LENGTH_MISMATCH",
      message: `Column order must include exactly ${info.columns.length} column(s).`,
      details: { expectedCount: info.columns.length, actualCount: columnOrder.length }
    });
  }
  const seen = new Set<number>();
  for (const requested of columnOrder) {
    const column =
      typeof requested === "number"
        ? info.columns.find((candidate) => candidate.index === requested)
        : info.columns.find((candidate) => candidate.name === requested);
    if (!column) {
      issues.push({
        code: "TABLE_COLUMN_NOT_FOUND",
        message: `Column ${String(requested)} is not present in table ${info.tableName}.`,
        details: { requested, availableColumns: info.columns.map((candidate) => candidate.name) }
      });
      continue;
    }
    if (seen.has(column.index)) {
      issues.push({
        code: "TABLE_COLUMN_ORDER_DUPLICATE",
        message: `Column ${column.name} appears more than once in the requested order.`,
        details: { requested, column }
      });
    }
    seen.add(column.index);
  }
  for (const column of info.columns) {
    if (!seen.has(column.index)) {
      issues.push({
        code: "TABLE_COLUMN_ORDER_MISSING_COLUMN",
        message: `Column ${column.name} is missing from the requested order.`,
        details: { column }
      });
    }
  }
  if (issues.length > 0) {
    return {
      ok: false,
      warnings: issues,
      error: runtimeError("INVALID_ARGUMENT", "Table column reorder requires a complete, unique column order.", {
        retryable: false,
        details: { issues }
      })
    };
  }
  return { ok: true };
}

function addExpectedPivotAxisIssues(
  issues: Array<{ code: string; severity: "info" | "warning" | "error"; message: string; details?: Record<string, unknown> }>,
  axis: "row" | "column" | "filter" | "data",
  expectedFields: string[] | undefined,
  actualFields: string[]
): void {
  if (!expectedFields?.length) {
    return;
  }
  for (const field of uniqueDefined(expectedFields)) {
    if (!actualFields.includes(field)) {
      issues.push({
        code: "PIVOT_EXPECTED_LAYOUT_MISMATCH",
        severity: "error",
        message: `Expected PivotTable ${axis} field is not present: ${field}`,
        details: { axis, field, actualFields }
      });
    }
  }
}

function addExpectedPivotDataSettingIssues(
  issues: Array<{ code: string; severity: "info" | "warning" | "error"; message: string; details?: Record<string, unknown> }>,
  expectedSettings: PivotValidateSourceRequest["expectedDataFieldSettings"],
  actualDataHierarchies: Array<{ name: string; summarizeBy?: string; numberFormat?: string; field?: { name: string } }>
): void {
  if (!expectedSettings?.length) {
    return;
  }
  for (const expected of expectedSettings) {
    const actual = actualDataHierarchies.find((hierarchy) => {
      const sourceMatches = expected.sourceFieldName === undefined || hierarchy.field?.name === expected.sourceFieldName;
      const nameMatches = expected.name === undefined || hierarchy.name === expected.name;
      return sourceMatches && nameMatches;
    });
    if (!actual) {
      issues.push({
        code: "PIVOT_EXPECTED_DATA_FIELD_MISSING",
        severity: "error",
        message: `Expected PivotTable data field is missing: ${expected.sourceFieldName ?? expected.name ?? "unknown"}`,
        details: { expected, actualDataFields: actualDataHierarchies.map((hierarchy) => ({ name: hierarchy.name, sourceFieldName: hierarchy.field?.name })) }
      });
      continue;
    }
    if (expected.summarizeBy !== undefined && actual.summarizeBy !== expected.summarizeBy) {
      issues.push({
        code: "PIVOT_EXPECTED_AGGREGATION_MISMATCH",
        severity: "error",
        message: `Expected PivotTable aggregation ${expected.summarizeBy} for ${actual.name}, found ${actual.summarizeBy ?? "unavailable"}.`,
        details: { expected, actual: { name: actual.name, sourceFieldName: actual.field?.name, summarizeBy: actual.summarizeBy } }
      });
    }
    if (expected.numberFormat !== undefined && actual.numberFormat !== expected.numberFormat) {
      issues.push({
        code: "PIVOT_EXPECTED_NUMBER_FORMAT_MISMATCH",
        severity: "error",
        message: `Expected PivotTable number format ${expected.numberFormat} for ${actual.name}, found ${actual.numberFormat ?? "unavailable"}.`,
        details: { expected, actual: { name: actual.name, sourceFieldName: actual.field?.name, numberFormat: actual.numberFormat } }
      });
    }
  }
}

function addExpectedPivotLayoutIssues(
  issues: Array<{ code: string; severity: "info" | "warning" | "error"; message: string; details?: Record<string, unknown> }>,
  expectedLayout: PivotValidateSourceRequest["expectedLayout"],
  actualLayout: PivotLayoutInfo | undefined
): void {
  if (!expectedLayout || Object.keys(expectedLayout).length === 0) {
    return;
  }
  if (!actualLayout) {
    issues.push({
      code: "PIVOT_LAYOUT_UNAVAILABLE",
      severity: "warning",
      message: "PivotTable layout metadata is unavailable from Office.js, so expected layout cannot be fully verified.",
      details: { expectedLayout }
    });
    return;
  }
  for (const [key, expectedValue] of Object.entries(expectedLayout)) {
    const actualValue = (actualLayout as Record<string, unknown>)[key];
    if (actualValue !== expectedValue) {
      issues.push({
        code: "PIVOT_EXPECTED_LAYOUT_SETTING_MISMATCH",
        severity: "error",
        message: `Expected PivotTable layout ${key} to be ${String(expectedValue)}, found ${String(actualValue)}.`,
        details: { key, expected: expectedValue, actual: actualValue }
      });
    }
  }
}

function nameMutationScopes(request: { workbookId: WorkbookId; name: string; sheetName?: string; reference?: string }): WorkbookScope[] {
  const scopes: WorkbookScope[] = [
    { type: "named_range", workbookId: request.workbookId, name: request.name, sheetName: request.sheetName }
  ];
  if (request.sheetName !== undefined && request.reference !== undefined) {
    scopes.push({ type: "range", workbookId: request.workbookId, sheetName: request.sheetName, address: request.reference });
  }
  return dedupeScopes(scopes);
}

function scopesFromOperation(workbookId: WorkbookId, operation: ExcelOperation): WorkbookScope[] {
  switch (operation.kind) {
    case "range.read_full":
    case "range.write_values":
    case "range.write_number_formats":
    case "range.write_styles":
    case "range.clear_style_dimensions":
    case "range.write_hyperlinks":
    case "range.write_comments":
    case "range.clear":
    case "range.clear_values":
    case "range.clear_formats":
    case "range.clear_values_keep_format":
    case "range.insert_rows":
    case "range.delete_rows":
    case "range.insert_columns":
    case "range.delete_columns":
    case "range.autofit_columns":
    case "range.autofit_rows":
    case "range.apply_autofilter":
    case "range.merge":
    case "range.unmerge":
    case "range.restore_snapshot":
      return [rangeScope(operation.target)];
    case "range.write_values_many":
      return operation.entries.map((entry) => rangeScope(entry.target));
    case "range.write_number_formats_many":
      return operation.entries.map((entry) => rangeScope(entry.target));
    case "range.write_styles_many":
      return operation.entries.map((entry) => rangeScope(entry.target));
    case "range.clear_style_dimensions_many":
      return operation.entries.map((entry) => rangeScope(entry.target));
    case "range.clear_many":
      return operation.entries.map((entry) => rangeScope(entry.target));
    case "range.clear_formats_many":
      return operation.targets.map(rangeScope);
    case "range.autofit_many":
      return operation.entries.map((entry) => rangeScope(entry.target));
    case "range.write_formulas":
      return [
        rangeScope(operation.target),
        { type: "formula", workbookId: operation.target.workbookId, sheetName: operation.target.sheetName, address: operation.target.address },
        ...formulaDependencyScopes(operation.target.workbookId, operation.target.sheetName, operation.formulas)
      ];
    case "range.copy":
    case "range.move":
      return [rangeScope(operation.source), rangeScope(operation.target)];
    case "sheet.create":
    case "template.create_sheet_from_template":
    case "workbook.calculate":
    case "workbook.save":
      return [{ type: "workbook", workbookId }];
    case "sheet.copy":
      return [
        { type: "sheet", workbookId, sheetName: operation.sourceSheetName },
        { type: "workbook", workbookId }
      ];
    case "sheet.copy_clean_data_regions":
      return [
        { type: "sheet", workbookId, sheetName: operation.sourceSheetName },
        { type: "workbook", workbookId }
      ];
    case "sheet.rename":
    case "sheet.delete":
    case "sheet.move":
    case "sheet.hide":
    case "sheet.unhide":
    case "sheet.protect":
    case "sheet.unprotect":
    case "sheet.clear":
    case "sheet.set_tab_color":
      return [{ type: "sheet", workbookId, sheetName: operation.sheetName }];
  }
}

function rangeScope(range: A1Range): WorkbookScope {
  return { type: "range", workbookId: range.workbookId, sheetName: range.sheetName, address: range.address };
}

function formulaDependencyScopes(workbookId: WorkbookId, sheetName: string, formulas: CellMatrix<string | null>): WorkbookScope[] {
  const scopes: WorkbookScope[] = [];
  for (const row of formulas) {
    for (const formula of row) {
      if (!formula) {
        continue;
      }
      for (const reference of extractFormulaReferences(workbookId, sheetName, formula)) {
        if (reference.kind === "range" && reference.sheetName !== undefined && reference.address !== undefined) {
          scopes.push({ type: "range", workbookId, sheetName: reference.sheetName, address: reference.address });
        }
        if (reference.kind === "table" && reference.tableName !== undefined) {
          scopes.push({ type: "table", workbookId, tableName: reference.tableName });
        }
      }
    }
  }
  return dedupeScopes(scopes);
}

function dedupeScopes(scopes: WorkbookScope[]): WorkbookScope[] {
  const seen = new Set<string>();
  const deduped: WorkbookScope[] = [];
  for (const scope of scopes) {
    const key = JSON.stringify(scope);
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(scope);
    }
  }
  return deduped;
}

function lockModeForDestructiveLevel(level: DestructiveLevel): LockMode {
  switch (level) {
    case "workbook":
      return "workbook";
    case "structure":
      return "structure";
    case "format":
      return "format_layout";
    case "values":
      return "write_values";
    case "none":
      return "read";
  }
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

function headerRowTarget(target: A1Range, headerRowIndex: number): A1Range {
  const parsed = parseA1Address(stripSheetName(target.address));
  const row = parsed.startRow + headerRowIndex;
  return {
    ...target,
    address: formatA1Address({ startColumn: parsed.startColumn, endColumn: parsed.endColumn, startRow: row, endRow: row })
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

function changedValueRunOperations(target: A1Range, before: CellMatrix, after: CellMatrix, reason: string): Array<Extract<ExcelOperation, { kind: "range.write_values" }>> {
  const parsed = parseA1Address(stripSheetName(target.address));
  const operations: Array<Extract<ExcelOperation, { kind: "range.write_values" }>> = [];
  const rowCount = Math.max(before.length, after.length);
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const columnCount = Math.max(before[rowIndex]?.length ?? 0, after[rowIndex]?.length ?? 0);
    let runStart: number | undefined;
    const flushRun = (exclusiveColumnIndex: number) => {
      if (runStart === undefined) {
        return;
      }
      const startColumn = parsed.startColumn + runStart;
      const endColumn = parsed.startColumn + exclusiveColumnIndex - 1;
      const row = parsed.startRow + rowIndex;
      operations.push({
        kind: "range.write_values",
        operationId: makeId<OperationId>("op"),
        workbookId: target.workbookId,
        destructiveLevel: "values",
        reason,
        target: {
          workbookId: target.workbookId,
          sheetName: target.sheetName,
          address: formatA1Address({ startRow: row, endRow: row, startColumn, endColumn })
        },
        values: [(after[rowIndex] ?? []).slice(runStart, exclusiveColumnIndex)],
        preserveFormats: true
      });
      runStart = undefined;
    };
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      if (before[rowIndex]?.[columnIndex] !== after[rowIndex]?.[columnIndex]) {
        runStart ??= columnIndex;
        continue;
      }
      flushRun(columnIndex);
    }
    flushRun(columnCount);
  }
  return operations;
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
  return rangesOverlapAddresses(range.address, region.address);
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

function clonePermissionStateForWorkbook(state: PermissionState, workbookId: WorkbookId): PermissionState {
  return {
    ...state,
    requireConfirmationFor: [...state.requireConfirmationFor],
    scope: { ...state.scope, workbookId },
    lockedRegions: state.lockedRegions.filter((region) => region.workbookId === workbookId).map((region) => ({ ...region }))
  };
}

function normalizeImportedTemplate(raw: Record<string, unknown>, workbookId: WorkbookId): TemplateRecord | undefined {
  if (!isRecord(raw)) {
    return undefined;
  }
  const templateId = raw.templateId;
  const name = raw.name;
  const scope = raw.scope;
  const sourceSheetName = raw.sourceSheetName;
  const fingerprint = raw.fingerprint;
  const fingerprintPayload = raw.fingerprintPayload;
  if (
    typeof templateId !== "string" ||
    typeof name !== "string" ||
    (scope !== "workbook" && scope !== "local") ||
    typeof sourceSheetName !== "string" ||
    !isRecord(fingerprint) ||
    !isRecord(fingerprintPayload)
  ) {
    return undefined;
  }
  const record: TemplateRecord = {
    templateId: templateId as TemplateId,
    name,
    scope,
    version: typeof raw.version === "number" ? raw.version : 1,
    sourceSheetName,
    fingerprint: fingerprint as unknown as TemplateRecord["fingerprint"],
    fingerprintPayload: fingerprintPayload as unknown as TemplateRecord["fingerprintPayload"],
    dataRegions: Array.isArray(raw.dataRegions) ? raw.dataRegions.filter((item): item is string => typeof item === "string") : [],
    createdAt: typeof raw.createdAt === "string" ? raw.createdAt : new Date().toISOString(),
    updatedAt: typeof raw.updatedAt === "string" ? raw.updatedAt : new Date().toISOString()
  };
  if (scope === "workbook") {
    record.workbookId = workbookId;
  } else if (typeof raw.workbookId === "string") {
    record.workbookId = raw.workbookId as WorkbookId;
  }
  return record;
}

function normalizeImportedRegion(region: WorkbookRegion, workbookId: WorkbookId): WorkbookRegion | undefined {
  if (!isRecord(region) || region.workbookId !== workbookId) {
    return undefined;
  }
  if (
    typeof region.regionId !== "string" ||
    typeof region.name !== "string" ||
    typeof region.sheetName !== "string" ||
    typeof region.address !== "string" ||
    typeof region.kind !== "string"
  ) {
    return undefined;
  }
  return {
    ...region,
    workbookId,
    createdAt: typeof region.createdAt === "string" ? region.createdAt : new Date().toISOString(),
    updatedAt: typeof region.updatedAt === "string" ? region.updatedAt : new Date().toISOString()
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function taskScheduleAction(
  taskStatus: TaskRecord["status"],
  state: "ready" | "waiting_dependencies" | "waiting_locks" | "blocked" | "done"
): "start" | "resume" | "wait" | "resolve_blockers" | "none" {
  if (state === "done") {
    return "none";
  }
  if (state === "waiting_dependencies" || state === "waiting_locks") {
    return "wait";
  }
  if (state === "blocked") {
    return "resolve_blockers";
  }
  return taskStatus === "blocked" ? "resume" : "start";
}

function taskScheduleMessage(goal: string, state: string, dependencyCount: number, lockConflictCount: number, blockerCount: number): string {
  switch (state) {
    case "ready":
      return `Ready: ${goal}`;
    case "waiting_dependencies":
      return `Waiting for ${dependencyCount} task dependenc${dependencyCount === 1 ? "y" : "ies"}: ${goal}`;
    case "waiting_locks":
      return `Waiting for ${lockConflictCount} lock conflict${lockConflictCount === 1 ? "" : "s"}: ${goal}`;
    case "blocked":
      return `Blocked by ${blockerCount} open blocker${blockerCount === 1 ? "" : "s"}: ${goal}`;
    case "done":
      return `Done: ${goal}`;
    default:
      return goal;
  }
}

function normalizeLockLeasePolicy(input: Partial<LockLeasePolicy>): LockLeasePolicy {
  const maxTtlMs = positiveInteger(input.maxTtlMs, 600_000);
  return {
    maxTtlMs,
    defaultTtlMs: Math.min(positiveInteger(input.defaultTtlMs, 120_000), maxTtlMs),
    transactionTtlMs: Math.min(positiveInteger(input.transactionTtlMs, 120_000), maxTtlMs),
    allowManualLocks: input.allowManualLocks ?? true
  };
}

function lockTtl(ttlMs: number, policy: LockLeasePolicy): number {
  return Math.min(positiveInteger(ttlMs, policy.defaultTtlMs), policy.maxTtlMs);
}

function disconnectedRuntimeCapabilities(): RuntimeCapabilities {
  return {
    engine: {
      name: "open-workbook-daemon",
      version: runtimeVersion(),
      platform: "unknown"
    },
    apiSets: [],
    capabilities: [
      {
        name: "mcp.catalog",
        supported: true,
        platforms: ["mac", "windows", "web"],
        notes: "Tool, resource, and prompt catalogs are available without a connected Excel host."
      },
      {
        name: "collaboration.runtime",
        supported: true,
        platforms: ["mac", "windows", "web"],
        notes: "Tasks, locks, transactions, conflict telemetry, and local state are daemon-side capabilities."
      },
      {
        name: "excel.office-js",
        supported: false,
        platforms: ["mac", "windows", "web"],
        notes: "Connect the Excel add-in to report real Office API set support."
      }
    ],
    hostCapabilities: [
      {
        name: "range-values-formulas-styles",
        supported: false,
        status: "unknown",
        reason: "No Excel add-in session is connected."
      },
      {
        name: "tables-filters-sorts",
        supported: false,
        status: "unknown",
        reason: "No Excel add-in session is connected."
      },
      {
        name: "pivots",
        supported: false,
        status: "unknown",
        reason: "No Excel add-in session is connected."
      },
      {
        name: "charts",
        supported: false,
        status: "unknown",
        reason: "No Excel add-in session is connected."
      }
    ]
  };
}

function positiveInteger(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? Math.round(value) : fallback;
}

function nextRetryAt(conflicts: ConflictRecord[]): string | undefined {
  const timestamps = conflicts
    .map((conflict) => conflict.lockExpiresAt)
    .filter((expiresAt): expiresAt is string => expiresAt !== undefined)
    .map((expiresAt) => Date.parse(expiresAt))
    .filter((timestamp) => Number.isFinite(timestamp));
  if (timestamps.length === 0) {
    return undefined;
  }
  return new Date(Math.min(...timestamps)).toISOString();
}

function telemetryBuckets(records: ConflictTelemetryRecord[], keySelector: (record: ConflictTelemetryRecord) => string[]): ConflictTelemetrySummary["byCode"] {
  const buckets = new Map<string, { count: number; openCount: number; lastSeenAt: string; codes: Set<string> }>();
  for (const record of records) {
    for (const key of keySelector(record)) {
      const bucket = buckets.get(key) ?? { count: 0, openCount: 0, lastSeenAt: record.createdAt, codes: new Set<string>() };
      bucket.count += 1;
      if (record.status === "open") {
        bucket.openCount += 1;
      }
      if (record.createdAt > bucket.lastSeenAt) {
        bucket.lastSeenAt = record.createdAt;
      }
      bucket.codes.add(record.code);
      buckets.set(key, bucket);
    }
  }
  return [...buckets.entries()]
    .map(([key, bucket]) => ({
      key,
      count: bucket.count,
      openCount: bucket.openCount,
      lastSeenAt: bucket.lastSeenAt,
      codes: [...bucket.codes].sort()
    }))
    .sort((a, b) => b.count - a.count || b.lastSeenAt.localeCompare(a.lastSeenAt));
}

function isTerminalTransactionStatus(status: TransactionRecord["status"]): boolean {
  return ["applied", "failed", "blocked", "rolled_back", "cancelled"].includes(status);
}

function isTerminalJobStatus(status: JobStatus): boolean {
  return ["applied", "failed", "partially_applied", "cancelled"].includes(status);
}

function jobProgressMessage(job: JobRecord, knownTransactionCount: number): string {
  switch (job.status) {
    case "queued":
      return `Workbook job is queued with ${knownTransactionCount || job.chunksTotal} chunk(s).`;
    case "applying":
      return `Workbook job is applying ${job.chunksCompleted} of ${job.chunksTotal} chunk(s).`;
    case "partially_applied":
      return job.errorMessage
        ? `Workbook job partially applied ${job.chunksCompleted} of ${job.chunksTotal} chunk(s): ${job.errorMessage}`
        : `Workbook job partially applied ${job.chunksCompleted} of ${job.chunksTotal} chunk(s).`;
    case "applied":
      return `Workbook job applied all ${job.chunksTotal} chunk(s).`;
    case "failed":
      return job.errorMessage ? `Workbook job failed: ${job.errorMessage}` : "Workbook job failed.";
    case "cancelled":
      return "Workbook job was cancelled before all chunks applied.";
  }
}

function planBatchChunks(operations: ExcelOperation[]): BatchChunkPlan {
  const chunkedOperations = chunkBatchOperations(operations);
  if (chunkedOperations.length <= 1 || chunkedOperations.length === operations.length) {
    return {
      strategy: "none",
      chunksTotal: 1,
      chunkSize: operations.length,
      operationCount: operations.length,
      chunkedOperationKinds: [],
      safeToAutoChunk: false
    };
  }
  const allChunkable = operations.every((operation) => isStyleChunkableOperation(operation) || isMatrixChunkableOperation(operation));
  if (!allChunkable) {
    return {
      strategy: "none",
      chunksTotal: 1,
      chunkSize: operations.length,
      operationCount: operations.length,
      chunkedOperationKinds: [],
      safeToAutoChunk: false
    };
  }
  const hasStyle = operations.some(isStyleChunkableOperation);
  const hasMatrix = operations.some(isMatrixChunkableOperation);
  const chunkSize = hasMatrix ? matrixChunkRowCount() : styleBatchChunkSize();
  return {
    strategy: hasStyle && hasMatrix ? "mixed" : hasMatrix ? "split_matrix_rows" : "split_style_entries",
    chunksTotal: chunkedOperations.length,
    chunkSize,
    operationCount: operations.length,
    chunkedOperationKinds: [...new Set(operations.map((operation) => operation.kind))],
    safeToAutoChunk: true
  };
}

function chunkBatchOperations(operations: ExcelOperation[]): ExcelOperation[][] {
  if (operations.length === 0) {
    return [];
  }
  if (operations.every(isStyleChunkableOperation)) {
    return chunkArray(operations, styleBatchChunkSize());
  }
  if (!operations.every((operation) => isStyleChunkableOperation(operation) || isMatrixChunkableOperation(operation))) {
    return [operations];
  }
  const chunks: ExcelOperation[][] = [];
  for (const operation of operations) {
    if (isMatrixChunkableOperation(operation)) {
      chunks.push(...chunkMatrixOperation(operation));
    } else {
      chunks.push([operation]);
    }
  }
  return chunks.length > 1 ? chunks : [operations];
}

function isStyleChunkableOperation(operation: ExcelOperation): boolean {
  return operation.kind === "range.write_styles";
}

function isMatrixChunkableOperation(operation: ExcelOperation): operation is Extract<ExcelOperation, { kind: "range.write_values" | "range.write_formulas" | "range.write_number_formats" }> {
  return operation.kind === "range.write_values" || operation.kind === "range.write_formulas" || operation.kind === "range.write_number_formats";
}

function chunkMatrixOperation(operation: Extract<ExcelOperation, { kind: "range.write_values" | "range.write_formulas" | "range.write_number_formats" }>): ExcelOperation[][] {
  const matrix = matrixForOperation(operation);
  const rowCount = matrix.length;
  const chunkRows = matrixChunkRowCount();
  if (rowCount <= chunkRows) {
    return [[operation]];
  }
  const chunks: ExcelOperation[][] = [];
  const parsed = parseA1Address(stripSheetName(operation.target.address));
  for (let start = 0; start < rowCount; start += chunkRows) {
    const rows = matrix.slice(start, start + chunkRows);
    const address = formatA1Address({
      ...parsed,
      startRow: parsed.startRow + start,
      endRow: parsed.startRow + start + rows.length - 1
    });
    chunks.push([cloneMatrixOperation(operation, rows, address)]);
  }
  return chunks;
}

function matrixForOperation(operation: Extract<ExcelOperation, { kind: "range.write_values" | "range.write_formulas" | "range.write_number_formats" }>): unknown[][] {
  switch (operation.kind) {
    case "range.write_values":
      return operation.values;
    case "range.write_formulas":
      return operation.formulas;
    case "range.write_number_formats":
      return operation.numberFormat;
  }
}

function cloneMatrixOperation(
  operation: Extract<ExcelOperation, { kind: "range.write_values" | "range.write_formulas" | "range.write_number_formats" }>,
  rows: unknown[][],
  address: string
): ExcelOperation {
  const target = { ...operation.target, address };
  switch (operation.kind) {
    case "range.write_values":
      return { ...operation, operationId: makeId<OperationId>("op"), target, values: rows as CellMatrix };
    case "range.write_formulas":
      return { ...operation, operationId: makeId<OperationId>("op"), target, formulas: rows as CellMatrix<string | null> };
    case "range.write_number_formats":
      return { ...operation, operationId: makeId<OperationId>("op"), target, numberFormat: rows as string[][] };
  }
}

function batchDirectOperationThreshold(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_BATCH_DIRECT_OPERATION_THRESHOLD", 25);
}

function batchDirectPayloadThresholdBytes(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_BATCH_DIRECT_PAYLOAD_BYTES", 512_000);
}

function batchDirectCellThreshold(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_BATCH_DIRECT_CELL_THRESHOLD", 50_000);
}

function styleBatchChunkSize(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_STYLE_BATCH_CHUNK_SIZE", 25);
}

function matrixChunkRowCount(): number {
  return positiveIntegerEnv("OPEN_WORKBOOK_MATRIX_CHUNK_ROWS", 500);
}

function positiveIntegerEnv(name: string, fallback: number): number {
  const value = Number(process.env[name] ?? fallback);
  return Number.isFinite(value) && value > 0 ? Math.round(value) : fallback;
}

function chunkArray<T>(items: T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function cancelledOperationResult(transactionId: TransactionId): OperationResult {
  return {
    ok: false,
    transactionId,
    transactionStatus: "cancelled",
    progressMessage: "Queued workbook mutation was cancelled before it reached Excel.",
    rollbackAvailable: false,
    backups: [],
    warnings: [],
    telemetry: { warningCount: 0 },
    error: runtimeError("TRANSACTION_CANCELLED", "Queued workbook mutation was cancelled before it reached Excel.", { retryable: false })
  };
}

function normalizeAgentExecutionContext(context: AgentRunExecutionContext | undefined): AgentRunExecutionContext | undefined {
  if (!context || typeof context.agentId !== "string" || context.agentId.length === 0) {
    return undefined;
  }
  return {
    agentId: context.agentId,
    ...(typeof context.agentName === "string" && context.agentName.length > 0 ? { agentName: context.agentName } : {}),
    clientType: context.clientType ?? "mcp"
  };
}

function queuedOperationResult(transaction: TransactionRecord): OperationResult {
  return {
    ok: true,
    transactionId: transaction.transactionId,
    transactionStatus: "queued",
    queuePosition: transaction.queuePosition,
    progressMessage: transaction.progressMessage ?? "Workbook mutation is queued and will apply when earlier workbook work finishes.",
    rollbackAvailable: false,
    backups: [],
    warnings: [],
    telemetry: { warningCount: 0 }
  };
}

function shouldRetryStyleBatch(request: BatchRequest): boolean {
  return request.retryStrategy !== "retry_timeout_split_style_entries"
    && request.operations.length > 1
    && request.operations.every((operation) => operation.kind === "range.write_styles");
}

function chunkOperationsForRetry(operations: ExcelOperation[]): ExcelOperation[][] {
  const chunkSize = Math.max(1, Math.ceil(operations.length / 2));
  const chunks: ExcelOperation[][] = [];
  for (let index = 0; index < operations.length; index += chunkSize) {
    chunks.push(operations.slice(index, index + chunkSize));
  }
  return chunks;
}

function batchInvalidatesWorkbookMetadata(request: BatchRequest): boolean {
  return operationsInvalidateWorkbookMetadata(request.operations);
}

function eventWorkbookId(params: unknown): WorkbookId | string | undefined {
  if (!params || typeof params !== "object") {
    return undefined;
  }
  const record = params as Record<string, unknown>;
  const direct = record.workbookId ?? record.workbookID;
  if (typeof direct === "string") {
    return direct;
  }
  const workbook = record.workbook;
  if (workbook && typeof workbook === "object") {
    const nested = (workbook as Record<string, unknown>).workbookId ?? (workbook as Record<string, unknown>).id;
    if (typeof nested === "string") {
      return nested;
    }
  }
  return undefined;
}

function operationsInvalidateWorkbookMetadata(operations: ExcelOperation[]): boolean {
  return operations.some((operation) => operationInvalidatesWorkbookMetadata(operation));
}

function operationInvalidatesWorkbookMetadata(operation: ExcelOperation): boolean {
  if (operation.destructiveLevel === "structure" || operation.destructiveLevel === "workbook") {
    return true;
  }
  if (operation.destructiveLevel === "values") {
    return false;
  }
  const kind = operation.kind as string;
  return /^(table|sheet|name|formula|pivot|chart|template)\./.test(kind)
    || operation.kind === "range.clear"
    || operation.kind === "range.merge";
}

function transactionInvalidatesWorkbookMetadata(input: { destructiveLevel: DestructiveLevel; scopes: WorkbookScope[] }): boolean {
  if (input.destructiveLevel === "structure" || input.destructiveLevel === "workbook") {
    return true;
  }
  return input.scopes.some((scope) =>
    scope.type === "sheet"
    || scope.type === "formula"
    || scope.type === "table"
    || scope.type === "named_range"
    || scope.type === "chart"
    || scope.type === "pivot"
    || scope.type === "template"
  );
}

function scopeTelemetryKey(scope: WorkbookScope): string {
  switch (scope.type) {
    case "workbook":
      return `${scope.workbookId}:workbook`;
    case "sheet":
      return `${scope.workbookId}:${scope.sheetName}`;
    case "range":
      return `${scope.workbookId}:${scope.sheetName}!${scope.address}`;
    case "formula":
      return `${scope.workbookId}:formula:${scope.sheetName}${scope.address ? `!${scope.address}` : ""}`;
    case "table":
      return `${scope.workbookId}:table:${scope.tableName}`;
    case "named_range":
      return `${scope.workbookId}:name:${scope.name}`;
    case "chart":
      return `${scope.workbookId}:chart:${scope.chartName}`;
    case "pivot":
      return `${scope.workbookId}:pivot:${scope.pivotName}`;
    case "template":
      return `${scope.workbookId}:template:${scope.templateId}`;
  }
}
