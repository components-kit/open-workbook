import type { AgentId, BackupId, JobId, LockId, PlanId, TaskId, TransactionId, WorkbookId } from "./ids.js";
import type { DestructiveLevel, DiffSummary, OperationTelemetry, OperationWarning } from "./operations.js";
import type { A1Range, RangeFingerprint } from "./workbook.js";

export type AgentStatus = "active" | "idle" | "disconnected";
export type TaskStatus = "open" | "claimed" | "planning" | "queued" | "applying" | "blocked" | "completed" | "failed" | "cancelled";
export type TransactionStatus = "queued" | "applying" | "applied" | "failed" | "rolled_back" | "blocked" | "cancelled";
export type JobStatus = "queued" | "applying" | "partially_applied" | "applied" | "failed" | "cancelled";
export type JobKind = "batch_chunked" | "style_chunked" | "matrix_chunked";
export type LockStatus = "active" | "released" | "expired";
export type TaskBlockerStatus = "open" | "resolved";
export type TaskBlockerSeverity = "info" | "warning" | "blocked";
export type LockMode =
  | "read"
  | "write_values"
  | "write_formulas"
  | "write_styles"
  | "format_layout"
  | "table"
  | "chart"
  | "pivot"
  | "structure"
  | "workbook";

export type WorkbookScope =
  | { type: "workbook"; workbookId: WorkbookId }
  | { type: "sheet"; workbookId: WorkbookId; sheetName: string }
  | { type: "range"; workbookId: WorkbookId; sheetName: string; address: string }
  | { type: "formula"; workbookId: WorkbookId; sheetName: string; address?: string | undefined }
  | { type: "table"; workbookId: WorkbookId; sheetName?: string | undefined; tableName: string }
  | { type: "named_range"; workbookId: WorkbookId; name: string; sheetName?: string | undefined }
  | { type: "chart"; workbookId: WorkbookId; sheetName?: string | undefined; chartName: string }
  | { type: "pivot"; workbookId: WorkbookId; sheetName?: string | undefined; pivotName: string }
  | { type: "template"; workbookId: WorkbookId; templateId: string };

export interface AgentRecord {
  agentId: AgentId;
  agentName?: string | undefined;
  clientType: "mcp" | "cli" | "daemon" | "unknown";
  pid?: number | undefined;
  status: AgentStatus;
  connectedAt: string;
  lastSeenAt: string;
}

export interface TaskRecord {
  taskId: TaskId;
  workbookId: WorkbookId;
  goal: string;
  role?: string | undefined;
  priority: "low" | "normal" | "high";
  status: TaskStatus;
  progress: number;
  currentStep?: string | undefined;
  blockers: TaskBlocker[];
  assignedAgentId?: AgentId | undefined;
  allowedScopes: WorkbookScope[];
  dependencies: TaskId[];
  planIds: PlanId[];
  transactionIds: TransactionId[];
  rollbackBackupIds: BackupId[];
  createdAt: string;
  updatedAt: string;
  completedAt?: string | undefined;
  failedAt?: string | undefined;
  errorMessage?: string | undefined;
}

export interface TaskBlocker {
  blockerId: string;
  severity: TaskBlockerSeverity;
  message: string;
  scope?: WorkbookScope | undefined;
  createdAt: string;
  resolvedAt?: string | undefined;
  status: TaskBlockerStatus;
}

export interface LockRecord {
  lockId: LockId;
  workbookId: WorkbookId;
  ownerAgentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  transactionId?: TransactionId | undefined;
  mode: LockMode;
  scope: WorkbookScope;
  status: LockStatus;
  acquiredAt: string;
  expiresAt: string;
  releasedAt?: string | undefined;
  reason: string;
}

export interface LockLeasePolicy {
  defaultTtlMs: number;
  transactionTtlMs: number;
  maxTtlMs: number;
  allowManualLocks: boolean;
}

export interface LockAcquireResponse {
  ok: boolean;
  locks: LockRecord[];
  conflicts: ConflictRecord[];
  policy: LockLeasePolicy;
}

export interface LockRenewResponse {
  ok: boolean;
  renewed: LockRecord[];
  missingLockIds: LockId[];
  policy: LockLeasePolicy;
}

export interface LockReleaseResponse {
  ok: boolean;
  released: LockRecord[];
  missingLockIds: LockId[];
}

export interface ConflictRecord {
  conflictId: string;
  code: string;
  message: string;
  workbookId: WorkbookId;
  scopes: WorkbookScope[];
  ownerAgentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  transactionId?: TransactionId | undefined;
  lockId?: LockId | undefined;
  lockExpiresAt?: string | undefined;
  retryable: boolean;
  suggestedAction: "wait_for_lock" | "refresh_plan" | "rebase_plan" | "split_scope" | "manual_review";
  guidance?: ConflictResolutionGuidance | undefined;
  createdAt: string;
}

export type ConflictResolutionAction =
  | "wait_for_lock"
  | "retry_after"
  | "renew_or_release_lock"
  | "split_scope"
  | "handoff_task"
  | "refresh_plan"
  | "rebase_plan"
  | "preview_rollback_chain"
  | "repair_from_backup"
  | "manual_review";

export interface ConflictResolutionStep {
  action: ConflictResolutionAction;
  title: string;
  description: string;
  toolName?: string | undefined;
  arguments?: Record<string, unknown> | undefined;
}

export interface ConflictResolutionGuidance {
  code: string;
  severity: "info" | "warning" | "blocked";
  retryable: boolean;
  primaryAction: ConflictResolutionAction;
  nextRetryAt?: string | undefined;
  ownerAgentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  transactionId?: TransactionId | undefined;
  lockId?: LockId | undefined;
  scopes: WorkbookScope[];
  steps: ConflictResolutionStep[];
}

export interface ConflictGuidanceResponse {
  ok: boolean;
  workbookId?: WorkbookId | undefined;
  guidance: ConflictResolutionGuidance[];
}

export type ConflictTelemetryStatus = "open" | "cleared";

export interface ConflictTelemetryRecord {
  telemetryId: string;
  conflictId: string;
  code: string;
  workbookId: WorkbookId;
  scopes: WorkbookScope[];
  scopeKeys: string[];
  primaryAction: ConflictResolutionAction;
  retryable: boolean;
  ownerAgentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  transactionId?: TransactionId | undefined;
  lockId?: LockId | undefined;
  status: ConflictTelemetryStatus;
  createdAt: string;
  clearedAt?: string | undefined;
  clearReason?: "lock_released" | "manual_clear" | "state_reset" | undefined;
}

export interface ConflictTelemetryBucket {
  key: string;
  count: number;
  openCount: number;
  lastSeenAt: string;
  codes: string[];
}

export interface ConflictTelemetrySummary {
  ok: boolean;
  workbookId?: WorkbookId | undefined;
  windowSize: number;
  totalCount: number;
  openCount: number;
  clearedCount: number;
  byCode: ConflictTelemetryBucket[];
  byPrimaryAction: ConflictTelemetryBucket[];
  hotScopes: ConflictTelemetryBucket[];
  hotTasks: ConflictTelemetryBucket[];
  hotAgents: ConflictTelemetryBucket[];
  recent: ConflictTelemetryRecord[];
}

export interface TransactionRecord {
  transactionId: TransactionId;
  workbookId: WorkbookId;
  agentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  planId?: PlanId | undefined;
  status: TransactionStatus;
  goal: string;
  scopes: WorkbookScope[];
  locks: LockId[];
  baseFingerprints: RangeFingerprint[];
  backups: BackupId[];
  warnings: OperationWarning[];
  diffSummary?: DiffSummary | undefined;
  telemetry?: OperationTelemetry | undefined;
  destructiveLevel: DestructiveLevel;
  queuedAt: string;
  queuePosition?: number | undefined;
  progressMessage?: string | undefined;
  queueWaitMs?: number | undefined;
  executionMs?: number | undefined;
  retryStrategy?: string | undefined;
  chunksTotal?: number | undefined;
  chunksCompleted?: number | undefined;
  startedAt?: string | undefined;
  finishedAt?: string | undefined;
  errorCode?: string | undefined;
  errorMessage?: string | undefined;
}

export interface JobRecord {
  jobId: JobId;
  workbookId: WorkbookId;
  agentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  planId?: PlanId | undefined;
  kind: JobKind;
  status: JobStatus;
  goal: string;
  transactionIds: TransactionId[];
  chunksTotal: number;
  chunksCompleted: number;
  progressMessage?: string | undefined;
  retryStrategy?: string | undefined;
  destructiveLevel: DestructiveLevel;
  warnings: OperationWarning[];
  queuedAt: string;
  startedAt?: string | undefined;
  finishedAt?: string | undefined;
  errorCode?: string | undefined;
  errorMessage?: string | undefined;
}

export interface CollaborationEvent {
  eventId: string;
  type:
    | "agent.registered"
    | "agent.heartbeat"
    | "task.created"
    | "task.updated"
    | "task.completed"
    | "task.failed"
    | "lock.acquired"
    | "lock.released"
    | "lock.policy_updated"
    | "transaction.queued"
    | "transaction.applying"
    | "transaction.applied"
    | "transaction.failed"
    | "transaction.cancelled"
    | "transaction.rollback_previewed"
    | "transaction.rolled_back"
    | "backup.created"
    | "backup.verified"
    | "backup.restored"
    | "backup.deleted"
    | "backup.pruned"
    | "backup.updated"
    | "permission.updated"
    | "conflict.detected";
  workbookId?: WorkbookId | undefined;
  agentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  transactionId?: TransactionId | undefined;
  lockId?: LockId | undefined;
  message: string;
  details?: Record<string, unknown> | undefined;
  createdAt: string;
}

export interface TransactionRollbackConflict {
  code: "ROLLBACK_CONFLICT" | "ROLLBACK_DEPENDENCY_CONFLICT" | "ROLLBACK_UNAVAILABLE";
  message: string;
  transactionId: TransactionId;
  conflictingTransactionId?: TransactionId | undefined;
  scopes: WorkbookScope[];
  suggestedAction: "manual_review" | "rollback_chain" | "repair_from_backup";
}

export interface TransactionRollbackPreview {
  ok: boolean;
  transactionId: TransactionId;
  workbookId?: WorkbookId | undefined;
  planId?: PlanId | undefined;
  taskId?: TaskId | undefined;
  rollbackAvailable: boolean;
  rollbackMethod?: "plan" | "backup" | undefined;
  scopes: WorkbookScope[];
  laterTransactions: TransactionRecord[];
  conflicts: TransactionRollbackConflict[];
  warnings: OperationWarning[];
}

export interface TransactionRollbackChainPreview {
  ok: boolean;
  rootTransactionId: TransactionId;
  workbookId?: WorkbookId | undefined;
  rollbackAvailable: boolean;
  rollbackOrder: TransactionRecord[];
  affectedTransactions: TransactionRecord[];
  conflicts: TransactionRollbackConflict[];
  warnings: OperationWarning[];
  requiresConfirmation: boolean;
  confirmationToken?: string | undefined;
}

export interface TaskScheduleDecision {
  taskId: TaskId;
  workbookId: WorkbookId;
  ready: boolean;
  state: "ready" | "waiting_dependencies" | "waiting_locks" | "blocked" | "done";
  waitingForTaskIds: TaskId[];
  lockConflicts: ConflictRecord[];
  blockers: TaskBlocker[];
  nextRetryAt?: string | undefined;
  suggestedAction: "start" | "resume" | "wait" | "resolve_blockers" | "none";
  message: string;
}

export interface TaskScheduleResponse {
  ok: boolean;
  workbookId?: WorkbookId | undefined;
  applied: boolean;
  decisions: TaskScheduleDecision[];
  updatedTasks: TaskRecord[];
}
