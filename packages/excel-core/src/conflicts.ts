import type {
  ConflictRecord,
  ConflictResolutionAction,
  ConflictResolutionGuidance,
  ConflictResolutionStep,
  TransactionRollbackConflict,
  TransactionId,
  WorkbookId,
  WorkbookScope
} from "@component-kit/open-workbook-protocol";
import { makeId } from "@component-kit/open-workbook-protocol";
import { scopesConflict } from "./lock-manager.js";

export type DependencyConflictCode =
  | "STRUCTURE_CONFLICT"
  | "TABLE_CONFLICT"
  | "FORMULA_DEPENDENCY_CONFLICT"
  | "DERIVED_OBJECT_CONFLICT"
  | "NAMED_RANGE_CONFLICT"
  | "LOCK_CONFLICT";

export interface ClassifiedScopeConflict {
  code: DependencyConflictCode;
  message: string;
  scopes: WorkbookScope[];
  suggestedAction: "wait_for_lock" | "refresh_plan" | "rebase_plan" | "split_scope" | "manual_review";
}

export function classifyScopeConflict(left: WorkbookScope, right: WorkbookScope): ClassifiedScopeConflict | undefined {
  if (!scopesConflict(left, right)) {
    return undefined;
  }
  const code = conflictCode(left, right);
  return {
    code,
    message: conflictMessage(code, left, right),
    scopes: [left, right],
    suggestedAction: code === "LOCK_CONFLICT" ? "wait_for_lock" : "manual_review"
  };
}

export function makeLockConflict(input: {
  workbookId: WorkbookId;
  left: WorkbookScope;
  right: WorkbookScope;
  ownerAgentId?: ConflictRecord["ownerAgentId"];
  taskId?: ConflictRecord["taskId"];
  transactionId?: ConflictRecord["transactionId"];
}): ConflictRecord | undefined {
  const classified = classifyScopeConflict(input.left, input.right);
  if (!classified) {
    return undefined;
  }
  const conflict: ConflictRecord = {
    conflictId: makeId<string>("conflict"),
    code: classified.code,
    message: classified.message,
    workbookId: input.workbookId,
    scopes: classified.scopes,
    retryable: true,
    suggestedAction: classified.suggestedAction,
    createdAt: new Date().toISOString()
  };
  if (input.ownerAgentId !== undefined) {
    conflict.ownerAgentId = input.ownerAgentId;
  }
  if (input.taskId !== undefined) {
    conflict.taskId = input.taskId;
  }
  if (input.transactionId !== undefined) {
    conflict.transactionId = input.transactionId;
  }
  return conflict;
}

export function makeRollbackConflict(input: {
  transactionId: TransactionId;
  conflictingTransactionId: TransactionId;
  left: WorkbookScope;
  right: WorkbookScope;
}): TransactionRollbackConflict | undefined {
  const classified = classifyScopeConflict(input.left, input.right);
  if (!classified) {
    return undefined;
  }
  return {
    code: rollbackCode(classified.code),
    message: `Rollback would conflict with later transaction ${input.conflictingTransactionId}: ${classified.message}`,
    transactionId: input.transactionId,
    conflictingTransactionId: input.conflictingTransactionId,
    scopes: classified.scopes,
    suggestedAction: "manual_review"
  };
}

export function makeConflictGuidance(conflict: ConflictRecord): ConflictResolutionGuidance {
  const primaryAction = primaryActionForConflict(conflict);
  const guidance: ConflictResolutionGuidance = {
    code: conflict.code,
    severity: severityForConflict(conflict),
    retryable: conflict.retryable,
    primaryAction,
    scopes: conflict.scopes,
    steps: guidanceSteps(conflict, primaryAction)
  };
  if (conflict.lockExpiresAt !== undefined) {
    guidance.nextRetryAt = conflict.lockExpiresAt;
  }
  if (conflict.ownerAgentId !== undefined) {
    guidance.ownerAgentId = conflict.ownerAgentId;
  }
  if (conflict.taskId !== undefined) {
    guidance.taskId = conflict.taskId;
  }
  if (conflict.transactionId !== undefined) {
    guidance.transactionId = conflict.transactionId;
  }
  if (conflict.lockId !== undefined) {
    guidance.lockId = conflict.lockId;
  }
  return guidance;
}

export function attachConflictGuidance(conflict: ConflictRecord): ConflictRecord {
  return {
    ...conflict,
    guidance: makeConflictGuidance(conflict)
  };
}

function conflictCode(left: WorkbookScope, right: WorkbookScope): DependencyConflictCode {
  if (left.type === "workbook" || right.type === "workbook" || left.type === "sheet" || right.type === "sheet") {
    return "STRUCTURE_CONFLICT";
  }
  if (left.type === "table" || right.type === "table") {
    return "TABLE_CONFLICT";
  }
  if (left.type === "formula" || right.type === "formula") {
    return "FORMULA_DEPENDENCY_CONFLICT";
  }
  if (left.type === "chart" || left.type === "pivot" || right.type === "chart" || right.type === "pivot") {
    return "DERIVED_OBJECT_CONFLICT";
  }
  if (left.type === "named_range" || right.type === "named_range") {
    return "NAMED_RANGE_CONFLICT";
  }
  if (left.type === "range" && right.type === "range") {
    return "LOCK_CONFLICT";
  }
  return "FORMULA_DEPENDENCY_CONFLICT";
}

function rollbackCode(code: DependencyConflictCode): TransactionRollbackConflict["code"] {
  if (code === "FORMULA_DEPENDENCY_CONFLICT" || code === "DERIVED_OBJECT_CONFLICT" || code === "NAMED_RANGE_CONFLICT" || code === "TABLE_CONFLICT") {
    return "ROLLBACK_DEPENDENCY_CONFLICT";
  }
  return "ROLLBACK_CONFLICT";
}

function conflictMessage(code: DependencyConflictCode, left: WorkbookScope, right: WorkbookScope): string {
  switch (code) {
    case "STRUCTURE_CONFLICT":
      return `${describeScope(left)} conflicts with workbook or sheet structure scope ${describeScope(right)}.`;
    case "TABLE_CONFLICT":
      return `${describeScope(left)} conflicts with table scope ${describeScope(right)}.`;
    case "DERIVED_OBJECT_CONFLICT":
      return `${describeScope(left)} conflicts with derived object scope ${describeScope(right)}.`;
    case "NAMED_RANGE_CONFLICT":
      return `${describeScope(left)} conflicts with named range scope ${describeScope(right)}.`;
    case "FORMULA_DEPENDENCY_CONFLICT":
      return `${describeScope(left)} may affect formula dependencies in ${describeScope(right)}.`;
    case "LOCK_CONFLICT":
      return `${describeScope(left)} overlaps ${describeScope(right)}.`;
  }
}

function primaryActionForConflict(conflict: ConflictRecord): ConflictResolutionAction {
  if (conflict.code === "LOCK_CONFLICT") {
    return conflict.lockExpiresAt ? "retry_after" : "wait_for_lock";
  }
  if (conflict.code === "TARGET_REGION_CHANGED") {
    return "refresh_plan";
  }
  if (conflict.code === "ROLLBACK_CONFLICT" || conflict.code === "ROLLBACK_DEPENDENCY_CONFLICT") {
    return "preview_rollback_chain";
  }
  if (conflict.code === "ROLLBACK_UNAVAILABLE") {
    return "repair_from_backup";
  }
  if (conflict.code === "STRUCTURE_CONFLICT") {
    return "manual_review";
  }
  if (conflict.code === "TABLE_CONFLICT" || conflict.code === "FORMULA_DEPENDENCY_CONFLICT" || conflict.code === "DERIVED_OBJECT_CONFLICT" || conflict.code === "NAMED_RANGE_CONFLICT") {
    return "handoff_task";
  }
  return conflict.suggestedAction;
}

function severityForConflict(conflict: ConflictRecord): ConflictResolutionGuidance["severity"] {
  if (!conflict.retryable || conflict.code === "STRUCTURE_CONFLICT" || conflict.code === "ROLLBACK_UNAVAILABLE") {
    return "blocked";
  }
  if (conflict.code === "LOCK_CONFLICT" || conflict.code === "TARGET_REGION_CHANGED") {
    return "warning";
  }
  return "blocked";
}

function guidanceSteps(conflict: ConflictRecord, primaryAction: ConflictResolutionAction): ConflictResolutionStep[] {
  const steps: ConflictResolutionStep[] = [];
  if (primaryAction === "retry_after" || primaryAction === "wait_for_lock") {
    steps.push({
      action: primaryAction,
      title: conflict.lockExpiresAt ? "Wait for lock expiry" : "Wait for active lock",
      description: conflict.lockExpiresAt
        ? `Retry after ${conflict.lockExpiresAt}, or coordinate with the lock owner if the task is urgent.`
        : "Wait for the active lock to be released, then evaluate the task schedule again.",
      toolName: "excel.task.evaluate_schedule",
      arguments: { workbookId: conflict.workbookId }
    });
    if (conflict.lockId !== undefined) {
      steps.push({
        action: "renew_or_release_lock",
        title: "Coordinate lock ownership",
        description: "Ask the lock owner to renew if still working or release the lock if the reserved work is done.",
        toolName: "excel.lock.release",
        arguments: { lockIds: [conflict.lockId] }
      });
    }
  }
  if (conflict.taskId !== undefined) {
    steps.push({
      action: "handoff_task",
      title: "Coordinate with task owner",
      description: "Review the owning task before retrying so two agents do not overwrite the same workbook area.",
      toolName: "excel.task.get",
      arguments: { taskId: conflict.taskId }
    });
  }
  if (canSuggestSplit(conflict)) {
    steps.push({
      action: "split_scope",
      title: "Split non-overlapping work",
      description: "Create a narrower task or plan that excludes the conflicting range/table/object scope.",
      toolName: "excel.task.create",
      arguments: { workbookId: conflict.workbookId, allowedScopes: conflict.scopes }
    });
  }
  if (primaryAction === "refresh_plan") {
    steps.push({
      action: "refresh_plan",
      title: "Refresh plan preview",
      description: "Refresh the plan preview and recreate the plan if target fingerprints changed.",
      toolName: "excel.plan.refresh_preview"
    });
  }
  if (primaryAction === "preview_rollback_chain") {
    steps.push({
      action: "preview_rollback_chain",
      title: "Preview rollback chain",
      description: "Check whether later dependent transactions must be rolled back together.",
      toolName: "excel.transaction.preview_rollback_chain",
      arguments: conflict.transactionId ? { transactionId: conflict.transactionId } : undefined
    });
  }
  if (primaryAction === "repair_from_backup") {
    steps.push({
      action: "repair_from_backup",
      title: "Repair from backup",
      description: "Use backup or template repair because this conflict cannot be resolved through transaction rollback metadata."
    });
  }
  steps.push({
    action: "manual_review",
    title: "Manual review",
    description: "If scope splitting or waiting is unsafe, ask the user to choose the winning change before applying more workbook mutations."
  });
  return dedupeSteps(steps);
}

function canSuggestSplit(conflict: ConflictRecord): boolean {
  return conflict.scopes.some((scope) => scope.type === "range" || scope.type === "table") && conflict.code !== "STRUCTURE_CONFLICT";
}

function dedupeSteps(steps: ConflictResolutionStep[]): ConflictResolutionStep[] {
  const seen = new Set<string>();
  const deduped: ConflictResolutionStep[] = [];
  for (const step of steps) {
    const key = `${step.action}:${step.toolName ?? ""}:${JSON.stringify(step.arguments ?? {})}`;
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(step);
    }
  }
  return deduped;
}

function describeScope(scope: WorkbookScope): string {
  switch (scope.type) {
    case "workbook":
      return "workbook";
    case "sheet":
      return `${scope.sheetName}`;
    case "range":
      return `${scope.sheetName}!${scope.address}`;
    case "formula":
      return `formulas ${scope.sheetName}${scope.address ? `!${scope.address}` : ""}`;
    case "table":
      return `table ${scope.tableName}`;
    case "named_range":
      return `named range ${scope.name}`;
    case "chart":
      return `chart ${scope.chartName}`;
    case "pivot":
      return `pivot ${scope.pivotName}`;
    case "template":
      return `template ${scope.templateId}`;
  }
}
