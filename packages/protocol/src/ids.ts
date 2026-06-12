export type WorkbookId = string & { readonly __brand: "WorkbookId" };
export type WorksheetId = string & { readonly __brand: "WorksheetId" };
export type SnapshotId = string & { readonly __brand: "SnapshotId" };
export type BackupId = string & { readonly __brand: "BackupId" };
export type TemplateId = string & { readonly __brand: "TemplateId" };
export type PlanId = string & { readonly __brand: "PlanId" };
export type OperationId = string & { readonly __brand: "OperationId" };
export type ConnectionId = string & { readonly __brand: "ConnectionId" };
export type AgentId = string & { readonly __brand: "AgentId" };
export type TaskId = string & { readonly __brand: "TaskId" };
export type LockId = string & { readonly __brand: "LockId" };
export type TransactionId = string & { readonly __brand: "TransactionId" };

export function makeId<T extends string>(prefix: string): T {
  const random =
    typeof crypto !== "undefined" && "randomUUID" in crypto
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  return `${prefix}_${random}` as T;
}
