import {
  makeId,
  type AgentId,
  type ConflictRecord,
  type LockId,
  type LockMode,
  type LockRecord,
  type TaskId,
  type TransactionId,
  type WorkbookId,
  type WorkbookScope
} from "@component-kit/open-workbook-protocol";
import { parseA1Address } from "./range-address.js";
import { makeLockConflict } from "./conflicts.js";

export interface AcquireLocksInput {
  workbookId: WorkbookId;
  ownerAgentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  transactionId?: TransactionId | undefined;
  scopes: WorkbookScope[];
  mode: LockMode;
  ttlMs?: number;
  reason: string;
}

export class LockManager {
  private readonly locks = new Map<LockId, LockRecord>();

  acquire(input: AcquireLocksInput): { ok: true; locks: LockRecord[] } | { ok: false; conflicts: ConflictRecord[] } {
    this.expireLocks();
    const conflicts = this.findConflicts(input.workbookId, input.scopes, input.mode, input.transactionId);
    if (conflicts.length > 0) {
      return { ok: false, conflicts };
    }

    const now = new Date();
    const expiresAt = new Date(now.getTime() + (input.ttlMs ?? 120_000)).toISOString();
    const locks = input.scopes.map((scope) => {
      const lock: LockRecord = {
        lockId: makeId<LockId>("lock"),
        workbookId: input.workbookId,
        mode: input.mode,
        scope,
        status: "active",
        acquiredAt: now.toISOString(),
        expiresAt,
        reason: input.reason
      };
      if (input.ownerAgentId !== undefined) {
        lock.ownerAgentId = input.ownerAgentId;
      }
      if (input.taskId !== undefined) {
        lock.taskId = input.taskId;
      }
      if (input.transactionId !== undefined) {
        lock.transactionId = input.transactionId;
      }
      this.locks.set(lock.lockId, lock);
      return lock;
    });
    return { ok: true, locks };
  }

  renew(lockIds: LockId[], ttlMs = 120_000): LockRecord[] {
    this.expireLocks();
    const expiresAt = new Date(Date.now() + ttlMs).toISOString();
    const renewed: LockRecord[] = [];
    for (const lockId of lockIds) {
      const lock = this.locks.get(lockId);
      if (lock?.status === "active") {
        lock.expiresAt = expiresAt;
        renewed.push(lock);
      }
    }
    return renewed;
  }

  renewWithMissing(lockIds: LockId[], ttlMs = 120_000): { renewed: LockRecord[]; missingLockIds: LockId[] } {
    const renewed = this.renew(lockIds, ttlMs);
    const renewedIds = new Set(renewed.map((lock) => lock.lockId));
    return {
      renewed,
      missingLockIds: lockIds.filter((lockId) => !renewedIds.has(lockId))
    };
  }

  release(lockIds: LockId[]): LockRecord[] {
    const releasedAt = new Date().toISOString();
    const released: LockRecord[] = [];
    for (const lockId of lockIds) {
      const lock = this.locks.get(lockId);
      if (lock?.status === "active") {
        lock.status = "released";
        lock.releasedAt = releasedAt;
        released.push(lock);
      }
    }
    return released;
  }

  releaseWithMissing(lockIds: LockId[]): { released: LockRecord[]; missingLockIds: LockId[] } {
    const released = this.release(lockIds);
    const releasedIds = new Set(released.map((lock) => lock.lockId));
    return {
      released,
      missingLockIds: lockIds.filter((lockId) => !releasedIds.has(lockId))
    };
  }

  list(workbookId?: WorkbookId): LockRecord[] {
    this.expireLocks();
    return [...this.locks.values()]
      .filter((lock) => workbookId === undefined || lock.workbookId === workbookId)
      .sort((a, b) => b.acquiredAt.localeCompare(a.acquiredAt));
  }

  load(records: LockRecord[]): void {
    this.locks.clear();
    for (const record of records) {
      this.locks.set(record.lockId, { ...record });
    }
    this.expireLocks();
  }

  dump(): LockRecord[] {
    this.expireLocks();
    return [...this.locks.values()].map((lock) => ({ ...lock }));
  }

  expireActive(reason = "Expired during runtime recovery."): LockRecord[] {
    const expiredAt = new Date().toISOString();
    const expired: LockRecord[] = [];
    for (const lock of this.locks.values()) {
      if (lock.status === "active") {
        lock.status = "expired";
        lock.releasedAt = expiredAt;
        lock.reason = `${lock.reason} ${reason}`.trim();
        expired.push(lock);
      }
    }
    return expired;
  }

  findConflicts(workbookId: WorkbookId, scopes: WorkbookScope[], mode: LockMode, transactionId?: TransactionId): ConflictRecord[] {
    this.expireLocks();
    if (mode === "read") {
      return [];
    }
    const conflicts: ConflictRecord[] = [];
    for (const existing of this.locks.values()) {
      if (existing.status !== "active" || existing.workbookId !== workbookId || (transactionId !== undefined && existing.transactionId === transactionId)) {
        continue;
      }
      if (existing.mode === "read") {
        continue;
      }
      for (const scope of scopes) {
        const conflict = makeLockConflict({
          workbookId,
          left: scope,
          right: existing.scope,
          ownerAgentId: existing.ownerAgentId,
          taskId: existing.taskId,
          transactionId: existing.transactionId
        });
        if (conflict) {
          if (conflict.code === "LOCK_CONFLICT") {
            conflict.message = `${conflict.message} Active ${existing.mode} lock is held on ${describeScope(existing.scope)}.`;
          }
          conflict.retryable = true;
          if (existing.ownerAgentId !== undefined) {
            conflict.ownerAgentId = existing.ownerAgentId;
          }
          if (existing.taskId !== undefined) {
            conflict.taskId = existing.taskId;
          }
          if (existing.transactionId !== undefined) {
            conflict.transactionId = existing.transactionId;
          }
          conflict.lockId = existing.lockId;
          conflict.lockExpiresAt = existing.expiresAt;
          conflicts.push(conflict);
        }
      }
    }
    return conflicts;
  }

  private expireLocks(): void {
    const now = Date.now();
    for (const lock of this.locks.values()) {
      if (lock.status === "active" && Date.parse(lock.expiresAt) <= now) {
        lock.status = "expired";
      }
    }
  }
}

export function scopesConflict(left: WorkbookScope, right: WorkbookScope): boolean {
  if (left.workbookId !== right.workbookId) {
    return false;
  }
  if (left.type === "workbook" || right.type === "workbook") {
    return true;
  }
  if (left.type === "sheet" || right.type === "sheet") {
    return scopeSheetName(left) === undefined || scopeSheetName(right) === undefined || scopeSheetName(left) === scopeSheetName(right);
  }
  if (left.type === "range" && right.type === "range") {
    return left.sheetName === right.sheetName && rangesOverlap(left.address, right.address);
  }
  if (left.type === "formula" && right.type === "formula") {
    return left.sheetName === right.sheetName && optionalFormulaRangesOverlap(left.address, right.address);
  }
  if (left.type === "formula" && right.type === "range") {
    return left.sheetName === right.sheetName && (left.address === undefined || rangesOverlap(left.address, right.address));
  }
  if (left.type === "range" && right.type === "formula") {
    return left.sheetName === right.sheetName && (right.address === undefined || rangesOverlap(left.address, right.address));
  }
  const leftSheet = scopeSheetName(left);
  const rightSheet = scopeSheetName(right);
  if (leftSheet !== undefined && rightSheet !== undefined && leftSheet !== rightSheet) {
    return false;
  }
  return left.type === right.type || left.type === "range" || right.type === "range";
}

function rangesOverlap(leftAddress: string, rightAddress: string): boolean {
  try {
    const left = parseA1Address(leftAddress);
    const right = parseA1Address(rightAddress);
    return left.startRow <= right.endRow && left.endRow >= right.startRow && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn;
  } catch {
    return true;
  }
}

function scopeSheetName(scope: WorkbookScope): string | undefined {
  switch (scope.type) {
    case "sheet":
    case "range":
    case "formula":
    case "table":
    case "named_range":
    case "chart":
    case "pivot":
      return scope.sheetName;
    case "workbook":
    case "template":
      return undefined;
  }
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

function optionalFormulaRangesOverlap(leftAddress: string | undefined, rightAddress: string | undefined): boolean {
  if (leftAddress === undefined || rightAddress === undefined) {
    return true;
  }
  return rangesOverlap(leftAddress, rightAddress);
}
