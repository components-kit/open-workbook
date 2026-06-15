import {
  makeId,
  type AgentId,
  type BackupId,
  type DiffSummary,
  type LockId,
  type OperationTelemetry,
  type OperationWarning,
  type PlanId,
  type RangeFingerprint,
  type TaskId,
  type TransactionId,
  type TransactionRecord,
  type TransactionStatus,
  type WorkbookId,
  type WorkbookScope,
  type DestructiveLevel
} from "@components-kit/open-workbook-protocol";

export interface CreateTransactionInput {
  workbookId: WorkbookId;
  agentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  planId?: PlanId | undefined;
  goal: string;
  scopes: WorkbookScope[];
  baseFingerprints?: RangeFingerprint[];
  destructiveLevel: DestructiveLevel;
  progressMessage?: string | undefined;
  retryStrategy?: string | undefined;
  chunksTotal?: number | undefined;
  chunksCompleted?: number | undefined;
}

export class TransactionManager {
  private readonly transactions = new Map<TransactionId, TransactionRecord>();

  create(input: CreateTransactionInput): TransactionRecord {
    const transaction: TransactionRecord = {
      transactionId: makeId<TransactionId>("txn"),
      workbookId: input.workbookId,
      status: "queued",
      goal: input.goal,
      scopes: input.scopes,
      locks: [],
      baseFingerprints: input.baseFingerprints ?? [],
      backups: [],
      warnings: [],
      destructiveLevel: input.destructiveLevel,
      queuedAt: new Date().toISOString()
    };
    if (input.progressMessage !== undefined) {
      transaction.progressMessage = input.progressMessage;
    }
    if (input.retryStrategy !== undefined) {
      transaction.retryStrategy = input.retryStrategy;
    }
    if (input.chunksTotal !== undefined) {
      transaction.chunksTotal = input.chunksTotal;
    }
    if (input.chunksCompleted !== undefined) {
      transaction.chunksCompleted = input.chunksCompleted;
    }
    if (input.agentId !== undefined) {
      transaction.agentId = input.agentId;
    }
    if (input.taskId !== undefined) {
      transaction.taskId = input.taskId;
    }
    if (input.planId !== undefined) {
      transaction.planId = input.planId;
    }
    this.transactions.set(transaction.transactionId, transaction);
    return transaction;
  }

  markApplying(transactionId: TransactionId, locks: LockId[]): TransactionRecord {
    const transaction = this.transactions.get(transactionId);
    const now = new Date().toISOString();
    return this.update(transactionId, {
      status: "applying",
      locks,
      startedAt: now,
      queuePosition: undefined,
      queueWaitMs: transaction ? Date.parse(now) - Date.parse(transaction.queuedAt) : undefined,
      progressMessage: applyingProgressMessage(transaction)
    });
  }

  markApplied(
    transactionId: TransactionId,
    result: {
      backups: BackupId[];
      warnings: OperationWarning[];
      diffSummary?: DiffSummary | undefined;
      telemetry?: OperationTelemetry | undefined;
    }
  ): TransactionRecord {
    const transaction = this.transactions.get(transactionId);
    const patch: Partial<TransactionRecord> = {
      status: "applied",
      backups: result.backups,
      warnings: result.warnings,
      finishedAt: new Date().toISOString(),
      progressMessage: appliedProgressMessage(transaction)
    };
    if (result.diffSummary !== undefined) {
      patch.diffSummary = result.diffSummary;
    }
    if (result.telemetry !== undefined) {
      patch.telemetry = result.telemetry;
    }
    if (transaction?.startedAt !== undefined) {
      patch.executionMs = Date.parse(patch.finishedAt!) - Date.parse(transaction.startedAt);
    }
    return this.update(transactionId, patch);
  }

  markFailed(transactionId: TransactionId, errorCode: string, errorMessage: string, warnings: OperationWarning[] = []): TransactionRecord {
    const finishedAt = new Date().toISOString();
    const transaction = this.transactions.get(transactionId);
    return this.update(transactionId, {
      status: "failed",
      errorCode,
      errorMessage,
      warnings,
      finishedAt,
      executionMs: transaction?.startedAt !== undefined ? Date.parse(finishedAt) - Date.parse(transaction.startedAt) : undefined,
      progressMessage: `Workbook mutation failed: ${errorMessage}`
    });
  }

  markBlocked(transactionId: TransactionId, errorCode: string, errorMessage: string, warnings: OperationWarning[] = []): TransactionRecord {
    return this.update(transactionId, {
      status: "blocked",
      errorCode,
      errorMessage,
      warnings,
      finishedAt: new Date().toISOString(),
      progressMessage: `Workbook mutation is blocked: ${errorMessage}`
    });
  }

  markRolledBack(transactionId: TransactionId): TransactionRecord {
    return this.update(transactionId, {
      status: "rolled_back",
      finishedAt: new Date().toISOString(),
      progressMessage: "Workbook mutation was rolled back."
    });
  }

  markCancelled(transactionId: TransactionId, message = "Queued workbook mutation was cancelled."): TransactionRecord {
    return this.update(transactionId, {
      status: "cancelled",
      errorCode: "TRANSACTION_CANCELLED",
      errorMessage: message,
      finishedAt: new Date().toISOString(),
      progressMessage: message
    });
  }

  get(transactionId: TransactionId): TransactionRecord | undefined {
    return this.transactions.get(transactionId);
  }

  getLaterApplied(transaction: TransactionRecord): TransactionRecord[] {
    const ordered = [...this.transactions.values()];
    const index = ordered.findIndex((candidate) => candidate.transactionId === transaction.transactionId);
    const candidates = index >= 0 ? ordered.slice(index + 1) : ordered;
    return candidates.filter((candidate) => candidate.workbookId === transaction.workbookId && candidate.status === "applied");
  }

  list(workbookId?: WorkbookId): TransactionRecord[] {
    return [...this.transactions.values()]
      .filter((transaction) => workbookId === undefined || transaction.workbookId === workbookId)
      .sort((a, b) => (b.finishedAt ?? b.startedAt ?? b.queuedAt).localeCompare(a.finishedAt ?? a.startedAt ?? a.queuedAt));
  }

  queuePosition(transactionId: TransactionId): number | undefined {
    const queued = [...this.transactions.values()]
      .filter((transaction) => transaction.status === "queued")
      .sort((a, b) => a.queuedAt.localeCompare(b.queuedAt));
    const index = queued.findIndex((transaction) => transaction.transactionId === transactionId);
    return index >= 0 ? index + 1 : undefined;
  }

  withQueueMetadata(transaction: TransactionRecord): TransactionRecord {
    return {
      ...transaction,
      queuePosition: transaction.status === "queued" ? this.queuePosition(transaction.transactionId) : undefined,
      progressMessage: transaction.progressMessage ?? progressMessageForStatus(transaction)
    };
  }

  load(records: TransactionRecord[]): void {
    this.transactions.clear();
    for (const record of records) {
      this.transactions.set(record.transactionId, { ...record });
    }
  }

  dump(): TransactionRecord[] {
    return [...this.transactions.values()].map((transaction) => ({ ...transaction }));
  }

  markInterrupted(message = "Daemon restarted before the transaction finished."): TransactionRecord[] {
    const interrupted: TransactionRecord[] = [];
    for (const transaction of this.transactions.values()) {
      if (transaction.status === "queued" || transaction.status === "applying") {
        transaction.status = "blocked";
        transaction.errorCode = "DAEMON_RESTARTED";
        transaction.errorMessage = message;
        transaction.finishedAt = new Date().toISOString();
        transaction.progressMessage = message;
        interrupted.push(transaction);
      }
    }
    return interrupted;
  }

  private update(transactionId: TransactionId, patch: Partial<TransactionRecord> & { status?: TransactionStatus }): TransactionRecord {
    const transaction = this.transactions.get(transactionId);
    if (!transaction) {
      throw new Error(`Transaction not found: ${transactionId}`);
    }
    Object.assign(transaction, patch);
    return transaction;
  }
}

function progressMessageForStatus(transaction: TransactionRecord): string {
  switch (transaction.status) {
    case "queued":
      return "Workbook mutation is queued and will apply when earlier workbook work finishes.";
    case "applying":
      return "Workbook mutation is applying in Excel.";
    case "applied":
      return "Workbook mutation applied successfully.";
    case "failed":
      return transaction.errorMessage ? `Workbook mutation failed: ${transaction.errorMessage}` : "Workbook mutation failed.";
    case "blocked":
      return transaction.errorMessage ? `Workbook mutation is blocked: ${transaction.errorMessage}` : "Workbook mutation is blocked.";
    case "rolled_back":
      return "Workbook mutation was rolled back.";
    case "cancelled":
      return transaction.errorMessage ?? "Workbook mutation was cancelled.";
  }
}

function applyingProgressMessage(transaction: TransactionRecord | undefined): string {
  if (transaction?.chunksTotal !== undefined && transaction.chunksCompleted !== undefined) {
    return `Workbook mutation is applying chunk ${transaction.chunksCompleted + 1} of ${transaction.chunksTotal}.`;
  }
  return "Workbook mutation is applying in Excel.";
}

function appliedProgressMessage(transaction: TransactionRecord | undefined): string {
  if (transaction?.chunksTotal !== undefined && transaction.chunksCompleted !== undefined) {
    return `Workbook mutation chunk ${transaction.chunksCompleted + 1} of ${transaction.chunksTotal} applied successfully.`;
  }
  return "Workbook mutation applied successfully.";
}
