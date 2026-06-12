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
} from "@open-workbook/protocol";

export interface CreateTransactionInput {
  workbookId: WorkbookId;
  agentId?: AgentId | undefined;
  taskId?: TaskId | undefined;
  planId?: PlanId | undefined;
  goal: string;
  scopes: WorkbookScope[];
  baseFingerprints?: RangeFingerprint[];
  destructiveLevel: DestructiveLevel;
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
    return this.update(transactionId, {
      status: "applying",
      locks,
      startedAt: new Date().toISOString()
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
    const patch: Partial<TransactionRecord> = {
      status: "applied",
      backups: result.backups,
      warnings: result.warnings,
      finishedAt: new Date().toISOString()
    };
    if (result.diffSummary !== undefined) {
      patch.diffSummary = result.diffSummary;
    }
    if (result.telemetry !== undefined) {
      patch.telemetry = result.telemetry;
    }
    return this.update(transactionId, patch);
  }

  markFailed(transactionId: TransactionId, errorCode: string, errorMessage: string, warnings: OperationWarning[] = []): TransactionRecord {
    return this.update(transactionId, {
      status: "failed",
      errorCode,
      errorMessage,
      warnings,
      finishedAt: new Date().toISOString()
    });
  }

  markBlocked(transactionId: TransactionId, errorCode: string, errorMessage: string, warnings: OperationWarning[] = []): TransactionRecord {
    return this.update(transactionId, {
      status: "blocked",
      errorCode,
      errorMessage,
      warnings,
      finishedAt: new Date().toISOString()
    });
  }

  markRolledBack(transactionId: TransactionId): TransactionRecord {
    return this.update(transactionId, { status: "rolled_back", finishedAt: new Date().toISOString() });
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
