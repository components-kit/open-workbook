import {
  type A1Range,
  type BackupId,
  type BackupRef,
  type OperationId,
  type WorkbookId,
  makeId
} from "@open-workbook/protocol";

export interface BackupRecord extends BackupRef {
  affectedRanges: A1Range[];
  retention: "session" | "persistent";
  payloadRef?: string;
  payload?: unknown;
}

export interface CreateBackupInput {
  workbookId: WorkbookId;
  kind: BackupRecord["kind"];
  reason: string;
  operationId?: OperationId;
  affectedRanges?: A1Range[];
  payloadRef?: string;
  payload?: unknown;
}

export class BackupManager {
  private readonly backups = new Map<BackupId, BackupRecord>();

  createBackup(input: CreateBackupInput): BackupRecord {
    const backupId = makeId<BackupId>("backup");
    const record: BackupRecord = {
      backupId,
      workbookId: input.workbookId,
      kind: input.kind,
      createdAt: new Date().toISOString(),
      reason: input.reason,
      affectedRanges: input.affectedRanges ?? [],
      retention: input.kind === "workbook-copy" ? "persistent" : "session"
    };
    if (input.operationId !== undefined) {
      record.operationId = input.operationId;
    }
    if (input.payloadRef !== undefined) {
      record.payloadRef = input.payloadRef;
    }
    if (input.payload !== undefined) {
      record.payload = input.payload;
    }
    this.backups.set(backupId, record);
    return record;
  }

  listBackups(workbookId: WorkbookId): BackupRecord[] {
    return [...this.backups.values()].filter((backup) => backup.workbookId === workbookId);
  }

  getBackup(backupId: BackupId): BackupRecord | undefined {
    return this.backups.get(backupId);
  }

  load(records: BackupRecord[]): void {
    this.backups.clear();
    for (const record of records) {
      this.backups.set(record.backupId, { ...record });
    }
  }

  dump(): BackupRecord[] {
    return [...this.backups.values()].map((backup) => ({ ...backup }));
  }

  assertRollbackAvailable(backupIds: BackupId[]): void {
    const missing = backupIds.filter((backupId) => !this.backups.has(backupId));
    if (missing.length > 0) {
      throw new Error(`Missing backup records: ${missing.join(", ")}`);
    }
  }
}
