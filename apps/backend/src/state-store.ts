import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import path from "node:path";
import type { BackupRecord, PlanRecord, TemplateRecord } from "@component-kit/open-workbook-excel-core";
import type {
  AgentRecord,
  CollaborationEvent,
  ConflictRecord,
  ConflictTelemetryRecord,
  LockLeasePolicy,
  LockRecord,
  PermissionState,
  TaskRecord,
  TransactionRecord,
  WorkbookRegion
} from "@component-kit/open-workbook-protocol";

export interface RuntimeStateSnapshot {
  version: 1;
  savedAt: string;
  agents: AgentRecord[];
  tasks: TaskRecord[];
  locks: LockRecord[];
  lockLeasePolicy?: LockLeasePolicy | undefined;
  transactions: TransactionRecord[];
  conflicts: ConflictRecord[];
  conflictTelemetry?: ConflictTelemetryRecord[] | undefined;
  collaborationEvents: CollaborationEvent[];
  templates?: TemplateRecord[] | undefined;
  regions?: WorkbookRegion[] | undefined;
  permissions?: PermissionState | undefined;
  plans?: PlanRecord[] | undefined;
  backups?: BackupRecord[] | undefined;
}

export class RuntimeStateStore {
  readonly filePath: string;

  constructor(stateDir = defaultStateDir()) {
    this.filePath = path.join(stateDir, "collaboration-state.json");
  }

  load(): RuntimeStateSnapshot | undefined {
    if (!existsSync(this.filePath)) {
      return undefined;
    }
    const parsed = JSON.parse(readFileSync(this.filePath, "utf8")) as Partial<RuntimeStateSnapshot>;
    if (parsed.version !== 1) {
      return undefined;
    }
    return {
      version: 1,
      savedAt: parsed.savedAt ?? new Date(0).toISOString(),
      agents: parsed.agents ?? [],
      tasks: parsed.tasks ?? [],
      locks: parsed.locks ?? [],
      lockLeasePolicy: parsed.lockLeasePolicy,
      transactions: parsed.transactions ?? [],
      conflicts: parsed.conflicts ?? [],
      conflictTelemetry: parsed.conflictTelemetry ?? [],
      collaborationEvents: parsed.collaborationEvents ?? [],
      templates: parsed.templates ?? [],
      regions: parsed.regions ?? [],
      permissions: parsed.permissions,
      plans: parsed.plans ?? [],
      backups: parsed.backups ?? []
    };
  }

  save(snapshot: RuntimeStateSnapshot): void {
    mkdirSync(path.dirname(this.filePath), { recursive: true });
    const tmpPath = `${this.filePath}.${process.pid}.tmp`;
    writeFileSync(tmpPath, `${JSON.stringify(snapshot, null, 2)}\n`, "utf8");
    renameSync(tmpPath, this.filePath);
  }
}

function defaultStateDir(): string {
  return process.env.OPEN_WORKBOOK_STATE_DIR ?? path.resolve(process.cwd(), ".open-workbook/state");
}
