import type { AgentId, TaskBlocker, TaskId, WorkbookId, WorkbookScope } from "@components-kit/open-workbook-protocol";
import { makeId, type TaskRecord, type TransactionId, type PlanId, type BackupId } from "@components-kit/open-workbook-protocol";

export interface CreateTaskInput {
  workbookId: WorkbookId;
  goal: string;
  role?: string | undefined;
  priority?: TaskRecord["priority"] | undefined;
  assignedAgentId?: AgentId | undefined;
  allowedScopes?: WorkbookScope[] | undefined;
  dependencies?: TaskId[] | undefined;
}

export class TaskRegistry {
  private readonly tasks = new Map<TaskId, TaskRecord>();

  create(input: CreateTaskInput): TaskRecord {
    const now = new Date().toISOString();
    const task: TaskRecord = {
      taskId: makeId<TaskId>("task"),
      workbookId: input.workbookId,
      goal: input.goal,
      priority: input.priority ?? "normal",
      status: input.assignedAgentId ? "claimed" : "open",
      progress: 0,
      blockers: [],
      allowedScopes: input.allowedScopes ?? [],
      dependencies: input.dependencies ?? [],
      planIds: [],
      transactionIds: [],
      rollbackBackupIds: [],
      createdAt: now,
      updatedAt: now
    };
    if (input.role !== undefined) {
      task.role = input.role;
    }
    if (input.assignedAgentId !== undefined) {
      task.assignedAgentId = input.assignedAgentId;
    }
    this.tasks.set(task.taskId, task);
    return task;
  }

  claim(taskId: TaskId, agentId: AgentId): TaskRecord {
    return this.update(taskId, { assignedAgentId: agentId, status: "claimed" });
  }

  update(
    taskId: TaskId,
    patch: Partial<
      Pick<
        TaskRecord,
        "goal" | "role" | "priority" | "status" | "progress" | "currentStep" | "blockers" | "assignedAgentId" | "allowedScopes" | "dependencies" | "errorMessage"
      >
    >
  ): TaskRecord {
    const task = this.require(taskId);
    Object.assign(task, patch, { updatedAt: new Date().toISOString() });
    if (patch.status === "completed") {
      task.completedAt = new Date().toISOString();
    }
    if (patch.status === "failed") {
      task.failedAt = new Date().toISOString();
    }
    return task;
  }

  setProgress(taskId: TaskId, progress: number, currentStep?: string | undefined): TaskRecord {
    const clamped = Math.max(0, Math.min(100, Math.round(progress)));
    return this.update(taskId, currentStep === undefined ? { progress: clamped } : { progress: clamped, currentStep });
  }

  addBlocker(
    taskId: TaskId,
    input: Pick<TaskBlocker, "message" | "severity"> & {
      scope?: TaskBlocker["scope"] | undefined;
    }
  ): TaskRecord {
    const task = this.require(taskId);
    const blocker: TaskBlocker = {
      blockerId: makeId<string>("blocker"),
      message: input.message,
      severity: input.severity,
      status: "open",
      createdAt: new Date().toISOString()
    };
    if (input.scope !== undefined) {
      blocker.scope = input.scope;
    }
    return this.update(taskId, { blockers: [...task.blockers, blocker], status: input.severity === "blocked" ? "blocked" : task.status });
  }

  resolveBlocker(taskId: TaskId, blockerId: string): TaskRecord {
    const task = this.require(taskId);
    const now = new Date().toISOString();
    return this.update(taskId, {
      blockers: task.blockers.map((blocker) =>
        blocker.blockerId === blockerId
          ? {
              ...blocker,
              status: "resolved",
              resolvedAt: now
            }
          : blocker
      )
    });
  }

  attachPlan(taskId: TaskId, planId: PlanId): TaskRecord | undefined {
    const task = this.tasks.get(taskId);
    if (!task) {
      return undefined;
    }
    task.planIds = [...new Set([...task.planIds, planId])];
    task.updatedAt = new Date().toISOString();
    return task;
  }

  attachTransaction(taskId: TaskId, transactionId: TransactionId): TaskRecord | undefined {
    const task = this.tasks.get(taskId);
    if (!task) {
      return undefined;
    }
    task.transactionIds = [...new Set([...task.transactionIds, transactionId])];
    task.updatedAt = new Date().toISOString();
    return task;
  }

  attachBackups(taskId: TaskId, backupIds: BackupId[]): TaskRecord | undefined {
    const task = this.tasks.get(taskId);
    if (!task) {
      return undefined;
    }
    task.rollbackBackupIds = [...new Set([...task.rollbackBackupIds, ...backupIds])];
    task.updatedAt = new Date().toISOString();
    return task;
  }

  complete(taskId: TaskId): TaskRecord {
    return this.update(taskId, { status: "completed" });
  }

  fail(taskId: TaskId, errorMessage: string): TaskRecord {
    return this.update(taskId, { status: "failed", errorMessage });
  }

  cancel(taskId: TaskId): TaskRecord {
    return this.update(taskId, { status: "cancelled" });
  }

  get(taskId: TaskId): TaskRecord | undefined {
    return this.tasks.get(taskId);
  }

  list(workbookId?: WorkbookId): TaskRecord[] {
    return [...this.tasks.values()]
      .filter((task) => workbookId === undefined || task.workbookId === workbookId)
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  }

  load(records: TaskRecord[]): void {
    this.tasks.clear();
    for (const record of records) {
      this.tasks.set(record.taskId, {
        ...record,
        progress: record.progress ?? 0,
        blockers: record.blockers ?? []
      });
    }
  }

  dump(): TaskRecord[] {
    return [...this.tasks.values()].map((task) => ({ ...task }));
  }

  private require(taskId: TaskId): TaskRecord {
    const task = this.tasks.get(taskId);
    if (!task) {
      throw new Error(`Task not found: ${taskId}`);
    }
    return task;
  }
}
