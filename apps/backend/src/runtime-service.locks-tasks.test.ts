import { describe, expect, it } from "vitest";
import {
  NativeFileBridge,
  RuntimeService,
  appliedTransaction,
  existsSync,
  mkdirSync,
  mkdtempSync,
  operationOk,
  path,
  readFileSync,
  runtimeWithDynamicSnapshotHash,
  runtimeWithExecutingAddin,
  runtimeWithFormulaGraph,
  runtimeWithPersistentAddin,
  runtimeWithSnapshotHash,
  sleepForTest,
  snapshotResponse,
  tmpdir,
  writeFileSync,
  writeFormulaOperation,
  writeStyleOperation,
  writeValuesOperation
} from "./runtime-service.test-support.js";
import type { AgentId, OperationId, PlanId, RuntimeCapabilities, WorkbookId } from "./runtime-service.test-support.js";

describe("RuntimeService task progress", () => {
  it("tracks progress and task blockers", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_tasks" as WorkbookId;
    const created = runtime.createTask({ workbookId, goal: "Build dashboard" });

    const progressed = runtime.setTaskProgress(created.task.taskId, 45, "Creating chart area");
    const blocked = runtime.addTaskBlocker(created.task.taskId, {
      severity: "blocked",
      message: "Waiting for Raw Data table lock",
      scope: { type: "table", workbookId, tableName: "RawData" }
    });
    const blockerId = blocked.blocker?.blockerId ?? "";
    const resolved = runtime.resolveTaskBlocker(created.task.taskId, blockerId);

    expect(progressed.task.progress).toBe(45);
    expect(progressed.task.currentStep).toBe("Creating chart area");
    expect(blocked.task.status).toBe("blocked");
    expect(resolved.task.blockers.find((blocker) => blocker.blockerId === blockerId)?.status).toBe("resolved");
  });

  it("evaluates task schedule from dependencies and locks", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_task_schedule" as WorkbookId;
    const dependency = runtime.createTask({ workbookId, goal: "Prepare raw data" });
    const waiting = runtime.createTask({
      workbookId,
      goal: "Build report",
      dependencies: [dependency.task.taskId],
      allowedScopes: [{ type: "range", workbookId, sheetName: "Report", address: "A1:F20" }]
    });
    runtime.locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Report", address: "C1:D5" }],
      mode: "write_values",
      reason: "Another agent is editing report"
    });

    const decisions = runtime.evaluateTaskSchedule({ workbookId }).decisions;
    const waitingDecision = decisions.find((decision) => decision.taskId === waiting.task.taskId);

    expect(waitingDecision?.state).toBe("waiting_dependencies");
    runtime.updateTask(dependency.task.taskId, { status: "completed" });
    const lockedDecision = runtime.evaluateTaskSchedule({ workbookId }).decisions.find((decision) => decision.taskId === waiting.task.taskId);

    expect(lockedDecision?.state).toBe("waiting_locks");
    expect(lockedDecision?.nextRetryAt).toBeDefined();
  });

  it("smokes multi-agent contention with one writer waiting on a lock", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_multi_agent_smoke" as WorkbookId;
    const agentA = "agent_cleaner" as AgentId;
    const agentB = "agent_reporter" as AgentId;
    runtime.registerAgent({ agentId: agentA, agentName: "Cleaner", clientType: "mcp" });
    runtime.registerAgent({ agentId: agentB, agentName: "Reporter", clientType: "mcp" });
    const cleaner = runtime.createTask({
      workbookId,
      goal: "Clean transaction log",
      assignedAgentId: agentA,
      allowedScopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }]
    }).task;
    const reporter = runtime.createTask({
      workbookId,
      goal: "Build formula summary",
      assignedAgentId: agentB,
      allowedScopes: [
        { type: "formula", workbookId, sheetName: "Transactions", address: "H1:H20" },
        { type: "range", workbookId, sheetName: "Transactions", address: "A1:A20" }
      ]
    }).task;
    const lock = runtime.acquireLocks({
      workbookId,
      ownerAgentId: agentA,
      taskId: cleaner.taskId,
      scopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }],
      mode: "write_values",
      reason: "Cleaner is updating transaction log"
    });
    expect(lock.ok).toBe(true);

    const schedule = runtime.evaluateTaskSchedule({ workbookId, lockMode: "write_formulas" });
    expect(schedule.decisions.find((decision) => decision.taskId === reporter.taskId)?.state).toBe("waiting_locks");

    const result = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      agentId: agentB,
      taskId: reporter.taskId,
      operations: [
        {
          kind: "range.write_formulas",
          operationId: "op_multi_agent_formula_lock" as OperationId,
          workbookId,
          destructiveLevel: "values",
          reason: "Write summary formula",
          target: { workbookId, sheetName: "Transactions", address: "H1" },
          formulas: [["=SUM(A1:A20)"]],
          preserveFormats: true
        }
      ]
    });

    expect(result.ok).toBe(false);
    expect(result.transactionId).toBeDefined();
    const transaction = runtime.transactions.get(result.transactionId!);
    expect(transaction?.status).toBe("blocked");
    expect(transaction?.agentId).toBe(agentB);
    expect(transaction?.taskId).toBe(reporter.taskId);
    expect(runtime.getTask(reporter.taskId).task?.status).toBe("blocked");
    const telemetry = runtime.getConflictTelemetry(workbookId);
    expect(telemetry.openCount).toBe(1);
    expect(telemetry.hotAgents.some((bucket) => bucket.key === agentA)).toBe(true);
    expect(telemetry.hotTasks.some((bucket) => bucket.key === cleaner.taskId)).toBe(true);
  });
});

describe("RuntimeService lock leases", () => {
  it("acquires, renews, and releases manual locks with policy ttl clamp", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_lock_leases" as WorkbookId;
    runtime.setLockPolicy({ defaultTtlMs: 1_000, maxTtlMs: 5_000 });

    const acquired = runtime.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:B2" }],
      mode: "write_values",
      reason: "Reserve edit range",
      ttlMs: 30_000
    });

    expect(acquired.ok).toBe(true);
    expect(Date.parse(acquired.locks[0]!.expiresAt) - Date.parse(acquired.locks[0]!.acquiredAt)).toBeLessThanOrEqual(5_000);
    const renewed = runtime.renewLocks([acquired.locks[0]!.lockId], 2_000);
    const released = runtime.releaseLocks([acquired.locks[0]!.lockId]);

    expect(renewed.ok).toBe(true);
    expect(renewed.renewed).toHaveLength(1);
    expect(released.ok).toBe(true);
    expect(released.released[0]?.status).toBe("released");
  });

  it("blocks manual locks when policy disables them", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_lock_policy" as WorkbookId;
    runtime.setLockPolicy({ allowManualLocks: false });

    const acquired = runtime.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:B2" }],
      mode: "write_values",
      reason: "Reserve edit range"
    });

    expect(acquired.ok).toBe(false);
    expect(acquired.conflicts[0]?.code).toBe("MANUAL_LOCKS_DISABLED");
  });

  it("blocks formula writes when a referenced range is locked", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_formula_lock" as WorkbookId;
    runtime.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:A10" }],
      mode: "write_values",
      reason: "Another agent is editing formula source"
    });

    const result = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      operations: [
        {
          kind: "range.write_formulas",
          operationId: "op_formula_lock" as OperationId,
          workbookId,
          destructiveLevel: "values",
          reason: "Write dependent formula",
          target: { workbookId, sheetName: "Sheet1", address: "B1" },
          formulas: [["=SUM(A1:A10)"]],
          preserveFormats: true
        }
      ]
    });

    expect(result.ok).toBe(false);
    expect(result.warnings.some((warning) => warning.code === "LOCK_CONFLICT")).toBe(true);
    const guidance = runtime.getConflictGuidance(workbookId);
    expect(guidance.guidance.some((item) => item.primaryAction === "retry_after" || item.primaryAction === "wait_for_lock")).toBe(true);
    const telemetry = runtime.getConflictTelemetry(workbookId);
    expect(telemetry.totalCount).toBe(1);
    expect(telemetry.openCount).toBe(1);
    expect(telemetry.hotScopes.some((bucket) => bucket.key.includes("Sheet1!A1:A10"))).toBe(true);
    runtime.releaseLocks(runtime.listLocks(workbookId).locks.map((lock) => lock.lockId));
    const afterRelease = runtime.getConflictTelemetry(workbookId);
    expect(afterRelease.clearedCount).toBe(1);
    expect(afterRelease.openCount).toBe(0);
  });

  it("clears conflict telemetry by workbook", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_conflict_telemetry_clear" as WorkbookId;
    runtime.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:A10" }],
      mode: "write_values",
      reason: "Reserve source"
    });
    await runtime.applyBatch({
      workbookId,
      mode: "apply",
      operations: [writeFormulaOperation(workbookId)]
    });

    expect(runtime.getConflictTelemetry(workbookId).totalCount).toBe(1);
    const cleared = runtime.clearConflictTelemetry(workbookId);

    expect(cleared.cleared).toBe(1);
    expect(runtime.getConflictTelemetry(workbookId).totalCount).toBe(0);
  });
});
