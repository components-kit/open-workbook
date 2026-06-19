import { describe, expect, it } from "vitest";
import { RuntimeService, appliedTransaction, operationOk, snapshotResponse, writeStyleOperation, writeValuesOperation } from "./runtime-service.test-support.js";
import type { AgentId, OperationId, PlanId, WorkbookId } from "./runtime-service.test-support.js";

describe("RuntimeService orchestration contracts", () => {
  it("covers runtime connection, capability, active workbook, active sheet, and ping contracts", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_runtime_contracts" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.sessions.update(session.connectionId, {
      activeWorkbook: { workbookId, name: "Runtime.xlsx", platform: "mac" } as any
    });
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "runtime.ping") {
          return { ok: true, pong: true, at: params.at };
        }
        if (method === "runtime.set_active_sheet") {
          return { ok: true, sheetName: params.sheetName };
        }
        throw new Error(`Unexpected method ${method}`);
      },
      close: () => undefined
    } as any);

    expect(runtime.connectAddinInfo()).toMatchObject({ ok: true, activeAddinConnected: true, activeWorkbookAvailable: true });
    expect(runtime.getCapabilities().internalCapabilities.total).toBeGreaterThan(0);
    expect(runtime.setActiveWorkbook("Runtime.xlsx").activeWorkbook?.workbookId).toBe(workbookId);
    await expect(runtime.setActiveSheet("Report")).resolves.toMatchObject({ ok: true, sheetName: "Report" });
    await expect(runtime.pingAddin()).resolves.toMatchObject({ ok: true, pong: true });
    expect(runtime.disconnectActiveAddin()).toMatchObject({ ok: true });
    expect(runtime.getStatus().connectionState).toBe("disconnected");
  });

  it("covers task lifecycle and collaboration list/view contracts", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_task_contracts" as WorkbookId;
    const agentId = "agent_task_contract" as AgentId;
    runtime.registerAgent({ agentId, agentName: "Task Contract", clientType: "mcp" });
    const created = runtime.createTask({ workbookId, goal: "Prepare report", priority: "high" });
    const claimed = runtime.claimTask(created.task.taskId, agentId);
    const updated = runtime.updateTask(created.task.taskId, { role: "writer", currentStep: "Planning" });
    const progressed = runtime.setTaskProgress(created.task.taskId, 30, "Reading source tables");
    const blocker = runtime.addTaskBlocker(created.task.taskId, { severity: "blocked", message: "Waiting for lock" });
    const resolved = runtime.resolveTaskBlocker(created.task.taskId, blocker.blocker!.blockerId);
    const schedule = runtime.evaluateTaskSchedule({ workbookId });
    const resumed = runtime.resumeReadyTasks(workbookId);

    expect(claimed.task.assignedAgentId).toBe(agentId);
    expect(updated.task.role).toBe("writer");
    expect(progressed.task.progress).toBe(30);
    expect(resolved.task.blockers[0]?.status).toBe("resolved");
    expect(schedule.decisions.some((decision) => decision.taskId === created.task.taskId)).toBe(true);
    expect(resumed.ok).toBe(true);
    expect(runtime.listTasks(workbookId).tasks).toHaveLength(1);
    expect(runtime.getTask(created.task.taskId).task?.goal).toBe("Prepare report");
    expect(runtime.completeTask(created.task.taskId).task.status).toBe("completed");

    const failed = runtime.createTask({ workbookId, goal: "Fail me" });
    const cancelled = runtime.createTask({ workbookId, goal: "Cancel me" });
    expect(runtime.failTask(failed.task.taskId, "contract failure").task.status).toBe("failed");
    expect(runtime.cancelTask(cancelled.task.taskId).task.status).toBe("cancelled");

    const collab = runtime.getCollaborationStatus(workbookId);
    expect(collab.agents.some((agent) => agent.agentId === agentId)).toBe(true);
    expect(collab.tasks).toHaveLength(3);
    expect(runtime.listAgents().agents.some((agent) => agent.agentId === agentId)).toBe(true);
  });

  it("runs agent requests under trusted execution identity", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const agentId = "agent_run_contract" as AgentId;
    const workbookId = "workbook_run_contract" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.sessions.update(session.connectionId, {
      activeWorkbook: { workbookId, name: "Run Contract.xlsx", platform: "mac" } as any
    });
    runtime.attachAddinClient(session.connectionId, { request: async () => ({ ok: true }) } as any);

    const result = await runtime.runAgent(
      { request: "Check connection", mode: "status" },
      { agentId, agentName: "Run Contract", clientType: "mcp" }
    );

    expect(result.status).toBe("SUCCESS");
    expect(runtime.currentAgentExecutionContext()).toBeUndefined();
    expect(runtime.listAgents().agents).toContainEqual(expect.objectContaining({ agentId, agentName: "Run Contract" }));
    expect((result.answer as any).collaboration.agents).toContainEqual(expect.objectContaining({ agentId, agentName: "Run Contract" }));
  });

  it("covers lock, conflict, transaction, and job read contracts", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_orchestration_contracts" as WorkbookId;
    runtime.setLockPolicy({ defaultTtlMs: 1_000, maxTtlMs: 5_000 });
    expect(runtime.getLockPolicy().policy.maxTtlMs).toBe(5_000);

    const acquired = runtime.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:A10" }],
      mode: "write_values",
      reason: "Contract lock"
    });
    expect(acquired.ok).toBe(true);
    expect(runtime.listLocks(workbookId).locks).toHaveLength(1);
    expect(runtime.renewLocks([acquired.locks[0]!.lockId], 2_000).renewed).toHaveLength(1);

    const blocked = runtime.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A2:A3" }],
      mode: "write_values",
      reason: "Conflicting lock"
    });
    expect(blocked.ok).toBe(false);
    expect(runtime.getConflictGuidance(workbookId).guidance.length).toBeGreaterThan(0);
    await runtime.applyBatch({
      workbookId,
      mode: "apply",
      operations: [{
        kind: "range.write_formulas",
        operationId: "op_explainable_conflict" as OperationId,
        workbookId,
        destructiveLevel: "values",
        reason: "Create explainable formula lock conflict",
        target: { workbookId, sheetName: "Sheet1", address: "B1" },
        formulas: [["=SUM(A1:A10)"]],
        preserveFormats: true
      }]
    });
    const explainableConflict = runtime.getCollaborationStatus(workbookId).conflicts.find((conflict) => Array.isArray(conflict.scopes));
    expect(explainableConflict).toBeDefined();
    expect(runtime.explainConflict(explainableConflict!.conflictId).ok).toBe(true);
    expect(runtime.getConflictTelemetry(workbookId).totalCount).toBeGreaterThan(0);
    expect(runtime.clearConflictTelemetry(workbookId).cleared).toBeGreaterThan(0);
    expect(runtime.releaseLocks([acquired.locks[0]!.lockId]).released).toHaveLength(1);

    const transaction = appliedTransaction(runtime, workbookId, "plan_contract" as PlanId, [
      { type: "range", workbookId, sheetName: "Sheet1", address: "B1:B2" }
    ]);
    expect(runtime.getTransaction(transaction.transactionId).transaction?.transactionId).toBe(transaction.transactionId);
    expect(runtime.listTransactions(workbookId).transactions.some((item) => item.transactionId === transaction.transactionId)).toBe(true);
    await expect(runtime.waitTransaction(transaction.transactionId, 100)).resolves.toMatchObject({ ok: true, completed: true });
    expect(runtime.cancelTransaction(transaction.transactionId).ok).toBe(false);
    expect(runtime.previewTransactionRollback(transaction.transactionId).rollbackAvailable).toBe(true);
    await expect(runtime.rollbackTransaction(transaction.transactionId)).resolves.toMatchObject({ ok: false });
    expect(runtime.previewTransactionRollbackChain(transaction.transactionId).rollbackOrder).toHaveLength(1);
    await expect(runtime.rollbackTransactionChain(transaction.transactionId)).resolves.toMatchObject({ ok: false });
  });

  it("covers batch, plan, job, and event contracts", async () => {
    const workbookId = "workbook_batch_plan_contracts" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.snapshot_ranges") {
          return snapshotResponse(workbookId, params.ranges);
        }
        if (method === "operation.execute_batch") {
          return operationOk();
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const operations = [
      writeValuesOperation(workbookId),
      {
        ...writeValuesOperation(workbookId),
        operationId: "op_contract_second" as OperationId,
        target: { workbookId, sheetName: "Sheet1", address: "A2" },
        values: [["second"]]
      }
    ];
    const preflight = runtime.preflightBatch({ workbookId, mode: "validate", operations });
    expect(preflight.ok).toBe(true);
    await expect(runtime.applyBatch({ workbookId, mode: "validate", operations })).resolves.toMatchObject({ ok: true });
    await expect(runtime.applyBatch({ workbookId, mode: "dry_run", operations })).resolves.toMatchObject({ ok: true });
    await expect(runtime.applyBatch({ workbookId, mode: "apply", operations })).resolves.toMatchObject({ ok: true });
    const submitted = runtime.submitBatch({ workbookId, mode: "apply", operations });
    expect(submitted.ok).toBe(true);
    await expect(runtime.waitTransaction(submitted.transactionId!, 1_000)).resolves.toMatchObject({ completed: true });

    const plan = runtime.createPlan({ workbookId, goal: "Plan contract", operations });
    const preview = await runtime.previewPlan(plan.planId);
    expect(preview.planId).toBe(plan.planId);
    await expect(runtime.refreshPlanPreview(plan.planId)).resolves.toMatchObject({ ok: true });
    await expect(runtime.rebasePlan(plan.planId)).resolves.toMatchObject({ ok: true });
    await expect(runtime.applyPlan(plan.planId)).resolves.toMatchObject({ ok: true });
    await expect(runtime.rollbackPlan(plan.planId)).resolves.toMatchObject({ ok: true });

    const manyOperations = Array.from({ length: 30 }, (_, index) => writeStyleOperation(workbookId, `C${index + 1}:C${index + 1}`));
    const job = runtime.submitChunkedBatch({ workbookId, mode: "apply", operations: manyOperations }, { goal: "Job contract" });
    expect(job.ok).toBe(true);
    expect(runtime.listJobs(workbookId).jobs.some((item) => item.jobId === job.jobId)).toBe(true);
    expect(runtime.getJob(job.jobId).job?.jobId).toBe(job.jobId);
    await expect(runtime.waitJob(job.jobId, 1_000)).resolves.toMatchObject({ ok: true, completed: true });
    expect(runtime.cancelJob(job.jobId).ok).toBe(true);

    expect(runtime.unsubscribeEvents()).toMatchObject({ ok: true, subscribed: false });
    expect(runtime.subscribeEvents()).toMatchObject({ ok: true, subscribed: true });
    runtime.recordAddinEvent(session.connectionId, "runtime.contract", { workbookId });
    expect(runtime.getRecentEvents(1).events[0]?.method).toBe("runtime.contract");
    expect(runtime.setEventDebounce(100_000).debounceMs).toBe(60_000);
    expect(runtime.clearEvents()).toMatchObject({ ok: true });
    expect(runtime.getRecentEvents().events).toHaveLength(0);
  });
});
