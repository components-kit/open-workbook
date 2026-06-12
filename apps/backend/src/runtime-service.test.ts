import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import type { AgentId, ExcelOperation, OperationId, PlanId, RuntimeCapabilities, TransactionRecord, WorkbookId, WorkbookScope } from "@open-workbook/protocol";
import { RuntimeService } from "./runtime-service.js";

describe("RuntimeService persistence", () => {
  it("restores agents and tasks from the state store", () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-state-"));
    const workbookId = "workbook_persist" as WorkbookId;
    const agentId = "agent_test" as AgentId;

    const first = new RuntimeService({ stateDir });
    first.registerAgent({ agentId, agentName: "persisted-agent", clientType: "mcp", pid: 1234 });
    const taskResult = first.createTask({
      workbookId,
      goal: "Clean transactions",
      assignedAgentId: agentId,
      allowedScopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }]
    });

    const second = new RuntimeService({ stateDir });
    const tasks = second.listTasks(workbookId).tasks;
    const agents = second.listAgents().agents;

    expect(tasks.some((task) => task.taskId === taskResult.task.taskId && task.goal === "Clean transactions")).toBe(true);
    expect(agents.some((agent) => agent.agentId === agentId && agent.agentName === "persisted-agent")).toBe(true);
  });

  it("persists blocked transactions caused by lock conflicts", async () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-state-lock-conflict-"));
    const workbookId = "workbook_persist_lock_conflict" as WorkbookId;
    const first = new RuntimeService({ stateDir });
    first.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:A10" }],
      mode: "write_values",
      reason: "Another agent reserved source cells"
    });

    const result = await first.applyBatch({
      workbookId,
      mode: "apply",
      operations: [writeFormulaOperation(workbookId)]
    });
    const transactionId = result.transactionId;
    const second = new RuntimeService({ stateDir });
    const transaction = transactionId === undefined ? undefined : second.transactions.get(transactionId);

    expect(result.ok).toBe(false);
    expect(transaction?.status).toBe("blocked");
    expect(transaction?.errorCode).toBe("LOCK_CONFLICT");
  });

  it("recovers active manual locks as expired after daemon restart", () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-state-lock-recovery-"));
    const workbookId = "workbook_persist_lock_recovery" as WorkbookId;
    const first = new RuntimeService({ stateDir });
    const acquired = first.acquireLocks({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Sheet1", address: "A1:B2" }],
      mode: "write_values",
      reason: "Reserve edit range"
    });
    expect(acquired.ok).toBe(true);

    const second = new RuntimeService({ stateDir });
    const locks = second.listLocks(workbookId).locks;

    expect(locks.some((lock) => lock.lockId === acquired.locks[0]?.lockId && lock.status === "expired")).toBe(true);
  });

  it("restores templates, regions, permissions, plans, and backup indexes", async () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-state-full-"));
    const previousBackupDir = process.env.OPEN_WORKBOOK_BACKUP_DIR;
    process.env.OPEN_WORKBOOK_BACKUP_DIR = path.join(stateDir, "backups");
    const workbookId = "workbook_persist_full" as WorkbookId;
    try {
      const first = runtimeWithPersistentAddin(stateDir, workbookId);

      first.setPermissions({ allowDestructiveActions: true, scope: { workbookId } });
      const plan = first.createPlan({
        workbookId,
        goal: "Persisted write plan",
        operations: [writeValuesOperation(workbookId)]
      });
      await first.previewPlan(plan.planId);
      const template = await first.registerTemplate({
        workbookId,
        name: "Monthly Report",
        scope: "workbook",
        sourceSheetName: "Template",
        dataRegions: ["A2:D20"]
      });
      await first.registerRegion({
        workbookId,
        name: "Data",
        sheetName: "Report",
        address: "A2:D20"
      });
      const backup = await first.createWorkbookBackup({
        workbookId,
        reason: "Persisted manual backup",
        ranges: [{ workbookId, sheetName: "Report", address: "A1:D20" }]
      });
      if (!("backup" in backup)) {
        throw new Error("Expected backup creation to succeed");
      }

      const second = new RuntimeService({ stateDir });

      expect(second.getPermissions().permissions.allowDestructiveActions).toBe(true);
      expect(second.getPermissions().permissions.scope.workbookId).toBe(workbookId);
      expect(second.listTemplates(workbookId).some((record) => record.templateId === template.templateId)).toBe(true);
      expect(second.listRegions(workbookId).regions.some((region) => region.name === "Data")).toBe(true);
      expect(second.plans.getPlan(plan.planId)?.goal).toBe("Persisted write plan");
      expect(second.backups.listBackups(workbookId).some((record) => record.backupId === backup.backup.backupId)).toBe(true);
    } finally {
      if (previousBackupDir === undefined) {
        delete process.env.OPEN_WORKBOOK_BACKUP_DIR;
      } else {
        process.env.OPEN_WORKBOOK_BACKUP_DIR = previousBackupDir;
      }
    }
  });

  it("exports and imports workbook local config metadata", async () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-local-config-"));
    const workbookId = "workbook_local_config" as WorkbookId;
    const source = runtimeWithPersistentAddin(stateDir, workbookId);

    source.setPermissions({
      allowDestructiveActions: true,
      scope: { workbookId, sheetNames: ["Report"], regionNames: ["Data"] }
    });
    const template = await source.registerTemplate({
      workbookId,
      name: "Monthly Report",
      scope: "workbook",
      sourceSheetName: "Template",
      dataRegions: ["A2:D20"]
    });
    await source.registerRegion({
      workbookId,
      name: "Data",
      sheetName: "Report",
      address: "A2:D20",
      templateId: template.templateId
    });
    await source.lockRegions({ workbookId, regions: [{ regionName: "Data", reason: "Protected output area" }] });

    const exported = source.exportWorkbookLocalConfig(workbookId);
    const target = new RuntimeService({ persistState: false });
    const imported = target.importWorkbookLocalConfig({
      workbookId,
      config: exported.config,
      overwrite: true
    });

    expect(imported.ok).toBe(true);
    expect(imported.imported.templates).toBe(1);
    expect(imported.imported.regions).toBe(1);
    expect(imported.imported.permissions).toBe(true);
    expect(target.listTemplates(workbookId).some((record) => record.templateId === template.templateId)).toBe(true);
    expect(target.listRegions(workbookId).regions.some((region) => region.name === "Data" && region.templateId === template.templateId)).toBe(true);
    expect(target.getPermissions().permissions.scope.workbookId).toBe(workbookId);
    expect(target.getPermissions().permissions.lockedRegions.some((region) => region.regionName === "Data")).toBe(true);
  });

  it("embeds local config through the add-in transaction path", async () => {
    const workbookId = "workbook_embed_config" as WorkbookId;
    const runtime = runtimeWithPersistentAddin(mkdtempSync(path.join(tmpdir(), "open-workbook-embed-config-")), workbookId);
    runtime.setPermissions({ allowDestructiveActions: true, allowWorkbookActions: true, requireConfirmationFor: [] });
    await runtime.registerRegion({
      workbookId,
      name: "Data",
      sheetName: "Report",
      address: "A2:D20"
    });
    let embeddedConfigWorkbookId: string | undefined;
    const session = runtime.sessions.getActive();
    if (!session) {
      throw new Error("Expected active add-in session");
    }
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.embed_local_config") {
          embeddedConfigWorkbookId = params.config.workbookId;
          return { ok: true, workbookId, embedded: true, partCount: 1 };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.embedWorkbookLocalConfig(workbookId);

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect(embeddedConfigWorkbookId).toBe(workbookId);
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.destructiveLevel === "workbook")).toBe(true);
  });

  it("imports embedded workbook local config from the add-in", async () => {
    const workbookId = "workbook_import_embedded_config" as WorkbookId;
    const source = runtimeWithPersistentAddin(mkdtempSync(path.join(tmpdir(), "open-workbook-import-embedded-source-")), workbookId);
    const template = await source.registerTemplate({
      workbookId,
      name: "Monthly Report",
      scope: "workbook",
      sourceSheetName: "Template",
      dataRegions: ["A2:D20"]
    });
    await source.registerRegion({
      workbookId,
      name: "Data",
      sheetName: "Report",
      address: "A2:D20",
      templateId: template.templateId
    });
    const exported = source.exportWorkbookLocalConfig(workbookId);
    const target = new RuntimeService({ persistState: false });
    const session = target.sessions.createSession();
    target.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "workbook.read_embedded_local_config") {
          return { ok: true, workbookId, embedded: true, partCount: 1, config: exported.config };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const imported = await target.importWorkbookEmbeddedLocalConfig({ workbookId, overwrite: true });

    expect(imported.ok).toBe(true);
    expect("imported" in imported && imported.imported.templates).toBe(1);
    expect(target.listTemplates(workbookId).some((record) => record.templateId === template.templateId)).toBe(true);
    expect(target.listRegions(workbookId).regions.some((region) => region.name === "Data")).toBe(true);
  });
});

describe("RuntimeService transaction rollback preview", () => {
  it("allows rollback preview when no later transaction overlaps", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_rollback" as WorkbookId;
    const planId = "plan_rollback" as PlanId;
    const scopes: WorkbookScope[] = [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }];
    const transaction = runtime.transactions.create({
      workbookId,
      planId,
      goal: "Clean transactions",
      scopes,
      destructiveLevel: "values"
    });
    runtime.transactions.markApplied(transaction.transactionId, { backups: [], warnings: [] });

    const preview = runtime.previewTransactionRollback(transaction.transactionId);

    expect(preview.ok).toBe(true);
    expect(preview.rollbackAvailable).toBe(true);
    expect(preview.rollbackMethod).toBe("plan");
  });

  it("blocks rollback when a later applied transaction overlaps the same scope", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_rollback_conflict" as WorkbookId;
    const earlier = appliedTransaction(runtime, workbookId, "plan_earlier" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);
    appliedTransaction(runtime, workbookId, "plan_later" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "D1:H20" }
    ]);

    const preview = runtime.previewTransactionRollback(earlier.transactionId);

    expect(preview.ok).toBe(false);
    expect(preview.conflicts.some((conflict) => conflict.code === "ROLLBACK_CONFLICT")).toBe(true);
  });

  it("blocks rollback when transaction has no plan metadata", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_no_plan" as WorkbookId;
    const transaction = appliedTransaction(runtime, workbookId, undefined, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);

    const preview = runtime.previewTransactionRollback(transaction.transactionId);

    expect(preview.ok).toBe(false);
    expect(preview.conflicts.some((conflict) => conflict.code === "ROLLBACK_UNAVAILABLE")).toBe(true);
  });

  it("reports dependency rollback conflict for later formula work", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_formula_dependency" as WorkbookId;
    const earlier = appliedTransaction(runtime, workbookId, "plan_values" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);
    appliedTransaction(runtime, workbookId, "plan_formulas" as PlanId, [
      { type: "formula", workbookId, sheetName: "Transactions", address: "D1:D20" }
    ]);

    const preview = runtime.previewTransactionRollback(earlier.transactionId);

    expect(preview.ok).toBe(false);
    expect(preview.conflicts.some((conflict) => conflict.code === "ROLLBACK_DEPENDENCY_CONFLICT")).toBe(true);
  });

  it("previews dependent rollback chains newest first", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_rollback_chain" as WorkbookId;
    const earlier = appliedTransaction(runtime, workbookId, "plan_chain_earlier" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);
    const later = appliedTransaction(runtime, workbookId, "plan_chain_later" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "D1:H20" }
    ]);

    const preview = runtime.previewTransactionRollbackChain(earlier.transactionId);

    expect(preview.ok).toBe(true);
    expect(preview.requiresConfirmation).toBe(true);
    expect(preview.rollbackOrder.map((transaction) => transaction.transactionId)).toEqual([later.transactionId, earlier.transactionId]);
    expect(preview.confirmationToken).toContain(earlier.transactionId);
  });
});

describe("RuntimeService plan refresh", () => {
  it("refreshes preview fingerprints when target ranges are unchanged", async () => {
    const runtime = runtimeWithSnapshotHash("hash_same");
    const workbookId = "workbook_refresh" as WorkbookId;
    const plan = runtime.createPlan({
      workbookId,
      goal: "Write values",
      operations: [writeValuesOperation(workbookId)]
    });
    await runtime.previewPlan(plan.planId);

    const result = await runtime.refreshPlanPreview(plan.planId);

    expect(result.ok).toBe(true);
    expect(result.refreshed).toBe(true);
    expect(result.conflicts).toHaveLength(0);
  });

  it("blocks refresh when target ranges changed since preview", async () => {
    let hash = "hash_before";
    const runtime = runtimeWithDynamicSnapshotHash(() => hash);
    const workbookId = "workbook_refresh_conflict" as WorkbookId;
    const plan = runtime.createPlan({
      workbookId,
      goal: "Write values",
      operations: [writeValuesOperation(workbookId)]
    });
    await runtime.previewPlan(plan.planId);
    hash = "hash_after";

    const result = await runtime.refreshPlanPreview(plan.planId);

    expect(result.ok).toBe(false);
    expect(result.refreshed).toBe(false);
    expect(result.conflicts.some((conflict) => conflict.code === "TARGET_REGION_CHANGED")).toBe(true);
  });
});

describe("RuntimeService formula dependency graph", () => {
  it("resolves structured references with table metadata from the add-in", async () => {
    const workbookId = "workbook_runtime_formula_graph" as WorkbookId;
    const runtime = runtimeWithFormulaGraph(workbookId);

    const result = await runtime.getFormulaDependencyGraph({
      workbookId,
      sheetName: "Report",
      address: "B2"
    });

    expect(result.ok).toBe(true);
    if (result.ok) {
      const rangeNodes = result.graph.nodes.filter((node) => node.kind === "range" && node.formula === undefined);
      expect(rangeNodes.some((node) => node.sheetName === "Transactions" && node.address === "C2:C10")).toBe(true);
    }
  });
});

describe("RuntimeService capabilities", () => {
  it("reports disconnected host capability fallback", () => {
    const runtime = new RuntimeService({ persistState: false });

    const capabilities = runtime.getCapabilities();

    expect(capabilities.activeHostCapabilities.engine.name).toBe("open-workbook-daemon");
    expect(capabilities.activeHostCapabilities.hostCapabilities?.some((capability) => capability.status === "unknown")).toBe(true);
    expect(capabilities.connectedHostCapabilities).toHaveLength(0);
  });

  it("includes active add-in Office API set support when connected", () => {
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    const reportedCapabilities: RuntimeCapabilities = {
      engine: {
        name: "office-js-addin",
        version: "0.1.0",
        platform: "mac",
        host: "Excel",
        officeVersion: "16.99"
      },
      apiSets: [
        { set: "ExcelApi", version: "1.9", supported: true },
        { set: "ExcelApi", version: "1.17", supported: false }
      ],
      capabilities: [
        {
          name: "range.batch.read_write",
          supported: true,
          platforms: ["mac", "windows", "web"],
          requires: [{ set: "ExcelApi", version: "1.9" }]
        }
      ],
      hostCapabilities: [
        {
          name: "range-values-formulas-styles",
          supported: true,
          status: "supported",
          requires: [{ set: "ExcelApi", version: "1.9" }]
        }
      ]
    };
    runtime.sessions.update(session.connectionId, { capabilities: reportedCapabilities });

    const capabilities = runtime.getCapabilities();

    expect(capabilities.activeHostCapabilities.engine.platform).toBe("mac");
    expect(capabilities.activeHostCapabilities.apiSets?.find((apiSet) => apiSet.version === "1.9")?.supported).toBe(true);
    const connectedHost = capabilities.connectedHostCapabilities[0];
    expect(connectedHost?.connectionId).toBe(session.connectionId);
    expect(connectedHost?.capabilities?.hostCapabilities?.[0]?.status).toBe("supported");
  });
});

describe("RuntimeService chart template copy", () => {
  it("records a backup and transaction for deterministic chart copy", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_chart_copy" as WorkbookId;
    runtime.allowDestructiveActions(true);
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.get_map") {
          return {
            sheets: [
              { name: "Template", usedRange: { address: "A1:D10" } },
              { name: "Report", usedRange: { address: "A1:D10" } }
            ]
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "chart_copy_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["snapshot"]],
              fingerprint: {
                range,
                hash: "chart_copy_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "chart.copy_from_template") {
          return {
            ok: true,
            copied: ["chartType", "style", "title", "position"],
            source: { workbookId, sheetName: params.templateSheetName, chartName: params.templateChartName },
            target: { workbookId, sheetName: params.sheetName, chartName: params.chartName }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.copyChartFromTemplate({
      workbookId,
      sheetName: "Report",
      chartName: "Revenue",
      templateSheetName: "Template",
      templateChartName: "TemplateRevenue"
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
  });
});

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

function appliedTransaction(runtime: RuntimeService, workbookId: WorkbookId, planId: PlanId | undefined, scopes: WorkbookScope[]): TransactionRecord {
  const transaction = runtime.transactions.create({
    workbookId,
    planId,
    goal: "Applied transaction",
    scopes,
    destructiveLevel: "values"
  });
  runtime.transactions.markApplying(transaction.transactionId, []);
  runtime.transactions.markApplied(transaction.transactionId, { backups: [], warnings: [] });
  return transaction;
}

function runtimeWithSnapshotHash(hash: string): RuntimeService {
  return runtimeWithDynamicSnapshotHash(() => hash);
}

function runtimeWithDynamicSnapshotHash(hash: () => string): RuntimeService {
  const runtime = new RuntimeService({ persistState: false });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (_method: string, params: any) => ({
      workbookFingerprint: {
        workbookId: params.workbookId,
        workbookHash: `workbook_${hash()}`,
        structureHash: "structure",
        capturedAt: new Date().toISOString()
      },
      rangeSnapshots: params.ranges.map((range: any) => ({
        fingerprint: {
          range,
          hash: hash(),
          cellCount: 1,
          capturedAt: new Date().toISOString()
        }
      }))
    })
  } as any);
  return runtime;
}

function runtimeWithPersistentAddin(stateDir: string, workbookId: WorkbookId): RuntimeService {
  const runtime = new RuntimeService({ stateDir });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (method: string, params: any) => {
      if (method === "template.capture") {
        return {
          sourceSheetName: params.sourceSheetName,
          dataRegions: params.dataRegions,
          fingerprintPayload: {
            structure: { sheets: ["Template", "Report"] },
            formulas: { range: "B2:D20" },
            styles: { theme: "default" },
            filters: {},
            tables: {},
            printLayout: {}
          }
        };
      }
      if (method === "workbook.snapshot_ranges") {
        return {
          workbookFingerprint: {
            workbookId,
            workbookHash: "workbook_persist_hash",
            structureHash: "structure",
            capturedAt: new Date().toISOString()
          },
          rangeSnapshots: params.ranges.map((range: any) => ({
            range,
            values: [["snapshot"]],
            fingerprint: {
              range,
              hash: "range_persist_hash",
              cellCount: 1,
              capturedAt: new Date().toISOString()
            }
          }))
        };
      }
      throw new Error(`Unexpected method ${method}`);
    }
  } as any);
  return runtime;
}

function runtimeWithFormulaGraph(workbookId: WorkbookId): RuntimeService {
  const runtime = new RuntimeService({ persistState: false });
  const session = runtime.sessions.createSession();
  runtime.attachAddinClient(session.connectionId, {
    request: async (method: string) => {
      if (method === "formula.read_patterns") {
        return {
          workbookId,
          sheetName: "Report",
          address: "B2:B2",
          capturedAt: new Date().toISOString(),
          rowCount: 1,
          columnCount: 1,
          formulaCount: 1,
          formulas: [["=SUM(Transactions[Amount])"]],
          patternMatrix: [["hash"]],
          patterns: [],
          cells: [{ rowIndex: 0, columnIndex: 0, formula: "=SUM(Transactions[Amount])", patternHash: "hash" }],
          warnings: []
        };
      }
      if (method === "table.list") {
        return {
          ok: true,
          tables: [
            {
              workbookId,
              tableName: "Transactions",
              sheetName: "Transactions",
              address: "A1:D11",
              rowCount: 11,
              columnCount: 4,
              showHeaders: true,
              showTotals: true,
              columns: [
                { index: 0, name: "Date" },
                { index: 1, name: "Status" },
                { index: 2, name: "Amount" },
                { index: 3, name: "Memo" }
              ]
            }
          ]
        };
      }
      throw new Error(`Unexpected method ${method}`);
    }
  } as any);
  return runtime;
}

function writeValuesOperation(workbookId: WorkbookId): ExcelOperation {
  return {
    kind: "range.write_values",
    operationId: "op_refresh" as OperationId,
    workbookId,
    destructiveLevel: "values",
    reason: "Write values",
    target: { workbookId, sheetName: "Sheet1", address: "A1" },
    values: [["ok"]],
    preserveFormats: true
  };
}

function writeFormulaOperation(workbookId: WorkbookId): ExcelOperation {
  return {
    kind: "range.write_formulas",
    operationId: "op_formula" as OperationId,
    workbookId,
    destructiveLevel: "values",
    reason: "Write dependent formula",
    target: { workbookId, sheetName: "Sheet1", address: "B1" },
    formulas: [["=SUM(A1:A10)"]],
    preserveFormats: true
  };
}
