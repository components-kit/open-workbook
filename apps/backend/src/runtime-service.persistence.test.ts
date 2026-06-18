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
