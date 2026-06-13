import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import type { AgentId, ExcelOperation, OperationId, PlanId, RuntimeCapabilities, TransactionRecord, WorkbookId, WorkbookScope } from "@components-kit/open-workbook-protocol";
import { NativeFileBridge } from "./native-file-bridge.js";
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
    expect(capabilities.fileBridge.available).toBe(false);
    expect(runtime.getStatus().fileBridge.available).toBe(false);
  });

  it("reports configured native file bridge status", () => {
    const runtime = new RuntimeService({
      persistState: false,
      fileBridge: new NativeFileBridge({ url: "http://127.0.0.1:37999" })
    });

    expect(runtime.getStatus().fileBridge).toMatchObject({
      available: true,
      url: "http://127.0.0.1:37999",
      path: "/v1/workbook-file"
    });
    expect(runtime.getCapabilities().fileBridge.available).toBe(true);
  });

  it("can probe configured native file bridge status", async () => {
    const runtime = new RuntimeService({
      persistState: false,
      fileBridge: new NativeFileBridge({
        url: "http://127.0.0.1:37999",
        fetchImpl: (async () => Response.json({
          ok: true,
          bridge: "open-workbook-native-file-bridge",
          route: "/v1/workbook-file",
          adapter: { platform: "win32", saveAsSupported: true }
        })) as typeof fetch
      })
    });

    const status = await runtime.getStatusWithFileBridgeProbe();

    expect(status.fileBridge).toMatchObject({
      available: true,
      reachable: true,
      bridge: "open-workbook-native-file-bridge",
      route: "/v1/workbook-file",
      adapter: { platform: "win32", saveAsSupported: true }
    });
  });

  it("includes active add-in Office API set support when connected", () => {
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    const reportedCapabilities: RuntimeCapabilities = {
      engine: {
        name: "office-js-addin",
        version: "0.1.1",
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

describe("RuntimeService selection", () => {
  it("passes through enriched selection metadata from the connected add-in", async () => {
    const workbookId = "workbook_runtime_selection" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        expect(method).toBe("runtime.get_selection");
        return {
          workbook: {
            workbookId,
            name: "Selection.xlsx",
            platform: "mac"
          },
          selection: {
            workbookId,
            sheetName: "Sheet1",
            address: "B4:D10",
            startCell: {
              workbookId,
              sheetName: "Sheet1",
              address: "B4",
              row: 4,
              column: 2,
              rowIndex: 3,
              columnIndex: 1
            },
            endCell: {
              workbookId,
              sheetName: "Sheet1",
              address: "D10",
              row: 10,
              column: 4,
              rowIndex: 9,
              columnIndex: 3
            },
            rowCount: 7,
            columnCount: 3,
            cellCount: 21,
            isSingleCell: false
          }
        };
      }
    } as any);

    const result = await runtime.getSelection();

    expect(result.selection?.startCell).toMatchObject({ address: "B4", row: 4, column: 2 });
    expect(result.selection?.endCell).toMatchObject({ address: "D10", row: 10, column: 4 });
    expect(result.selection).toMatchObject({ rowCount: 7, columnCount: 3, cellCount: 21, isSingleCell: false });
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

describe("RuntimeService PivotTable template copy", () => {
  it("records a backup and transaction for deterministic PivotTable copy", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_copy" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: params.pivotTableName === "TemplatePivot" ? "Template" : "Report",
              source: "Transactions",
              sourceType: "Table"
            }
          };
        }
        if (method === "workbook.get_map") {
          return {
            sheets: [
              { name: "Template", usedRange: { address: "A1:H20" } },
              { name: "Report", usedRange: { address: "A1:H20" } }
            ]
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "pivot_copy_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["snapshot"]],
              fingerprint: {
                range,
                hash: "pivot_copy_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "pivot.copy_from_template") {
          return {
            ok: true,
            copied: ["layout", "rowHierarchyPositions", "dataHierarchySettings"],
            source: { workbookId, pivotTableName: params.templatePivotTableName },
            target: { workbookId, pivotTableName: params.pivotTableName }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.copyPivotFromTemplate({
      workbookId,
      pivotTableName: "ReportPivot",
      templatePivotTableName: "TemplatePivot"
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect((result as { capabilityStatus?: { capabilities?: Array<{ capability: string; status: string }> } }).capabilityStatus?.capabilities).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ capability: "template_copy", status: "partial" }),
        expect.objectContaining({ capability: "source_reassignment", status: "unsupported" })
      ])
    );
    expect((result as { warnings?: Array<{ code: string }> }).warnings).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: "PIVOT_TEMPLATE_COPY_PARTIAL" })])
    );
    expect((result as { result?: { capabilityStatus?: unknown; warnings?: Array<{ code: string }> } }).result?.capabilityStatus).toBeDefined();
    expect((result as { result?: { warnings?: Array<{ code: string }> } }).result?.warnings).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: "PIVOT_TEMPLATE_COPY_PARTIAL" })])
    );
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.warnings.some((warning) => warning.code === "PIVOT_TEMPLATE_COPY_PARTIAL"))).toBe(true);
  });

  it("blocks PivotTable template copy when target source fields are incompatible", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_copy_incompatible" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: string[] = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push(method);
        if (method === "pivot.get_info") {
          if (params.pivotTableName === "TemplatePivot") {
            return {
              ok: true,
              info: {
                workbookId,
                pivotTableName: "TemplatePivot",
                sheetName: "Template",
                source: "TemplateTransactions",
                sourceType: "Table",
                hierarchies: [{ name: "Region" }, { name: "Amount" }],
                rowHierarchies: [{ name: "Region" }],
                columnHierarchies: [],
                filterHierarchies: [],
                dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" } }]
              }
            };
          }
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "ReportPivot",
              sheetName: "Report",
              source: "ReportTransactions",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }],
              rowHierarchies: [],
              columnHierarchies: [],
              filterHierarchies: [],
              dataHierarchies: []
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.copyPivotFromTemplate({
      workbookId,
      pivotTableName: "ReportPivot",
      templatePivotTableName: "TemplatePivot"
    });

    expect((result as { ok?: boolean }).ok).toBe(false);
    expect((result as { error?: { code?: string } }).error?.code).toBe("TEMPLATE_MISMATCH");
    expect((result as { issues?: Array<{ code: string; details?: Record<string, unknown> }> }).issues?.some((issue) => issue.code === "PIVOT_TEMPLATE_SOURCE_FIELD_MISSING" && issue.details?.field === "Amount")).toBe(true);
    expect(calls).toEqual(["pivot.get_info", "pivot.get_info"]);
    expect(runtime.transactions.list(workbookId)).toHaveLength(0);
  });
});

describe("RuntimeService PivotTable validation", () => {
  it("records a scoped backup and transaction before deleting a PivotTable", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_delete" as WorkbookId;
    runtime.allowDestructiveActions(true);
    const session = runtime.sessions.createSession();
    const calls: Array<{ method: string; params: any }> = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push({ method, params });
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: "Report",
              range: { address: "Report!B4:F20", rowCount: 17, columnCount: 5 }
            }
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "pivot_delete_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["pivot"]],
              fingerprint: {
                range,
                hash: "pivot_delete_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "pivot.delete") {
          return { ok: true, deleted: true };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.deletePivotTable({
      workbookId,
      pivotTableName: "ReportPivot"
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { transactionId?: string }).transactionId).toBeDefined();
    expect(calls.some((call) => call.method === "workbook.snapshot_ranges" && call.params.ranges[0].address === "B4:F20")).toBe(true);
    expect(calls.some((call) => call.method === "pivot.delete")).toBe(true);
    expect(runtime.transactions.list(workbookId).some((transaction) => transaction.status === "applied" && transaction.backups.length === 1)).toBe(true);
  });

  it("reports useful PivotTable metadata issues", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "EmptyPivot",
              sheetName: "Report",
              range: { address: "Report!A3:C10" },
              sourceType: "Table",
              dataHierarchies: []
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "EmptyPivot"
    });

    expect(result.ok).toBe(true);
    expect(result.summary.hasOutputRange).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_SOURCE_UNAVAILABLE")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_HAS_NO_DATA_FIELDS")).toBe(true);
  });

  it("validates expected PivotTable source and layout fields", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate_fields" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "SalesPivot",
              sheetName: "Report",
              range: { address: "Report!A3:E20" },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }, { name: "Month" }, { name: "Amount" }],
              rowHierarchies: [{ name: "Region" }],
              columnHierarchies: [{ name: "Month" }],
              filterHierarchies: [],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" }, summarizeBy: "sum", numberFormat: "$#,##0" }],
              layout: { showRowGrandTotals: true }
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      expectedFields: ["Region", "Month", "Amount"],
      expectedRowFields: ["Region"],
      expectedColumnFields: ["Month"],
      expectedDataFields: ["Amount"],
      expectedDataFieldSettings: [{ sourceFieldName: "Amount", summarizeBy: "sum", numberFormat: "$#,##0" }],
      expectedLayout: { showRowGrandTotals: true }
    });

    expect(result.ok).toBe(true);
    expect(result.summary.sourceFieldCount).toBe(3);
    expect(result.summary.rowFields).toEqual(["Region"]);
    expect(result.summary.dataFields).toEqual(["Amount"]);
    expect(result.issues).toHaveLength(0);
  });

  it("reports missing and misplaced expected PivotTable fields", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate_bad_fields" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "SalesPivot",
              sheetName: "Report",
              range: { address: "Report!A3:E20" },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }, { name: "Amount" }],
              rowHierarchies: [{ name: "Region" }],
              columnHierarchies: [],
              filterHierarchies: [],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" } }]
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      expectedFields: ["Region", "Customer"],
      expectedColumnFields: ["Region"],
      expectedDataFields: ["Customer"]
    });

    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_FIELD_MISSING" && issue.details?.field === "Customer")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_LAYOUT_MISMATCH" && issue.details?.axis === "column")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_LAYOUT_MISMATCH" && issue.details?.axis === "data")).toBe(true);
  });

  it("reports PivotTable aggregation, number format, and layout mismatches", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_validate_settings" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: "SalesPivot",
              sheetName: "Report",
              range: { address: "Report!A3:E20" },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Amount" }],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" }, summarizeBy: "sum", numberFormat: "$#,##0" }],
              layout: { showRowGrandTotals: true }
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      expectedDataFieldSettings: [{ sourceFieldName: "Amount", summarizeBy: "average", numberFormat: "0.00%" }],
      expectedLayout: { showRowGrandTotals: false }
    });

    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_AGGREGATION_MISMATCH")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_NUMBER_FORMAT_MISMATCH")).toBe(true);
    expect(result.issues.some((issue) => issue.code === "PIVOT_EXPECTED_LAYOUT_SETTING_MISMATCH")).toBe(true);
  });

  it("marks missing PivotTables as validation errors", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_missing" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        if (method === "pivot.get_info") {
          return { ok: false };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.validatePivotSource({
      workbookId,
      pivotTableName: "MissingPivot"
    });

    expect(result.ok).toBe(false);
    expect(result.issues.some((issue) => issue.code === "PIVOT_NOT_FOUND" && issue.severity === "error")).toBe(true);
  });

  it("captures and diffs deterministic PivotTable fingerprints", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_fingerprint" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "pivot.get_info") {
          const isTemplate = params.pivotTableName === "TemplatePivot";
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: "Report",
              range: { address: "Report!A3:E20", rowCount: 18, columnCount: 5 },
              source: "SalesTable",
              sourceType: "Table",
              hierarchies: [{ name: "Region" }, { name: "Month" }, { name: "Amount" }],
              rowHierarchies: [{ name: "Region" }],
              columnHierarchies: isTemplate ? [{ name: "Month" }] : [],
              filterHierarchies: [],
              dataHierarchies: [{ name: "Sum of Amount", field: { name: "Amount" }, summarizeBy: "sum", numberFormat: "$#,##0" }],
              layout: { showRowGrandTotals: true }
            }
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const fingerprint = await runtime.getPivotFingerprint({ workbookId, pivotTableName: "TemplatePivot" });
    const diff = await runtime.diffPivotTables({
      workbookId,
      pivotTableName: "TemplatePivot",
      targetPivotTableName: "ReportPivot"
    });

    expect((fingerprint as { ok?: boolean }).ok).toBe(true);
    expect((fingerprint as { fingerprint?: { hash?: string } }).fingerprint?.hash).toMatch(/^[a-f0-9]{16}$/);
    expect(diff.ok).toBe(false);
    expect(diff.changes.some((change) => change.path === "layout.columnFields")).toBe(true);
  });

  it("reports PivotTable source reassignment as an explicit capability limit", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_pivot_update_source" as WorkbookId;

    const result = runtime.updatePivotSource({
      workbookId,
      pivotTableName: "SalesPivot",
      sourceSheetName: "Data",
      sourceAddress: "A1:D100",
      destinationSheetName: "Report",
      destinationAddress: "B4"
    });

    expect(result.ok).toBe(false);
    expect(result.error.code).toBe("CAPABILITY_UNAVAILABLE");
    expect(result.capabilityStatus.fallback).toBe("excel.pivot.rebuild_with_source");
    expect(result.capabilityStatus.capabilities).toEqual(
      expect.arrayContaining([expect.objectContaining({ capability: "source_reassignment", status: "unsupported" })])
    );
    expect(result.warnings).toEqual(expect.arrayContaining([expect.objectContaining({ code: "PIVOT_SOURCE_REASSIGNMENT_UNSUPPORTED" })]));
  });

  it("rebuilds an existing PivotTable through explicit delete and create steps", async () => {
    const runtime = new RuntimeService({ persistState: false });
    runtime.allowDestructiveActions(true);
    const workbookId = "workbook_pivot_rebuild_replace" as WorkbookId;
    const session = runtime.sessions.createSession();
    const calls: string[] = [];
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        calls.push(method);
        if (method === "pivot.get_info") {
          return {
            ok: true,
            info: {
              workbookId,
              pivotTableName: params.pivotTableName,
              sheetName: "Report",
              range: { address: "Report!B4:F20", rowCount: 17, columnCount: 5 }
            }
          };
        }
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "pivot_rebuild_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["pivot"]],
              fingerprint: { range, hash: "pivot_rebuild_range", cellCount: 1, capturedAt: new Date().toISOString() }
            }))
          };
        }
        if (method === "pivot.delete") {
          return { ok: true, deleted: true };
        }
        if (method === "pivot.create") {
          return { ok: true, info: { workbookId, pivotTableName: params.pivotTableName, sheetName: params.destinationSheetName } };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.rebuildPivotWithSource({
      workbookId,
      pivotTableName: "SalesPivot",
      sourceSheetName: "Data",
      sourceAddress: "A1:D100",
      destinationSheetName: "Report",
      destinationAddress: "B4",
      rowFields: ["Region"],
      dataFields: [{ sourceFieldName: "Amount", summarizeBy: "sum" }],
      replaceExisting: true
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect(calls).toContain("pivot.delete");
    expect(calls).toContain("pivot.create");
    expect(runtime.transactions.list(workbookId).filter((transaction) => transaction.status === "applied")).toHaveLength(2);
  });
});

describe("RuntimeService durable file backups", () => {
  it("creates, verifies, pins, and prunes durable file backup manifests", async () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-file-backup-"));
    const previousBackupDir = process.env.OPEN_WORKBOOK_BACKUP_DIR;
    process.env.OPEN_WORKBOOK_BACKUP_DIR = path.join(stateDir, "backups");
    const workbookId = "workbook_file_backup" as WorkbookId;
    try {
      const bridge = new NativeFileBridge({
        url: "http://127.0.0.1:1",
        fetchImpl: async (_input: RequestInfo | URL, init?: RequestInit) => {
          const request = JSON.parse(String(init?.body ?? "{}")) as { operation?: string; targetPath?: string; restoreTargetPath?: string };
          if (request.targetPath) {
            mkdirSync(path.dirname(request.targetPath), { recursive: true });
            writeFileSync(request.targetPath, "xlsx backup payload", "utf8");
          }
          return new Response(
            JSON.stringify({
              ok: true,
              operation: request.operation ?? "workbook.export_copy",
              workbookId,
              targetPath: request.targetPath,
              filePath: request.restoreTargetPath ?? request.targetPath
            }),
            { status: 200, headers: { "content-type": "application/json" } }
          );
        }
      });
      const runtime = new RuntimeService({ persistState: false, fileBridge: bridge });
      runtime.setPermissions({ allowDestructiveActions: true, allowWorkbookActions: true });
      const session = runtime.sessions.createSession();
      runtime.attachAddinClient(session.connectionId, {
        request: async (method: string, params: any) => {
          if (method === "workbook.snapshot_ranges") {
            return {
              workbookFingerprint: {
                workbookId,
                workbookHash: "file_backup_workbook",
                structureHash: "structure",
                capturedAt: new Date().toISOString()
              },
              rangeSnapshots: (params.ranges ?? []).map((range: any) => ({
                range,
                values: [["snapshot"]],
                fingerprint: { range, hash: "file_backup_range", cellCount: 1, capturedAt: new Date().toISOString() }
              }))
            };
          }
          if (method === "workbook.get_map") {
            return { sheets: [{ name: "Sheet1", usedRange: { address: "A1:B2" } }] };
          }
          throw new Error(`Unexpected method ${method}`);
        }
      } as any);

      const created = await runtime.createFileBackup({ workbookId, reason: "Before risky report edit", pin: true });
      const backupId = (created as { manifest?: { backupId?: string } }).manifest?.backupId as any;
      const verified = await runtime.verifyFileBackup(backupId);
      const pinnedDelete = runtime.deleteFileBackup(backupId);
      const restored = await runtime.restoreFileBackup({
        workbookId,
        backupId,
        mode: "replace-open-workbook",
        force: true,
        restoreTargetPath: path.join(stateDir, "restored.xlsx")
      });
      const auditEvents = runtime.getCollaborationStatus(workbookId).events;
      const unpinned = runtime.pinFileBackup(backupId, false);
      const prunedDryRun = runtime.pruneFileBackups({ workbookId, maxBackupsPerWorkbook: 0, dryRun: true });

      expect((created as { ok?: boolean }).ok).toBe(true);
      expect((created as { manifest?: { checksum?: string; size?: number; pinned?: boolean } }).manifest?.checksum).toMatch(/^sha256:/);
      expect((created as { manifest?: { size?: number } }).manifest?.size).toBeGreaterThan(0);
      expect((verified as { ok?: boolean }).ok).toBe(true);
      expect((pinnedDelete as { ok?: boolean }).ok).toBe(false);
      expect((restored as { ok?: boolean }).ok).toBe(true);
      expect((restored as { emergencyBackup?: { ok?: boolean } }).emergencyBackup?.ok).toBe(true);
      expect(auditEvents.some((event) => event.type === "backup.created")).toBe(true);
      expect(auditEvents.some((event) => event.type === "backup.verified")).toBe(true);
      expect(auditEvents.some((event) => event.type === "backup.restored")).toBe(true);
      expect((unpinned as { ok?: boolean }).ok).toBe(true);
      expect((prunedDryRun as { candidates?: unknown[] }).candidates).toHaveLength(1);
    } finally {
      if (previousBackupDir === undefined) {
        delete process.env.OPEN_WORKBOOK_BACKUP_DIR;
      } else {
        process.env.OPEN_WORKBOOK_BACKUP_DIR = previousBackupDir;
      }
    }
  });
});

describe("RuntimeService native file bridge", () => {
  it("uses the configured bridge for workbook save_as", async () => {
    const workbookId = "workbook_file_bridge" as WorkbookId;
    let requestBody: any;
    const bridge = new NativeFileBridge({
      url: "http://127.0.0.1:37999",
      fetchImpl: (async (_url: string | URL | Request, init?: RequestInit) => {
        requestBody = JSON.parse(String(init?.body));
        return new Response(JSON.stringify({
          ok: true,
          operation: "workbook.save_as",
          workbookId,
          targetPath: "/tmp/report.xlsx",
          filePath: "/tmp/report.xlsx"
        }), { status: 200, headers: { "content-type": "application/json" } });
      }) as typeof fetch
    });
    const runtime = new RuntimeService({ persistState: false, fileBridge: bridge });

    const result = await runtime.saveWorkbookAs(workbookId, "/tmp/report.xlsx");

    expect(result.ok).toBe(true);
    expect((result as { targetPath?: string }).targetPath).toBe("/tmp/report.xlsx");
    expect(requestBody).toMatchObject({
      operation: "workbook.save_as",
      workbookId,
      targetPath: "/tmp/report.xlsx"
    });
  });

  it("writes workbook export copies from add-in compressed file payloads", async () => {
    const workbookId = "workbook_file_export" as WorkbookId;
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-file-export-"));
    const targetPath = path.join(stateDir, "exports", "report.xlsx");
    const runtime = new RuntimeService({ stateDir, persistState: false });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "file_export_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["snapshot"]],
              fingerprint: {
                range,
                hash: "file_export_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        if (method === "workbook.get_file") {
          return {
            ok: true,
            workbookId,
            fileType: "compressed",
            size: 8,
            sliceCount: 1,
            base64: Buffer.from("xlsxdata").toString("base64"),
            capturedAt: new Date().toISOString()
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.exportWorkbookCopy({
      workbookId,
      targetPath,
      ranges: [{ workbookId, sheetName: "Report", address: "A1:B2" }]
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect(readFileSync(targetPath, "utf8")).toBe("xlsxdata");
    expect((result as { file?: { method?: string } }).file?.method).toBe("office-js-compressed-file");
  });

  it("returns the native bridge file path for workbook export copies", async () => {
    const workbookId = "workbook_file_bridge_export" as WorkbookId;
    const bridgeTargetPath = "/tmp/open-workbook/report-copy.xlsx";
    let bridgeRequest: any;
    const bridge = new NativeFileBridge({
      url: "http://127.0.0.1:37999",
      fetchImpl: (async (_url: string | URL | Request, init?: RequestInit) => {
        bridgeRequest = JSON.parse(String(init?.body));
        return Response.json({
          ok: true,
          operation: "workbook.export_copy",
          workbookId,
          targetPath: bridgeTargetPath,
          filePath: bridgeTargetPath,
          sourceBackupId: bridgeRequest.sourceBackupId
        });
      }) as typeof fetch
    });
    const runtime = new RuntimeService({ persistState: false, fileBridge: bridge });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.snapshot_ranges") {
          return {
            workbookFingerprint: {
              workbookId,
              workbookHash: "bridge_export_workbook",
              structureHash: "structure",
              capturedAt: new Date().toISOString()
            },
            rangeSnapshots: params.ranges.map((range: any) => ({
              range,
              values: [["snapshot"]],
              fingerprint: {
                range,
                hash: "bridge_export_range",
                cellCount: 1,
                capturedAt: new Date().toISOString()
              }
            }))
          };
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.exportWorkbookCopy({
      workbookId,
      targetPath: "relative-report-copy.xlsx",
      ranges: [{ workbookId, sheetName: "Report", address: "A1:B2" }]
    });

    expect((result as { ok?: boolean }).ok).toBe(true);
    expect((result as { targetPath?: string }).targetPath).toBe(bridgeTargetPath);
    expect((result as { bridge?: { filePath?: string; sourceBackupId?: string } }).bridge?.filePath).toBe(bridgeTargetPath);
    expect((result as { bridge?: { sourceBackupId?: string } }).bridge?.sourceBackupId).toBeDefined();
    expect(bridgeRequest).toMatchObject({
      operation: "workbook.export_copy",
      workbookId,
      targetPath: "relative-report-copy.xlsx"
    });
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
