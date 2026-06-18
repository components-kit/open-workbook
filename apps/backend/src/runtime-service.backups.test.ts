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
      const pinnedDelete = await runtime.deleteFileBackup(backupId);
      const restored = await runtime.restoreFileBackup({
        workbookId,
        backupId,
        mode: "replace-open-workbook",
        force: true,
        restoreTargetPath: path.join(stateDir, "restored.xlsx")
      });
      const auditEvents = runtime.getCollaborationStatus(workbookId).events;
      const unpinned = runtime.pinFileBackup(backupId, false);
      const prunedDryRun = await runtime.pruneFileBackups({ workbookId, kind: "file-copy", maxBackupsPerWorkbook: 0, dryRun: true });

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

  it("prunes and deletes persisted JSON snapshot backup payloads", async () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-json-backup-"));
    const previousBackupDir = process.env.OPEN_WORKBOOK_BACKUP_DIR;
    const previousDisabled = process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED;
    process.env.OPEN_WORKBOOK_BACKUP_DIR = path.join(stateDir, "backups");
    process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED = "1";
    const workbookId = "workbook_json_backup" as WorkbookId;
    try {
      const runtime = runtimeWithPersistentAddin(stateDir, workbookId);
      const first = await runtime.createWorkbookBackup({
        workbookId,
        reason: "First JSON backup",
        ranges: [{ workbookId, sheetName: "Report", address: "A1:B2" }]
      });
      const second = await runtime.createWorkbookBackup({
        workbookId,
        reason: "Second JSON backup",
        ranges: [{ workbookId, sheetName: "Report", address: "C1:D2" }]
      });
      if (!("backup" in first) || !("backup" in second)) {
        throw new Error("Expected JSON backups to be created");
      }
      const firstPath = first.backup.payloadRef;
      const secondPath = second.backup.payloadRef;
      const firstExistedBefore = Boolean(firstPath && existsSync(firstPath));
      const secondExistedBefore = Boolean(secondPath && existsSync(secondPath));
      const dryRun = await runtime.pruneFileBackups({ workbookId, kind: "snapshot-json", maxBackupsPerWorkbook: 1, dryRun: true });
      const pruned = await runtime.pruneFileBackups({ workbookId, kind: "snapshot-json", maxBackupsPerWorkbook: 1 });
      const prunedId = (pruned as { pruned?: string[] }).pruned?.[0];
      const remaining = prunedId === first.backup.backupId ? second : first;
      const remainingPath = prunedId === first.backup.backupId ? secondPath : firstPath;
      const prunedPath = prunedId === first.backup.backupId ? firstPath : secondPath;
      const deleted = await runtime.deleteFileBackup(remaining.backup.backupId);

      expect(firstExistedBefore).toBe(true);
      expect(secondExistedBefore).toBe(true);
      expect((dryRun as { candidates?: Array<{ reasons?: string[] }> }).candidates?.[0]?.reasons).toContain("count");
      expect([first.backup.backupId, second.backup.backupId]).toContain(prunedId);
      expect(prunedPath && existsSync(prunedPath)).toBe(false);
      expect((deleted as { ok?: boolean }).ok).toBe(true);
      expect(remainingPath && existsSync(remainingPath)).toBe(false);
    } finally {
      if (previousBackupDir === undefined) {
        delete process.env.OPEN_WORKBOOK_BACKUP_DIR;
      } else {
        process.env.OPEN_WORKBOOK_BACKUP_DIR = previousBackupDir;
      }
      if (previousDisabled === undefined) {
        delete process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED;
      } else {
        process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED = previousDisabled;
      }
    }
  });

  it("reports byte-budget prune candidates across persisted backup kinds", async () => {
    const stateDir = mkdtempSync(path.join(tmpdir(), "open-workbook-byte-backup-"));
    const previousBackupDir = process.env.OPEN_WORKBOOK_BACKUP_DIR;
    const previousDisabled = process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED;
    process.env.OPEN_WORKBOOK_BACKUP_DIR = path.join(stateDir, "backups");
    process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED = "1";
    const workbookId = "workbook_byte_backup" as WorkbookId;
    try {
      const runtime = runtimeWithPersistentAddin(stateDir, workbookId);
      const first = await runtime.createWorkbookBackup({
        workbookId,
        reason: "Large JSON backup",
        ranges: [{ workbookId, sheetName: "Report", address: "A1:D20" }]
      });
      const second = await runtime.createWorkbookBackup({
        workbookId,
        reason: "Another JSON backup",
        ranges: [{ workbookId, sheetName: "Report", address: "A1:D20" }]
      });
      if (!("backup" in first) || !("backup" in second)) {
        throw new Error("Expected JSON backups to be created");
      }
      const dryRun = await runtime.pruneFileBackups({ workbookId, maxTotalBytes: 1, dryRun: true });

      expect((dryRun as { candidates?: Array<{ reasons?: string[]; bytes?: number }>; reclaimedBytes?: number }).candidates?.length).toBeGreaterThan(0);
      expect((dryRun as { candidates?: Array<{ reasons?: string[] }> }).candidates?.some((candidate) => candidate.reasons?.includes("size"))).toBe(true);
      expect((dryRun as { reclaimedBytes?: number }).reclaimedBytes).toBeGreaterThan(0);
      expect(first.backup.payloadRef && existsSync(first.backup.payloadRef)).toBe(true);
      expect(second.backup.payloadRef && existsSync(second.backup.payloadRef)).toBe(true);
    } finally {
      if (previousBackupDir === undefined) {
        delete process.env.OPEN_WORKBOOK_BACKUP_DIR;
      } else {
        process.env.OPEN_WORKBOOK_BACKUP_DIR = previousBackupDir;
      }
      if (previousDisabled === undefined) {
        delete process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED;
      } else {
        process.env.OPEN_WORKBOOK_BACKUP_RETENTION_DISABLED = previousDisabled;
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
