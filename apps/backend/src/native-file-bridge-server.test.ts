import path from "node:path";
import { tmpdir } from "node:os";
import { describe, expect, it } from "vitest";
import type { WorkbookFileBridgeRequest, WorkbookId } from "@open-workbook/protocol";
import {
  createPlatformNativeHostBridgeAdapter,
  handleWorkbookFileBridgeRequest,
  startNativeFileBridgeServer
} from "./native-file-bridge-server.js";

describe("Native file bridge server", () => {
  const request: WorkbookFileBridgeRequest = {
    operation: "workbook.save_as",
    workbookId: "Book1.xlsx" as WorkbookId,
    targetPath: path.join(tmpdir(), "open-workbook-save-as.xlsx")
  };
  const exportRequest: WorkbookFileBridgeRequest = {
    operation: "workbook.export_copy",
    workbookId: "Book1.xlsx" as WorkbookId,
    targetPath: path.join(tmpdir(), "open-workbook-export-copy.xlsx"),
    sourceBackupId: "backup_1" as WorkbookFileBridgeRequest["sourceBackupId"]
  };

  it("routes save_as requests through the configured adapter", async () => {
    const result = await handleWorkbookFileBridgeRequest(request, {
      getStatus: () => ({ test: true }),
      saveAs: async (input) => ({
        ok: true,
        operation: input.operation,
        workbookId: input.workbookId,
        targetPath: input.targetPath,
        filePath: input.targetPath
      })
    });

    expect(result.ok).toBe(true);
    expect(result.filePath).toBe(request.targetPath);
  });

  it("runs the macOS Save As adapter through osascript", async () => {
    let command: string | undefined;
    let args: string[] = [];
    const adapter = createPlatformNativeHostBridgeAdapter("darwin", async (nextCommand, nextArgs) => {
      command = nextCommand;
      args = nextArgs;
      return { code: 0, stdout: "saved", stderr: "" };
    });

    const result = await adapter.saveAs(request);

    expect(result.ok).toBe(true);
    expect(command).toBe("osascript");
    expect(args).toContain(request.workbookId);
    expect(args).toContain(request.targetPath);
  });

  it("routes export_copy requests through the configured adapter", async () => {
    const result = await handleWorkbookFileBridgeRequest(exportRequest, {
      getStatus: () => ({ test: true }),
      saveAs: async () => {
        throw new Error("saveAs should not be called for export_copy");
      },
      exportCopy: async (input) => ({
        ok: true,
        operation: input.operation,
        workbookId: input.workbookId,
        targetPath: input.targetPath,
        sourceBackupId: input.sourceBackupId,
        filePath: input.targetPath
      })
    });

    expect(result.ok).toBe(true);
    expect(result.operation).toBe("workbook.export_copy");
    expect(result.filePath).toBe(exportRequest.targetPath);
    expect(result.sourceBackupId).toBe(exportRequest.sourceBackupId);
  });

  it("reports platform operation support in status", () => {
    const adapter = createPlatformNativeHostBridgeAdapter("win32", async () => ({ code: 0, stdout: "", stderr: "" }));
    const status = adapter.getStatus() as { operations?: Record<string, boolean>; exportCopySupported?: boolean };

    expect(status.exportCopySupported).toBe(true);
    expect(status.operations?.["workbook.save_as"]).toBe(true);
    expect(status.operations?.["workbook.export_copy"]).toBe(true);
    expect(status.operations?.["workbook.restore_file_backup"]).toBe(false);
  });

  it("runs the macOS Export Copy adapter through SaveCopyAs", async () => {
    let command: string | undefined;
    let args: string[] = [];
    const adapter = createPlatformNativeHostBridgeAdapter("darwin", async (nextCommand, nextArgs) => {
      command = nextCommand;
      args = nextArgs;
      return { code: 0, stdout: "copied", stderr: "" };
    });

    const result = await adapter.exportCopy?.(exportRequest);

    expect(result?.ok).toBe(true);
    expect(command).toBe("osascript");
    expect(args.join("\n")).toContain("SaveCopyAs");
    expect(args).toContain(exportRequest.workbookId);
    expect(args).toContain(exportRequest.targetPath);
  });

  it("blocks Save As targets outside configured allowed directories", async () => {
    const previous = process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS;
    process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS = path.join(tmpdir(), "allowed-open-workbook-exports");
    try {
      const adapter = createPlatformNativeHostBridgeAdapter("darwin", async () => ({ code: 0, stdout: "", stderr: "" }));

      const result = await adapter.saveAs(request);

      expect(result.ok).toBe(false);
      expect(result.error).toContain("outside allowed bridge directories");
    } finally {
      if (previous === undefined) {
        delete process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS;
      } else {
        process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS = previous;
      }
    }
  });

  it("blocks Export Copy targets outside configured allowed directories", async () => {
    const previous = process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS;
    process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS = path.join(tmpdir(), "allowed-open-workbook-exports");
    try {
      const adapter = createPlatformNativeHostBridgeAdapter("win32", async () => ({ code: 0, stdout: "", stderr: "" }));

      const result = await adapter.exportCopy?.(exportRequest);

      expect(result?.ok).toBe(false);
      expect(result?.error).toContain("outside allowed bridge directories");
    } finally {
      if (previous === undefined) {
        delete process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS;
      } else {
        process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS = previous;
      }
    }
  });

  it("serves status and workbook file requests over HTTP", async () => {
    const server = await startNativeFileBridgeServer({
      host: "127.0.0.1",
      port: 0,
      adapter: {
        getStatus: () => ({ test: true }),
        saveAs: async (input) => ({
          ok: true,
          operation: input.operation,
          workbookId: input.workbookId,
          targetPath: input.targetPath,
          filePath: input.targetPath
        })
      }
    });
    try {
      const statusResponse = await fetch(`${server.url}/status`);
      const status = await statusResponse.json() as { ok?: boolean; adapter?: { test?: boolean } };
      const saveResponse = await fetch(`${server.url}${server.route}`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(request)
      });
      const save = await saveResponse.json() as { ok?: boolean; filePath?: string };

      expect(status.ok).toBe(true);
      expect(status.adapter?.test).toBe(true);
      expect(save.ok).toBe(true);
      expect(save.filePath).toBe(request.targetPath);
    } finally {
      await server.close();
    }
  });
});
