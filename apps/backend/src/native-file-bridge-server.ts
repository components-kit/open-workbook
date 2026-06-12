import { spawn } from "node:child_process";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { homedir } from "node:os";
import path from "node:path";
import {
  runtimeError,
  type WorkbookFileBridgeRequest,
  type WorkbookFileBridgeResponse,
  type WorkbookId
} from "@open-workbook/protocol";

export interface NativeFileBridgeServerOptions {
  host: string;
  port: number;
  path?: string;
  shutdownPath?: string;
  adapter?: NativeHostFileBridgeAdapter;
}

export interface NativeFileBridgeServerHandle {
  url: string;
  route: string;
  close(): Promise<void>;
}

export interface NativeHostFileBridgeAdapter {
  getStatus(): Record<string, unknown>;
  saveAs(request: WorkbookFileBridgeRequest): Promise<WorkbookFileBridgeResponse>;
  exportCopy?(request: WorkbookFileBridgeRequest): Promise<WorkbookFileBridgeResponse>;
  restoreFileBackup?(request: WorkbookFileBridgeRequest): Promise<WorkbookFileBridgeResponse>;
}

export type NativeCommandExecutor = (command: string, args: string[]) => Promise<{ code: number; stdout: string; stderr: string }>;

export function startNativeFileBridgeServer(options: NativeFileBridgeServerOptions): Promise<NativeFileBridgeServerHandle> {
  const route = options.path ?? "/v1/workbook-file";
  const shutdownPath = options.shutdownPath ?? "/shutdown";
  const adapter = options.adapter ?? createPlatformNativeHostBridgeAdapter();
  let handle: NativeFileBridgeServerHandle | undefined;
  const server = createServer((request, response) => {
    if (request.method === "GET" && request.url === "/status") {
      sendJson(response, 200, {
        ok: true,
        bridge: "open-workbook-native-file-bridge",
        route,
        adapter: adapter.getStatus()
      });
      return;
    }
    if (request.method === "POST" && request.url === shutdownPath) {
      sendJson(response, 200, { ok: true });
      setTimeout(() => {
        void handle?.close();
      }, 25);
      return;
    }
    if (request.method === "POST" && request.url === route) {
      void handleWorkbookFileBridgeHttpRequest(request, response, adapter);
      return;
    }
    sendJson(response, 404, { ok: false, error: "not_found" });
  });

  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(options.port, options.host, () => {
      server.off("error", reject);
      const address = server.address();
      const port = typeof address === "object" && address !== null ? address.port : options.port;
      handle = {
        url: `http://${options.host}:${port}`,
        route,
        close: async () => {
          await new Promise<void>((closeResolve) => server.close(() => closeResolve()));
        }
      };
      resolve(handle);
    });
  });
}

export async function handleWorkbookFileBridgeRequest(
  request: WorkbookFileBridgeRequest,
  adapter: NativeHostFileBridgeAdapter
): Promise<WorkbookFileBridgeResponse> {
  const validationError = validateBridgeRequest(request);
  if (validationError) {
    return bridgeError(request, validationError);
  }

  switch (request.operation) {
    case "workbook.save_as":
      return adapter.saveAs(request);
    case "workbook.export_copy":
      return adapter.exportCopy ? adapter.exportCopy(request) : bridgeError(request, "Native export_copy is not implemented by this bridge adapter.");
    case "workbook.restore_file_backup":
      return adapter.restoreFileBackup ? adapter.restoreFileBackup(request) : bridgeError(request, "Native restore_file_backup is not implemented by this bridge adapter.");
  }
}

export function createPlatformNativeHostBridgeAdapter(
  platform: NodeJS.Platform = process.platform,
  execute: NativeCommandExecutor = executeNativeCommand
): NativeHostFileBridgeAdapter {
  return {
    getStatus: () => ({
      platform,
      saveAsSupported: platform === "darwin" || platform === "win32",
      allowedDirs: allowedTargetDirectories()
    }),
    saveAs: async (request) => {
      if (!request.targetPath) {
        return bridgeError(request, "targetPath is required for workbook.save_as.");
      }
      const pathError = validateTargetPath(request.targetPath);
      if (pathError) {
        return bridgeError(request, pathError);
      }
      if (platform === "darwin") {
        return saveAsWithAppleScript(request, execute);
      }
      if (platform === "win32") {
        return saveAsWithPowerShell(request, execute);
      }
      return bridgeError(request, `Native Save As is not supported on ${platform}.`);
    }
  };
}

async function handleWorkbookFileBridgeHttpRequest(
  request: IncomingMessage,
  response: ServerResponse,
  adapter: NativeHostFileBridgeAdapter
): Promise<void> {
  try {
    const payload = await readJsonBody<WorkbookFileBridgeRequest>(request, 1024 * 1024);
    const result = await handleWorkbookFileBridgeRequest(payload, adapter);
    sendJson(response, 200, result);
  } catch (error) {
    sendJson(response, 400, {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

async function saveAsWithAppleScript(request: WorkbookFileBridgeRequest, execute: NativeCommandExecutor): Promise<WorkbookFileBridgeResponse> {
  const script = `
on run argv
  set workbookName to item 1 of argv
  set targetPath to item 2 of argv
  tell application "Microsoft Excel"
    if it is not running then error "Microsoft Excel is not running"
    set targetWorkbook to missing value
    repeat with candidateWorkbook in workbooks
      if (name of candidateWorkbook as string) is workbookName or (full name of candidateWorkbook as string) is workbookName then
        set targetWorkbook to candidateWorkbook
        exit repeat
      end if
    end repeat
    if targetWorkbook is missing value then error "Workbook not found: " & workbookName
    save workbook as targetWorkbook filename targetPath
  end tell
end run
`;
  return nativeSaveAsResult(request, await execute("osascript", ["-e", script, request.workbookId, request.targetPath!]));
}

async function saveAsWithPowerShell(request: WorkbookFileBridgeRequest, execute: NativeCommandExecutor): Promise<WorkbookFileBridgeResponse> {
  const script = `
$workbookName = $args[0]
$targetPath = $args[1]
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$targetWorkbook = $null
foreach ($workbook in $excel.Workbooks) {
  if ($workbook.Name -eq $workbookName -or $workbook.FullName -eq $workbookName) {
    $targetWorkbook = $workbook
    break
  }
}
if ($null -eq $targetWorkbook) {
  throw "Workbook not found: $workbookName"
}
$targetWorkbook.SaveAs($targetPath)
`;
  return nativeSaveAsResult(request, await execute("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script, request.workbookId, request.targetPath!]));
}

function nativeSaveAsResult(request: WorkbookFileBridgeRequest, result: Awaited<ReturnType<NativeCommandExecutor>>): WorkbookFileBridgeResponse {
  if (result.code === 0) {
    return {
      ok: true,
      operation: request.operation,
      workbookId: request.workbookId,
      targetPath: request.targetPath!,
      filePath: request.targetPath!,
      metadata: {
        stdout: result.stdout.trim()
      }
    };
  }
  return bridgeError(request, result.stderr.trim() || result.stdout.trim() || `Native command exited with code ${result.code}.`);
}

function validateBridgeRequest(request: WorkbookFileBridgeRequest): string | undefined {
  if (!request || typeof request !== "object") {
    return "Bridge request must be a JSON object.";
  }
  if (request.operation !== "workbook.save_as" && request.operation !== "workbook.export_copy" && request.operation !== "workbook.restore_file_backup") {
    return "Unsupported workbook file bridge operation.";
  }
  if (!request.workbookId || typeof request.workbookId !== "string") {
    return "workbookId is required.";
  }
  return undefined;
}

function validateTargetPath(targetPath: string): string | undefined {
  const resolved = path.resolve(targetPath);
  const allowedDirs = allowedTargetDirectories();
  if (allowedDirs.length === 0) {
    return undefined;
  }
  if (allowedDirs.some((directory) => isPathWithin(resolved, directory))) {
    return undefined;
  }
  return `targetPath is outside allowed bridge directories: ${allowedDirs.join(", ")}`;
}

function allowedTargetDirectories(): string[] {
  const raw = process.env.OPEN_WORKBOOK_FILE_BRIDGE_ALLOWED_DIRS;
  if (!raw) {
    return [];
  }
  return raw
    .split(path.delimiter)
    .map((entry) => path.resolve(entry.replace(/^~(?=$|\/|\\)/, homedir())))
    .filter(Boolean);
}

function isPathWithin(targetPath: string, directory: string): boolean {
  const relative = path.relative(directory, targetPath);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function bridgeError(request: Partial<WorkbookFileBridgeRequest>, message: string): WorkbookFileBridgeResponse {
  return {
    ok: false,
    operation: request.operation ?? "workbook.save_as",
    workbookId: (request.workbookId ?? "unknown") as WorkbookId,
    ...(request.targetPath !== undefined ? { targetPath: request.targetPath } : {}),
    ...(request.sourceBackupId !== undefined ? { sourceBackupId: request.sourceBackupId } : {}),
    error: runtimeError("OPERATION_FAILED", message, { retryable: false }).message
  };
}

function readJsonBody<T>(request: IncomingMessage, maxBytes: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let bytes = 0;
    request.on("data", (chunk: Buffer) => {
      bytes += chunk.length;
      if (bytes > maxBytes) {
        reject(new Error("Bridge request body is too large."));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });
    request.on("end", () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")) as T);
      } catch {
        reject(new Error("Bridge request body must be valid JSON."));
      }
    });
    request.on("error", reject);
  });
}

function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  response.writeHead(status, { "content-type": "application/json" });
  response.end(JSON.stringify(payload));
}

function executeNativeCommand(command: string, args: string[]): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const child = spawn(command, args, { stdio: ["ignore", "pipe", "pipe"] });
    const stdout: Buffer[] = [];
    const stderr: Buffer[] = [];
    child.stdout.on("data", (chunk: Buffer) => stdout.push(chunk));
    child.stderr.on("data", (chunk: Buffer) => stderr.push(chunk));
    child.on("error", (error) => {
      resolve({ code: 1, stdout: "", stderr: error.message });
    });
    child.on("close", (code) => {
      resolve({
        code: code ?? 1,
        stdout: Buffer.concat(stdout).toString("utf8"),
        stderr: Buffer.concat(stderr).toString("utf8")
      });
    });
  });
}
