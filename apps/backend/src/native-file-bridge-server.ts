import { spawn } from "node:child_process";
import { access, mkdir } from "node:fs/promises";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { homedir } from "node:os";
import path from "node:path";
import {
  runtimeError,
  type WorkbookFileBridgeRequest,
  type WorkbookFileBridgeResponse,
  type WorkbookId
} from "@components-kit/open-workbook-protocol";

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
      exportCopySupported: platform === "darwin" || platform === "win32",
      restoreFileBackupSupported: platform === "darwin" || platform === "win32",
      operations: {
        "workbook.save_as": platform === "darwin" || platform === "win32",
        "workbook.export_copy": platform === "darwin" || platform === "win32",
        "workbook.restore_file_backup": platform === "darwin" || platform === "win32"
      },
      allowedDirs: allowedTargetDirectories()
    }),
    saveAs: async (request) => {
      const prepared = await prepareTargetFileRequest(request, "workbook.save_as");
      if (!prepared.ok) {
        return prepared.response;
      }
      if (platform === "darwin") {
        return saveAsWithAppleScript(prepared.request, execute);
      }
      if (platform === "win32") {
        return saveAsWithPowerShell(prepared.request, execute);
      }
      return bridgeError(request, `Native Save As is not supported on ${platform}.`);
    },
    exportCopy: async (request) => {
      const prepared = await prepareTargetFileRequest(request, "workbook.export_copy");
      if (!prepared.ok) {
        return prepared.response;
      }
      if (platform === "darwin") {
        return saveCopyAsWithAppleScript(prepared.request, execute);
      }
      if (platform === "win32") {
        return saveCopyAsWithPowerShell(prepared.request, execute);
      }
      return bridgeError(request, `Native Export Copy is not supported on ${platform}.`);
    },
    restoreFileBackup: async (request) => {
      const prepared = await prepareRestoreFileRequest(request);
      if (!prepared.ok) {
        return prepared.response;
      }
      if (platform === "darwin") {
        return restoreFileBackupWithAppleScript(prepared.request, execute);
      }
      if (platform === "win32") {
        return restoreFileBackupWithPowerShell(prepared.request, execute);
      }
      return bridgeError(request, `Native Restore File Backup is not supported on ${platform}.`);
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

type PreparedBridgeTarget =
  | { ok: true; request: WorkbookFileBridgeRequest & { targetPath: string } }
  | { ok: false; response: WorkbookFileBridgeResponse };

type PreparedBridgeRestore =
  | { ok: true; request: WorkbookFileBridgeRequest & { backupPath: string; restoreMode: NonNullable<WorkbookFileBridgeRequest["restoreMode"]> } }
  | { ok: false; response: WorkbookFileBridgeResponse };

async function prepareTargetFileRequest(
  request: WorkbookFileBridgeRequest,
  operation: "workbook.save_as" | "workbook.export_copy"
): Promise<PreparedBridgeTarget> {
  if (!request.targetPath) {
    return { ok: false, response: bridgeError(request, `targetPath is required for ${operation}.`) };
  }
  const resolvedTargetPath = path.resolve(request.targetPath);
  const pathError = validateTargetPath(resolvedTargetPath);
  if (pathError) {
    return { ok: false, response: bridgeError(request, pathError) };
  }
  try {
    await mkdir(path.dirname(resolvedTargetPath), { recursive: true });
  } catch (error) {
    return {
      ok: false,
      response: bridgeError(request, `Failed to create target directory for ${operation}: ${error instanceof Error ? error.message : String(error)}`)
    };
  }
  return {
    ok: true,
    request: {
      ...request,
      targetPath: resolvedTargetPath
    }
  };
}

async function prepareRestoreFileRequest(request: WorkbookFileBridgeRequest): Promise<PreparedBridgeRestore> {
  const backupPath = request.backupPath ?? request.targetPath;
  if (!backupPath) {
    return { ok: false, response: bridgeError(request, "backupPath or targetPath is required for workbook.restore_file_backup.") };
  }
  const restoreMode = request.restoreMode ?? "open-as-new";
  if (restoreMode === "restore-into-open-workbook") {
    return { ok: false, response: bridgeError(request, "restore-into-open-workbook is not supported by the native bridge. Use open-as-new or replace-open-workbook.") };
  }
  const resolvedBackupPath = path.resolve(backupPath);
  const backupPathError = validateTargetPath(resolvedBackupPath);
  if (backupPathError) {
    return { ok: false, response: bridgeError(request, backupPathError.replace("targetPath", "backupPath")) };
  }
  try {
    await access(resolvedBackupPath);
  } catch {
    return { ok: false, response: bridgeError(request, `Backup file does not exist: ${resolvedBackupPath}`) };
  }
  const prepared: WorkbookFileBridgeRequest & { backupPath: string; restoreMode: NonNullable<WorkbookFileBridgeRequest["restoreMode"]> } = {
    ...request,
    backupPath: resolvedBackupPath,
    targetPath: resolvedBackupPath,
    restoreMode
  };
  if (request.restoreTargetPath !== undefined) {
    const resolvedRestoreTargetPath = path.resolve(request.restoreTargetPath);
    const targetPathError = validateTargetPath(resolvedRestoreTargetPath);
    if (targetPathError) {
      return { ok: false, response: bridgeError(request, targetPathError.replace("targetPath", "restoreTargetPath")) };
    }
    prepared.restoreTargetPath = resolvedRestoreTargetPath;
  }
  return { ok: true, request: prepared };
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

async function saveCopyAsWithAppleScript(request: WorkbookFileBridgeRequest, execute: NativeCommandExecutor): Promise<WorkbookFileBridgeResponse> {
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
    activate object targetWorkbook
    do Visual Basic "ActiveWorkbook.SaveCopyAs " & quote & my escapeForVbaString(targetPath) & quote
  end tell
end run

on escapeForVbaString(valueText)
  set oldDelimiters to AppleScript's text item delimiters
  set AppleScript's text item delimiters to quote
  set valueParts to text items of valueText
  set AppleScript's text item delimiters to quote & quote
  set escapedText to valueParts as text
  set AppleScript's text item delimiters to oldDelimiters
  return escapedText
end escapeForVbaString
`;
  return nativeSaveAsResult(request, await execute("osascript", ["-e", script, request.workbookId, request.targetPath!]));
}

async function saveCopyAsWithPowerShell(request: WorkbookFileBridgeRequest, execute: NativeCommandExecutor): Promise<WorkbookFileBridgeResponse> {
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
$targetWorkbook.SaveCopyAs($targetPath)
`;
  return nativeSaveAsResult(request, await execute("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script, request.workbookId, request.targetPath!]));
}

async function restoreFileBackupWithAppleScript(
  request: WorkbookFileBridgeRequest & { backupPath: string; restoreMode: NonNullable<WorkbookFileBridgeRequest["restoreMode"]> },
  execute: NativeCommandExecutor
): Promise<WorkbookFileBridgeResponse> {
  const script = `
on run argv
  set workbookName to item 1 of argv
  set backupPath to item 2 of argv
  set restoreMode to item 3 of argv
  set explicitTargetPath to item 4 of argv
  tell application "Microsoft Excel"
    if it is not running then error "Microsoft Excel is not running"
    if restoreMode is "open-as-new" then
      open workbook workbook file name backupPath
      return backupPath
    end if
    if restoreMode is not "replace-open-workbook" then error "Unsupported restore mode: " & restoreMode
    set targetWorkbook to missing value
    repeat with candidateWorkbook in workbooks
      if (name of candidateWorkbook as string) is workbookName or (full name of candidateWorkbook as string) is workbookName then
        set targetWorkbook to candidateWorkbook
        exit repeat
      end if
    end repeat
    if targetWorkbook is missing value then error "Workbook not found: " & workbookName
    if explicitTargetPath is "" then
      set restoreTargetPath to full name of targetWorkbook as string
    else
      set restoreTargetPath to explicitTargetPath
    end if
    close targetWorkbook saving no
    do shell script "/bin/cp -f " & quoted form of backupPath & " " & quoted form of restoreTargetPath
    open workbook workbook file name restoreTargetPath
    return restoreTargetPath
  end tell
end run
`;
  return nativeRestoreResult(
    request,
    await execute("osascript", ["-e", script, request.workbookId, request.backupPath, request.restoreMode, request.restoreTargetPath ?? ""])
  );
}

async function restoreFileBackupWithPowerShell(
  request: WorkbookFileBridgeRequest & { backupPath: string; restoreMode: NonNullable<WorkbookFileBridgeRequest["restoreMode"]> },
  execute: NativeCommandExecutor
): Promise<WorkbookFileBridgeResponse> {
  const script = `
$workbookName = $args[0]
$backupPath = $args[1]
$restoreMode = $args[2]
$explicitTargetPath = $args[3]
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
if ($restoreMode -eq "open-as-new") {
  $excel.Workbooks.Open($backupPath) | Out-Null
  Write-Output $backupPath
  exit 0
}
if ($restoreMode -ne "replace-open-workbook") {
  throw "Unsupported restore mode: $restoreMode"
}
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
if ([string]::IsNullOrWhiteSpace($explicitTargetPath)) {
  $restoreTargetPath = $targetWorkbook.FullName
} else {
  $restoreTargetPath = $explicitTargetPath
}
$targetWorkbook.Close($false)
Copy-Item -LiteralPath $backupPath -Destination $restoreTargetPath -Force
$excel.Workbooks.Open($restoreTargetPath) | Out-Null
Write-Output $restoreTargetPath
`;
  return nativeRestoreResult(
    request,
    await execute("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script, request.workbookId, request.backupPath, request.restoreMode, request.restoreTargetPath ?? ""])
  );
}

function nativeSaveAsResult(request: WorkbookFileBridgeRequest, result: Awaited<ReturnType<NativeCommandExecutor>>): WorkbookFileBridgeResponse {
  if (result.code === 0) {
    return {
      ok: true,
      operation: request.operation,
      workbookId: request.workbookId,
      targetPath: request.targetPath!,
      ...(request.sourceBackupId !== undefined ? { sourceBackupId: request.sourceBackupId } : {}),
      filePath: request.targetPath!,
      metadata: {
        stdout: result.stdout.trim()
      }
    };
  }
  return bridgeError(request, result.stderr.trim() || result.stdout.trim() || `Native command exited with code ${result.code}.`);
}

function nativeRestoreResult(request: WorkbookFileBridgeRequest, result: Awaited<ReturnType<NativeCommandExecutor>>): WorkbookFileBridgeResponse {
  if (result.code === 0) {
    const restoredPath = result.stdout.trim() || request.restoreTargetPath || request.backupPath || request.targetPath;
    const response: WorkbookFileBridgeResponse = {
      ok: true,
      operation: request.operation,
      workbookId: request.workbookId,
      ...(request.targetPath !== undefined ? { targetPath: request.targetPath } : {}),
      ...(request.backupPath !== undefined ? { backupPath: request.backupPath } : {}),
      ...(request.restoreTargetPath !== undefined ? { restoreTargetPath: request.restoreTargetPath } : {}),
      ...(request.restoreMode !== undefined ? { restoreMode: request.restoreMode } : {}),
      ...(request.sourceBackupId !== undefined ? { sourceBackupId: request.sourceBackupId } : {}),
      metadata: {
        stdout: result.stdout.trim()
      }
    };
    if (restoredPath !== undefined) {
      response.filePath = restoredPath;
    }
    return response;
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
    ...(request.backupPath !== undefined ? { backupPath: request.backupPath } : {}),
    ...(request.restoreTargetPath !== undefined ? { restoreTargetPath: request.restoreTargetPath } : {}),
    ...(request.restoreMode !== undefined ? { restoreMode: request.restoreMode } : {}),
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
