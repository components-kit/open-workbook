import type {
  A1Range,
  BackupId,
  WorkbookFileBridgeOperation,
  WorkbookFileBridgeResponse,
  WorkbookFileBridgeStatus,
  WorkbookId
} from "@open-workbook/protocol";

export interface NativeFileBridgeOptions {
  url?: string;
  timeoutMs?: number;
  fetchImpl?: typeof fetch;
}

export interface NativeFileBridgeRequest {
  operation: WorkbookFileBridgeOperation;
  workbookId: WorkbookId;
  targetPath?: string;
  sourceBackupId?: BackupId;
  ranges?: A1Range[];
  reason?: string;
}

export class NativeFileBridge {
  private readonly url: string | undefined;
  private readonly timeoutMs: number;
  private readonly fetchImpl: typeof fetch | undefined;

  constructor(options: NativeFileBridgeOptions = {}) {
    this.url = options.url ?? process.env.OPEN_WORKBOOK_FILE_BRIDGE_URL;
    this.timeoutMs = options.timeoutMs ?? Number(process.env.OPEN_WORKBOOK_FILE_BRIDGE_TIMEOUT_MS ?? 30000);
    this.fetchImpl = options.fetchImpl ?? globalThis.fetch;
  }

  getStatus(): WorkbookFileBridgeStatus {
    if (!this.url) {
      return {
        available: false,
        reason: "not_configured"
      };
    }
    return {
      available: true,
      url: this.url,
      reason: "configured"
    };
  }

  async request(input: NativeFileBridgeRequest): Promise<WorkbookFileBridgeResponse> {
    if (!this.url || !this.fetchImpl) {
      return {
        ok: false,
        operation: input.operation,
        workbookId: input.workbookId,
        ...(input.targetPath !== undefined ? { targetPath: input.targetPath } : {}),
        ...(input.sourceBackupId !== undefined ? { sourceBackupId: input.sourceBackupId } : {}),
        error: "Native file bridge is not configured."
      };
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await this.fetchImpl(this.operationUrl(), {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(input),
        signal: controller.signal
      });
      const payload = await response.json().catch(() => ({})) as Partial<WorkbookFileBridgeResponse>;
      const result: WorkbookFileBridgeResponse = {
        ok: response.ok && payload.ok === true,
        operation: input.operation,
        workbookId: input.workbookId,
        ...(input.targetPath !== undefined ? { targetPath: input.targetPath } : {}),
        ...(input.sourceBackupId !== undefined ? { sourceBackupId: input.sourceBackupId } : {}),
        ...payload
      };
      const error = response.ok ? payload.error : payload.error ?? `Native file bridge returned HTTP ${response.status}.`;
      if (error !== undefined) {
        result.error = error;
      }
      return result;
    } catch (error) {
      return {
        ok: false,
        operation: input.operation,
        workbookId: input.workbookId,
        ...(input.targetPath !== undefined ? { targetPath: input.targetPath } : {}),
        ...(input.sourceBackupId !== undefined ? { sourceBackupId: input.sourceBackupId } : {}),
        error: error instanceof Error ? error.message : String(error)
      };
    } finally {
      clearTimeout(timeout);
    }
  }

  private operationUrl(): string {
    const base = this.url ?? "";
    return base.endsWith("/") ? `${base}v1/workbook-file` : `${base}/v1/workbook-file`;
  }
}
