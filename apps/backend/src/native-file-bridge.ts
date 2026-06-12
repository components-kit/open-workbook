import type {
  WorkbookFileBridgeRequest,
  WorkbookFileBridgeResponse,
  WorkbookFileBridgeStatus
} from "@open-workbook/protocol";

export interface NativeFileBridgeOptions {
  url?: string;
  path?: string;
  timeoutMs?: number;
  fetchImpl?: typeof fetch;
}

export class NativeFileBridge {
  private readonly url: string | undefined;
  private readonly path: string;
  private readonly timeoutMs: number;
  private readonly fetchImpl: typeof fetch | undefined;

  constructor(options: NativeFileBridgeOptions = {}) {
    this.url = options.url ?? process.env.OPEN_WORKBOOK_FILE_BRIDGE_URL;
    this.path = normalizeBridgePath(options.path ?? process.env.OPEN_WORKBOOK_FILE_BRIDGE_PATH ?? "/v1/workbook-file");
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
      path: this.path,
      reason: "configured"
    };
  }

  async probeStatus(): Promise<WorkbookFileBridgeStatus> {
    const status = this.getStatus();
    if (!this.url || !this.fetchImpl) {
      return {
        ...status,
        reachable: false,
        checkedAt: new Date().toISOString(),
        error: "Native file bridge is not configured."
      };
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await this.fetchImpl(this.statusUrl(), {
        method: "GET",
        signal: controller.signal
      });
      const payload = await response.json().catch(() => ({})) as {
        ok?: boolean;
        bridge?: string;
        route?: string;
        adapter?: Record<string, unknown>;
        error?: string;
      };
      return {
        ...status,
        reachable: response.ok && payload.ok === true,
        checkedAt: new Date().toISOString(),
        statusCode: response.status,
        ...(payload.bridge !== undefined ? { bridge: payload.bridge } : {}),
        ...(payload.route !== undefined ? { route: payload.route } : {}),
        ...(payload.adapter !== undefined ? { adapter: payload.adapter } : {}),
        ...(payload.error !== undefined ? { error: payload.error } : {})
      };
    } catch (error) {
      return {
        ...status,
        reachable: false,
        checkedAt: new Date().toISOString(),
        error: error instanceof Error ? error.message : String(error)
      };
    } finally {
      clearTimeout(timeout);
    }
  }

  async request(input: WorkbookFileBridgeRequest): Promise<WorkbookFileBridgeResponse> {
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
    return base.endsWith("/") ? `${base.slice(0, -1)}${this.path}` : `${base}${this.path}`;
  }

  private statusUrl(): string {
    const base = this.url ?? "";
    return base.endsWith("/") ? `${base.slice(0, -1)}/status` : `${base}/status`;
  }
}

function normalizeBridgePath(path: string): string {
  if (!path) {
    return "/v1/workbook-file";
  }
  return path.startsWith("/") ? path : `/${path}`;
}
