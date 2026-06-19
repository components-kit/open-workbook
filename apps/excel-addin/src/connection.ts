import type { JsonRpcMessage, JsonRpcRequest } from "@components-kit/open-workbook-protocol";
import { getActiveWorkbookContext } from "./excel-executor.js";
import { getRuntimeHelloPayload, handleAddinRpcRequest } from "./rpc-methods.js";

export interface AddinConnectionOptions {
  backendUrl: string;
  heartbeatMs: number;
  reconnectMs?: number;
  onStatus?: (status: string) => void;
}

export class AddinConnection {
  private socket?: WebSocket;
  private heartbeat: number | undefined;
  private reconnect: number | undefined;
  private closedByUser = false;

  constructor(private readonly options: AddinConnectionOptions) {}

  connect(): void {
    this.closedByUser = false;
    this.stopReconnect();
    this.options.onStatus?.(`Connecting to ${this.options.backendUrl}...`);
    this.socket = new WebSocket(this.options.backendUrl);
    this.socket.addEventListener("open", () => {
      this.options.onStatus?.("Connected to local Open Workbook runtime. Checking active workbook...");
      void this.sendHello();
      this.startHeartbeat();
    });
    this.socket.addEventListener("message", (event) => this.handleMessage(JSON.parse(String(event.data))));
    this.socket.addEventListener("error", () => {
      this.options.onStatus?.(`Could not connect to ${this.options.backendUrl}. Retrying...`);
    });
    this.socket.addEventListener("close", () => {
      this.stopHeartbeat();
      if (!this.closedByUser) {
        this.options.onStatus?.("Disconnected from local runtime. Retrying...");
        this.scheduleReconnect();
      }
    });
  }

  disconnect(): void {
    this.closedByUser = true;
    this.stopHeartbeat();
    this.stopReconnect();
    this.socket?.close();
  }

  private scheduleReconnect(): void {
    this.stopReconnect();
    this.reconnect = window.setTimeout(() => this.connect(), this.options.reconnectMs ?? 2_000);
  }

  private stopReconnect(): void {
    if (this.reconnect !== undefined) {
      window.clearTimeout(this.reconnect);
      this.reconnect = undefined;
    }
  }

  private startHeartbeat(): void {
    this.stopHeartbeat();
    this.heartbeat = window.setInterval(() => {
      this.sendNotification("addin.heartbeat", { at: new Date().toISOString() });
    }, this.options.heartbeatMs);
  }

  private stopHeartbeat(): void {
    if (this.heartbeat !== undefined) {
      window.clearInterval(this.heartbeat);
      this.heartbeat = undefined;
    }
  }

  private handleMessage(message: JsonRpcMessage): void {
    if ("method" in message && "id" in message) {
      void this.handleRequest(message);
    }
  }

  private async handleRequest(request: JsonRpcRequest): Promise<void> {
    try {
      const response = await handleAddinRpcRequest(request);
      if (!response.ok) {
        this.send({ jsonrpc: "2.0", id: request.id, error: { code: -32601, message: response.message } });
        return;
      }
      this.sendSuccess(request.id, response.result);
      if (request.method === "runtime.disconnect") {
        this.disconnect();
      }
    } catch (error) {
      this.send({
        jsonrpc: "2.0",
        id: request.id,
        error: {
          code: -32000,
          message: error instanceof Error ? error.message : String(error)
        }
      });
    }
  }

  private async sendHello(): Promise<void> {
    let activeWorkbook;
    try {
      activeWorkbook = await getActiveWorkbookContext();
    } catch {
      activeWorkbook = undefined;
    }
    this.sendNotification("addin.hello", {
      ...getRuntimeHelloPayload(),
      ...(activeWorkbook ? { activeWorkbook } : {}),
      connectedAt: new Date().toISOString()
    });
    this.options.onStatus?.(activeWorkbook ? "Workbook ready for local Open Workbook runtime." : "Connected to local Open Workbook runtime. Open a workbook to continue.");
  }

  private sendNotification(method: string, params: unknown): void {
    this.send({ jsonrpc: "2.0", method, params });
  }

  private sendSuccess(id: string | number, result: unknown): void {
    this.send({ jsonrpc: "2.0", id, result });
  }

  private send(message: JsonRpcMessage): void {
    if (this.socket?.readyState === WebSocket.OPEN) {
      this.socket.send(JSON.stringify(message));
    }
  }
}
