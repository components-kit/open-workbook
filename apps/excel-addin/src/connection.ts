import type { JsonRpcMessage, JsonRpcRequest } from "@open-workbook/protocol";
import { executeBatch, getActiveWorkbookContext, snapshotRanges } from "./excel-executor.js";

export interface AddinConnectionOptions {
  backendUrl: string;
  heartbeatMs: number;
}

export class AddinConnection {
  private socket?: WebSocket;
  private heartbeat: number | undefined;

  constructor(private readonly options: AddinConnectionOptions) {}

  connect(): void {
    this.socket = new WebSocket(this.options.backendUrl);
    this.socket.addEventListener("open", () => {
      this.sendNotification("addin.hello", {
        host: "excel",
        runtime: "office-js",
        connectedAt: new Date().toISOString()
      });
      this.startHeartbeat();
    });
    this.socket.addEventListener("message", (event) => this.handleMessage(JSON.parse(String(event.data))));
    this.socket.addEventListener("close", () => this.stopHeartbeat());
  }

  disconnect(): void {
    this.stopHeartbeat();
    this.socket?.close();
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
      switch (request.method) {
        case "runtime.get_active_context":
          this.sendSuccess(request.id, await getActiveWorkbookContext());
          break;
        case "workbook.snapshot_ranges": {
          const params = request.params as { workbookId: string; ranges: Parameters<typeof snapshotRanges>[1] };
          this.sendSuccess(request.id, await snapshotRanges(params.workbookId, params.ranges));
          break;
        }
        case "operation.execute_batch":
          this.sendSuccess(request.id, await executeBatch(request.params as Parameters<typeof executeBatch>[0]));
          break;
        default:
          this.send({
            jsonrpc: "2.0",
            id: request.id,
            error: {
              code: -32601,
              message: `Method not implemented in add-in: ${request.method}`
            }
          });
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
