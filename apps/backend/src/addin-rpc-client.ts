import type { WebSocket } from "ws";
import type { JsonRpcFailure, JsonRpcMessage, JsonRpcRequest, JsonRpcSuccess } from "@open-workbook/protocol";

export interface AddinRpcClientOptions {
  timeoutMs: number;
}

interface PendingRequest {
  resolve(value: unknown): void;
  reject(error: Error): void;
  timeout: NodeJS.Timeout;
}

export class AddinRpcClient {
  private nextId = 1;
  private readonly pending = new Map<string | number, PendingRequest>();

  constructor(
    private readonly websocket: WebSocket,
    private readonly options: AddinRpcClientOptions
  ) {}

  request<TResult>(method: string, params?: unknown): Promise<TResult> {
    const id = this.nextId++;
    const message: JsonRpcRequest = {
      jsonrpc: "2.0",
      id,
      method,
      params
    };

    return new Promise<TResult>((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Timed out waiting for add-in method: ${method}`));
      }, this.options.timeoutMs);

      this.pending.set(id, {
        resolve: (value) => resolve(value as TResult),
        reject,
        timeout
      });
      this.websocket.send(JSON.stringify(message));
    });
  }

  handleMessage(message: JsonRpcMessage): boolean {
    if (!("id" in message) || "method" in message) {
      return false;
    }

    if (message.id === null) {
      return false;
    }

    const pending = this.pending.get(message.id);
    if (!pending) {
      return false;
    }

    clearTimeout(pending.timeout);
    this.pending.delete(message.id);

    if (isFailure(message)) {
      pending.reject(new Error(message.error.message));
      return true;
    }

    pending.resolve((message as JsonRpcSuccess).result);
    return true;
  }

  rejectAll(reason: string): void {
    for (const [id, pending] of this.pending.entries()) {
      clearTimeout(pending.timeout);
      pending.reject(new Error(reason));
      this.pending.delete(id);
    }
  }

  close(): void {
    this.rejectAll("Excel add-in connection closed by backend");
    this.websocket.close();
  }
}

function isFailure(message: JsonRpcMessage): message is JsonRpcFailure {
  return "error" in message;
}
