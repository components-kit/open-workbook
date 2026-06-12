import { createServer, type IncomingMessage } from "node:http";
import { WebSocketServer, type WebSocket } from "ws";
import type { ConnectionId, JsonRpcMessage, JsonRpcNotification } from "@open-workbook/protocol";
import { RuntimeService } from "./runtime-service.js";
import type { AddinSession } from "./session-registry.js";
import { AddinRpcClient } from "./addin-rpc-client.js";

export interface BackendServerOptions {
  host: string;
  port: number;
  addinPath: string;
  rpcPath?: string;
  shutdownPath?: string;
}

export interface BackendServerHandle {
  close(): Promise<void>;
}

export function startBackendServer(runtime: RuntimeService, options: BackendServerOptions): Promise<BackendServerHandle> {
  const rpcPath = options.rpcPath ?? "/rpc";
  const shutdownPath = options.shutdownPath ?? "/shutdown";
  let handle: BackendServerHandle | undefined;
  const httpServer = createServer((request, response) => {
    if (request.url === "/status") {
      response.writeHead(200, { "content-type": "application/json" });
      response.end(JSON.stringify(runtime.getStatus()));
      return;
    }
    if (request.url === rpcPath && request.method === "POST") {
      handleRuntimeRpc(runtime, request, response);
      return;
    }
    if (request.url === shutdownPath && request.method === "POST") {
      response.writeHead(200, { "content-type": "application/json" });
      response.end(JSON.stringify({ ok: true }));
      setTimeout(() => {
        void handle?.close();
      }, 25);
      return;
    }

    response.writeHead(404, { "content-type": "application/json" });
    response.end(JSON.stringify({ error: "not_found" }));
  });

  const websocketServer = new WebSocketServer({
    noServer: true
  });

  httpServer.on("upgrade", (request, socket, head) => {
    if (!isAddinUpgrade(request, options.addinPath)) {
      socket.destroy();
      return;
    }

    websocketServer.handleUpgrade(request, socket, head, (websocket) => {
      websocketServer.emit("connection", websocket, request);
    });
  });

  websocketServer.on("connection", (websocket) => {
    const session = runtime.sessions.createSession();
    const rpcClient = new AddinRpcClient(websocket, { timeoutMs: 30_000 });
    runtime.attachAddinClient(session.connectionId, rpcClient);

    websocket.on("message", (raw) => {
      const parsed = parseJsonRpcMessage(raw.toString());
      if (!parsed) {
        sendError(websocket, null, -32700, "Invalid JSON-RPC message");
        return;
      }
      if (rpcClient.handleMessage(parsed)) {
        return;
      }
      handleAddinMessage(runtime, session.connectionId, parsed);
    });

    websocket.on("close", () => {
      rpcClient.rejectAll("Excel add-in disconnected");
      runtime.detachAddinClient(session.connectionId);
      runtime.sessions.remove(session.connectionId);
    });

    send(websocket, {
      jsonrpc: "2.0",
      method: "backend.connected",
      params: {
        connectionId: session.connectionId,
        connectedAt: session.connectedAt
      }
    });
  });

  return new Promise((resolve, reject) => {
    httpServer.once("error", reject);
    httpServer.listen(options.port, options.host, () => {
      httpServer.off("error", reject);
      handle = {
        close: async () => {
          await new Promise<void>((closeResolve) => websocketServer.close(() => closeResolve()));
          await new Promise<void>((closeResolve) => httpServer.close(() => closeResolve()));
        }
      };
      resolve(handle);
    });
  });
}

function handleRuntimeRpc(runtime: RuntimeService, request: IncomingMessage, response: import("node:http").ServerResponse): void {
  const chunks: Buffer[] = [];
  let bytes = 0;
  request.on("data", (chunk: Buffer) => {
    bytes += chunk.length;
    if (bytes > 10 * 1024 * 1024) {
      response.writeHead(413, { "content-type": "application/json" });
      response.end(JSON.stringify({ ok: false, error: { code: "PAYLOAD_TOO_LARGE", message: "RPC payload is too large." } }));
      request.destroy();
      return;
    }
    chunks.push(chunk);
  });
  request.on("end", async () => {
    try {
      const payload = JSON.parse(Buffer.concat(chunks).toString("utf8")) as { method?: string; args?: unknown[] };
      if (!payload.method || !Array.isArray(payload.args)) {
        response.writeHead(400, { "content-type": "application/json" });
        response.end(JSON.stringify({ ok: false, error: { code: "RANGE_INVALID", message: "RPC payload must include method and args array." } }));
        return;
      }
      const target = (runtime as unknown as Record<string, unknown>)[payload.method];
      if (typeof target !== "function") {
        response.writeHead(404, { "content-type": "application/json" });
        response.end(JSON.stringify({ ok: false, error: { code: "NOT_FOUND", message: `Runtime method not found: ${payload.method}` } }));
        return;
      }
      const result = await target.apply(runtime, payload.args);
      response.writeHead(200, { "content-type": "application/json" });
      response.end(JSON.stringify({ ok: true, result }));
    } catch (error) {
      response.writeHead(500, { "content-type": "application/json" });
      response.end(
        JSON.stringify({
          ok: false,
          error: {
            code: "OPERATION_FAILED",
            message: error instanceof Error ? error.message : String(error)
          }
        })
      );
    }
  });
}

function handleAddinMessage(runtime: RuntimeService, connectionId: ConnectionId, message: JsonRpcMessage): void {
  if (!("method" in message)) {
    return;
  }
  runtime.recordAddinEvent(connectionId, message.method, message.params);

  switch (message.method) {
    case "addin.hello":
      runtime.sessions.update(connectionId, createSessionPatch(message, ["capabilities", "activeWorkbook"]));
      break;
    case "addin.heartbeat":
      runtime.sessions.touch(connectionId);
      break;
    case "workbook.contextChanged":
      runtime.sessions.update(connectionId, createSessionPatch(message, ["activeWorkbook"]));
      break;
    default:
      runtime.sessions.touch(connectionId);
  }
}

function createSessionPatch(
  message: JsonRpcNotification,
  keys: Array<"capabilities" | "activeWorkbook">
): Partial<Omit<AddinSession, "connectionId" | "connectedAt">> {
  const patch: Partial<Omit<AddinSession, "connectionId" | "connectedAt">> = {};
  if (keys.includes("capabilities")) {
    const capabilities = readParam<AddinSession["capabilities"]>(message, "capabilities");
    if (capabilities !== undefined) {
      patch.capabilities = capabilities;
    }
  }
  if (keys.includes("activeWorkbook")) {
    const activeWorkbook = readParam<AddinSession["activeWorkbook"]>(message, "activeWorkbook");
    if (activeWorkbook !== undefined) {
      patch.activeWorkbook = activeWorkbook;
    }
  }
  return patch;
}

function isAddinUpgrade(request: IncomingMessage, addinPath: string): boolean {
  return request.url === addinPath;
}

function parseJsonRpcMessage(raw: string): JsonRpcMessage | undefined {
  try {
    const message = JSON.parse(raw) as JsonRpcMessage;
    if (message && typeof message === "object" && "jsonrpc" in message && message.jsonrpc === "2.0") {
      return message;
    }
    return undefined;
  } catch {
    return undefined;
  }
}

function readParam<T>(message: JsonRpcNotification, key: string): T | undefined {
  const params = message.params;
  if (!params || typeof params !== "object" || !(key in params)) {
    return undefined;
  }
  return (params as Record<string, T>)[key];
}

function send(websocket: WebSocket, message: JsonRpcMessage): void {
  if (websocket.readyState === websocket.OPEN) {
    websocket.send(JSON.stringify(message));
  }
}

function sendError(websocket: WebSocket, id: string | number | null, code: number, message: string): void {
  send(websocket, {
    jsonrpc: "2.0",
    id,
    error: {
      code,
      message
    }
  });
}
