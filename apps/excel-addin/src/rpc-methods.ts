import type { JsonRpcRequest } from "@components-kit/open-workbook-protocol";
import { getRuntimeCapabilities } from "./excel-executor.js";
import { getHostMethod } from "./host/registry.js";

export function getRuntimeHelloPayload() {
  return {
    host: "excel",
    runtime: "office-js",
    capabilities: getRuntimeCapabilities()
  };
}

export async function handleAddinRpcRequest(request: JsonRpcRequest): Promise<{ ok: true; result: unknown } | { ok: false; message: string }> {
  if (request.method === "runtime.disconnect") {
    return { ok: true, result: { ok: true } };
  }
  const entry = getHostMethod(request.method);
  if (!entry) {
    return { ok: false, message: `Method not implemented in add-in: ${request.method}` };
  }
  return { ok: true, result: await entry.handler(request.params) };
}
