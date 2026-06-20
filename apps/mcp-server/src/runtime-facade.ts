import { RuntimeService } from "@components-kit/open-workbook-backend/runtime";
import { startBackendServer } from "@components-kit/open-workbook-backend/server";
import type { AgentRunExecutionContext, BatchRequest } from "@components-kit/open-workbook-protocol";
import type { McpServerConfig } from "./config.js";

export type RuntimeFacade = RuntimeService & {
  compileBatch(request: BatchRequest): unknown;
  getAgentSemanticIndexResource(workbookContextId: string): unknown;
  getCompactResource(resourceId: string, options?: { view?: "summary" | "full"; maxBytes?: number }): unknown;
};

export interface RuntimeFacadeHandle {
  runtime: RuntimeFacade;
  agentExecutionContext?: AgentRunExecutionContext | undefined;
}

export async function createRuntimeFacade(config: McpServerConfig): Promise<RuntimeFacadeHandle> {
  if (!config.standalone && await daemonAvailable(config.daemonUrl)) {
    const proxy = createDaemonRuntimeProxy(config.daemonUrl) as RuntimeFacade;
    const registration = await proxy.registerAgent({ agentName: config.agentName, clientType: "mcp", pid: process.pid });
    const agentExecutionContext = agentContextFromRegistration(registration, config.agentName);
    console.error(`open-workbook MCP adapter connected to ${config.daemonUrl}${agentExecutionContext?.agentId ? ` as ${agentExecutionContext.agentId}` : ""}`);
    return { runtime: proxy, agentExecutionContext };
  }

  const localRuntime = new RuntimeService() as RuntimeFacade;
  const registration = localRuntime.registerAgent({ agentName: config.agentName, clientType: "mcp", pid: process.pid });
  const agentExecutionContext = agentContextFromRegistration(registration, config.agentName);
  await startBackendServer(localRuntime, { host: config.host, port: config.port, addinPath: config.addinPath });
  console.error(`open-workbook MCP standalone backend listening on ws://${config.host}:${config.port}${config.addinPath}`);
  return { runtime: localRuntime, agentExecutionContext };
}

async function daemonAvailable(baseUrl: string): Promise<boolean> {
  try {
    const response = await fetch(`${baseUrl}/status`);
    return response.ok;
  } catch {
    return false;
  }
}

function createDaemonRuntimeProxy(baseUrl: string): unknown {
  const call = async (method: string, args: unknown[]) => {
    const response = await fetch(`${baseUrl}/rpc`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ method, args })
    });
    const payload = await response.json() as { ok: boolean; result?: unknown; error?: unknown };
    if (!response.ok || !payload.ok) {
      throw new Error(JSON.stringify(payload.error ?? { code: "OPERATION_FAILED", message: `Daemon RPC failed: ${method}` }));
    }
    return payload.result;
  };
  return new Proxy(
    {},
    {
      get(_target, property) {
        if (typeof property !== "string" || property === "then") {
          return undefined;
        }
        return (...args: unknown[]) => call(property, args);
      }
    }
  );
}

function agentContextFromRegistration(registration: unknown, fallbackAgentName?: string): AgentRunExecutionContext | undefined {
  const registeredAgent = (registration as { agent?: { agentId?: string; agentName?: string } } | undefined)?.agent;
  if (!registeredAgent?.agentId) {
    return undefined;
  }
  const agentName = registeredAgent.agentName ?? fallbackAgentName;
  return {
    agentId: registeredAgent.agentId,
    ...(agentName !== undefined ? { agentName } : {}),
    clientType: "mcp"
  };
}
