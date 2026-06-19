#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { readConfig } from "./config.js";
import { createRuntimeFacade } from "./runtime-facade.js";
import { registerPrompts } from "./prompts.js";
import { registerResources } from "./resources.js";
import { registerAgentTools } from "./tools/agent-run.js";

const config = readConfig();
const { runtime, agentExecutionContext } = await createRuntimeFacade(config);

const server = new McpServer({
  name: "open-workbook",
  version: config.runtimeVersion
});

registerAgentTools(server, runtime, agentExecutionContext);
registerResources(server, runtime);
registerPrompts(server);

await server.connect(new StdioServerTransport());
