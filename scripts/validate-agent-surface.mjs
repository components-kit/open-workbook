#!/usr/bin/env node
import { readFileSync } from "node:fs";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");

const checks = [
  {
    ok: source.includes('const mcpSurface = process.env.OPEN_WORKBOOK_MCP_SURFACE ?? "agent"'),
    message: "OPEN_WORKBOOK_MCP_SURFACE must default to agent."
  },
  {
    ok: source.includes('return name === "excel.agent.run"'),
    message: "Agent surface must expose only excel.agent.run by default."
  },
  {
    ok: source.includes('mcpSurface === "agent"') && source.includes('return ["excel.agent.run"]'),
    message: "Runtime capabilities must report only excel.agent.run in agent mode."
  },
  {
    ok: source.includes("registerAgentTools(server)") && source.includes('"excel.agent.run"'),
    message: "excel.agent.run must be registered with the MCP server."
  }
];

const failed = checks.filter((check) => !check.ok);
if (failed.length > 0) {
  console.error("Agent MCP surface validation failed.");
  for (const check of failed) {
    console.error(`- ${check.message}`);
  }
  process.exit(1);
}

console.log("Agent MCP surface check passed: default surface exposes excel.agent.run only.");
