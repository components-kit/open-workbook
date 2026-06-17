#!/usr/bin/env node
import { readFileSync } from "node:fs";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");

const checks = [
  {
    ok: source.includes('const exposeInternalToolSurface = process.env.OPEN_WORKBOOK_INTERNAL_TOOL_SURFACE === "1"'),
    message: "Internal primitive tool exposure must be gated by OPEN_WORKBOOK_INTERNAL_TOOL_SURFACE."
  },
  {
    ok: source.includes('name === "excel.agent.run"'),
    message: "Public agent surface must expose only excel.agent.run."
  },
  {
    ok: source.includes('exposeInternalToolSurface') && source.includes('["excel.agent.run"]'),
    message: "Runtime capabilities must report only excel.agent.run unless internal tooling is enabled."
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

console.log("Agent MCP surface check passed: public surface exposes excel.agent.run only.");
