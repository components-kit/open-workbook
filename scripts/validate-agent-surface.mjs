#!/usr/bin/env node
import { readFileSync } from "node:fs";

const entrySource = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");
const agentToolSource = readFileSync(new URL("../apps/mcp-server/src/tools/agent-run.ts", import.meta.url), "utf8");
const combinedSource = `${entrySource}\n${agentToolSource}`;
const deprecatedInternalSurfaceEnv = "OPEN_WORKBOOK_" + "INTERNAL_TOOL_SURFACE";
const deprecatedInternalSurfaceFlag = "expose" + "InternalToolSurface";
const registeredToolNames = [...agentToolSource.matchAll(/\bmcp\.registerTool\s+as\s+any\)\(\s*"([^"]+)"/g)].map((match) => match[1]);
const forbiddenPrimitiveRegistrations = ["Runtime", "Workbook", "Range"].map((name) => `register${name}Tools(server)`);

const checks = [
  {
    ok: !combinedSource.includes(deprecatedInternalSurfaceEnv) && !combinedSource.includes(deprecatedInternalSurfaceFlag),
    message: "MCP must not include an internal primitive tool exposure gate."
  },
  {
    ok: registeredToolNames.length === 1 && registeredToolNames[0] === "excel.agent.run",
    message: "Public agent surface must expose only excel.agent.run."
  },
  {
    ok: entrySource.includes("registerAgentTools(server") && agentToolSource.includes('"excel.agent.run"'),
    message: "excel.agent.run must be registered with the MCP server."
  },
  {
    ok: registeredToolNames.length === 1 && registeredToolNames[0] === "excel.agent.run",
    message: `MCP must register exactly one tool; found ${registeredToolNames.join(", ") || "none"}.`
  },
  {
    ok: forbiddenPrimitiveRegistrations.every((snippet) => !combinedSource.includes(snippet)),
    message: "Primitive MCP tool registration groups must not be registered."
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
