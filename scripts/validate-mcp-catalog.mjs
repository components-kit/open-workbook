#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getExposedToolCatalog, getInternalCapabilityCatalog } from "../packages/protocol/dist/tools.js";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");

const callableTools = getExposedToolCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const registeredTools = source.includes("registerAgentTools(server)") ? ["excel.agent.run"] : [];
const missing = callableTools.filter((name) => !registeredTools.includes(name));
const forbiddenActiveRegistrations = ["Runtime", "Workbook", "Range", "Batch", "Workflow", "Table", "Chart", "Pivot"]
  .map((name) => `register${name}Tools(server)`)
  .filter((snippet) => source.includes(snippet));
const internalCapabilities = getInternalCapabilityCatalog({ includePreview: true });

if (missing.length > 0 || forbiddenActiveRegistrations.length > 0 || internalCapabilities.length <= 300) {
  console.error("MCP tool catalog drift detected.");
  if (missing.length > 0) {
    console.error(`Missing public MCP registrations:\n${missing.map((name) => `- ${name}`).join("\n")}`);
  }
  if (forbiddenActiveRegistrations.length > 0) {
    console.error(`Primitive MCP registration groups must not be active:\n${forbiddenActiveRegistrations.map((name) => `- ${name}`).join("\n")}`);
  }
  if (internalCapabilities.length <= 300) {
    console.error(`Internal capability catalog is unexpectedly small: ${internalCapabilities.length}`);
  }
  process.exit(1);
}

console.log(`MCP tool catalog check passed: ${callableTools.length} public tool, ${internalCapabilities.length} internal backend capabilities.`);
