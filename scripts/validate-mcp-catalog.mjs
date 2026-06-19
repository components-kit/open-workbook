#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getInternalCapabilityCatalog, getPublicAgentToolCatalog } from "../packages/protocol/dist/tools.js";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");

const publicTools = getPublicAgentToolCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const registeredTools = source.includes("registerAgentTools(server)") ? ["excel.agent.run"] : [];
const missing = publicTools.filter((name) => !registeredTools.includes(name));
const forbiddenActiveRegistrations = ["Runtime", "Workbook", "Range", "Batch", "Workflow", "Table", "Chart", "Pivot"]
  .map((name) => `register${name}Tools(server)`)
  .filter((snippet) => source.includes(snippet));
const internalCapabilities = getInternalCapabilityCatalog({ includePreview: true });

const expectedInternalCapabilityCount = 294;

if (missing.length > 0 || forbiddenActiveRegistrations.length > 0 || internalCapabilities.length !== expectedInternalCapabilityCount) {
  console.error("Public MCP tool contract drift detected.");
  if (missing.length > 0) {
    console.error(`Missing public MCP registrations:\n${missing.map((name) => `- ${name}`).join("\n")}`);
  }
  if (forbiddenActiveRegistrations.length > 0) {
    console.error(`Primitive MCP registration groups must not be active:\n${forbiddenActiveRegistrations.map((name) => `- ${name}`).join("\n")}`);
  }
  if (internalCapabilities.length !== expectedInternalCapabilityCount) {
    console.error(`Internal capability catalog count changed: expected ${expectedInternalCapabilityCount}, got ${internalCapabilities.length}`);
  }
  process.exit(1);
}

console.log(`Public MCP tool contract check passed: ${publicTools.length} public tool, ${internalCapabilities.length} internal backend capabilities.`);
