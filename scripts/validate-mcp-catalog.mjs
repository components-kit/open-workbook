#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getExposedToolCatalog } from "../packages/protocol/dist/tools.js";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");

const callableTools = getExposedToolCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const sourceTools = [...source.matchAll(/"(excel\.[a-z0-9_.]+)"/g)]
  .map((match) => match[1])
  .filter((name) => !name.startsWith("excel.prompts."))
  .filter((name) =>
    /^excel\.(?:runtime|workbook|sheet|range|batch|plan|template|style|formula|table|filter|sort|pivot|chart|names|region|task|collab|lock|conflict|transaction|permissions|clean|validate|repair|snapshot|diff|events)\./.test(name)
  );
const registeredTools = [...new Set(sourceTools)].sort();

const missing = callableTools.filter((name) => !registeredTools.includes(name));
const extra = registeredTools.filter((name) => !callableTools.includes(name));

if (missing.length > 0 || extra.length > 0) {
  console.error("MCP tool catalog drift detected.");
  if (missing.length > 0) {
    console.error(`Missing MCP registrations:\n${missing.map((name) => `- ${name}`).join("\n")}`);
  }
  if (extra.length > 0) {
    console.error(`Unexpected MCP registrations:\n${extra.map((name) => `- ${name}`).join("\n")}`);
  }
  process.exit(1);
}

console.log(`MCP tool catalog check passed: ${callableTools.length} callable tools registered.`);
