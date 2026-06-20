#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getInternalCapabilityCatalog, getPublicAgentToolCatalog } from "../../packages/protocol/dist/tools.js";
import { createValidation } from "../lib/validation.mjs";

const validation = createValidation("MCP surface validation");
const entrySource = readFileSync(new URL("../../apps/mcp-server/src/index.ts", import.meta.url), "utf8");
const agentToolSource = readFileSync(new URL("../../apps/mcp-server/src/tools/agent-run.ts", import.meta.url), "utf8");
const combinedSource = `${entrySource}\n${agentToolSource}`;

const publicTools = getPublicAgentToolCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const internalCapabilities = getInternalCapabilityCatalog({ includePreview: true });
const internalCapabilityNames = new Set(internalCapabilities.map((tool) => tool.name));
const exposedTools = new Set(publicTools);
const registeredToolNames = [...agentToolSource.matchAll(/\bmcp\.registerTool\s+as\s+any\)\(\s*"([^"]+)"/g)].map((match) => match[1]);

const deprecatedInternalSurfaceEnv = "OPEN_WORKBOOK_" + "INTERNAL_TOOL_SURFACE";
const deprecatedInternalSurfaceFlag = "expose" + "InternalToolSurface";
const forbiddenPrimitiveRegistrations = ["Runtime", "Workbook", "Range", "Batch", "Workflow", "Table", "Chart", "Pivot"]
  .map((name) => `register${name}Tools(server)`);
const expectedInternalCapabilityCount = 303;

const requiredInternalCapabilities = [
  "excel.workflow.prepare_session",
  "excel.workflow.inspect_analyze",
  "excel.workflow.preview_risky_edit",
  "excel.range.read_compact",
  "excel.table.read_compact",
  "excel.validate.compact",
  "excel.snapshot.get_compact",
  "excel.snapshot.compare_compact",
  "excel.batch.apply",
  "excel.range.write_values",
  "excel.range.write_formulas",
  "excel.range.write_styles_many",
  "excel.table.apply_filters",
  "excel.table.sort"
];

const forbiddenInternalCapabilities = [
  "excel.range.read_values",
  "excel.range.read_formulas",
  "excel.range.read_number_formats",
  "excel.range.read_display_text",
  "excel.range.read_styles",
  "excel.range.read_full",
  "excel.table.read",
  "excel.snapshot.get",
  "excel.snapshot.compare",
  "excel.diff.get_details",
  "excel.diff.export_json",
  "excel.diff.export_html",
  "excel.filter.get_filters",
  "excel.filter.apply",
  "excel.filter.clear",
  "excel.filter.preserve_from_template",
  "excel.filter.validate",
  "excel.sort.apply",
  "excel.sort.clear",
  "excel.sort.preserve_from_template"
];

validation.expect(!combinedSource.includes(deprecatedInternalSurfaceEnv) && !combinedSource.includes(deprecatedInternalSurfaceFlag), "MCP must not include an internal primitive tool exposure gate.");
validation.expect(registeredToolNames.length === 1 && registeredToolNames[0] === "excel.agent.run", `MCP must register exactly one tool; found ${registeredToolNames.join(", ") || "none"}.`);
validation.expect(entrySource.includes("registerAgentTools(server") && agentToolSource.includes('"excel.agent.run"'), "excel.agent.run must be registered with the MCP server.");

for (const toolName of publicTools) {
  validation.expect(registeredToolNames.includes(toolName), `Missing public MCP registration: ${toolName}`);
}
for (const toolName of [...exposedTools].filter((name) => name !== "excel.agent.run")) {
  validation.fail(`Public MCP tools leaked beyond excel.agent.run: ${toolName}`);
}
for (const registration of forbiddenPrimitiveRegistrations) {
  validation.expect(!combinedSource.includes(registration), `Primitive MCP registration group must not be active: ${registration}`);
}

validation.expect(internalCapabilities.length === expectedInternalCapabilityCount, `Internal capability catalog count changed: expected ${expectedInternalCapabilityCount}, got ${internalCapabilities.length}`);
for (const capability of requiredInternalCapabilities) {
  validation.expect(internalCapabilityNames.has(capability), `Missing required internal capability: ${capability}`);
}
for (const capability of forbiddenInternalCapabilities) {
  validation.expect(!internalCapabilityNames.has(capability), `Raw/full or duplicate capability leaked into optimized internal catalog: ${capability}`);
}

validation.finish(`MCP surface check passed: ${publicTools.length} public tool, ${internalCapabilities.length} internal backend capabilities.`);
