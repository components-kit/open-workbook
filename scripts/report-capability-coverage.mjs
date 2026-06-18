#!/usr/bin/env node
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const protocolToolsPath = path.join(repoRoot, "packages/protocol/src/tools.ts");
const backendCapabilitiesPath = path.join(repoRoot, "apps/backend/src/excel-capabilities.ts");
const agentHandlersPath = path.join(repoRoot, "apps/backend/src/agent-action-handlers.ts");

const args = process.argv.slice(2);
const json = args.includes("--json");
const outIndex = args.indexOf("--out");
const outPath = outIndex >= 0 ? args[outIndex + 1] : undefined;

const toolsSource = readFileSync(protocolToolsPath, "utf8");
const capabilitiesSource = readFileSync(backendCapabilitiesPath, "utf8");
const handlersSource = readFileSync(agentHandlersPath, "utf8");

const capabilityNames = parseToolNames(toolsSource);
const groups = parseGroups(capabilitiesSource);
const handlerCapabilities = new Set([...handlersSource.matchAll(/capabilityName:\s*"([^"]+)"/g)].map((match) => match[1]));
const hostLimitedCapabilities = new Set([
  "excel.workbook.save_as",
  "excel.workbook.export_copy",
  "excel.formula.find_circular_references",
  "excel.style.copy_freeze_panes",
  "excel.style.copy_print_settings",
  "excel.style.copy_page_layout",
  "excel.style.copy_hidden_rows_columns"
]);
const unitContractGroups = new Set([
  "runtime",
  "lookup",
  "batch",
  "plan",
  "job",
  "task",
  "collaboration",
  "lock",
  "conflict",
  "transaction",
  "diff",
  "events",
  "compact_resource",
  "permissions"
]);
const statuses = ["covered", "needs_unit_contract", "future_orchestration_candidate", "host_limited", "defer"];

const entries = capabilityNames.map((name) => {
  const group = resolveGroup(name, groups);
  if (!group) {
    return { name, group: "unclassified", agentStatus: "internal_capability", planningStatus: "defer" };
  }
  const agentStatus = name === "excel.agent.run" ? "agent_entrypoint" : handlerCapabilities.has(name) ? "agent_action_handler" : "internal_capability";
  return { name, group: group.group, agentStatus, planningStatus: planningStatus(name, group.group, agentStatus) };
});

const unclassified = entries.filter((entry) => entry.group === "unclassified");
if (unclassified.length > 0) {
  console.error(`Capability coverage report failed: ${unclassified.length} unclassified capability(s).`);
  for (const entry of unclassified) {
    console.error(`- ${entry.name}`);
  }
  process.exit(1);
}

const report = {
  total: entries.length,
  byPlanningStatus: countBy(entries, "planningStatus"),
  byGroup: groups.map((group) => {
    const groupEntries = entries.filter((entry) => entry.group === group.group);
    return {
      group: group.group,
      label: group.label,
      total: groupEntries.length,
      byPlanningStatus: countBy(groupEntries, "planningStatus")
    };
  }),
  entries
};

const output = json ? `${JSON.stringify(report, null, 2)}\n` : renderMarkdown(report);
if (outPath) {
  writeFileSync(path.resolve(repoRoot, outPath), output);
} else {
  process.stdout.write(output);
}

function parseToolNames(source) {
  const match = source.match(/const TOOL_NAMES = \[([\s\S]*?)\] as const;/);
  if (!match) {
    throw new Error("Could not find TOOL_NAMES in protocol catalog.");
  }
  return [...match[1].matchAll(/"([^"]+)"/g)].map((item) => item[1]).filter((name) => name.startsWith("excel."));
}

function parseGroups(source) {
  return [...source.matchAll(/\{ group: "([^"]+)", label: "([^"]+)", description: "[^"]+", prefixes: \[([^\]]+)\] \}/g)].map((match) => ({
    group: match[1],
    label: match[2],
    prefixes: [...match[3].matchAll(/"([^"]+)"/g)].map((prefix) => prefix[1])
  }));
}

function resolveGroup(name, groups) {
  const matches = groups.filter((group) => group.prefixes.some((prefix) => name.startsWith(prefix)));
  return matches.length === 1 ? matches[0] : undefined;
}

function planningStatus(name, group, agentStatus) {
  if (agentStatus !== "internal_capability") return "covered";
  if (group === "pivot" || group === "chart" || hostLimitedCapabilities.has(name)) return "host_limited";
  if (unitContractGroups.has(group)) return "needs_unit_contract";
  return "future_orchestration_candidate";
}

function countBy(entries, key) {
  return Object.fromEntries(statuses.map((status) => [status, entries.filter((entry) => entry[key] === status).length]));
}

function renderMarkdown(data) {
  const lines = [];
  lines.push("# Capability Coverage Matrix");
  lines.push("");
  lines.push(`Total internal capabilities: ${data.total}`);
  lines.push("");
  lines.push("## Planning Status");
  lines.push("");
  lines.push("| Status | Count |");
  lines.push("| --- | ---: |");
  for (const status of statuses) {
    lines.push(`| ${status} | ${data.byPlanningStatus[status]} |`);
  }
  lines.push("");
  lines.push("## Groups");
  lines.push("");
  lines.push("| Group | Total | Covered | Unit Contract | Future Orchestration | Host Limited | Defer |");
  lines.push("| --- | ---: | ---: | ---: | ---: | ---: | ---: |");
  for (const group of data.byGroup) {
    lines.push(`| ${group.label} | ${group.total} | ${group.byPlanningStatus.covered} | ${group.byPlanningStatus.needs_unit_contract} | ${group.byPlanningStatus.future_orchestration_candidate} | ${group.byPlanningStatus.host_limited} | ${group.byPlanningStatus.defer} |`);
  }
  return `${lines.join("\n")}\n`;
}
