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
const orchestratedCapabilities = parseOrchestratedCapabilities(capabilitiesSource, handlerCapabilities);
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
  "events",
  "pivot",
  "chart",
  "permissions"
]);
const unitContractCapabilities = new Set([
  "excel.workbook.save_as",
  "excel.workbook.export_copy"
]);
const contractTestedCapabilities = new Set([
  "excel.runtime.connect_addin",
  "excel.runtime.disconnect_addin",
  "excel.runtime.ping_addin",
  "excel.runtime.get_capabilities",
  "excel.runtime.set_active_workbook",
  "excel.runtime.set_active_sheet",
  "excel.workbook.save_as",
  "excel.workbook.export_copy",
  "excel.batch.apply",
  "excel.batch.submit",
  "excel.batch.submit_chunked",
  "excel.batch.preflight",
  "excel.batch.validate",
  "excel.batch.dry_run",
  "excel.plan.create",
  "excel.plan.preview",
  "excel.plan.refresh_preview",
  "excel.plan.rebase",
  "excel.plan.apply",
  "excel.plan.rollback",
  "excel.task.create",
  "excel.task.claim",
  "excel.task.update",
  "excel.task.set_progress",
  "excel.task.add_blocker",
  "excel.task.resolve_blocker",
  "excel.task.evaluate_schedule",
  "excel.task.resume_ready",
  "excel.task.complete",
  "excel.task.fail",
  "excel.task.cancel",
  "excel.task.list",
  "excel.task.get",
  "excel.collab.get_status",
  "excel.collab.list_agents",
  "excel.collab.list_tasks",
  "excel.collab.list_locks",
  "excel.collab.list_transactions",
  "excel.collab.get_conflicts",
  "excel.collab.get_recent_events",
  "excel.lock.get_policy",
  "excel.lock.set_policy",
  "excel.lock.acquire",
  "excel.lock.renew",
  "excel.lock.release",
  "excel.conflict.get_guidance",
  "excel.conflict.explain",
  "excel.conflict.get_telemetry",
  "excel.conflict.clear_telemetry",
  "excel.transaction.get",
  "excel.transaction.list",
  "excel.transaction.wait",
  "excel.transaction.cancel",
  "excel.transaction.preview_rollback",
  "excel.transaction.rollback",
  "excel.transaction.preview_rollback_chain",
  "excel.transaction.rollback_chain",
  "excel.job.list",
  "excel.job.get",
  "excel.job.wait",
  "excel.job.cancel",
  "excel.events.subscribe",
  "excel.events.unsubscribe",
  "excel.events.get_recent",
  "excel.events.clear",
  "excel.events.set_debounce",
  "excel.pivot.list",
  "excel.pivot.get_info",
  "excel.pivot.create",
  "excel.pivot.refresh",
  "excel.pivot.refresh_all",
  "excel.pivot.update_source",
  "excel.pivot.copy_from_template",
  "excel.pivot.delete",
  "excel.pivot.validate_source",
  "excel.pivot.get_capability_matrix",
  "excel.pivot.get_fingerprint",
  "excel.pivot.compare_fingerprint",
  "excel.pivot.diff",
  "excel.pivot.repair_from_template",
  "excel.pivot.rebuild_with_source",
  "excel.chart.list",
  "excel.chart.get_info",
  "excel.chart.create",
  "excel.chart.update_data_source",
  "excel.chart.copy_from_template",
  "excel.chart.refresh",
  "excel.chart.delete",
  "excel.chart.validate_against_template",
  "excel.permissions.get",
  "excel.permissions.set",
  "excel.permissions.require_confirmation",
  "excel.permissions.set_scope",
  "excel.permissions.allow_destructive_actions",
  "excel.permissions.allow_macro_execution",
  "excel.permissions.lock_regions",
  "excel.permissions.unlock_regions"
]);
const statuses = ["covered", "needs_unit_contract", "future_orchestration_candidate", "host_limited", "defer"];

const entries = capabilityNames.map((name) => {
  const group = resolveGroup(name, groups);
  if (!group) {
    return { name, group: "unclassified", agentStatus: "internal_capability", planningStatus: "defer" };
  }
  const agentStatus = name === "excel.agent.run" ? "agent_entrypoint" : orchestratedCapabilities.has(name) ? "agent_action_handler" : "internal_capability";
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

function parseOrchestratedCapabilities(source, handlerCapabilities) {
  const match = source.match(/const AGENT_ORCHESTRATED_CAPABILITIES = new Set\(\[([\s\S]*?)\]\);/);
  if (!match) {
    return handlerCapabilities;
  }
  return new Set([
    ...handlerCapabilities,
    ...[...match[1].matchAll(/"([^"]+)"/g)].map((item) => item[1])
  ]);
}

function resolveGroup(name, groups) {
  const matches = groups.filter((group) => group.prefixes.some((prefix) => name.startsWith(prefix)));
  return matches.length === 1 ? matches[0] : undefined;
}

function planningStatus(name, group, agentStatus) {
  if (agentStatus !== "internal_capability") return "covered";
  if (contractTestedCapabilities.has(name)) return "covered";
  if (unitContractGroups.has(group) || unitContractCapabilities.has(name)) return "needs_unit_contract";
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
