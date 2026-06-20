#!/usr/bin/env node
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { findRepoRoot } from "../lib/repo-root.mjs";

const repoRoot = findRepoRoot(import.meta.url);
const paths = {
  protocolCatalog: "packages/protocol/src/catalog/index.ts",
  protocolAgent: "packages/protocol/src/agent.ts",
  protocolOperations: "packages/protocol/src/operations.ts",
  capabilities: "apps/backend/src/capabilities/registry.ts",
  domains: "apps/backend/src/capabilities/domains/metadata.ts",
  handlers: "apps/backend/src/agent-action-handlers.ts",
  policy: "apps/backend/src/agent-action-policy.ts",
  compiler: "packages/excel-core/src/range/batch-compiler.ts",
  executor: "apps/excel-addin/src/host/executor-core.ts",
  hostRegistry: "apps/excel-addin/src/host/registry.ts",
  prompts: "apps/mcp-server/src/prompts.ts"
};

const args = process.argv.slice(2);
const json = args.includes("--json");
const check = args.includes("--check");
const outIndex = args.indexOf("--out");
const outPath = outIndex >= 0 ? args[outIndex + 1] : undefined;

const sources = Object.fromEntries(Object.entries(paths).map(([key, relative]) => [key, read(relative)]));
const capabilityNames = parseToolNames(sources.protocolCatalog);
const groups = parseGroups(sources.domains);
const modes = parseConstArray(sources.protocolAgent, "AGENT_RUN_MODES");
const intentActions = parseConstArray(sources.protocolAgent, "AGENT_INTENT_ACTIONS");
const handlerEntries = parseAgentHandlers(sources.handlers);
const policyEntries = parseActionPolicy(sources.policy);
const operationKinds = operationKindsFromProtocol(sources.protocolOperations);
const compilerKinds = caseKinds(sources.compiler);
const executorKinds = caseKinds(sources.executor);
const hostMethodsByCapability = parseHostMethodsByCapability(sources.hostRegistry);
const backendMethodsByCapability = parseBackendMethodsByCapability(sources.capabilities);
const documentedWorkflows = parsePromptWorkflowHints(sources.prompts);

const handlersByCapability = groupBy(handlerEntries, "capabilityName");
const policiesByKind = new Map(policyEntries.map((entry) => [entry.kind, entry]));
const backendByCapability = new Map(backendMethodsByCapability.map((entry) => [entry.capability, entry.methods]));
const hostByCapability = new Map(hostMethodsByCapability.map((entry) => [entry.capability, entry.methods]));

const entries = capabilityNames.map((capability) => {
  const group = resolveGroup(capability, groups);
  const handlers = handlersByCapability.get(capability) ?? [];
  const relatedKind = relatedOperationKind(capability, operationKinds);
  const relatedPolicy = relatedKind ? policiesByKind.get(relatedKind) : undefined;
  const backendMethods = backendByCapability.get(capability) ?? [];
  const hostMethods = hostByCapability.get(capability) ?? [];
  return {
    capability,
    group: group?.group ?? "unclassified",
    publicMcpTool: capability === "excel.agent.run",
    agentModes: capability === "excel.agent.run" ? modes : [],
    agentHandlers: handlers.map((handler) => handler.id),
    intentActions: unique(handlers.map((handler) => handler.intentAction).filter(Boolean)),
    risk: highestRisk([...handlers.map((handler) => handler.riskKind), relatedPolicy?.risk].filter(Boolean)),
    previewRequired: handlers.some((handler) => handler.riskKind !== "read_only") || Boolean(relatedPolicy?.previewRequired),
    confirmationRequired: handlers.some((handler) => handler.riskKind !== "read_only") || Boolean(relatedPolicy?.confirmationRequired),
    backendMethods,
    hostMethods,
    protocolOperationKind: relatedKind,
    batchCompilerCovered: relatedKind ? compilerKinds.has(relatedKind) : undefined,
    addinExecutorCovered: relatedKind ? executorKinds.has(relatedKind) : undefined,
    docsAnchor: `docs/tool-surface.md#${capability.replace(/\./g, "")}`,
    workflowHints: documentedWorkflows.filter((workflow) => workflow.includes(capability) || capability.includes(workflow))
  };
});

const issues = [];
for (const entry of entries) {
  if (entry.group === "unclassified") {
    issues.push(`${entry.capability}: missing capability domain`);
  }
  if (entry.protocolOperationKind && !entry.batchCompilerCovered) {
    issues.push(`${entry.capability}: protocol operation ${entry.protocolOperationKind} is not covered by the batch compiler`);
  }
  if (entry.protocolOperationKind && !entry.addinExecutorCovered && !knownExecutorUnsupported(entry.protocolOperationKind)) {
    issues.push(`${entry.capability}: protocol operation ${entry.protocolOperationKind} is not covered by the add-in executor`);
  }
  if (entry.previewRequired && entry.backendMethods.length === 0 && entry.agentHandlers.length === 0 && !entry.publicMcpTool) {
    issues.push(`${entry.capability}: mutating capability has no backend method or agent handler mapping`);
  }
}

const manifest = {
  generatedAt: new Date().toISOString(),
  sources: paths,
  totals: {
    capabilities: entries.length,
    publicMcpTools: entries.filter((entry) => entry.publicMcpTool).length,
    agentHandled: entries.filter((entry) => entry.agentHandlers.length > 0).length,
    batchOperations: entries.filter((entry) => entry.protocolOperationKind).length,
    issues: issues.length
  },
  agentModes: modes,
  intentActions,
  entries,
  issues
};

if (check && issues.length > 0) {
  console.error("Operation manifest check failed:");
  console.error(issues.map((issue) => `- ${issue}`).join("\n"));
  process.exit(1);
}
if (check && !json && !outPath) {
  console.log(`Operation manifest check passed: ${entries.length} capability entries, ${manifest.totals.batchOperations} batch operation mappings, ${manifest.totals.agentHandled} agent-handled capabilities.`);
  process.exit(0);
}

const output = json ? `${JSON.stringify(manifest, null, 2)}\n` : renderMarkdown(manifest);
if (outPath) {
  writeFileSync(path.resolve(repoRoot, outPath), output);
} else {
  process.stdout.write(output);
}

function read(relativePath) {
  return readFileSync(path.join(repoRoot, relativePath), "utf8");
}

function parseToolNames(source) {
  const match = source.match(/const TOOL_NAMES = \[([\s\S]*?)\] as const;/);
  if (!match) throw new Error("Could not find TOOL_NAMES.");
  return [...match[1].matchAll(/"([^"]+)"/g)].map((item) => item[1]).filter((name) => name.startsWith("excel."));
}

function parseConstArray(source, name) {
  const match = source.match(new RegExp(`export const ${name} = \\[([\\s\\S]*?)\\] as const;`));
  if (!match) return [];
  return [...match[1].matchAll(/"([^"]+)"/g)].map((item) => item[1]);
}

function parseGroups(source) {
  return [...source.matchAll(/\{ group: "([^"]+)", label: "([^"]+)", description: "[^"]+", prefixes: \[([^\]]+)\][^}]*\}/g)].map((match) => ({
    group: match[1],
    label: match[2],
    prefixes: [...match[3].matchAll(/"([^"]+)"/g)].map((prefix) => prefix[1])
  }));
}

function parseAgentHandlers(source) {
  const entries = [];
  for (const block of source.matchAll(/\{\s*id:\s*"([^"]+)",([\s\S]*?)matches:/g)) {
    const body = block[2];
    entries.push({
      id: block[1],
      capabilityName: capture(body, /capabilityName:\s*"([^"]+)"/),
      intentAction: capture(body, /intentAction:\s*"([^"]+)"/),
      riskKind: capture(body, /riskKind:\s*"([^"]+)"/),
      requiresResolvedTarget: capture(body, /requiresResolvedTarget:\s*(true|false)/) === "true"
    });
  }
  return entries.filter((entry) => entry.capabilityName);
}

function parseActionPolicy(source) {
  return [...source.matchAll(/\{ kind: "([^"]+)", risk: "([^"]+)", previewRequired: (true|false), confirmationRequired: (true|false) \}/g)].map((match) => ({
    kind: match[1],
    risk: match[2],
    previewRequired: match[3] === "true",
    confirmationRequired: match[4] === "true"
  }));
}

function operationKindsFromProtocol(source) {
  const definitions = source.slice(0, source.indexOf("export type ExcelOperation"));
  return new Set([...definitions.matchAll(/kind:\s*"([^"]+)"/g)].map((match) => match[1]));
}

function caseKinds(source) {
  return new Set([...source.matchAll(/case\s+"([^"]+)"/g)].map((match) => match[1]));
}

function parseHostMethodsByCapability(source) {
  const entries = [];
  for (const match of source.matchAll(/\["([^"]+)",\s*[^,\]]+,\s*\[([^\]]*)\]\]/g)) {
    for (const capability of [...match[2].matchAll(/"([^"]+)"/g)].map((item) => item[1])) {
      entries.push({ capability, methods: [match[1]] });
    }
  }
  return mergeMethodEntries(entries);
}

function parseBackendMethodsByCapability(source) {
  const entries = [];
  for (const match of source.matchAll(/\["([^"]+)",\s*\[([^\]]*)\]\]/g)) {
    entries.push({ capability: match[1], methods: [...match[2].matchAll(/"([^"]+)"/g)].map((item) => item[1]) });
  }
  return mergeMethodEntries(entries);
}

function parsePromptWorkflowHints(source) {
  return unique([...source.matchAll(/replace_range_with_styled_table|operation_status|cancel_operation|preview_update|apply_update/g)].map((match) => match[0]));
}

function mergeMethodEntries(entries) {
  const byCapability = new Map();
  for (const entry of entries) {
    byCapability.set(entry.capability, unique([...(byCapability.get(entry.capability) ?? []), ...entry.methods]));
  }
  return [...byCapability.entries()].map(([capability, methods]) => ({ capability, methods }));
}

function relatedOperationKind(capability, operationKinds) {
  const suffix = capability.replace(/^excel\./, "");
  if (operationKinds.has(suffix)) return suffix;
  if (suffix === "template.create_sheet_from_template") return "template.create_sheet_from_template";
  return undefined;
}

function knownExecutorUnsupported(kind) {
  return kind === "range.write_hyperlinks" || kind === "range.write_comments" || kind === "sheet.move";
}

function resolveGroup(name, groups) {
  const matches = groups.filter((group) => group.prefixes.some((prefix) => name.startsWith(prefix)));
  return matches.length === 1 ? matches[0] : undefined;
}

function groupBy(entries, key) {
  const grouped = new Map();
  for (const entry of entries) {
    const value = entry[key];
    if (!grouped.has(value)) grouped.set(value, []);
    grouped.get(value).push(entry);
  }
  return grouped;
}

function capture(source, re) {
  return source.match(re)?.[1];
}

function unique(values) {
  return [...new Set(values)];
}

function highestRisk(risks) {
  const order = ["read_only", "safe_format", "small_value_write", "table_append", "formula_write", "broad_range_write", "structure_change", "destructive"];
  return risks.sort((a, b) => order.indexOf(b) - order.indexOf(a))[0] ?? "unknown";
}

function renderMarkdown(data) {
  const lines = [];
  lines.push("# Operation Manifest");
  lines.push("");
  lines.push(`Generated: ${data.generatedAt}`);
  lines.push("");
  lines.push("## Totals");
  lines.push("");
  lines.push("| Metric | Count |");
  lines.push("| --- | ---: |");
  for (const [key, value] of Object.entries(data.totals)) {
    lines.push(`| ${key} | ${value} |`);
  }
  lines.push("");
  lines.push("## Agent Surface");
  lines.push("");
  lines.push(`Modes: ${data.agentModes.map((mode) => `\`${mode}\``).join(", ")}`);
  lines.push("");
  lines.push("## Capabilities");
  lines.push("");
  lines.push("| Capability | Group | Public | Agent Handlers | Risk | Backend Methods | Host Methods | Operation |");
  lines.push("| --- | --- | --- | --- | --- | --- | --- | --- |");
  for (const entry of data.entries) {
    lines.push(`| \`${entry.capability}\` | ${entry.group} | ${entry.publicMcpTool ? "yes" : ""} | ${entry.agentHandlers.join(", ")} | ${entry.risk} | ${entry.backendMethods.join(", ")} | ${entry.hostMethods.join(", ")} | ${entry.protocolOperationKind ?? ""} |`);
  }
  if (data.issues.length > 0) {
    lines.push("");
    lines.push("## Issues");
    lines.push("");
    lines.push(...data.issues.map((issue) => `- ${issue}`));
  }
  return `${lines.join("\n")}\n`;
}
