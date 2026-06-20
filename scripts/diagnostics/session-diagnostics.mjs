#!/usr/bin/env node
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { findRepoRoot } from "../lib/repo-root.mjs";

const repoRoot = findRepoRoot(import.meta.url);
const args = process.argv.slice(2);
const json = args.includes("--json");
const inputPath = args.find((arg, index) => !arg.startsWith("--") && args[index - 1] !== "--out");
const outIndex = args.indexOf("--out");
const outPath = outIndex >= 0 ? args[outIndex + 1] : undefined;

if (!inputPath) {
  console.error("Usage: node scripts/diagnostics/session-diagnostics.mjs <log-file> [--json] [--out path]");
  process.exit(1);
}

const source = readFileSync(path.resolve(repoRoot, inputPath), "utf8");
const events = extractEvents(source);
const agentCalls = events.filter((event) => event.tool === "excel.agent.run");
const modeCounts = countBy(agentCalls.map((event) => event.mode ?? "unknown"));
const actionCounts = countBy(agentCalls.map((event) => event.intentAction ?? "unknown"));
const duplicateRuns = consecutiveDuplicates(agentCalls);
const previewCalls = agentCalls.filter((event) => event.mode === "preview_update");
const applyCalls = agentCalls.filter((event) => event.mode === "apply_update");
const operationStatusCalls = agentCalls.filter((event) => event.mode === "operation_status");
const cancelCalls = agentCalls.filter((event) => event.mode === "cancel_operation");
const styledWorkflowCalls = agentCalls.filter((event) => event.intentAction === "replace_range_with_styled_table" || event.raw.includes("replace_range_with_styled_table"));
const tableViewCalls = agentCalls.filter((event) => event.intentAction === "apply_table_view" || event.raw.includes("apply_table_view"));
const styleSignals = events.filter((event) => /style|format|autofit|header|border|fill|font/i.test(event.raw));
const valueSignals = events.filter((event) => /write_values|values|table|booking|image|ocr|extracted/i.test(event.raw));
const tableFilterSignals = events.filter((event) => /filter_range|apply_filters|\bfilters?\b/i.test(event.raw));
const tableSortSignals = events.filter((event) => /sort_table|\bsort(?:ing|ed)?\b/i.test(event.raw));
const recommendations = [];

if (previewCalls.length > 1 && styledWorkflowCalls.length === 0 && styleSignals.length > 0 && valueSignals.length > 0) {
  recommendations.push("Use one preview_update with intent.action=replace_range_with_styled_table instead of splitting value writes, style copies, autofit, and clears.");
}
if (previewCalls.length > 1 && tableViewCalls.length === 0 && tableFilterSignals.length > 0 && tableSortSignals.length > 0) {
  recommendations.push("Use one preview_update with intent.action=apply_table_view when one table request combines filters and sorting.");
}
if (duplicateRuns.length > 0) {
  recommendations.push("Collapse repeated identical excel.agent.run calls or use operation_status for the returned operationId after a long preview/apply.");
}
if (applyCalls.length > previewCalls.length) {
  recommendations.push("Apply only returned preview operationIds; avoid repeated apply_update unless the first apply returned IN_PROGRESS or retryable status.");
}
if (operationStatusCalls.length === 0 && /IN_PROGRESS|queued|timeout|long-running/i.test(source)) {
  recommendations.push("Poll with mode=operation_status for long-running operations instead of reissuing preview/apply calls.");
}
if (cancelCalls.length > 0 && applyCalls.length > 0) {
  recommendations.push("cancel_operation only works before apply starts; after apply, use rollback or backup restore guidance.");
}
if (recommendations.length === 0) {
  recommendations.push("No obvious batching or lifecycle issue detected from this log.");
}

const report = {
  inputPath,
  totals: {
    events: events.length,
    agentCalls: agentCalls.length,
    previewUpdateCalls: previewCalls.length,
    applyUpdateCalls: applyCalls.length,
    operationStatusCalls: operationStatusCalls.length,
    cancelOperationCalls: cancelCalls.length,
    duplicateRuns: duplicateRuns.length
  },
  modeCounts,
  actionCounts,
  callSequence: agentCalls.map((event, index) => ({
    index: index + 1,
    mode: event.mode,
    intentAction: event.intentAction,
    operationId: event.operationId,
    confirmationTokenPresent: event.confirmationTokenPresent
  })),
  duplicateRuns,
  recommendations
};

const output = json ? `${JSON.stringify(report, null, 2)}\n` : renderMarkdown(report);
if (outPath) {
  writeFileSync(path.resolve(repoRoot, outPath), output);
} else {
  process.stdout.write(output);
}

function extractEvents(sourceText) {
  const events = [];
  const lines = sourceText.split(/\r?\n/);
  for (const [index, line] of lines.entries()) {
    if (!line.trim()) continue;
    const parsed = parseJsonish(line);
    const raw = parsed ? JSON.stringify(parsed) : line;
    if (!raw.includes("excel.agent.run")) continue;
    const payload = parsed ?? parseInlinePayload(line);
    events.push({
      line: index + 1,
      tool: "excel.agent.run",
      raw,
      mode: findValue(payload, "mode") ?? captureText(raw, /mode["'=:\s]+([a-z_]+)/i),
      intentAction: findValue(payload, "action") ?? captureText(raw, /intent\.action["'=:\s]+([a-z_]+)/i) ?? captureText(raw, /replace_range_with_styled_table/),
      operationId: findValue(payload, "operationId") ?? captureText(raw, /operationId["'=:\s]+([A-Za-z0-9:_-]+)/),
      confirmationTokenPresent: Boolean(findValue(payload, "confirmationToken") ?? /confirmationToken/.test(raw))
    });
  }
  return events;
}

function parseJsonish(line) {
  const trimmed = line.trim();
  const candidates = [trimmed, trimmed.slice(trimmed.indexOf("{"))].filter((candidate) => candidate.startsWith("{"));
  for (const candidate of candidates) {
    try {
      return JSON.parse(candidate);
    } catch {
      // Keep trying less structured log formats.
    }
  }
  return undefined;
}

function parseInlinePayload(line) {
  return {
    mode: captureText(line, /mode=([a-z_]+)/i),
    action: captureText(line, /(?:intent\.)?action=([a-z_]+)/i),
    operationId: captureText(line, /operationId=([A-Za-z0-9:_-]+)/),
    confirmationToken: captureText(line, /confirmationToken=([A-Za-z0-9:_-]+)/)
  };
}

function findValue(value, key) {
  if (value === null || value === undefined) return undefined;
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findValue(item, key);
      if (found !== undefined) return found;
    }
    return undefined;
  }
  if (typeof value === "object") {
    if (Object.prototype.hasOwnProperty.call(value, key) && value[key] !== undefined) return String(value[key]);
    for (const item of Object.values(value)) {
      const found = findValue(item, key);
      if (found !== undefined) return found;
    }
  }
  return undefined;
}

function captureText(text, re) {
  return text.match(re)?.[1] ?? (re.source.includes("replace_range_with_styled_table") && re.test(text) ? "replace_range_with_styled_table" : undefined);
}

function countBy(values) {
  const counts = {};
  for (const value of values) {
    counts[value] = (counts[value] ?? 0) + 1;
  }
  return counts;
}

function consecutiveDuplicates(calls) {
  const duplicates = [];
  for (let index = 1; index < calls.length; index += 1) {
    const previous = signature(calls[index - 1]);
    const current = signature(calls[index]);
    if (previous === current) {
      duplicates.push({ firstCall: index, secondCall: index + 1, signature: current });
    }
  }
  return duplicates;
}

function signature(call) {
  return [call.mode ?? "", call.intentAction ?? "", call.operationId ?? "", call.raw.replace(/\s+/g, " ").slice(0, 240)].join("|");
}

function renderMarkdown(data) {
  const lines = [];
  lines.push("# Session Diagnostics");
  lines.push("");
  lines.push(`Input: \`${data.inputPath}\``);
  lines.push("");
  lines.push("## Totals");
  lines.push("");
  lines.push("| Metric | Count |");
  lines.push("| --- | ---: |");
  for (const [key, value] of Object.entries(data.totals)) {
    lines.push(`| ${key} | ${value} |`);
  }
  lines.push("");
  lines.push("## Mode Counts");
  lines.push("");
  lines.push("| Mode | Count |");
  lines.push("| --- | ---: |");
  for (const [mode, count] of Object.entries(data.modeCounts)) {
    lines.push(`| ${mode} | ${count} |`);
  }
  lines.push("");
  lines.push("## Recommendations");
  lines.push("");
  lines.push(...data.recommendations.map((item) => `- ${item}`));
  return `${lines.join("\n")}\n`;
}
