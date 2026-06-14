#!/usr/bin/env node
import { mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const platform = readArg("--platform") ?? "unknown";
const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
const port = process.env.OPEN_WORKBOOK_PORT ?? "37845";
const statusUrl = readArg("--status-url") ?? process.env.OPEN_WORKBOOK_LIVE_E2E_STATUS_URL ?? `http://${host}:${port}/status`;
const artifactDir = readArg("--artifact-dir") ?? process.env.OPEN_WORKBOOK_LIVE_E2E_ARTIFACT_DIR ?? path.join(tmpdir(), `open-workbook-live-${platform}`);
const optIn = process.env.OPEN_WORKBOOK_LIVE_E2E === "1" || hasArg("--run");
const dryRun = hasArg("--dry-run");
const allowDisconnected = hasArg("--allow-disconnected") || process.env.OPEN_WORKBOOK_LIVE_E2E_ALLOW_DISCONNECTED === "1";
const deep = hasArg("--deep") || process.env.OPEN_WORKBOOK_LIVE_E2E_DEEP === "1";
const scratchSheet = readArg("--scratch-sheet") ?? process.env.OPEN_WORKBOOK_LIVE_E2E_SHEET ?? `OWB_${Date.now().toString(36)}`;

mkdirSync(artifactDir, { recursive: true });

if (dryRun || !optIn) {
  const report = {
    ok: dryRun,
    platform,
    statusUrl,
    artifactDir,
    mode: dryRun ? "dry-run" : "not-opted-in",
    requiredBeforeRun: [
      "Open desktop Excel.",
      "Load the Open Workbook add-in.",
      "Start the local Open Workbook MCP/backend runtime.",
      "Confirm the add-in is connected to the backend.",
      "Set OPEN_WORKBOOK_LIVE_E2E=1 or pass --run."
    ]
  };
  writeReports(report);
  const message = [
    `Live Excel E2E gate (${platform}) is host-driven.`,
    `Backend status URL: ${statusUrl}`,
    `Artifacts: ${artifactDir}`,
    "Set OPEN_WORKBOOK_LIVE_E2E=1 or pass --run after Excel and the add-in are connected.",
    "Use --dry-run to print this contract without failing."
  ].join("\n");
  if (dryRun) {
    console.log(message);
    process.exit(0);
  }
  console.error(message);
  process.exit(1);
}

const started = performance.now();
let report;
try {
  const status = await fetchJson(statusUrl);
  const checks = [
    check(status?.ok === true, "backend status ok"),
    check(Boolean(status?.runtime?.service), "backend runtime metadata present", status?.runtime?.service),
    check(Boolean(status?.activeAddinConnected) || allowDisconnected, "Excel add-in connected", allowDisconnected ? "allow-disconnected enabled" : undefined),
    check(Boolean(status?.activeWorkbook) || allowDisconnected, "active workbook available", allowDisconnected ? "allow-disconnected enabled" : undefined)
  ];
  const deepChecks = [];
  if (deep && status?.activeWorkbook?.workbookId) {
    deepChecks.push(...await runDeepSmoke(status.activeWorkbook.workbookId));
  }
  const allChecks = [...checks, ...deepChecks];
  report = {
    ok: allChecks.every((item) => item.ok),
    platform,
    statusUrl,
    artifactDir,
    elapsedMs: Math.round(performance.now() - started),
    checks: allChecks,
    deep,
    scratchSheet,
    status
  };
} catch (error) {
  report = {
    ok: false,
    platform,
    statusUrl,
    artifactDir,
    elapsedMs: Math.round(performance.now() - started),
    checks: [check(false, "backend status reachable", error instanceof Error ? error.message : String(error))]
  };
}

writeReports(report);
console.log(renderMarkdown(report));
console.log(`\nSaved live E2E artifacts:\n- ${path.join(artifactDir, "live-smoke.md")}\n- ${path.join(artifactDir, "live-smoke.json")}`);
if (!report.ok) {
  process.exit(1);
}

async function fetchJson(url) {
  const response = await fetch(url);
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${text}`);
  }
  return JSON.parse(text);
}

function check(ok, label, details) {
  return { ok, label, ...(details !== undefined ? { details } : {}) };
}

function writeReports(report) {
  writeFileSync(path.join(artifactDir, "live-smoke.json"), JSON.stringify(report, null, 2));
  writeFileSync(path.join(artifactDir, "live-smoke.md"), renderMarkdown(report));
}

function renderMarkdown(report) {
  const lines = [
    "# Open Workbook Live Excel Smoke",
    "",
    `- Platform: ${report.platform}`,
    `- Status: ${report.ok ? "PASS" : "FAIL"}`,
    `- Mode: ${report.mode ?? "run"}`,
    `- Deep: ${report.deep ? "yes" : "no"}`,
    `- Scratch sheet: ${report.scratchSheet ?? scratchSheet}`,
    `- Status URL: ${report.statusUrl}`,
    `- Elapsed: ${report.elapsedMs ?? 0} ms`,
    `- Artifacts: ${report.artifactDir}`,
    "",
    "## Checks"
  ];
  if (report.checks?.length) {
    for (const item of report.checks) {
      lines.push(`- ${item.ok ? "PASS" : "FAIL"} ${item.label}${item.details ? `: ${item.details}` : ""}`);
    }
  } else {
    lines.push("- Not run.");
  }
  if (report.requiredBeforeRun?.length) {
    lines.push("");
    lines.push("## Required Before Run");
    for (const item of report.requiredBeforeRun) {
      lines.push(`- ${item}`);
    }
  }
  return lines.join("\n");
}

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function hasArg(name) {
  return process.argv.includes(name);
}

async function runDeepSmoke(workbookId) {
  const checks = [];
  const target = { workbookId, sheetName: scratchSheet, address: "A1:C4" };
  const formulaTarget = { workbookId, sheetName: scratchSheet, address: "C2:C4" };
  const createSheet = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_smoke_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName: scratchSheet,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live smoke scratch sheet"
      }
    ]
  }]);
  checks.push(check(createSheet?.ok === true, "deep scratch sheet create", createSheet?.error?.message));
  if (createSheet?.ok !== true) {
    return checks;
  }

  const beforeSnapshot = await rpc("createWorkbookSnapshot", [{ workbookId, reason: "Live smoke before", ranges: [target] }]);
  checks.push(check(beforeSnapshot?.ok === true, "deep before snapshot", beforeSnapshot?.error?.message));

  const apply = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_smoke_values",
        kind: "range.write_values",
        workbookId,
        target,
        values: [
          ["Input", "Tax", "Total"],
          [100, 8, null],
          [125, 10, null],
          [150, 12, null]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live smoke values"
      },
      {
        operationId: "live_smoke_formulas",
        kind: "range.write_formulas",
        workbookId,
        target: formulaTarget,
        formulas: [["=A2+B2"], ["=A3+B3"], ["=A4+B4"]],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live smoke formulas"
      }
    ]
  }]);
  checks.push(check(apply?.ok === true, "deep scoped batch apply", apply?.error?.message));

  const read = await rpc("readRangeMetadata", ["range.read_full", { ...target, includeFormulas: true }]);
  checks.push(check(Boolean(read), "deep scoped range read"));

  const formulaValidation = await rpc("validateFormulas", [{ workbookId, sheetName: scratchSheet, address: formulaTarget.address }]);
  checks.push(check(formulaValidation?.ok !== false, "deep formula validation", formulaValidation?.error?.message));

  const afterSnapshot = await rpc("createWorkbookSnapshot", [{ workbookId, reason: "Live smoke after", ranges: [target] }]);
  checks.push(check(afterSnapshot?.ok === true, "deep after snapshot"));

  const leftSnapshotId = beforeSnapshot?.snapshot?.snapshotId;
  const rightSnapshotId = afterSnapshot?.snapshot?.snapshotId;
  if (leftSnapshotId && rightSnapshotId) {
    const diff = await rpc("compareSnapshots", [leftSnapshotId, rightSnapshotId]);
    checks.push(check(diff?.ok === true, "deep snapshot diff", diff?.error?.message));
  } else {
    checks.push(check(false, "deep snapshot diff", "missing snapshot IDs"));
  }

  if (apply?.transactionId) {
    const rollbackPreview = await rpc("previewTransactionRollback", [apply.transactionId]);
    checks.push(check(Boolean(rollbackPreview), "deep rollback preview"));
  } else {
    checks.push(check(false, "deep rollback preview", "missing transaction ID"));
  }
  return checks;
}

async function rpc(method, args) {
  const response = await fetch(statusUrl.replace(/\/status$/, "/rpc"), {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ method, args })
  });
  const payload = await response.json();
  if (!response.ok || !payload.ok) {
    return { ok: false, error: payload.error ?? { message: `RPC ${method} failed with HTTP ${response.status}` } };
  }
  return payload.result;
}
