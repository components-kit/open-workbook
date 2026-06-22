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
const requiredTaskpaneBundleVersion = readArg("--taskpane-bundle-version") ?? process.env.OPEN_WORKBOOK_LIVE_E2E_TASKPANE_BUNDLE_VERSION;
const scratchSheet = readArg("--scratch-sheet") ?? process.env.OPEN_WORKBOOK_LIVE_E2E_SHEET ?? `OWB_${Date.now().toString(36)}`;
const scenarioGroups = scenarioGroupsFromInput(readArg("--scenario-group") ?? process.env.OPEN_WORKBOOK_LIVE_E2E_SCENARIO_GROUPS, deep);

mkdirSync(artifactDir, { recursive: true });

if (dryRun || !optIn) {
  const report = {
    ok: dryRun,
    platform,
    statusUrl,
    artifactDir,
    mode: dryRun ? "dry-run" : "not-opted-in",
    scenarioGroups,
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
    check(Boolean(status?.activeWorkbook) || allowDisconnected, "active workbook available", allowDisconnected ? "allow-disconnected enabled" : undefined),
    ...taskpaneBundleChecks(status, requiredTaskpaneBundleVersion)
  ];
  const scenarioChecks = [];
  if (status?.activeWorkbook?.workbookId) {
    scenarioChecks.push(...await runScenarioGroupsWithScratchPermissions(status.activeWorkbook.workbookId, scenarioGroups));
  } else if (scenarioGroups.length > 0 && !allowDisconnected) {
    scenarioChecks.push(check(false, "live scenario groups", "missing active workbook"));
  }
  const allChecks = [...checks, ...scenarioChecks];
  report = {
    ok: allChecks.every((item) => item.ok),
    platform,
    statusUrl,
    artifactDir,
    elapsedMs: Math.round(performance.now() - started),
    checks: allChecks,
    deep,
    scenarioGroups,
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
    `- Scenario groups: ${(report.scenarioGroups ?? []).join(", ") || "connection"}`,
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

function taskpaneBundleChecks(status, expectedVersion) {
  if (!expectedVersion) {
    return [];
  }
  const versions = Array.isArray(status?.sessions)
    ? status.sessions
        .map((session) => session?.capabilities?.engine?.taskpaneBundleVersion)
        .filter((version) => typeof version === "string")
    : [];
  const activeVersion = status?.sessions?.find((session) => session?.activeWorkbook?.workbookId === status?.activeWorkbook?.workbookId)
    ?.capabilities?.engine?.taskpaneBundleVersion;
  const version = activeVersion ?? versions[0];
  return [
    check(
      version === expectedVersion,
      "Excel taskpane bundle version",
      version ? `expected ${expectedVersion}, connected ${version}` : `expected ${expectedVersion}, connected version unavailable`
    )
  ];
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

  const read = await readFullRange({ ...target, includeFormulas: true });
  checks.push(check(read?.ok === true, "deep scoped range read", read?.error?.message));

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

async function runScenarioGroups(workbookId, groups) {
  const checks = [];
  for (const group of groups) {
    if (group === "scratch-core") {
      checks.push(...await runDeepSmoke(workbookId));
    } else if (group === "regression-pack") {
      checks.push(...await runRegressionPack(workbookId));
    } else if (group === "template-formula-repair") {
      checks.push(...await runTemplateFormulaRepair(workbookId));
    } else if (group === "pivot-chart-core") {
      checks.push(...await runPivotChartCore(workbookId));
    } else if (group === "pivot-template-repair") {
      checks.push(...await runPivotTemplateRepair(workbookId));
    } else if (group === "chart-template-copy") {
      checks.push(...await runChartTemplateCopy(workbookId));
    } else {
      checks.push(check(false, `live scenario group ${group}`, "unknown scenario group"));
    }
  }
  return checks;
}

async function runScenarioGroupsWithScratchPermissions(workbookId, groups) {
  if (groups.length === 0) {
    return [];
  }
  const checks = [];
  const previous = await rpc("getPermissions", []);
  checks.push(check(previous?.ok === true && Boolean(previous.permissions), "live scratch permission policy captured", previous?.error?.message));
  const previousPermissions = previous?.permissions;
  const enabled = await rpc("setPermissions", [{
    allowWrites: true,
    allowDestructiveActions: true,
    allowWorkbookActions: true,
    requireConfirmationFor: [],
    scope: { workbookId }
  }]);
  checks.push(check(enabled?.ok === true, "live scratch permission policy enabled", enabled?.error?.message));
  if (enabled?.ok !== true) {
    return checks;
  }

  try {
    checks.push(...await runScenarioGroups(workbookId, groups));
  } finally {
    if (previousPermissions) {
      const restored = await rpc("setPermissions", [previousPermissions]);
      checks.push(check(restored?.ok === true, "live scratch permission policy restored", restored?.error?.message));
    }
  }
  return checks;
}

async function runRegressionPack(workbookId) {
  const checks = [];
  const sheetName = `${scratchSheet.slice(0, 22)}_REG`;
  const createSheet = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_regression_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live regression scratch sheet"
      }
    ]
  }]);
  checks.push(check(createSheet?.ok === true, "regression scratch sheet create", createSheet?.error?.message));
  if (createSheet?.ok !== true) {
    return checks;
  }

  const seed = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_regression_seed",
        kind: "range.write_values",
        workbookId,
        target: { workbookId, sheetName, address: "A1:E6" },
        values: [
          ["Date", "Customer", "Product", "Amount", "Status"],
          ["2026-01-03", "Acme Co", "Consulting", 1200, "Open"],
          ["2026-01-04", "Northwind", "Support", 450, "Closed"],
          ["2026-01-08", "Contoso", "Implementation", 3200, "Open"],
          ["2026-01-10", "Fabrikam", "Training", 800, "Open"],
          ["2026-01-12", "Tailspin", "Support", 650, "Closed"]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live regression seed data"
      }
    ]
  }]);
  checks.push(check(seed?.ok === true, "regression seed values", seed?.error?.message));
  if (seed?.ok !== true) {
    return checks;
  }

  const apply = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_regression_header_style",
        kind: "range.write_styles",
        workbookId,
        target: { workbookId, sheetName, address: "A1:E1" },
        style: { fillColor: "#000000", fontColor: "#FFFFFF", fontBold: true, horizontalAlignment: "center" },
        preserveValues: true,
        destructiveLevel: "format",
        reason: "Live regression header style"
      },
      {
        operationId: "live_regression_insert_column",
        kind: "range.insert_columns",
        workbookId,
        target: { workbookId, sheetName, address: "F1:F6" },
        destructiveLevel: "structure",
        reason: "Live regression insert column"
      },
      {
        operationId: "live_regression_reorder_columns",
        kind: "range.reorder_columns",
        workbookId,
        target: { workbookId, sheetName, address: "A1:B6" },
        columnOrder: [2, 1],
        destructiveLevel: "structure",
        reason: "Live regression reorder columns"
      },
      {
        operationId: "live_regression_validation",
        kind: "range.write_data_validation",
        workbookId,
        target: { workbookId, sheetName, address: "E2:E6" },
        validation: { type: "list", source: ["Open", "Reviewed", "Closed"], inCellDropDown: true, ignoreBlanks: true },
        destructiveLevel: "format",
        reason: "Live regression data validation"
      },
      {
        operationId: "live_regression_conditional_formatting",
        kind: "range.write_conditional_formatting",
        workbookId,
        target: { workbookId, sheetName, address: "A2:E6" },
        rule: { type: "custom", formula: "=$E2=\"Open\"", style: { fillColor: "#FFFF00" } },
        destructiveLevel: "format",
        reason: "Live regression conditional formatting"
      }
    ]
  }]);
  checks.push(check(apply?.ok === true, "regression operation batch apply", apply?.error?.message));

  const read = await readFullRange({ workbookId, sheetName, address: "A1:B6", includeFormulas: true });
  checks.push(check(read?.ok === true, "regression reordered range read", read?.error?.message));
  const values = extractValues(read);
  if (values) {
    checks.push(check(values[0]?.[0] === "Customer" && values[0]?.[1] === "Date", "regression reordered header values", JSON.stringify(values[0] ?? [])));
  }
  return checks;
}

async function runTemplateFormulaRepair(workbookId) {
  const checks = [];
  const sourceSheetName = `${scratchSheet.slice(0, 18)}_TPL_SRC`;
  const targetSheetName = `${scratchSheet.slice(0, 18)}_TPL_TGT`;
  const createSheets = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_template_source_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName: sourceSheetName,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live template formula repair source sheet"
      },
      {
        operationId: "live_template_target_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName: targetSheetName,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live template formula repair target sheet"
      }
    ]
  }]);
  checks.push(check(createSheets?.ok === true, "template formula scratch sheets create", createSheets?.error?.message));
  if (createSheets?.ok !== true) {
    return checks;
  }

  const seed = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_template_source_seed_values",
        kind: "range.write_values",
        workbookId,
        target: { workbookId, sheetName: sourceSheetName, address: "A1:D4" },
        values: [
          ["Item", "Amount", "Tax", "Total"],
          ["A", 100, 8, null],
          ["B", 125, 10, null],
          ["C", 150, 12, null]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live template formula source seed values"
      },
      {
        operationId: "live_template_source_seed_formulas",
        kind: "range.write_formulas",
        workbookId,
        target: { workbookId, sheetName: sourceSheetName, address: "D2:D4" },
        formulas: [["=B2+C2"], ["=B3+C3"], ["=B4+C4"]],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live template formula source formulas"
      },
      {
        operationId: "live_template_target_seed_values",
        kind: "range.write_values",
        workbookId,
        target: { workbookId, sheetName: targetSheetName, address: "A1:D4" },
        values: [
          ["Item", "Amount", "Tax", "Total"],
          ["A", 100, 8, null],
          ["B", 125, 10, null],
          ["C", 150, 12, null]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live template formula target seed values"
      },
      {
        operationId: "live_template_target_bad_formulas",
        kind: "range.write_formulas",
        workbookId,
        target: { workbookId, sheetName: targetSheetName, address: "D2:D4" },
        formulas: [["=B2-C2"], ["=B3-C3"], ["=B4-C4"]],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live template formula target incorrect formulas"
      }
    ]
  }]);
  checks.push(check(seed?.ok === true, "template formula seed source and target", seed?.error?.message));
  if (seed?.ok !== true) {
    return checks;
  }

  const registered = await rpc("registerTemplate", [{
    workbookId,
    name: "Live Formula Repair Template",
    scope: "workbook",
    sourceSheetName,
    dataRegions: []
  }]);
  checks.push(check(Boolean(registered?.templateId), "template formula register template", registered?.error?.message));
  if (!registered?.templateId) {
    return checks;
  }

  const repaired = await rpc("repairFormulasFromTemplate", [{
    workbookId,
    templateId: registered.templateId,
    targetSheetName
  }]);
  checks.push(check(repaired?.ok === true, "template formula repair apply", repaired?.result?.error?.message ?? repaired?.error?.message));

  const read = await readFullRange({ workbookId, sheetName: targetSheetName, address: "D2:D4", includeFormulas: true });
  checks.push(check(read?.ok === true, "template formula repaired range read", read?.error?.message));
  const formulas = extractFormulas(read);
  if (formulas) {
    checks.push(check(
      formulas[0]?.[0] === "=B2+C2" && formulas[2]?.[0] === "=B4+C4",
      "template formula repaired formulas",
      JSON.stringify(formulas)
    ));
  } else {
    checks.push(check(false, "template formula repaired formulas", "missing formula matrix"));
  }
  return checks;
}

async function runPivotChartCore(workbookId) {
  const checks = [];
  const sheetName = `${scratchSheet.slice(0, 20)}_PVT`;
  const sourceTableName = `${sheetName}_Source`;
  const pivotTableName = `${sheetName}_Pivot`;
  const chartName = `${sheetName}_Chart`;
  const createSheet = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_pivot_chart_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live PivotTable/chart scratch sheet"
      }
    ]
  }]);
  checks.push(check(createSheet?.ok === true, "pivot chart scratch sheet create", createSheet?.error?.message));
  if (createSheet?.ok !== true) {
    return checks;
  }

  const seed = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_pivot_chart_seed",
        kind: "range.write_values",
        workbookId,
        target: { workbookId, sheetName, address: "A1:C7" },
        values: [
          ["Status", "Region", "Amount"],
          ["Open", "North", 1200],
          ["Closed", "North", 450],
          ["Open", "South", 3200],
          ["Open", "East", 800],
          ["Closed", "South", 650],
          ["Reviewed", "East", 1100]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live PivotTable/chart source data"
      }
    ]
  }]);
  checks.push(check(seed?.ok === true, "pivot chart seed values", seed?.error?.message));
  if (seed?.ok !== true) {
    return checks;
  }

  const sourceTable = await rpc("createTable", [{
    workbookId,
    sheetName,
    address: "A1:C7",
    tableName: sourceTableName,
    hasHeaders: true,
    style: "TableStyleMedium2"
  }]);
  checks.push(check(sourceTable?.ok === true, "pivot chart create source table", sourceTable?.error?.message ?? sourceTable?.result?.error?.message));
  if (sourceTable?.ok !== true) {
    return checks;
  }

  const pivot = await rpc("createPivotTable", [{
    workbookId,
    pivotTableName,
    sourceTableName,
    destinationSheetName: sheetName,
    destinationAddress: "E3",
    rowFields: ["Status"],
    dataFields: [{ sourceFieldName: "Amount", summarizeBy: "sum", numberFormat: "#,##0" }],
    layout: { showRowGrandTotals: true },
    refresh: true
  }]);
  checks.push(check(pivot?.ok === true, "pivot chart create PivotTable", pivot?.error?.message ?? pivot?.result?.error?.message));
  if (pivot?.ok !== true) {
    return checks;
  }

  const pivotInfo = await rpc("getPivotTableInfo", [{ workbookId, pivotTableName }]);
  checks.push(check(pivotInfo?.ok === true, "pivot chart read PivotTable info", pivotInfo?.error?.message));

  const pivotValidation = await rpc("validatePivotSource", [{
    workbookId,
    pivotTableName,
    expectedFields: ["Status", "Amount"],
    expectedRowFields: ["Status"],
    expectedDataFields: ["Amount"],
    expectedDataFieldSettings: [{ sourceFieldName: "Amount", summarizeBy: "sum", numberFormat: "#,##0" }],
    expectedLayout: { showRowGrandTotals: true }
  }]);
  checks.push(check(pivotValidation?.ok !== false, "pivot chart validate PivotTable", summarizeIssues(pivotValidation?.issues)));

  const refreshedPivot = await rpc("refreshPivotTable", [{ workbookId, pivotTableName }]);
  checks.push(check(refreshedPivot?.ok === true, "pivot chart refresh PivotTable", refreshedPivot?.error?.message));

  const chart = await rpc("createChart", [{
    workbookId,
    sheetName,
    chartName,
    sourceAddress: "A1:C7",
    chartType: "ColumnClustered",
    seriesBy: "Columns",
    title: "Live Sales Status",
    position: { startCell: "J3", endCell: "P18" },
    style: 4
  }]);
  checks.push(check(chart?.ok === true, "pivot chart create chart", chart?.error?.message ?? chart?.result?.error?.message));
  if (chart?.ok !== true) {
    return checks;
  }

  const chartInfo = await rpc("getChartInfo", [{ workbookId, sheetName, chartName }]);
  checks.push(check(chartInfo?.ok === true, "pivot chart read chart info", chartInfo?.error?.message));
  const info = chartInfo?.info ?? chartInfo?.result?.info;
  if (info) {
    checks.push(check(info.chartName === chartName && info.title === "Live Sales Status", "pivot chart chart metadata", JSON.stringify({ chartName: info.chartName, title: info.title, chartType: info.chartType })));
  }

  const refreshedChart = await rpc("refreshChart", [{ workbookId, sheetName, chartName }]);
  checks.push(check(refreshedChart?.ok === true, "pivot chart refresh chart", refreshedChart?.error?.message));
  return checks;
}

async function runPivotTemplateRepair(workbookId) {
  const checks = [];
  const sheetName = `${scratchSheet.slice(0, 18)}_PVT_TPL`;
  const sourceTableName = `${sheetName}_Source`;
  const templatePivotTableName = `${sheetName}_Template`;
  const targetPivotTableName = `${sheetName}_Target`;
  const createSheet = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_pivot_template_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live PivotTable template repair scratch sheet"
      }
    ]
  }]);
  checks.push(check(createSheet?.ok === true, "pivot template scratch sheet create", createSheet?.error?.message));
  if (createSheet?.ok !== true) {
    return checks;
  }

  const seed = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_pivot_template_seed",
        kind: "range.write_values",
        workbookId,
        target: { workbookId, sheetName, address: "A1:C7" },
        values: [
          ["Status", "Region", "Amount"],
          ["Open", "North", 1200],
          ["Closed", "North", 450],
          ["Open", "South", 3200],
          ["Open", "East", 800],
          ["Closed", "South", 650],
          ["Reviewed", "East", 1100]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live PivotTable template repair source data"
      }
    ]
  }]);
  checks.push(check(seed?.ok === true, "pivot template seed values", seed?.error?.message));
  if (seed?.ok !== true) {
    return checks;
  }

  const sourceTable = await rpc("createTable", [{
    workbookId,
    sheetName,
    address: "A1:C7",
    tableName: sourceTableName,
    hasHeaders: true,
    style: "TableStyleMedium2"
  }]);
  checks.push(check(sourceTable?.ok === true, "pivot template create source table", sourceTable?.error?.message ?? sourceTable?.result?.error?.message));
  if (sourceTable?.ok !== true) {
    return checks;
  }

  const templatePivot = await rpc("createPivotTable", [{
    workbookId,
    pivotTableName: templatePivotTableName,
    sourceTableName,
    destinationSheetName: sheetName,
    destinationAddress: "E3",
    rowFields: ["Region"],
    dataFields: [{ sourceFieldName: "Amount", summarizeBy: "sum", numberFormat: "#,##0" }],
    layout: { showRowGrandTotals: true },
    refresh: true
  }]);
  checks.push(check(templatePivot?.ok === true, "pivot template create source PivotTable", templatePivot?.error?.message ?? templatePivot?.result?.error?.message));

  const targetPivot = await rpc("createPivotTable", [{
    workbookId,
    pivotTableName: targetPivotTableName,
    sourceTableName,
    destinationSheetName: sheetName,
    destinationAddress: "J3",
    rowFields: ["Status"],
    dataFields: [{ sourceFieldName: "Amount", summarizeBy: "sum", numberFormat: "#,##0" }],
    layout: { showRowGrandTotals: true },
    refresh: true
  }]);
  checks.push(check(targetPivot?.ok === true, "pivot template create target PivotTable", targetPivot?.error?.message ?? targetPivot?.result?.error?.message));
  if (templatePivot?.ok !== true || targetPivot?.ok !== true) {
    return checks;
  }

  const beforeDiff = await rpc("diffPivotTables", [{
    workbookId,
    pivotTableName: templatePivotTableName,
    targetPivotTableName
  }]);
  checks.push(check(hasPivotDiffPath(beforeDiff, "layout.rowFields"), "pivot template diff detects row field mismatch", summarizePivotDiff(beforeDiff)));

  const repaired = await rpc("repairPivotFromTemplate", [{
    workbookId,
    pivotTableName: targetPivotTableName,
    templatePivotTableName,
    dimensions: ["layout", "fields", "dataFields", "numberFormats", "refresh"],
    strict: false
  }]);
  checks.push(check(repaired?.ok === true, "pivot template repair apply", repaired?.error?.message ?? repaired?.copy?.error?.message));
  if (repaired?.ok !== true) {
    return checks;
  }

  const targetInfo = await rpc("getPivotTableInfo", [{ workbookId, pivotTableName: targetPivotTableName }]);
  checks.push(check(targetInfo?.ok === true, "pivot template read repaired target info", targetInfo?.error?.message));
  const rowFields = pivotRowFieldNames(targetInfo?.info ?? targetInfo?.result?.info);
  checks.push(check(rowFields.includes("Region") && !rowFields.includes("Status"), "pivot template repaired row field", JSON.stringify(rowFields)));

  const afterDiff = await rpc("diffPivotTables", [{
    workbookId,
    pivotTableName: templatePivotTableName,
    targetPivotTableName
  }]);
  checks.push(check(!hasPivotDiffPath(afterDiff, "layout.rowFields"), "pivot template row field diff cleared", summarizePivotDiff(afterDiff)));
  return checks;
}

async function runChartTemplateCopy(workbookId) {
  const checks = [];
  const sheetName = `${scratchSheet.slice(0, 18)}_CHT_TPL`;
  const templateChartName = `${sheetName}_Template`;
  const targetChartName = `${sheetName}_Target`;
  const createSheet = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_chart_template_sheet",
        kind: "sheet.create",
        workbookId,
        sheetName,
        activate: false,
        destructiveLevel: "structure",
        reason: "Live chart template copy scratch sheet"
      }
    ]
  }]);
  checks.push(check(createSheet?.ok === true, "chart template scratch sheet create", createSheet?.error?.message));
  if (createSheet?.ok !== true) {
    return checks;
  }

  const seed = await rpc("applyBatch", [{
    workbookId,
    mode: "apply",
    operations: [
      {
        operationId: "live_chart_template_seed",
        kind: "range.write_values",
        workbookId,
        target: { workbookId, sheetName, address: "A1:C5" },
        values: [
          ["Month", "Revenue", "Expense"],
          ["Jan", 1200, 800],
          ["Feb", 1450, 900],
          ["Mar", 1600, 1100],
          ["Apr", 1750, 1150]
        ],
        preserveFormats: true,
        destructiveLevel: "values",
        reason: "Live chart template source data"
      }
    ]
  }]);
  checks.push(check(seed?.ok === true, "chart template seed values", seed?.error?.message));
  if (seed?.ok !== true) {
    return checks;
  }

  const templateChart = await rpc("createChart", [{
    workbookId,
    sheetName,
    chartName: templateChartName,
    sourceAddress: "A1:C5",
    chartType: "Line",
    seriesBy: "Columns",
    title: "Template Revenue Trend",
    position: { startCell: "E2", endCell: "L16" },
    style: 5
  }]);
  checks.push(check(templateChart?.ok === true, "chart template create template chart", templateChart?.error?.message ?? templateChart?.result?.error?.message));

  const targetChart = await rpc("createChart", [{
    workbookId,
    sheetName,
    chartName: targetChartName,
    sourceAddress: "A1:C5",
    chartType: "ColumnClustered",
    seriesBy: "Columns",
    title: "Target Before Copy",
    position: { startCell: "N2", endCell: "U16" },
    style: 2
  }]);
  checks.push(check(targetChart?.ok === true, "chart template create target chart", targetChart?.error?.message ?? targetChart?.result?.error?.message));
  if (templateChart?.ok !== true || targetChart?.ok !== true) {
    return checks;
  }

  const copied = await rpc("copyChartFromTemplate", [{
    workbookId,
    templateSheetName: sheetName,
    templateChartName,
    sheetName,
    chartName: targetChartName
  }]);
  checks.push(check(copied?.ok === true, "chart template copy apply", copied?.error?.message ?? copied?.result?.error?.message));
  if (copied?.ok !== true) {
    return checks;
  }

  const copiedFields = copied?.result?.copied ?? copied?.copied ?? [];
  checks.push(check(["chartType", "style", "title", "position"].every((field) => copiedFields.includes(field)), "chart template copied dimensions", JSON.stringify(copiedFields)));

  const targetInfo = await rpc("getChartInfo", [{ workbookId, sheetName, chartName: targetChartName }]);
  checks.push(check(targetInfo?.ok === true, "chart template read copied target", targetInfo?.error?.message));
  const info = targetInfo?.info ?? targetInfo?.result?.info;
  if (info) {
    checks.push(check(
      info.title === "Template Revenue Trend" && String(info.chartType).toLowerCase().includes("line"),
      "chart template copied metadata",
      JSON.stringify({ title: info.title, chartType: info.chartType, style: info.style })
    ));
  }
  return checks;
}

function scenarioGroupsFromInput(raw, deepMode) {
  const groups = new Set();
  if (deepMode) groups.add("scratch-core");
  if (typeof raw === "string") {
    for (const item of raw.split(",")) {
      const group = item.trim();
      if (group) groups.add(group);
    }
  }
  return [...groups];
}

function extractValues(readResult) {
  const candidates = [
    readResult?.values,
    readResult?.data?.values,
    readResult?.data?.[0]?.snapshot?.values,
    readResult?.readData?.[0]?.snapshot?.values,
    readResult?.snapshot?.values,
    readResult?.range?.values,
    readResult?.answer?.values
  ];
  return candidates.find((value) => Array.isArray(value));
}

function extractFormulas(readResult) {
  const candidates = [
    readResult?.formulas,
    readResult?.data?.formulas,
    readResult?.data?.[0]?.snapshot?.formulas,
    readResult?.readData?.[0]?.snapshot?.formulas,
    readResult?.snapshot?.formulas,
    readResult?.range?.formulas,
    readResult?.answer?.formulas
  ];
  return candidates.find((value) => Array.isArray(value));
}

function summarizeIssues(issues) {
  return Array.isArray(issues) && issues.length > 0
    ? issues.map((issue) => `${issue.code}:${issue.severity}`).join(", ")
    : undefined;
}

function pivotRowFieldNames(info) {
  return Array.isArray(info?.rowHierarchies) ? info.rowHierarchies.map((hierarchy) => hierarchy.name).filter(Boolean) : [];
}

function hasPivotDiffPath(diffResult, pathName) {
  const changes = diffResult?.changes ?? diffResult?.diff?.changes ?? [];
  return Array.isArray(changes) && changes.some((change) => change.path === pathName);
}

function summarizePivotDiff(diffResult) {
  const changes = diffResult?.changes ?? diffResult?.diff?.changes ?? [];
  return Array.isArray(changes) && changes.length > 0
    ? changes.map((change) => change.path).join(", ")
    : undefined;
}

async function readFullRange(request) {
  return rpc("applyBatch", [{
    workbookId: request.workbookId,
    mode: "dry_run",
    operations: [
      {
        operationId: `live_read_${request.sheetName}_${request.address}`.replace(/[^A-Za-z0-9_]/g, "_").slice(0, 80),
        kind: "range.read_full",
        workbookId: request.workbookId,
        target: {
          workbookId: request.workbookId,
          sheetName: request.sheetName,
          address: request.address
        },
        includeFormulas: request.includeFormulas,
        destructiveLevel: "none",
        reason: "Live smoke read"
      }
    ]
  }]);
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
