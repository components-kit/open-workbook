#!/usr/bin/env node
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-e2e-fast-"));
const artifactsDir = path.join(tempRoot, "artifacts");
mkdirSync(artifactsDir, { recursive: true });

const workbookId = "workbook_e2e_fast";
const backendPort = 37880 + Math.floor(Math.random() * 500);
const backendUrl = `http://127.0.0.1:${backendPort}`;
const backendWsUrl = `ws://127.0.0.1:${backendPort}/addin`;
const maxFastMs = Number(process.env.OPEN_WORKBOOK_E2E_FAST_BUDGET_MS ?? 300_000);

const transcript = [];

async function main() {
  const started = performance.now();
  const server = spawn(process.execPath, ["apps/mcp-server/dist/index.js", "--standalone", "--agent-name", "e2e-fast"], {
    cwd: repoRoot,
    stdio: ["pipe", "pipe", "pipe"],
    env: {
      ...process.env,
      OPEN_WORKBOOK_HOST: "127.0.0.1",
      OPEN_WORKBOOK_PORT: String(backendPort),
      OPEN_WORKBOOK_ADDIN_PATH: "/addin",
      OPEN_WORKBOOK_STATE_DIR: path.join(tempRoot, "state"),
      OPEN_WORKBOOK_BACKUP_DIR: path.join(tempRoot, "backups"),
      OPEN_WORKBOOK_MCP_SURFACE: "advanced",
      OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1"
    }
  });

  let serverStderr = "";
  server.stderr.on("data", (chunk) => {
    serverStderr += String(chunk);
  });

  const mcp = new McpClient(server);
  let fakeAddin;

  try {
    await waitForHttp(`${backendUrl}/status`, 15_000);
    fakeAddin = await FakeAddin.connect(backendWsUrl, createWorkbookFixture(workbookId));

    await mcp.request("initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "open-workbook-e2e-fast", version: "0.0.0" }
    });
    mcp.notify("notifications/initialized", {});

    await runProtocolSweep(mcp);
    await runCoreWorkbookWorkflow(mcp);
    await runTableLargeWorkflow(mcp);
    await runSafetyAndRollbackWorkflow(mcp);
    await runStableToolGroupSweep(mcp);
    await runMultiAgentWorkflow(mcp);

    const elapsedMs = performance.now() - started;
    assert(elapsedMs <= maxFastMs, `fast E2E exceeded budget: ${Math.round(elapsedMs)}ms > ${maxFastMs}ms`);
    writeArtifact("e2e-fast-transcript.json", { elapsedMs, transcript, fakeAddin: fakeAddin.summary(), summary: summarizeTranscript(transcript) });
    console.log(`E2E fast passed in ${Math.round(elapsedMs)}ms. Artifacts: ${artifactsDir}`);
  } catch (error) {
    writeArtifact("e2e-fast-failure.json", {
      error: error instanceof Error ? { message: error.message, stack: error.stack } : String(error),
      transcript,
      summary: summarizeTranscript(transcript),
      serverStderr,
      fakeAddin: fakeAddin?.summary()
    });
    console.error(`E2E fast failed. Artifacts: ${artifactsDir}`);
    throw error;
  } finally {
    fakeAddin?.close();
    mcp.close();
    server.kill();
  }
}

async function runProtocolSweep(client) {
  const tools = await client.request("tools/list", {});
  const names = tools.tools.map((tool) => tool.name);
  for (const name of [
    "excel.runtime.get_status",
    "excel.runtime.get_active_context",
    "excel.workbook.get_workbook_map",
    "excel.range.write_values",
    "excel.table.reorder_columns",
    "excel.workflow.prepare_session",
    "excel.workflow.create_formula_sheet",
    "excel.workflow.create_template_report",
    "excel.workflow.create_pivot_chart_summary",
    "excel.workflow.repair_formula_errors",
    "excel.workflow.preview_risky_edit",
    "excel.transaction.list"
  ]) {
    assert(names.includes(name), `expected MCP tool ${name}`);
  }
  const resources = await client.request("resources/list", {});
  assert(Array.isArray(resources.resources), "resources/list should return resources");
  const prompts = await client.request("prompts/list", {});
  assert(Array.isArray(prompts.prompts), "prompts/list should return prompts");
}

async function runCoreWorkbookWorkflow(client) {
  const status = await callTool(client, "excel.runtime.get_status", {});
  assert(status.sessions.length >= 1, "runtime should have connected fake add-in session");

  const context = await callTool(client, "excel.runtime.get_active_context", {});
  assert(context.ok && context.activeWorkbook?.workbookId === workbookId, "active context should return fixture workbook");

  const capabilities = await callTool(client, "excel.runtime.get_capabilities", {});
  assert(capabilities.catalog.tools.length > 100, "capabilities should include optimized tool catalog");

  const map = await callTool(client, "excel.workbook.get_workbook_map", {});
  assert(map.ok && map.map.sheets.some((sheet) => sheet.name === "Data"), "workbook map should include Data sheet");

  const permissions = await callTool(client, "excel.permissions.set", {
    allowWrites: true,
    allowDestructiveActions: true,
    requireConfirmationFor: [],
    scope: { workbookId }
  });
  assert(permissions.ok, "E2E permissions should allow scoped mutations");

  const created = await callTool(client, "excel.sheet.create", { workbookId, sheetName: "Scratch" });
  assertApplied(created, "sheet create");

  const write = await callTool(client, "excel.range.write_values", {
    workbookId,
    sheetName: "Scratch",
    address: "A1:C3",
    values: [
      ["Account", "Amount", "Status"],
      ["A-100", 125, "Open"],
      ["A-200", 300, "Closed"]
    ]
  });
  assertApplied(write, "range write values");
  assert(writtenCellCount(write) === 9, "range write should report cells written");

  const read = await callTool(client, "excel.range.read_compact", {
    workbookId,
    sheetName: "Scratch",
    address: "A1:C3",
    includeValues: true,
    responseMode: "verbose"
  });
  assert(read.ok && read.values?.[1]?.[1] === 125, "range read should return written values");

  const formulas = await callTool(client, "excel.range.write_formulas", {
    workbookId,
    sheetName: "Scratch",
    address: "D2:D3",
    formulas: [["=B2*2"], ["=B3*2"]]
  });
  assertApplied(formulas, "range write formulas");
  assert(formulas.transactionId, "formula write should include transaction id");

  const formats = await callTool(client, "excel.range.write_number_formats", {
    workbookId,
    sheetName: "Scratch",
    address: "B2:B3",
    numberFormat: [["$#,##0"], ["$#,##0"]]
  });
  assertApplied(formats, "range write number formats");

  const copied = await callTool(client, "excel.range.copy", {
    workbookId,
    sourceSheetName: "Scratch",
    sourceAddress: "A1:C3",
    targetSheetName: "Scratch",
    targetAddress: "F1:H3",
    copyType: "all"
  });
  assertApplied(copied, "range copy");

  const cleared = await callTool(client, "excel.range.clear_values_keep_format", {
    workbookId,
    sheetName: "Scratch",
    address: "C2:C3"
  });
  assertApplied(cleared, "clear values keep format");

  const formulaSheet = await callTool(client, "excel.workflow.create_formula_sheet", {
    workbookId,
    sheetName: "QuickEntryWorkflow",
    valuesAddress: "A1:C4",
    values: [
      ["Account", "Amount", "Tax"],
      ["A-100", 125, 10],
      ["A-200", 300, 24],
      ["A-300", 90, 7.2]
    ],
    formulasAddress: "D1:D4",
    formulas: [["Total"], ["=B2+C2"], ["=B3+C3"], ["=B4+C4"]],
    numberFormatAddress: "B2:D4",
    numberFormat: [
      ["$#,##0.00", "$#,##0.00", "$#,##0.00"],
      ["$#,##0.00", "$#,##0.00", "$#,##0.00"],
      ["$#,##0.00", "$#,##0.00", "$#,##0.00"]
    ],
    validateAddress: "D2:D4"
  });
  assert(formulaSheet.ok && hasTruthyKey(formulaSheet, "formulaValidationOk"), "formula sheet workflow should validate formulas");

  const formulaRepair = await callTool(client, "excel.workflow.repair_formula_errors", {
    workbookId,
    sheetName: "QuickEntryWorkflow",
    errorAddress: "D2:D4",
    patternAddress: "D2:D4",
    sourceAddress: "D2:D2",
    targetAddress: "D2:D4",
    direction: "down"
  });
  assert(formulaRepair.ok && hasTruthyKey(formulaRepair, "formulaValidationOk"), "formula repair workflow should validate after repair");
}

async function runTableLargeWorkflow(client) {
  const rows = generateRows(5000);
  const created = await callTool(client, "excel.table.create", {
    workbookId,
    sheetName: "Large",
    address: `A1:E${rows.length + 1}`,
    tableName: "TransactionsLarge",
    hasHeaders: true,
    values: [["Date", "Account", "Amount", "Region", "Status"], ...rows],
    style: "TableStyleMedium2"
  });
  assert(created.ok && created.transactionId && created.backup, "table create should transact and back up");

  const page = await callTool(client, "excel.table.read_compact", {
    workbookId,
    tableName: "TransactionsLarge",
    includeValues: true,
    columns: ["Account", "Amount", "Status"],
    rowOffset: 100,
    maxRows: 25,
    responseMode: "verbose"
  });
  assert(page.values.length === 25, "large table read should return requested row window");
  assert(page.truncated === true, "large table read should mark truncated page");

  const reordered = await callTool(client, "excel.table.reorder_columns", {
    workbookId,
    tableName: "TransactionsLarge",
    columnOrder: ["Status", "Date", "Account", "Amount", "Region"]
  });
  assert(reordered.ok && reordered.transactionId && reordered.backup, "table reorder should transact and back up");

  const filtered = await callTool(client, "excel.table.apply_filters", {
    workbookId,
    tableName: "TransactionsLarge",
    filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }]
  });
  assert(filtered.ok && filtered.transactionId, "table filter should transact");

  const sorted = await callTool(client, "excel.table.sort", {
    workbookId,
    tableName: "TransactionsLarge",
    fields: [{ key: 3, ascending: false }]
  });
  assert(sorted.ok && sorted.transactionId, "table sort should transact");

  const info = await callTool(client, "excel.table.get_info", { workbookId, tableName: "TransactionsLarge" });
  assert(info.info.columns[0].name === "Status", "table column order should persist");
}

async function runSafetyAndRollbackWorkflow(client) {
  const backupResult = await callTool(client, "excel.workbook.create_backup", {
    workbookId,
    reason: "E2E rollback backup",
    ranges: [{ workbookId, sheetName: "Scratch", address: "A1:D3" }]
  });
  assert(backupResult.ok && backupResult.backup?.backupId, "manual workbook backup should be created");

  const write = await callTool(client, "excel.range.write_values", {
    workbookId,
    sheetName: "Scratch",
    address: "B2:B2",
    values: [[999]]
  });
  assertApplied(write, "pre-rollback write");

  const restored = await callTool(client, "excel.workbook.restore_backup", {
    backupId: backupResult.backup.backupId,
    confirmationToken: "confirm"
  });
  assertApplied(restored, "backup restore");

  const txs = await callTool(client, "excel.transaction.list", { workbookId });
  assert(txs.transactions.some((tx) => tx.transactionId === write.transactionId), "transaction list should include write transaction");

  const prepared = await callTool(client, "excel.workflow.prepare_session", { workbookId });
  assert(prepared.ok && prepared.activeContext?.activeWorkbook?.workbookId === workbookId, "workflow prepare session should return active workbook context");

  const workflow = await callTool(client, "excel.workflow.preview_risky_edit", {
    workbookId,
    reason: "E2E combined risky edit",
    ranges: [{ workbookId, sheetName: "Scratch", address: "B2:B2" }],
    operations: [
      {
        tool: "excel.range.write_values",
        args: {
          sheetName: "Scratch",
          address: "B2:B2",
          values: [[321]]
        }
      }
    ]
  });
  assert(workflow.ok && workflow.applied === true, "workflow risky edit should apply");
  assert(workflow.diff?.ok === true, "workflow risky edit should include snapshot diff");
  assert(workflow.rollbackPreview?.rollbackAvailable === true, "workflow risky edit should include rollback preview");
  assert(hasTruthyKey(workflow, "previewed"), "workflow risky edit should include rollback summary");

  const sparseBlocked = await callTool(client, "excel.workflow.preview_risky_edit", {
    workbookId,
    reason: "E2E sparse write guard",
    ranges: [{ workbookId, sheetName: "Scratch", address: "A1:D4" }],
    operations: [
      {
        tool: "excel.range.write_values",
        args: {
          sheetName: "Scratch",
          address: "A1:D4",
          values: [
            ["only one cell", null, null, null],
            [null, null, null, null],
            [null, null, null, null],
            [null, null, null, null]
          ]
        }
      }
    ]
  });
  assert(sparseBlocked.ok === false && sparseBlocked.errorStep === "sparse_write_guard", "workflow sparse range write should be blocked");

  const telemetry = await callTool(client, "excel.conflict.get_telemetry", { workbookId });
  assert(telemetry.ok, "conflict telemetry should be readable");
}

async function runStableToolGroupSweep(client) {
  const template = await runTemplateStyleFormulaSweep(client);
  await runNamesRegionsCleaningSweep(client, template.templateId);
  await runPivotChartSweep(client);
  await runSnapshotsDiffsEventsPermissionsSweep(client);
}

async function runTemplateStyleFormulaSweep(client) {
  const detected = await callTool(client, "excel.template.detect_templates", { workbookId });
  assert(detected.ok && detected.candidates.length >= 1, "template detection should return candidates");

  const registered = await callTool(client, "excel.template.register", {
    workbookId,
    name: "Monthly Template",
    scope: "workbook",
    sourceSheetName: "Template",
    dataRegions: ["A2:C3"]
  });
  assert(registered.templateId, "template register should return templateId");

  const templateId = registered.templateId;
  assert((await callTool(client, "excel.template.get", { templateId })).ok, "template get should succeed");
  assert((await callTool(client, "excel.template.list", { workbookId })).some((template) => template.templateId === templateId), "template list should include registered template");
  assert((await callTool(client, "excel.template.infer_regions", { templateId })).ok, "template infer regions should succeed");

  const created = await callTool(client, "excel.template.create_sheet_from_template", {
    workbookId,
    templateId,
    newSheetName: "Report_Jan",
    clearDataRegions: true
  });
  assertApplied(created, "template sheet create");

  const fill = await callTool(client, "excel.template.fill_regions", {
    workbookId,
    targetSheetName: "Report_Jan",
    regions: [{ address: "A2:C3", values: [["Revenue", 1200, 1300], ["Cost", 500, 550]] }]
  });
  assertApplied(fill, "template fill regions");

  const validation = await callTool(client, "excel.template.validate_sheet_against_template", {
    workbookId,
    templateId,
    targetSheetName: "Report_Jan"
  });
  assert(validation.ok === true, "template validation should pass in fake host");

  const styleFingerprint = await callTool(client, "excel.style.get_fingerprint", {
    workbookId,
    sheetName: "Report_Jan",
    address: "A1:C3",
    maxCellSamples: 9
  });
  assert(styleFingerprint.ok && styleFingerprint.fingerprint.dimensions, "style fingerprint should be captured");

  const styleCompare = await callTool(client, "excel.style.compare_fingerprint", {
    workbookId,
    sourceSheetName: "Template",
    targetSheetName: "Report_Jan",
    sourceAddress: "A1:C3",
    targetAddress: "A1:C3",
    dimensions: ["fills", "fonts", "numberFormats"]
  });
  assert(styleCompare.ok === true, "style compare should pass in fake host");

  const styleApply = await callTool(client, "excel.style.apply_style", {
    workbookId,
    sheetName: "Report_Jan",
    address: "A1:C1",
    style: { fill: { color: "#D9EAF7" }, font: { bold: true } }
  });
  assertApplied(styleApply, "style apply");

  const styleCopy = await callTool(client, "excel.style.copy_fills", {
    workbookId,
    sourceSheetName: "Template",
    targetSheetName: "Report_Jan",
    sourceAddress: "A1:C3",
    targetAddress: "A1:C3"
  });
  assert(styleCopy.ok === true && styleCopy.backup?.ok === true, "style copy should create backup and validate");

  const formulaPatterns = await callTool(client, "excel.formula.read_patterns", {
    workbookId,
    sheetName: "Scratch",
    address: "D2:D3"
  });
  assert(formulaPatterns.ok && formulaPatterns.patterns.cells.length >= 1, "formula patterns should be readable");

  const formulaGraph = await callTool(client, "excel.formula.get_dependency_graph", {
    workbookId,
    sheetName: "Scratch",
    address: "D2:D3"
  });
  assert(formulaGraph.ok && Array.isArray(formulaGraph.graph.edges), "formula dependency graph should be returned");

  assert((await callTool(client, "excel.formula.trace_precedents", { workbookId, sheetName: "Scratch", address: "D2" })).ok, "formula precedent trace should pass");
  assert((await callTool(client, "excel.formula.trace_dependents", { workbookId, sheetName: "Scratch", address: "B2" })).ok, "formula dependent trace should pass");
  assert((await callTool(client, "excel.formula.validate", { workbookId, sheetName: "Scratch", address: "D2:D3" })).ok, "formula validation should pass");
  assert((await callTool(client, "excel.formula.explain", { workbookId, formula: "=SUM(B2:B3)" })).ok, "formula explain should pass");

  const copyFormula = await callTool(client, "excel.formula.copy_patterns", {
    workbookId,
    sourceSheetName: "Scratch",
    targetSheetName: "Report_Jan",
    sourceAddress: "D2:D3",
    targetAddress: "D2:D3"
  });
  assert(copyFormula.ok === true && copyFormula.backup?.ok === true, "formula copy should create backup");

  const fillDown = await callTool(client, "excel.formula.fill_down", {
    workbookId,
    sheetName: "Report_Jan",
    sourceAddress: "D2:D2",
    targetAddress: "D2:D3"
  });
  assert(fillDown.ok === true && fillDown.backup?.ok === true, "formula fill down should create backup");

  const repair = await callTool(client, "excel.formula.repair_patterns", {
    workbookId,
    templateId,
    targetSheetName: "Report_Jan"
  });
  assert(repair.ok === true && repair.backup?.ok === true, "formula repair should create backup");

  const styleRepair = await callTool(client, "excel.repair.style_from_template", {
    workbookId,
    templateId,
    targetSheetName: "Report_Jan"
  });
  assert(styleRepair.ok === true && Array.isArray(styleRepair.backups) && styleRepair.backups.length > 0, "style repair should create backup");

  const workflowReport = await callTool(client, "excel.workflow.create_template_report", {
    workbookId,
    templateId,
    newSheetName: "Report_Workflow",
    fillRegions: [
      {
        address: "A2:C3",
        values: [
          ["North", 1200, "Open"],
          ["South", 980, "Closed"]
        ]
      }
    ]
  });
  assert(workflowReport.ok && hasTruthyKey(workflowReport, "styleCompared") && hasTruthyKey(workflowReport, "validated"), "template workflow should compare styles and validate");

  return { templateId };
}

async function runNamesRegionsCleaningSweep(client, templateId) {
  const nameCreated = await callTool(client, "excel.names.create", {
    workbookId,
    name: "InputBlock",
    sheetName: "Report_Jan",
    reference: "A2:C3",
    comment: "E2E named range"
  });
  assert(nameCreated.ok === true && nameCreated.transactionId, "name create should transact");

  assert((await callTool(client, "excel.names.list", { workbookId })).ok, "names list should pass");
  assert((await callTool(client, "excel.names.get", { workbookId, name: "InputBlock", sheetName: "Report_Jan" })).ok, "names get should pass");
  assert((await callTool(client, "excel.names.update", { workbookId, name: "InputBlock", sheetName: "Report_Jan", reference: "A2:C3", visible: true })).ok, "names update should pass");

  const region = await callTool(client, "excel.region.register", {
    workbookId,
    name: "InputRegion",
    sheetName: "Report_Jan",
    address: "A2:C3",
    kind: "data",
    templateId
  });
  assert(region.ok && region.region?.regionId, "region register should pass");
  assert((await callTool(client, "excel.region.detect", { workbookId })).ok, "region detect should pass");
  assert((await callTool(client, "excel.region.list", { workbookId })).ok, "region list should pass");
  assert((await callTool(client, "excel.region.get", { workbookId, regionName: "InputRegion" })).ok, "region get should pass");

  const regionFill = await callTool(client, "excel.region.fill", {
    workbookId,
    regionName: "InputRegion",
    clearFirst: true,
    values: [[" Revenue ", " 1400 ", "1500"], ["Cost", "", " 650 "]]
  });
  assertApplied(regionFill, "region fill");

  assert((await callTool(client, "excel.clean.detect_header_row", { workbookId, sheetName: "Report_Jan", address: "A1:C3" })).ok, "clean detect header should pass");
  assert((await callTool(client, "excel.clean.trim_whitespace", { workbookId, sheetName: "Report_Jan", address: "A2:C3" })).ok, "clean trim should pass");
  assert((await callTool(client, "excel.clean.parse_numbers", { workbookId, sheetName: "Report_Jan", address: "B2:C3" })).ok, "clean parse numbers should pass");
  assert((await callTool(client, "excel.clean.fill_missing_values", { workbookId, sheetName: "Report_Jan", address: "B2:C3", strategy: "zero" })).ok, "clean fill missing should pass");
  assert((await callTool(client, "excel.clean.remove_duplicates", { workbookId, sheetName: "Report_Jan", address: "A1:C3", hasHeader: true })).ok, "clean remove duplicates should pass");
  assert((await callTool(client, "excel.clean.split_column", { workbookId, sheetName: "Report_Jan", address: "A2:A3", columnIndex: 0, delimiter: "e", targetAddress: "E2:F3" })).ok, "clean split column should pass");
  assert((await callTool(client, "excel.clean.merge_columns", { workbookId, sheetName: "Report_Jan", address: "A2:C3", columnIndexes: [0, 1], separator: "-", targetAddress: "G2:G3" })).ok, "clean merge columns should pass");
  assert((await callTool(client, "excel.clean.detect_outliers", { workbookId, sheetName: "Report_Jan", address: "B2:C3", columnIndex: 0 })).ok, "clean detect outliers should pass");
  assert((await callTool(client, "excel.clean.fuzzy_match", { workbookId, sheetName: "Report_Jan", address: "A2:A3", lookupValues: ["Revenue", "Cost"], threshold: 0.5 })).ok, "clean fuzzy match should pass");
}

async function runPivotChartSweep(client) {
  const pivotCreated = await callTool(client, "excel.pivot.create", {
    workbookId,
    pivotTableName: "PivotSales",
    sourceTableName: "TransactionsLarge",
    destinationSheetName: "Report_Jan",
    destinationAddress: "J2",
    rowFields: ["Status"],
    dataFields: [{ sourceFieldName: "Amount", name: "Total Amount", summarizeBy: "Sum", numberFormat: "$#,##0" }]
  });
  assert(pivotCreated.ok === true && pivotCreated.transactionId, "pivot create should transact");

  assert((await callTool(client, "excel.pivot.list", { workbookId })).ok, "pivot list should pass");
  assert((await callTool(client, "excel.pivot.get_info", { workbookId, pivotTableName: "PivotSales" })).ok, "pivot info should pass");
  assert((await callTool(client, "excel.pivot.validate_source", { workbookId, pivotTableName: "PivotSales", expectedFields: ["Status", "Amount"] })).ok, "pivot source validation should pass");
  assert((await callTool(client, "excel.pivot.get_capability_matrix", { workbookId })).ok, "pivot capability matrix should pass");
  assert((await callTool(client, "excel.pivot.get_fingerprint", { workbookId, pivotTableName: "PivotSales" })).ok, "pivot fingerprint should pass");
  assert((await callTool(client, "excel.pivot.refresh", { workbookId, pivotTableName: "PivotSales" })).ok, "pivot refresh should pass");

  const chartCreated = await callTool(client, "excel.chart.create", {
    workbookId,
    sheetName: "Report_Jan",
    chartName: "SalesChart",
    sourceAddress: "A1:C3",
    chartType: "ColumnClustered",
    title: "Sales E2E"
  });
  assert(chartCreated.ok === true && chartCreated.transactionId, "chart create should transact");
  assert((await callTool(client, "excel.chart.list", { workbookId })).ok, "chart list should pass");
  assert((await callTool(client, "excel.chart.get_info", { workbookId, sheetName: "Report_Jan", chartName: "SalesChart" })).ok, "chart info should pass");
  assert((await callTool(client, "excel.chart.update_data_source", { workbookId, sheetName: "Report_Jan", chartName: "SalesChart", sourceAddress: "A1:C3" })).ok, "chart source update should pass");
  assert((await callTool(client, "excel.chart.refresh", { workbookId, sheetName: "Report_Jan", chartName: "SalesChart" })).ok, "chart refresh should pass");
  assert((await callTool(client, "excel.chart.validate_against_template", { workbookId, sheetName: "Report_Jan", chartName: "SalesChart" })).ok, "chart validation should pass");

  const workflow = await callTool(client, "excel.workflow.create_pivot_chart_summary", {
    workbookId,
    pivotTableName: "PivotWorkflow",
    sourceTableName: "TransactionsLarge",
    pivotDestinationSheetName: "Report_Jan",
    pivotDestinationAddress: "J2",
    rowFields: ["Status"],
    columnFields: ["Region"],
    dataFields: [{ sourceFieldName: "Amount", summarizeBy: "Sum", name: "Sum of Amount" }],
    chartSheetName: "Report_Jan",
    chartName: "WorkflowSalesChart",
    chartSourceAddress: "J2:M8",
    chartType: "ColumnClustered",
    chartTitle: "Workflow Sales"
  });
  assert(workflow.ok && hasTruthyKey(workflow, "refreshed") && hasTruthyKey(workflow, "validated"), "pivot/chart workflow should create, refresh, and validate");
}

async function runSnapshotsDiffsEventsPermissionsSweep(client) {
  const left = await callTool(client, "excel.snapshot.create", {
    workbookId,
    reason: "E2E left snapshot",
    ranges: [{ workbookId, sheetName: "Report_Jan", address: "A1:C3" }]
  });
  assert(left.ok && left.snapshot?.snapshotId, "snapshot create should pass");

  const mutate = await callTool(client, "excel.range.write_values", {
    workbookId,
    sheetName: "Report_Jan",
    address: "C2:C2",
    values: [[1600]]
  });
  assertApplied(mutate, "snapshot diff mutation");

  const right = await callTool(client, "excel.snapshot.create", {
    workbookId,
    reason: "E2E right snapshot",
    ranges: [{ workbookId, sheetName: "Report_Jan", address: "A1:C3" }]
  });
  assert(right.ok && right.snapshot?.snapshotId, "second snapshot create should pass");

  assert((await callTool(client, "excel.snapshot.get_compact", { snapshotId: left.snapshot.snapshotId })).ok, "snapshot get should pass");
  assert((await callTool(client, "excel.snapshot.list", { workbookId })).ok, "snapshot list should pass");
  assert((await callTool(client, "excel.snapshot.compare_compact", { leftSnapshotId: left.snapshot.snapshotId, rightSnapshotId: right.snapshot.snapshotId })).ok, "snapshot compare should pass");
  assert((await callTool(client, "excel.diff.create", { leftSnapshotId: left.snapshot.snapshotId, rightSnapshotId: right.snapshot.snapshotId })).ok, "diff create should pass");
  assert((await callTool(client, "excel.diff.get_compact", { leftSnapshotId: left.snapshot.snapshotId, rightSnapshotId: right.snapshot.snapshotId })).ok, "diff compact get should pass");

  assert((await callTool(client, "excel.events.subscribe", {})).ok, "events subscribe should pass");
  assert((await callTool(client, "excel.events.get_recent", { limit: 20 })).ok, "events get recent should pass");
  assert((await callTool(client, "excel.events.set_debounce", { debounceMs: 100 })).ok, "events debounce should pass");
  assert((await callTool(client, "excel.events.clear", {})).ok, "events clear should pass");
  assert((await callTool(client, "excel.events.unsubscribe", {})).ok, "events unsubscribe should pass");

  assert((await callTool(client, "excel.permissions.get", {})).ok, "permissions get should pass");
  assert((await callTool(client, "excel.permissions.set", { allowWrites: true, allowDestructiveActions: true, scope: { workbookId } })).ok, "permissions set should pass");
  assert((await callTool(client, "excel.permissions.require_confirmation", { levels: ["workbook"] })).ok, "permission confirmation config should pass");
  assert((await callTool(client, "excel.permissions.set_scope", { workbookId, sheetNames: ["Report_Jan"] })).ok, "permission scope set should pass");
  assert((await callTool(client, "excel.permissions.allow_destructive_actions", { allow: true })).ok, "permission destructive toggle should pass");
  assert((await callTool(client, "excel.permissions.allow_macro_execution", { allow: false })).ok, "permission macro toggle should pass");
  assert((await callTool(client, "excel.permissions.lock_regions", { workbookId, regions: [{ regionName: "InputRegion", reason: "E2E lock" }] })).ok, "permission lock regions should pass");
  assert((await callTool(client, "excel.permissions.unlock_regions", { workbookId, regionNames: ["InputRegion"] })).ok, "permission unlock regions should pass");
  assert((await callTool(client, "excel.permissions.set", { allowWrites: true, allowDestructiveActions: true, requireConfirmationFor: [], scope: { workbookId } })).ok, "permissions reset should pass");
}

async function runMultiAgentWorkflow(client) {
  const lock = await callTool(client, "excel.lock.acquire", {
    workbookId,
    scopes: [{ type: "range", workbookId, sheetName: "Scratch", address: "A1:B3" }],
    mode: "write_values",
    reason: "E2E reserved range"
  });
  assert(lock.ok && lock.locks.length === 1, "manual lock should be acquired");

  const blocked = await callTool(client, "excel.range.write_values", {
    workbookId,
    sheetName: "Scratch",
    address: "A2:A2",
    values: [["Blocked"]]
  });
  assert(blocked.ok === false && blocked.error?.code === "LOCK_CONFLICT", "overlapping write should be blocked by lock conflict");

  const guidance = await callTool(client, "excel.conflict.get_guidance", { workbookId });
  assert(guidance.ok, "conflict guidance should be available");

  const released = await callTool(client, "excel.lock.release", { lockIds: [lock.locks[0].lockId] });
  assert(released.ok, "manual lock should release");

  const retry = await callTool(client, "excel.range.write_values", {
    workbookId,
    sheetName: "Scratch",
    address: "A2:A2",
    values: [["Retried"]]
  });
  assertApplied(retry, "post-lock retry");
}

async function callTool(client, name, args) {
  const startedAt = performance.now();
  const result = await client.request("tools/call", { name, arguments: args });
  const elapsedMs = Math.round(performance.now() - startedAt);
  if (result.isError) {
    throw new Error(`${name} returned MCP error: ${JSON.stringify(result)}`);
  }
  const text = result.content?.find((item) => item.type === "text")?.text;
  assert(text, `${name} returned no text content`);
  const parsed = JSON.parse(text);
  transcript.push({ tool: name, elapsedMs, telemetry: parsed.telemetry, args, result: parsed });
  return parsed;
}

function assertApplied(result, label) {
  assert(result.ok === true, `${label} should succeed: ${JSON.stringify(result.error ?? result)}`);
  assert(result.transactionId, `${label} should include transactionId`);
  assert(Array.isArray(result.backups), `${label} should include backups array`);
}

class McpClient {
  constructor(child) {
    this.child = child;
    this.nextId = 1;
    this.pending = new Map();
    this.buffer = Buffer.alloc(0);
    child.stdout.on("data", (chunk) => this.read(chunk));
    child.on("exit", (code, signal) => {
      for (const pending of this.pending.values()) {
        pending.reject(new Error(`MCP server exited code=${code} signal=${signal}`));
      }
      this.pending.clear();
    });
  }

  request(method, params) {
    const id = this.nextId++;
    this.write({ jsonrpc: "2.0", id, method, params });
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Timed out waiting for MCP method ${method}`));
      }, 30_000);
      this.pending.set(id, { resolve, reject, timeout });
    });
  }

  notify(method, params) {
    this.write({ jsonrpc: "2.0", method, params });
  }

  write(message) {
    this.child.stdin.write(`${JSON.stringify(message)}\n`);
  }

  read(chunk) {
    this.buffer = Buffer.concat([this.buffer, Buffer.from(chunk)]);
    while (this.buffer.length > 0) {
      const parsed = this.readFramed() ?? this.readLineDelimited();
      if (!parsed) {
        return;
      }
      this.handle(parsed);
    }
  }

  readFramed() {
    const marker = this.buffer.indexOf("\r\n\r\n");
    if (marker < 0) {
      return undefined;
    }
    const header = this.buffer.slice(0, marker).toString("utf8");
    const match = /^Content-Length:\s*(\d+)/im.exec(header);
    if (!match) {
      return undefined;
    }
    const length = Number(match[1]);
    const bodyStart = marker + 4;
    const bodyEnd = bodyStart + length;
    if (this.buffer.length < bodyEnd) {
      return undefined;
    }
    const body = this.buffer.slice(bodyStart, bodyEnd).toString("utf8");
    this.buffer = this.buffer.slice(bodyEnd);
    return JSON.parse(body);
  }

  readLineDelimited() {
    const newline = this.buffer.indexOf("\n");
    if (newline < 0) {
      return undefined;
    }
    const line = this.buffer.slice(0, newline).toString("utf8").trim();
    this.buffer = this.buffer.slice(newline + 1);
    if (!line) {
      return undefined;
    }
    return JSON.parse(line);
  }

  handle(message) {
    if (!("id" in message) || "method" in message) {
      return;
    }
    const pending = this.pending.get(message.id);
    if (!pending) {
      return;
    }
    clearTimeout(pending.timeout);
    this.pending.delete(message.id);
    if (message.error) {
      pending.reject(new Error(message.error.message));
    } else {
      pending.resolve(message.result);
    }
  }

  close() {
    this.child.stdin.destroy();
  }
}

class FakeAddin {
  static async connect(url, workbook) {
    const { WebSocket } = await import("../apps/backend/node_modules/ws/wrapper.mjs");
    const socket = new WebSocket(url);
    const addin = new FakeAddin(socket, workbook);
    await new Promise((resolve, reject) => {
      socket.once("open", resolve);
      socket.once("error", reject);
    });
    socket.on("message", (raw) => addin.onMessage(JSON.parse(String(raw))));
    await addin.waitForConnectionId();
    addin.sendNotification("addin.hello", {
      capabilities: {
        platform: "mac",
        officeVersion: "fake-e2e",
        apiSets: { ExcelApi: "1.16" },
        features: {
          ranges: "supported",
          tables: "supported",
          formulas: "supported",
          pivots: "partial",
          charts: "partial"
        }
      },
      activeWorkbook: workbook.ref
    });
    return addin;
  }

  constructor(socket, workbook) {
    this.socket = socket;
    this.workbook = workbook;
    this.connectionId = undefined;
    this.calls = [];
    this.connectedResolvers = [];
  }

  waitForConnectionId() {
    if (this.connectionId) {
      return Promise.resolve();
    }
    return new Promise((resolve) => this.connectedResolvers.push(resolve));
  }

  onMessage(message) {
    if (message.method === "backend.connected") {
      this.connectionId = message.params.connectionId;
      for (const resolve of this.connectedResolvers.splice(0)) {
        resolve();
      }
      return;
    }
    if (!("id" in message) || !message.method) {
      return;
    }
    Promise.resolve()
      .then(() => this.handleRequest(message.method, message.params ?? {}))
      .then((result) => {
        this.socket.send(JSON.stringify({ jsonrpc: "2.0", id: message.id, result }));
      })
      .catch((error) => {
        this.socket.send(JSON.stringify({
          jsonrpc: "2.0",
          id: message.id,
          error: { code: -32000, message: error instanceof Error ? error.message : String(error) }
        }));
      });
  }

  handleRequest(method, params) {
    this.calls.push({ method, params });
    switch (method) {
      case "runtime.ping":
        return { ok: true, at: params.at };
      case "runtime.get_active_context":
        return this.workbook.ref;
      case "workbook.get_info":
        return { ...this.workbook.ref, sheetCount: this.workbook.sheets.size };
      case "workbook.get_map":
        return this.workbook.getMap();
      case "workbook.snapshot_ranges":
        return this.workbook.snapshotRanges(params.ranges);
      case "operation.execute_batch":
        return this.workbook.executeBatch(params.request, params.compiled);
      case "template.capture":
        return this.workbook.captureTemplate(params);
      case "template.capture_sheet":
        return this.workbook.captureTemplateSheet(params);
      case "template.repair":
        return { ok: true, repaired: params.repair ?? ["styles", "formulas", "dataRegions"], warnings: [] };
      case "style.capture_fingerprint":
        return this.workbook.styleFingerprint(params);
      case "style.copy_dimensions":
        return { ok: true, copied: params.dimensions ?? [], warnings: [] };
      case "formula.read_patterns":
        return this.workbook.formulaPatterns(params);
      case "formula.copy_patterns":
        return this.workbook.copyFormulaPatterns(params);
      case "formula.fill_pattern":
        return this.workbook.fillFormulaPattern(params);
      case "formula.convert_to_values":
        return { ok: true, converted: true, warnings: [] };
      case "table.list":
        return { ok: true, tables: [...this.workbook.tables.values()].map((table) => table.info()) };
      case "table.get_info":
        return { ok: true, info: this.workbook.table(params.tableName).info() };
      case "table.read":
        return this.workbook.table(params.tableName).read(params);
      case "table.create":
        return this.workbook.createTable(params);
      case "table.reorder_columns":
        return this.workbook.table(params.tableName).reorder(params.columnOrder);
      case "table.apply_filters":
        return this.workbook.table(params.tableName).setFilters(params.filters);
      case "table.clear_filters":
        return this.workbook.table(params.tableName).setFilters([]);
      case "table.sort":
        return this.workbook.table(params.tableName).sort(params.fields);
      case "table.clear_sort":
        return this.workbook.table(params.tableName).sort([]);
      case "names.list":
        return { ok: true, names: [...this.workbook.names.values()] };
      case "names.get":
        return this.workbook.getName(params);
      case "names.create":
        return this.workbook.createName(params);
      case "names.update":
        return this.workbook.createName(params);
      case "names.delete":
        return this.workbook.deleteName(params);
      case "pivot.list":
        return { ok: true, pivots: [...this.workbook.pivots.values()] };
      case "pivot.get_info":
        return this.workbook.getPivot(params);
      case "pivot.create":
        return this.workbook.createPivot(params);
      case "pivot.refresh":
      case "pivot.refresh_all":
        return { ok: true, refreshed: true };
      case "pivot.copy_from_template":
        return { ok: true, warnings: [{ code: "FAKE_HOST_PARTIAL", message: "Fake host replays deterministic pivot dimensions only." }] };
      case "pivot.delete":
        return this.workbook.deletePivot(params);
      case "chart.list":
        return { ok: true, charts: [...this.workbook.charts.values()] };
      case "chart.get_info":
        return this.workbook.getChart(params);
      case "chart.create":
        return this.workbook.createChart(params);
      case "chart.update_data_source":
        return this.workbook.updateChart(params);
      case "chart.copy_from_template":
        return { ok: true, warnings: [] };
      case "chart.refresh":
        return this.workbook.getChart(params);
      case "chart.delete":
        return this.workbook.deleteChart(params);
      case "workbook.calculate":
      case "workbook.save":
        return { ok: true, workbookId: this.workbook.ref.workbookId };
      default:
        if (method.startsWith("pivot.") || method.startsWith("chart.")) {
          return { ok: true, capabilityStatus: "partial", warnings: [{ code: "FAKE_HOST_PARTIAL", message: `${method} is partial in fake host.` }] };
        }
        return { ok: true, capabilityStatus: "unsupported", warnings: [{ code: "FAKE_HOST_UNIMPLEMENTED", message: `${method} returned deterministic fake-host fallback.` }] };
    }
  }

  sendNotification(method, params) {
    this.socket.send(JSON.stringify({ jsonrpc: "2.0", method, params }));
  }

  close() {
    this.socket.close();
  }

  summary() {
    return {
      connectionId: this.connectionId,
      calls: this.calls.map((call) => call.method),
      workbook: this.workbook.getMap()
    };
  }
}

class FakeWorkbook {
  constructor(id) {
    this.ref = { workbookId: id, name: "E2E Fast.xlsx", path: path.join(tempRoot, "E2E Fast.xlsx"), platform: "mac" };
    this.sheets = new Map();
    this.tables = new Map();
    this.names = new Map();
    this.pivots = new Map();
    this.charts = new Map();
  }

  addSheet(name) {
    const sheet = new FakeSheet(this.ref.workbookId, name);
    this.sheets.set(name, sheet);
    return sheet;
  }

  sheet(name) {
    const sheet = this.sheets.get(name);
    if (!sheet) {
      throw new Error(`Unknown fake sheet: ${name}`);
    }
    return sheet;
  }

  table(name) {
    const table = this.tables.get(name);
    if (!table) {
      throw new Error(`Unknown fake table: ${name}`);
    }
    return table;
  }

  getMap() {
    return {
      workbook: this.ref,
      sheets: [...this.sheets.values()].map((sheet) => ({
        workbookId: this.ref.workbookId,
        worksheetId: `sheet_${sheet.name}`,
        name: sheet.name,
        usedRange: sheet.usedRange(),
        tables: [...this.tables.values()].filter((table) => table.sheetName === sheet.name).map((table) => table.info())
      }))
    };
  }

  snapshotRanges(ranges) {
    return {
      workbookId: this.ref.workbookId,
      capturedAt: fixedNow(),
      workbookFingerprint: {
        workbookId: this.ref.workbookId,
        workbookHash: stableHash(this.getMap()),
        structureHash: stableHash([...this.sheets.keys(), ...this.tables.keys()]),
        capturedAt: fixedNow()
      },
      rangeSnapshots: ranges.map((range) => this.sheet(range.sheetName).snapshot(range))
    };
  }

  executeBatch(request, compiled = {}) {
    const startedAt = performance.now();
    let cellsRead = 0;
    let cellsWritten = 0;
    let sheetsChanged = 0;
    const readData = [];
    for (const operation of request.operations) {
      switch (operation.kind) {
        case "range.read_full": {
          const snapshot = this.sheet(operation.target.sheetName).snapshot(operation.target);
          cellsRead += snapshot.fingerprint.cellCount;
          readData.push({ operationId: operation.operationId, snapshot });
          break;
        }
        case "range.write_values":
          cellsWritten += this.sheet(operation.target.sheetName).writeValues(operation.target.address, operation.values);
          break;
        case "range.write_formulas":
          cellsWritten += this.sheet(operation.target.sheetName).writeFormulas(operation.target.address, operation.formulas);
          break;
        case "range.write_number_formats":
          cellsWritten += this.sheet(operation.target.sheetName).writeNumberFormats(operation.target.address, operation.numberFormat);
          break;
        case "range.write_styles":
          cellsWritten += parseRange(operation.target.address).rowCount * parseRange(operation.target.address).columnCount;
          break;
        case "range.clear_values_keep_format":
        case "range.clear_values":
        case "range.clear":
          cellsWritten += this.sheet(operation.target.sheetName).clearValues(operation.target.address);
          break;
        case "range.copy":
          cellsWritten += this.copyRange(operation.source, operation.target);
          break;
        case "range.restore_snapshot":
          cellsWritten += this.sheet(operation.target.sheetName).restore(operation.target.address, operation.snapshot);
          break;
        case "sheet.create":
          this.addSheet(operation.sheetName);
          sheetsChanged += 1;
          break;
        case "sheet.rename": {
          const sheet = this.sheet(operation.sheetName);
          this.sheets.delete(operation.sheetName);
          sheet.name = operation.newSheetName;
          this.sheets.set(operation.newSheetName, sheet);
          sheetsChanged += 1;
          break;
        }
        case "template.create_sheet_from_template": {
          const sourceName = this.findTemplateSource(operation.templateId) ?? "Template";
          const source = this.sheet(sourceName);
          const created = this.addSheet(operation.newSheetName);
          const sourceRange = source.usedRange().address;
          created.restore(sourceRange, source.snapshot({ workbookId: this.ref.workbookId, sheetName: sourceName, address: sourceRange }));
          sheetsChanged += 1;
          break;
        }
        default:
          break;
      }
    }
    return {
      ok: true,
      rollbackAvailable: request.mode === "apply",
      backups: [],
      warnings: [],
      readData,
      diffSummary: {
        title: "Fake Excel batch applied",
        changedRanges: compiled.targetFingerprints?.map((fingerprint) => fingerprint.range) ?? [],
        cellsChanged: cellsWritten,
        formulasChanged: 0,
        stylesChanged: 0,
        tablesChanged: 0,
        sheetsChanged,
        destructiveLevel: compiled.destructiveLevel ?? "none"
      },
      telemetry: {
        durationMs: Math.round(performance.now() - startedAt),
        syncCount: 1,
        cellsRead,
        cellsWritten,
        rangeCount: request.operations.length,
        chunkCount: cellsWritten > 1000 ? Math.ceil(cellsWritten / 1000) : 1,
        engineName: "fake-excel-e2e",
        warningCount: 0
      }
    };
  }

  copyRange(source, target) {
    const snapshot = this.sheet(source.sheetName).snapshot(source);
    return this.sheet(target.sheetName).restore(target.address, snapshot);
  }

  createTable(request) {
    if (request.values) {
      this.sheet(request.sheetName).writeValues(request.address, request.values);
    }
    const tableName = request.tableName ?? `Table${this.tables.size + 1}`;
    const table = new FakeTable(this, tableName, request.sheetName, request.address, request.style);
    this.tables.set(tableName, table);
    return { ok: true, table: table.info() };
  }

  captureTemplate(request) {
    this.lastTemplateSource = request.sourceSheetName;
    return {
      workbookId: request.workbookId,
      sourceSheetName: request.sourceSheetName,
      dataRegions: request.dataRegions ?? [],
      fingerprintPayload: this.templateFingerprintPayload(request.sourceSheetName, request.dataRegions ?? [])
    };
  }

  captureTemplateSheet(request) {
    return {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      dataRegions: request.dataRegions ?? [],
      fingerprintPayload: this.templateFingerprintPayload(request.sheetName, request.dataRegions ?? [])
    };
  }

  findTemplateSource() {
    return this.lastTemplateSource;
  }

  templateFingerprintPayload(sheetName, dataRegions) {
    const sheet = this.sheet(sheetName);
    const { address, rowCount, columnCount } = sheet.usedRange();
    return {
      structure: { usedRange: { address, rowCount, columnCount }, dataRegions },
      formulas: this.formulaPatterns({ workbookId: this.ref.workbookId, sheetName }).formulas,
      styles: this.styleFingerprint({ workbookId: this.ref.workbookId, sheetName }).dimensions,
      filters: [...this.tables.values()].filter((table) => table.sheetName === sheetName).map((table) => table.filters),
      tables: [...this.tables.values()].filter((table) => table.sheetName === sheetName).map((table) => table.info()),
      printLayout: { orientation: "portrait", fakeHost: true }
    };
  }

  styleFingerprint(request) {
    const target = request.address ?? this.sheet(request.sheetName).usedRange().address;
    const parsed = parseRange(target);
    return {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      address: target,
      capturedAt: fixedNow(),
      rowCount: parsed.rowCount,
      columnCount: parsed.columnCount,
      truncated: false,
      dimensions: {
        fills: stableHash({ address: target, dimension: "fills" }),
        fonts: stableHash({ address: target, dimension: "fonts" }),
        numberFormats: stableHash({ address: target, dimension: "numberFormats" }),
        alignment: stableHash({ address: target, dimension: "alignment" })
      },
      warnings: []
    };
  }

  formulaPatterns(request) {
    const address = request.address ?? this.sheet(request.sheetName).usedRange().address;
    const snapshot = this.sheet(request.sheetName).snapshot({ workbookId: request.workbookId, sheetName: request.sheetName, address });
    const formulas = snapshot.formulas;
    const patternMatrix = formulas.map((row) => row.map((formula) => formula ? stableHash(formula) : null));
    const cells = [];
    for (let row = 0; row < formulas.length; row += 1) {
      for (let col = 0; col < formulas[row].length; col += 1) {
        const formula = formulas[row][col];
        if (formula) {
          cells.push({ rowIndex: row, columnIndex: col, formula, formulaR1C1: formula, patternHash: stableHash(formula) });
        }
      }
    }
    const range = parseRange(address);
    return {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      address,
      capturedAt: fixedNow(),
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      formulaCount: cells.length,
      formulas,
      formulasR1C1: formulas,
      patternMatrix,
      patterns: [...new Map(cells.map((cell) => [cell.patternHash, cell])).values()].map((cell) => ({
        patternHash: cell.patternHash,
        formulaR1C1: cell.formulaR1C1,
        count: cells.filter((candidate) => candidate.patternHash === cell.patternHash).length,
        cells: cells.filter((candidate) => candidate.patternHash === cell.patternHash).map((candidate) => ({ rowIndex: candidate.rowIndex, columnIndex: candidate.columnIndex }))
      })),
      cells,
      warnings: []
    };
  }

  copyFormulaPatterns(request) {
    const sourceAddress = request.sourceAddress ?? this.sheet(request.sourceSheetName).usedRange().address;
    const targetAddress = request.targetAddress ?? sourceAddress;
    const source = this.sheet(request.sourceSheetName).snapshot({ workbookId: request.workbookId, sheetName: request.sourceSheetName, address: sourceAddress });
    this.sheet(request.targetSheetName).writeFormulas(targetAddress, source.formulas);
    return { ok: true, copied: true, warnings: [] };
  }

  fillFormulaPattern(request) {
    const source = this.sheet(request.sheetName).snapshot({ workbookId: request.workbookId, sheetName: request.sheetName, address: request.sourceAddress });
    this.sheet(request.sheetName).writeFormulas(request.targetAddress, repeatFirstFormula(source.formulas, parseRange(request.targetAddress)));
    return { ok: true, filled: true, warnings: [] };
  }

  createName(request) {
    const name = {
      workbookId: request.workbookId,
      name: request.name,
      scope: request.sheetName ? "worksheet" : "workbook",
      sheetName: request.sheetName,
      type: request.reference ? "Range" : "Formula",
      formula: request.formula,
      comment: request.comment,
      visible: request.visible ?? true,
      address: request.reference
    };
    this.names.set(`${request.sheetName ?? ""}:${request.name}`, name);
    return { ok: true, name };
  }

  getName(request) {
    const name = this.names.get(`${request.sheetName ?? ""}:${request.name}`) ?? this.names.get(`:${request.name}`);
    return name ? { ok: true, name } : { ok: false, error: { code: "NOT_FOUND", message: `Name not found: ${request.name}` } };
  }

  deleteName(request) {
    this.names.delete(`${request.sheetName ?? ""}:${request.name}`);
    this.names.delete(`:${request.name}`);
    return { ok: true };
  }

  createPivot(request) {
    const pivot = {
      workbookId: request.workbookId,
      pivotTableName: request.pivotTableName,
      id: `pivot_${request.pivotTableName}`,
      sheetName: request.destinationSheetName,
      range: { address: request.destinationAddress, rowCount: 4, columnCount: 3 },
      source: request.sourceTableName ?? `${request.sourceSheetName}!${request.sourceAddress}`,
      sourceType: request.sourceTableName ? "table" : "range",
      rowHierarchies: (request.rowFields ?? []).map((name) => ({ name })),
      columnHierarchies: (request.columnFields ?? []).map((name) => ({ name })),
      filterHierarchies: (request.filterFields ?? []).map((name) => ({ name })),
      dataHierarchies: (request.dataFields ?? []).map((field) => ({ name: field.name ?? field.sourceFieldName, field: { name: field.sourceFieldName }, summarizeBy: field.summarizeBy, numberFormat: field.numberFormat })),
      hierarchies: ["Date", "Account", "Amount", "Region", "Status"].map((name) => ({ name })),
      layout: request.layout ?? {}
    };
    this.pivots.set(request.pivotTableName, pivot);
    return { ok: true, info: pivot, warnings: [] };
  }

  getPivot(request) {
    const info = this.pivots.get(request.pivotTableName);
    return info ? { ok: true, info } : { ok: false, error: { code: "NOT_FOUND", message: `Pivot not found: ${request.pivotTableName}` } };
  }

  deletePivot(request) {
    this.pivots.delete(request.pivotTableName);
    return { ok: true };
  }

  createChart(request) {
    const chartName = request.chartName ?? `Chart${this.charts.size + 1}`;
    const chart = {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      chartName,
      id: `chart_${chartName}`,
      chartType: request.chartType,
      title: request.title,
      style: request.style,
      plotBy: request.seriesBy ?? "Auto",
      sourceAddress: request.sourceAddress
    };
    this.charts.set(`${request.sheetName}:${chartName}`, chart);
    return { ok: true, info: chart, warnings: [] };
  }

  getChart(request) {
    const info = this.charts.get(`${request.sheetName}:${request.chartName}`);
    return info ? { ok: true, info } : { ok: false, error: { code: "NOT_FOUND", message: `Chart not found: ${request.chartName}` } };
  }

  updateChart(request) {
    const current = this.charts.get(`${request.sheetName}:${request.chartName}`) ?? {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      chartName: request.chartName,
      id: `chart_${request.chartName}`,
      chartType: "ColumnClustered"
    };
    current.sourceAddress = request.sourceAddress;
    current.plotBy = request.seriesBy ?? current.plotBy ?? "Auto";
    this.charts.set(`${request.sheetName}:${request.chartName}`, current);
    return { ok: true, info: current, warnings: [] };
  }

  deleteChart(request) {
    this.charts.delete(`${request.sheetName}:${request.chartName}`);
    return { ok: true };
  }
}

class FakeSheet {
  constructor(workbookId, name) {
    this.workbookId = workbookId;
    this.name = name;
    this.cells = new Map();
  }

  writeValues(address, values) {
    const range = parseRange(address);
    forEachMatrix(values, (value, row, col) => {
      this.cell(range.startRow + row, range.startCol + col).value = value;
    });
    return matrixCellCount(values);
  }

  writeFormulas(address, formulas) {
    const range = parseRange(address);
    forEachMatrix(formulas, (formula, row, col) => {
      this.cell(range.startRow + row, range.startCol + col).formula = formula;
    });
    return matrixCellCount(formulas);
  }

  writeNumberFormats(address, formats) {
    const range = parseRange(address);
    forEachMatrix(formats, (format, row, col) => {
      this.cell(range.startRow + row, range.startCol + col).numberFormat = format;
    });
    return matrixCellCount(formats);
  }

  clearValues(address) {
    const range = parseRange(address);
    let count = 0;
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.cell(row, col);
        cell.value = null;
        cell.formula = null;
        count += 1;
      }
    }
    return count;
  }

  restore(address, snapshot) {
    const range = parseRange(address);
    const values = snapshot.values ?? [];
    const formulas = snapshot.formulas ?? [];
    for (let row = 0; row < values.length; row += 1) {
      for (let col = 0; col < values[row].length; col += 1) {
        const cell = this.cell(range.startRow + row, range.startCol + col);
        cell.value = values[row][col];
        cell.formula = formulas[row]?.[col] ?? null;
      }
    }
    return matrixCellCount(values);
  }

  snapshot(rangeRef) {
    const range = parseRange(rangeRef.address);
    const values = [];
    const formulas = [];
    const numberFormat = [];
    const text = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const valueRow = [];
      const formulaRow = [];
      const formatRow = [];
      const textRow = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.cell(row, col);
        valueRow.push(cell.value);
        formulaRow.push(cell.formula);
        formatRow.push(cell.numberFormat);
        textRow.push(cell.formula ?? (cell.value === null || cell.value === undefined ? "" : String(cell.value)));
      }
      values.push(valueRow);
      formulas.push(formulaRow);
      numberFormat.push(formatRow);
      text.push(textRow);
    }
    return {
      fingerprint: {
        range: rangeRef,
        hash: stableHash({ values, formulas, numberFormat }),
        cellCount: range.rowCount * range.columnCount,
        capturedAt: fixedNow()
      },
      values,
      formulas,
      numberFormat,
      text
    };
  }

  usedRange() {
    let maxRow = 1;
    let maxCol = 1;
    for (const key of this.cells.keys()) {
      const [row, col] = key.split(":").map(Number);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }
    return { workbookId: this.workbookId, sheetName: this.name, address: `A1:${columnName(maxCol)}${maxRow}`, rowCount: maxRow, columnCount: maxCol };
  }

  cell(row, col) {
    const key = `${row}:${col}`;
    if (!this.cells.has(key)) {
      this.cells.set(key, { value: null, formula: null, numberFormat: "General" });
    }
    return this.cells.get(key);
  }
}

class FakeTable {
  constructor(workbook, tableName, sheetName, address, style = "TableStyleMedium2") {
    this.workbook = workbook;
    this.tableName = tableName;
    this.sheetName = sheetName;
    this.address = address;
    this.style = style;
    this.filters = [];
    this.sortFields = [];
  }

  info() {
    const range = parseRange(this.address);
    const headers = this.headers();
    return {
      workbookId: this.workbook.ref.workbookId,
      tableName: this.tableName,
      id: `table_${this.tableName}`,
      sheetName: this.sheetName,
      address: this.address,
      headerAddress: `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.startRow}`,
      rowCount: Math.max(0, range.rowCount - 1),
      columnCount: range.columnCount,
      columns: headers.map((name, index) => ({ id: index + 1, index, name })),
      style: this.style,
      showHeaders: true,
      showTotals: false,
      showFilterButton: true,
      filters: this.filters,
      sort: this.sortFields
    };
  }

  headers() {
    return this.workbook.sheet(this.sheetName).snapshot({
      workbookId: this.workbook.ref.workbookId,
      sheetName: this.sheetName,
      address: this.headerAddress()
    }).values[0].map(String);
  }

  headerAddress() {
    const range = parseRange(this.address);
    return `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.startRow}`;
  }

  read(request) {
    const info = this.info();
    const range = parseRange(this.address);
    const headers = [this.headers()];
    const columnIndexes = request.columns?.map((column) => typeof column === "number" ? column : info.columns.find((item) => item.name === column)?.index).filter((value) => value !== undefined) ?? info.columns.map((column) => column.index);
    const offset = request.rowOffset ?? 0;
    const limit = request.rowLimit ?? info.rowCount;
    const values = [];
    const sheet = this.workbook.sheet(this.sheetName);
    for (let rowOffset = offset; rowOffset < Math.min(info.rowCount, offset + limit); rowOffset += 1) {
      const rowNumber = range.startRow + 1 + rowOffset;
      values.push(columnIndexes.map((columnIndex) => sheet.cell(rowNumber, range.startCol + columnIndex).value));
    }
    return {
      info,
      headers: [columnIndexes.map((index) => headers[0][index])],
      values,
      rowOffset: offset,
      rowLimit: limit,
      rowCount: values.length,
      truncated: offset + values.length < info.rowCount,
      projectedColumns: columnIndexes.map((index) => info.columns[index])
    };
  }

  reorder(columnOrder) {
    const info = this.info();
    const indexes = columnOrder.map((column) => typeof column === "number" ? column : info.columns.find((item) => item.name === column)?.index);
    if (indexes.some((index) => index === undefined) || indexes.length !== info.columnCount) {
      return { ok: false, error: { code: "TABLE_COLUMN_MISMATCH", message: "Invalid column order.", retryable: false } };
    }
    const range = parseRange(this.address);
    const sheet = this.workbook.sheet(this.sheetName);
    const rows = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      rows.push(indexes.map((index) => sheet.cell(row, range.startCol + index).value));
    }
    sheet.writeValues(this.address, rows);
    return { ok: true, table: this.info() };
  }

  setFilters(filters) {
    this.filters = filters;
    return { ok: true, filters };
  }

  sort(fields) {
    this.sortFields = fields;
    return { ok: true, sort: fields };
  }
}

function createWorkbookFixture(id) {
  const workbook = new FakeWorkbook(id);
  const data = workbook.addSheet("Data");
  data.writeValues("A1:D5", [
    ["Date", "Account", "Amount", "Status"],
    ["2026-01-01", "A-100", 100, "Open"],
    ["2026-01-02", "A-200", 200, "Closed"],
    ["2026-01-03", "A-300", 300, "Open"],
    ["2026-01-04", "A-400", 400, "Open"]
  ]);
  workbook.addSheet("Template").writeValues("A1:C3", [["Metric", "Jan", "Feb"], ["Revenue", 1000, 1100], ["Cost", 400, 450]]);
  workbook.addSheet("Large");
  workbook.createTable({ workbookId: id, sheetName: "Data", address: "A1:D5", tableName: "Transactions", hasHeaders: true });
  return workbook;
}

function generateRows(count) {
  const rows = [];
  for (let index = 0; index < count; index += 1) {
    rows.push([
      `2026-01-${String((index % 28) + 1).padStart(2, "0")}`,
      `A-${String(index).padStart(5, "0")}`,
      index * 3,
      ["NA", "EU", "APAC"][index % 3],
      index % 4 === 0 ? "Closed" : "Open"
    ]);
  }
  return rows;
}

function parseRange(address) {
  const [start, end = start] = address.replace(/\$/g, "").split(":");
  const startCell = parseCell(start);
  const endCell = parseCell(end);
  return {
    startRow: startCell.row,
    startCol: startCell.col,
    endRow: endCell.row,
    endCol: endCell.col,
    rowCount: endCell.row - startCell.row + 1,
    columnCount: endCell.col - startCell.col + 1
  };
}

function parseCell(cell) {
  const match = /^([A-Z]+)(\d+)$/i.exec(cell);
  if (!match) {
    throw new Error(`Unsupported A1 cell: ${cell}`);
  }
  return { col: columnIndex(match[1].toUpperCase()), row: Number(match[2]) };
}

function columnIndex(name) {
  let value = 0;
  for (const char of name) {
    value = value * 26 + char.charCodeAt(0) - 64;
  }
  return value;
}

function columnName(index) {
  let value = "";
  let remaining = index;
  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    value = String.fromCharCode(65 + mod) + value;
    remaining = Math.floor((remaining - mod) / 26);
  }
  return value || "A";
}

function forEachMatrix(matrix, callback) {
  for (let row = 0; row < matrix.length; row += 1) {
    for (let col = 0; col < matrix[row].length; col += 1) {
      callback(matrix[row][col], row, col);
    }
  }
}

function matrixCellCount(matrix) {
  return matrix.reduce((total, row) => total + row.length, 0);
}

function repeatFirstFormula(formulas, range) {
  const formula = formulas.flat().find((value) => value) ?? null;
  return Array.from({ length: range.rowCount }, () => Array.from({ length: range.columnCount }, () => formula));
}

function stableHash(value) {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function fixedNow() {
  return "2026-01-01T00:00:00.000Z";
}

async function waitForHttp(url, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url);
      if (response.ok) {
        return;
      }
    } catch {
      // keep polling
    }
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function writeArtifact(name, value) {
  writeFileSync(path.join(artifactsDir, name), JSON.stringify(value, null, 2));
}

function summarizeTranscript(entries) {
  const ranked = [...entries].sort((left, right) => right.elapsedMs - left.elapsedMs);
  const byTool = new Map();
  for (const entry of entries) {
    const current = byTool.get(entry.tool) ?? { tool: entry.tool, calls: 0, elapsedMs: 0, cellsRead: 0, cellsWritten: 0, chunkCount: 0 };
    current.calls += 1;
    current.elapsedMs += entry.elapsedMs;
    current.cellsRead += Number(entry.telemetry?.cellsRead ?? 0);
    current.cellsWritten += writtenCellCount(entry.result ?? entry);
    current.chunkCount += Number(entry.telemetry?.chunkCount ?? 0);
    byTool.set(entry.tool, current);
  }
  return {
    toolCallCount: entries.length,
    slowestToolCalls: ranked.slice(0, 15).map(({ tool, elapsedMs, telemetry }) => ({ tool, elapsedMs, telemetry })),
    toolTotals: [...byTool.values()].sort((left, right) => right.elapsedMs - left.elapsedMs),
    helperCandidates: inferHelperCandidates(entries)
  };
}

function inferHelperCandidates(entries) {
  const candidates = [];
  for (const entry of entries) {
    const cellsRead = Number(entry.telemetry?.cellsRead ?? 0);
    const cellsWritten = writtenCellCount(entry.result ?? entry);
    if (cellsRead > 10_000) {
      candidates.push({ tool: entry.tool, reason: `large read touched ${cellsRead} cells; prefer projected reads or a narrower helper` });
    }
    if (cellsWritten > 10_000) {
      candidates.push({ tool: entry.tool, reason: `large write touched ${cellsWritten} cells; consider a table-native or structural helper` });
    }
  }
  return candidates;
}

function writtenCellCount(result) {
  return Number(result?.telemetry?.cellsWritten ?? result?.compactProof?.cellsChanged ?? 0);
}

function hasTruthyKey(value, key) {
  if (!value || typeof value !== "object") {
    return false;
  }
  if (value[key]) {
    return true;
  }
  if (Array.isArray(value)) {
    return value.some((item) => hasTruthyKey(item, key));
  }
  return Object.values(value).some((item) => hasTruthyKey(item, key));
}

await main();
