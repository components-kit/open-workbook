#!/usr/bin/env node
import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdtempSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-e2e-codex-"));
const artifactsDir = process.env.OPEN_WORKBOOK_E2E_AGENT_ARTIFACT_DIR ?? path.join(tempRoot, "artifacts");
mkdirSync(artifactsDir, { recursive: true });

const model = readArg("--model") ?? process.env.OPEN_WORKBOOK_E2E_CODEX_MODEL ?? "gpt-5.4-mini";
const reasoning = readArg("--reasoning") ?? process.env.OPEN_WORKBOOK_E2E_CODEX_REASONING ?? "low";
const timeoutMs = Number(readArg("--timeout-ms") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_TIMEOUT_MS ?? 180_000);
const repeat = Number(readArg("--repeat") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_REPEAT ?? 1);
const reportOnly = hasArg("--report-only") || process.env.OPEN_WORKBOOK_E2E_AGENT_REPORT_ONLY === "1";
const skillGuidance = loadSkillGuidance();

const discoveryTools = [
  "excel.workflow.prepare_session"
];

const selfPreparingWorkflowTools = new Set([
  "excel.workflow.create_formula_sheet",
  "excel.workflow.create_template_report",
  "excel.workflow.create_pivot_chart_summary",
  "excel.workflow.repair_formula_errors",
  "excel.workflow.preview_risky_edit"
]);

const coreScenarioIds = [
  "large-table-reorder",
  "large-table-filter-sort",
  "large-table-append",
  "large-table-update",
  "multi-agent-locking"
];

const qualityScenarioIds = [
  "create-sheet-range-formula",
  "formula-repair",
  "snapshots-diffs-rollback",
  "template-style-repair",
  "pivot-chart"
];

const selectedScenarioIds = parseScenarioSelection(readArg("--scenarios") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_SCENARIOS);

const globalForbidden = [
  {
    name: "no formula to values unless explicitly requested",
    match: (call) => call.name === "excel.formula.convert_to_values"
  },
  {
    name: "no sheet clear/delete for ordinary edits",
    match: (call) => ["excel.sheet.clear", "excel.sheet.delete"].includes(call.name)
  }
];

const scenarios = [
  {
    id: "create-sheet-range-formula",
    title: "Create sheet, update cells, formulas, and formats",
    task: "After the required discovery call, prefer excel.workflow.create_formula_sheet to create a new sheet named QuickEntry, write a compact 4 column input block with headers Account, Amount, Tax, Total, add formulas in Total for three rows, apply currency number formats to Amount, Tax, and Total, and validate formulas on that sheet. If using separate tools, complete the same sequence.",
    requiredAny: [["excel.workflow.create_formula_sheet", "excel.sheet.create"], ["excel.workflow.create_formula_sheet", "excel.range.write_values"], ["excel.workflow.create_formula_sheet", "excel.range.write_formulas"], ["excel.workflow.create_formula_sheet", "excel.range.write_number_formats", "excel.style.apply_style"], ["excel.workflow.create_formula_sheet", "excel.validate.formulas", "excel.formula.validate", "excel.validate.no_formula_errors", "excel.validate.sheet"]],
    forbidden: [
      { name: "cell-by-cell writes", match: hasPerCellWriteLoop }
    ]
  },
  {
    id: "large-table-reorder",
    title: "Reorder columns in a large table safely",
    task: "In the 5,000 row TransactionsLarge table, first inspect the table with excel.table.get_info or excel.workbook.get_workbook_map, then swap the Status column to the first position and keep all rows, formulas, filters, and style intact. Use excel.table.reorder_columns, not a manual column rewrite. After mutation, validate with excel.validate.tables or excel.workbook.get_workbook_map.",
    requiredTools: ["excel.table.reorder_columns"],
    requiredAny: [["excel.table.get_info", "excel.workbook.get_workbook_map", "excel.workflow.prepare_session"], ["excel.validate.tables", "excel.table.validate_against_template", "excel.workbook.get_workbook_map", "excel.table.get_info", "excel.workflow.prepare_session"]],
    forbidden: [
      { name: "delete or insert columns for table reordering", match: (call) => ["excel.range.delete_columns", "excel.range.insert_columns"].includes(call.name) },
      { name: "manual range value rewrite for large table reorder", match: (call) => call.name === "excel.range.write_values" },
      { name: "whole table read for reorder", match: isLargeFullRead }
    ]
  },
  {
    id: "large-table-filter-sort",
    title: "Filter and sort a large table",
    task: "First inspect TransactionsLarge with excel.table.get_info or excel.workbook.get_workbook_map. Then use excel.table.apply_filters to filter Status = Open and use excel.table.sort to sort Amount descending. Preserve existing table structure and avoid reading or rewriting the whole table. Validate the filter/table state afterward with excel.validate.filters, excel.validate.tables, excel.table.get_info, or excel.workbook.get_workbook_map.",
    requiredTools: ["excel.table.apply_filters"],
    requiredAny: [["excel.table.sort"], ["excel.validate.filters", "excel.validate.tables", "excel.table.get_info", "excel.workbook.get_workbook_map", "excel.workbook.list_open_workbooks"]],
    forbidden: [
      { name: "whole table read for filter/sort", match: isLargeFullRead },
      { name: "range rewrite for filter/sort", match: (call) => call.name === "excel.range.write_values" }
    ]
  },
  {
    id: "large-table-append",
    title: "Append rows to a structured table",
    task: "Inspect TransactionsLarge with excel.table.get_info or excel.workbook.get_workbook_map, then append 25 new June transaction rows to TransactionsLarge using excel.table.append_rows. Generate rows with the existing columns: Date values from 2026-06-01 onward, Account values A-6001 through A-6025, Amount numeric values, Region values such as North/South/East/West, and Status Open. Keep the table style, filters, totals, and formulas intact. Do not manually paste into a guessed range. After append, you must call excel.validate.tables before finishing; this validation call is required.",
    requiredTools: ["excel.table.append_rows"],
    requiredAny: [["excel.validate.tables", "excel.table.get_info", "excel.workbook.get_workbook_map", "excel.workflow.prepare_session"]],
    forbidden: [
      { name: "manual range append", match: (call) => call.name === "excel.range.write_values" },
      { name: "whole table read before append", match: isLargeFullRead }
    ]
  },
  {
    id: "large-table-update",
    title: "Update matching rows without rewriting the table",
    task: "For TransactionsLarge, inspect the table with excel.table.get_info or excel.workbook.get_workbook_map, locate only accounts A-1010 through A-1020 with bounded excel.table.read_compact or excel.range.search, and update those structured table rows to Status = Reviewed without rewriting the entire table. Validate afterward with excel.validate.tables, excel.table.get_info, or excel.workbook.get_workbook_map.",
    requiredTools: ["excel.table.update_rows"],
    requiredAny: [["excel.table.read_compact", "excel.range.search"], ["excel.validate.tables", "excel.table.get_info", "excel.validate.filters", "excel.workbook.get_workbook_map", "excel.workflow.prepare_session"]],
    forbidden: [
      { name: "whole table rewrite for row update", match: (call) => call.name === "excel.range.write_values" },
      { name: "whole table read for targeted update", match: isLargeFullRead }
    ]
  },
  {
    id: "clean-import-range",
    title: "Clean imported data",
    task: "Clean ImportRaw!A1:F250: detect the header row, normalize headers, trim whitespace, parse numbers and dates, remove duplicate rows, and flag likely outliers. Use cleaning tools instead of hand-editing cells.",
    requiredTools: ["excel.clean.detect_header_row", "excel.clean.normalize_headers", "excel.clean.trim_whitespace", "excel.clean.remove_duplicates"],
    requiredAny: [["excel.clean.parse_numbers", "excel.clean.parse_dates"], ["excel.clean.detect_outliers", "excel.clean.fuzzy_match"], ["excel.validate.sheet", "excel.validate.no_unintended_changes"]],
    forbidden: [
      { name: "manual cell writes for cleaning", match: hasPerCellWriteLoop }
    ]
  },
  {
    id: "formula-repair",
    title: "Repair formula errors using patterns",
    task: "Report_Jan has broken formulas in D2:D20. Prefer excel.workflow.repair_formula_errors with errorAddress D2:D20, sourceAddress D2:D2, targetAddress D2:D20, and direction down. Do the full repair sequence, not diagnosis only: find the formula errors, read neighboring formula patterns, inspect dependency graph or precedents, repair formulas with the workflow, excel.formula.repair_patterns, excel.formula.fill_down, or excel.range.write_formulas, then validate there are no formula errors. Do not convert formulas to static values.",
    requiredAny: [["excel.workflow.repair_formula_errors", "excel.formula.find_errors", "excel.range.find_errors", "excel.repair.formula_errors"], ["excel.workflow.repair_formula_errors", "excel.formula.read_patterns"], ["excel.workflow.repair_formula_errors", "excel.formula.get_dependency_graph", "excel.formula.trace_precedents"], ["excel.workflow.repair_formula_errors", "excel.formula.repair_patterns", "excel.formula.fill_down", "excel.range.write_formulas", "excel.repair.formula_errors"], ["excel.workflow.repair_formula_errors", "excel.validate.no_formula_errors", "excel.formula.validate", "excel.validate.workbook"]],
    forbidden: [
      { name: "formula conversion instead of repair", match: (call) => call.name === "excel.formula.convert_to_values" }
    ]
  },
  {
    id: "template-style-repair",
    title: "Create from template and repair styles",
    task: "Create Report_Feb from the workbook's monthly template using excel.workflow.create_template_report if available, not a generic sheet copy. The workflow must clear only declared data regions, fill declared regions with the new period values, compare style consistency against the template, repair style drift with style repair/copy tools, and validate the sheet against the template. If using separate tools, complete the same sequence.",
    requiredAny: [["excel.workflow.create_template_report", "excel.template.list", "excel.template.detect_templates"], ["excel.workflow.create_template_report", "excel.template.create_sheet_from_template"], ["excel.workflow.create_template_report", "excel.template.clear_data_regions", "excel.range.clear_values_keep_format"], ["excel.workflow.create_template_report", "excel.template.fill_regions", "excel.region.fill"], ["excel.workflow.create_template_report", "excel.style.compare_fingerprint", "excel.style.validate_consistency"], ["excel.workflow.create_template_report", "excel.style.repair_consistency", "excel.style.copy_from_template"], ["excel.workflow.create_template_report", "excel.template.validate_sheet_against_template"]],
    forbidden: [
      { name: "whole sheet clear for template data reset", match: (call) => call.name === "excel.sheet.clear" },
      { name: "manual style repaint loop", match: hasPerCellStyleLoop }
    ]
  },
  {
    id: "names-regions",
    title: "Named range and region workflow",
    task: "Create a workbook name Inputs_CurrentMonth for Report_Jan!B2:C12, register that same block as an input region, fill it with a 2D values matrix, then list names and regions to verify the setup.",
    requiredTools: ["excel.names.create", "excel.region.register", "excel.region.fill"],
    requiredAny: [["excel.names.list", "excel.names.get"], ["excel.region.list", "excel.region.get"]],
    forbidden: [
      { name: "manual range write instead of region fill", match: (call) => call.name === "excel.range.write_values" }
    ]
  },
  {
    id: "pivot-chart",
    title: "Create summary pivot and chart",
    task: "After the required discovery call, prefer excel.workflow.create_pivot_chart_summary to validate capability, create a pivot table summarizing Amount by Status and Region, refresh it, create a chart from that summary on Report_Jan, refresh or update the chart data source, and validate the pivot source. If using separate tools, complete the same sequence. If a capability is partial, report the warning honestly and use the supported fallback.",
    requiredAny: [["excel.workflow.create_pivot_chart_summary", "excel.pivot.create"], ["excel.workflow.create_pivot_chart_summary", "excel.chart.create"], ["excel.workflow.create_pivot_chart_summary", "excel.pivot.refresh", "excel.pivot.refresh_all"], ["excel.workflow.create_pivot_chart_summary", "excel.chart.refresh", "excel.chart.update_data_source"], ["excel.workflow.create_pivot_chart_summary", "excel.pivot.validate_source", "excel.pivot.get_capability_matrix"]],
    forbidden: [
      { name: "range-only pivot replacement", match: (call) => call.name === "excel.range.write_values" }
    ]
  },
  {
    id: "snapshots-diffs-rollback",
    title: "Snapshots, diffs, backup, and rollback preview",
    task: "After the required discovery calls, use excel.workflow.preview_risky_edit with a non-empty scoped write operation for the risky edit to Report_Jan!B2:C12, or manually create a before snapshot or backup, make a small scoped value change, create an after snapshot, call a diff tool with the two snapshot IDs, then call a rollback preview tool. Do not set apply=false unless only a preview was requested. Do not stop after the write or plan. Do not actually roll back.",
    requiredAny: [["excel.workflow.preview_risky_edit", "excel.snapshot.create", "excel.workbook.snapshot", "excel.workbook.create_backup", "excel.backup.create_file"], ["excel.workflow.preview_risky_edit", "excel.range.write_values", "excel.region.write_values", "excel.region.fill"], ["excel.workflow.preview_risky_edit", "excel.diff.create", "excel.diff.summarize", "excel.snapshot.compare_compact"], ["excel.workflow.preview_risky_edit", "excel.transaction.preview_rollback", "excel.transaction.preview_rollback_chain"]],
    forbidden: [
      { name: "actual rollback when only preview requested", match: (call) => ["excel.transaction.rollback", "excel.transaction.rollback_chain", "excel.workbook.restore_backup", "excel.backup.restore_file"].includes(call.name) }
    ]
  },
  {
    id: "events-permissions",
    title: "Events and permissions controls",
    task: "Set a scoped write permission for Report_Jan, require confirmation for workbook-level destructive changes, subscribe to workbook events with a debounce, read recent events, then unsubscribe. Do not perform any destructive workbook operation.",
    requiredTools: ["excel.permissions.get", "excel.permissions.set_scope", "excel.permissions.require_confirmation", "excel.events.subscribe", "excel.events.get_recent", "excel.events.unsubscribe"],
    forbidden: [
      { name: "destructive operation during permission/event audit", match: (call) => ["excel.sheet.delete", "excel.sheet.clear", "excel.range.delete_rows", "excel.range.delete_columns"].includes(call.name) }
    ]
  },
  {
    id: "multi-agent-locking",
    title: "Multi-agent lock and conflict flow",
    task: "Acquire a write lock for Report_Jan!A1:C20 for a planned update, check collaboration status and conflict guidance, then release the lock. Do not write values in this scenario.",
    requiredTools: ["excel.lock.acquire", "excel.lock.release"],
    requiredAny: [["excel.collab.get_status", "excel.collab.list_locks"], ["excel.conflict.get_guidance", "excel.collab.get_status"]],
    forbidden: [
      { name: "write while lock scenario only asks for coordination", match: (call) => call.name.startsWith("excel.range.write_") || call.name.startsWith("excel.table.update") || call.name.startsWith("excel.region.write") || call.name === "excel.region.fill" }
    ]
  }
];

async function main() {
  const selectedScenarios = selectedScenarioIds.length > 0 ? scenarios.filter((scenario) => selectedScenarioIds.includes(scenario.id)) : scenarios;
  if (selectedScenarios.length === 0) {
    fail(`No Codex E2E scenarios matched OPEN_WORKBOOK_E2E_AGENT_SCENARIOS=${process.env.OPEN_WORKBOOK_E2E_AGENT_SCENARIOS}`);
  }

  const codex = spawnSync("codex", ["--version"], { encoding: "utf8" });
  if (codex.status !== 0) {
    fail("Codex CLI is not available. Install/sign in to Codex before running test:e2e:agent.", codex);
  }

  for (const required of ["packages/cli/dist/index.js", "apps/mcp-server/dist/index.js"]) {
    if (!existsSync(path.join(repoRoot, required))) {
      fail(`Required build artifact is missing: ${required}. Run \`corepack pnpm build\` before test:e2e:agent.`);
    }
  }

  const suiteStarted = performance.now();
  const results = [];
  for (let iteration = 1; iteration <= repeat; iteration += 1) {
    for (const scenario of selectedScenarios) {
      results.push(await runScenario(scenario, iteration));
    }
  }

  const suite = {
    ok: results.every((result) => result.ok),
    model,
    reasoning,
    repeat,
    reportOnly,
    selectedScenarioIds: selectedScenarios.map((scenario) => scenario.id),
    elapsedMs: Math.round(performance.now() - suiteStarted),
    artifactDir: artifactsDir,
    failureCategoryCounts: countFailureCategories(results),
    results
  };
  writeFileSync(path.join(artifactsDir, "codex-agent-suite.json"), JSON.stringify(suite, null, 2));
  writeFileSync(path.join(artifactsDir, "codex-agent-report.md"), renderSuiteMarkdown(suite));

  if (!suite.ok) {
    console.error(renderSuiteMarkdown(suite));
    if (reportOnly) {
      console.error(`Codex agent decision E2E completed in report-only mode with failures. Artifacts: ${artifactsDir}`);
      return;
    }
    console.error(`Codex agent decision E2E failed. Artifacts: ${artifactsDir}`);
    process.exit(1);
  }

  console.log(renderSuiteMarkdown(suite));
  console.log(`Codex agent decision E2E passed. Artifacts: ${artifactsDir}`);
}

async function runScenario(scenario, iteration) {
  const scenarioSlug = `${String(iteration).padStart(2, "0")}-${scenario.id}`;
  const scenarioDir = path.join(artifactsDir, scenarioSlug);
  mkdirSync(scenarioDir, { recursive: true });

  const backendPort = 38380 + Math.floor(Math.random() * 1500);
  const backendUrl = `http://127.0.0.1:${backendPort}`;
  const backendWsUrl = `ws://127.0.0.1:${backendPort}/addin`;
  const prompt = buildPrompt(scenario);
  writeFileSync(path.join(scenarioDir, "prompt.txt"), prompt);

  const codexProcess = spawn(
    "codex",
    [
      "exec",
      "--json",
      "-m",
      model,
      "-c",
      `model_reasoning_effort="${reasoning}"`,
      "-c",
      'sandbox_mode="read-only"',
      "-c",
      'mcp_servers.open-workbook.command="node"',
      "-c",
      'mcp_servers.open-workbook.args=["packages/cli/dist/index.js","mcp","--standalone","--agent-name","e2e-codex"]',
      "-c",
      `mcp_servers.open-workbook.env={OPEN_WORKBOOK_HOST="127.0.0.1",OPEN_WORKBOOK_PORT="${backendPort}",OPEN_WORKBOOK_ADDIN_PATH="/addin",OPEN_WORKBOOK_MCP_SURFACE="advanced",OPEN_WORKBOOK_DISABLE_UPDATE_CHECK="1",OPEN_WORKBOOK_STATE_DIR="${path.join(tempRoot, scenarioSlug, "state")}",OPEN_WORKBOOK_BACKUP_DIR="${path.join(tempRoot, scenarioSlug, "backups")}"}`,
      "-c",
      "mcp_servers.open-workbook.required=true",
      prompt
    ],
    {
      cwd: repoRoot,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
      env: {
        ...process.env,
        OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1"
      }
    }
  );
  const exitPromise = waitForExit(codexProcess);

  let stdout = "";
  let stderr = "";
  codexProcess.stdout.on("data", (chunk) => {
    stdout += String(chunk);
  });
  codexProcess.stderr.on("data", (chunk) => {
    stderr += String(chunk);
  });

  let fakeAddin;
  const started = performance.now();
  let exitStatus = undefined;
  let error = undefined;
  try {
    exitStatus = await waitForHttpOrEarlyExit(`${backendUrl}/status`, 60_000, exitPromise);
    if (exitStatus !== undefined) {
      throw new Error(`Codex exited before Open Workbook backend became ready with status ${exitStatus}.`);
    }
    fakeAddin = await MinimalFakeAddin.connect(backendWsUrl, scenario.id);
    exitStatus = await waitForExitWithTimeout(exitPromise, timeoutMs);
  } catch (caught) {
    error = caught;
  } finally {
    fakeAddin?.close();
    codexProcess.kill();
  }

  writeFileSync(path.join(scenarioDir, "codex-agent.jsonl"), stdout);
  writeFileSync(path.join(scenarioDir, "codex-agent.stderr.log"), stderr);
  writeFileSync(path.join(scenarioDir, "fake-addin.json"), JSON.stringify(fakeAddin?.summary() ?? null, null, 2));

  const events = parseJsonLines(stdout);
  const calls = extractMcpCalls(events);
  const analysis = analyzeScenario(scenario, calls, exitStatus, error);
  const failureCategories = failureCategoriesForChecks(analysis.checks);
  const result = {
    scenarioId: scenario.id,
    title: scenario.title,
    iteration,
    ok: analysis.ok,
    elapsedMs: Math.round(performance.now() - started),
    exitStatus,
    error: error instanceof Error ? { message: error.message, stack: error.stack } : error ? String(error) : undefined,
    toolCalls: calls,
    uniqueTools: [...new Set(calls.map((call) => call.name))],
    checks: analysis.checks,
    failureCategories,
    artifactDir: scenarioDir
  };
  writeFileSync(path.join(scenarioDir, "analysis.json"), JSON.stringify(result, null, 2));
  writeFileSync(path.join(scenarioDir, "analysis.md"), renderScenarioMarkdown(result));
  return result;
}

function buildPrompt(scenario) {
  return [
    "You are running an Open Workbook MCP agent decision E2E scenario.",
    "Use only the configured Open Workbook MCP server. Do not run shell commands. Do not edit repository files.",
    "",
    "Start the workbook session before any mutation with this MCP call:",
    ...discoveryTools.map((tool, index) => `${index + 1}. ${tool}`),
    "",
    "Choose the fastest reliable Open Workbook MCP tool for the task. Prefer table, template, style, formula, region, workflow, clean, pivot, chart, snapshot, diff, permission, lock, and transaction tools over manual range rewrites when they match the task.",
    "For large tables, use structured table operations and bounded reads. Never loop cell by cell. Preserve formulas, filters, styles, templates, and rollback metadata.",
    "If a mutating MCP call is cancelled in this noninteractive run, do not retry by changing permissions. Instead create an excel.plan.create draft that contains the exact safe operations you intended to apply, then report that mutation was blocked.",
    "Use the specific diagnostic tool when the task names one: formula error finding should use excel.formula.find_errors, excel.range.find_errors, or excel.repair.formula_errors; formula repair should use pattern/dependency tools; diff summaries should use excel.diff.summarize, excel.diff.get_compact, or excel.snapshot.compare_compact.",
    "For snapshot/diff risky edits, after excel.workflow.prepare_session prefer excel.workflow.preview_risky_edit with one minimal scoped operation and default apply behavior because it combines scoped apply, before/after snapshots, diff, and rollback preview. Do not write a large null-padded matrix when one cell or a smaller rectangle is enough. If you do not use it, capture before and after states, then call a diff tool before rollback preview.",
    "For sheet/formula creation tasks, prefer excel.workflow.create_formula_sheet because it combines sheet creation, value writes, formula writes, number formats, and formula validation.",
    "For formula repair tasks, prefer excel.workflow.repair_formula_errors because it combines validation, pattern read, dependency graph inspection, scoped repair, and after validation.",
    "For pivot/chart summary tasks, prefer excel.workflow.create_pivot_chart_summary because it combines capability check, pivot create, pivot refresh, chart create, chart update/refresh, and pivot source validation.",
    "For template report tasks, prefer excel.workflow.create_template_report because it combines template sheet creation, declared region clear/fill, style comparison, style repair, and template validation.",
    "For quality workflow tasks, complete the post-action tool sequence even if an earlier tool returns ok: formula repair must include a repair tool, template work must fill/repair/validate, pivot/chart work must refresh and validate, and snapshot work must diff before rollback preview.",
    "After risky mutations, validate the affected area and mention backup, transaction, warning, diff, or rollback metadata that the MCP returns.",
    "",
    "Bundled Open Workbook Excel skill guidance:",
    skillGuidance,
    "",
    "Workbook fixture summary:",
    "- Workbook: Agent Decision E2E.xlsx",
    "- Sheets: Data, Large, Template, Report_Jan, ImportRaw",
    "- Large table: TransactionsLarge, 5,000 rows, columns Date, Account, Amount, Region, Status",
    "- Template sheet: Template with registered-looking monthly report structure",
    "- Existing report: Report_Jan with formulas and styles",
    "",
    `Scenario: ${scenario.title}`,
    `Task: ${scenario.task}`,
    "",
    "Return a concise JSON object with fields ok, toolsUsed, safetyNotes, validationNotes, and capabilityWarnings."
  ].join("\n");
}

function analyzeScenario(scenario, calls, exitStatus, error) {
  const checks = [];
  const toolNames = calls.map((call) => call.name);
  const uniqueTools = new Set(toolNames);

  checks.push(check(exitStatus === 0 && !error, "codex process completed", error instanceof Error ? error.message : undefined));

  const hasWorkbookIdentityDiscovery = ["excel.workflow.prepare_session", "excel.workbook.get_workbook_map", "excel.runtime.get_active_context", "excel.workbook.list_open_workbooks"].some((tool) => uniqueTools.has(tool));
  checks.push(check(hasWorkbookIdentityDiscovery, "required workbook identity discovery"));
  checks.push(check(["excel.workflow.prepare_session", "excel.runtime.get_status", "excel.runtime.get_active_context", "excel.runtime.get_capabilities", "excel.workbook.list_open_workbooks", "excel.workbook.get_workbook_map", "excel.collab.get_status"].some((tool) => uniqueTools.has(tool)), "required runtime discovery"));
  const firstMutationIndex = calls.findIndex((call) => isMutationTool(call.name));
  if (firstMutationIndex >= 0) {
    const beforeMutation = new Set(calls.slice(0, firstMutationIndex).map((call) => call.name));
    const firstMutation = calls[firstMutationIndex];
    const firstMutationHasInternalPreflight = selfPreparingWorkflowTools.has(firstMutation.name);
    checks.push(check(firstMutationHasInternalPreflight || ["excel.workflow.prepare_session", "excel.workbook.get_workbook_map", "excel.runtime.get_active_context", "excel.workbook.list_open_workbooks"].some((tool) => beforeMutation.has(tool)), "workbook identity discovery happened before first mutation"));
    checks.push(check(firstMutationHasInternalPreflight || ["excel.workflow.prepare_session", "excel.runtime.get_status", "excel.runtime.get_active_context", "excel.runtime.get_capabilities"].some((tool) => beforeMutation.has(tool)), "runtime discovery happened before first mutation"));
  }

  for (const tool of scenario.requiredTools ?? []) {
    checks.push(check(uniqueTools.has(tool), `required tool ${tool}`));
  }
  for (const group of scenario.requiredAny ?? []) {
    checks.push(check(group.some((tool) => uniqueTools.has(tool)), `one of required tools: ${group.join(", ")}`));
  }
  if (scenario.id === "snapshots-diffs-rollback" && uniqueTools.has("excel.workflow.preview_risky_edit")) {
    const validWorkflow = calls.some((call) =>
      call.name === "excel.workflow.preview_risky_edit" &&
      call.arguments?.apply !== false &&
      Array.isArray(call.arguments?.operations) &&
      call.arguments.operations.length > 0
    );
    checks.push(check(validWorkflow, "valid combined risky workflow call"));
  }

  for (const rule of [...globalForbidden, ...(scenario.forbidden ?? [])]) {
    const offending = calls.filter((call) => rule.match(call, calls));
    checks.push(check(offending.length === 0, `forbidden behavior: ${rule.name}`, offending.map((call) => call.name).join(", ")));
  }

  checks.push(check(!hasPerCellWriteLoop(undefined, calls), "no repeated single-cell write loop"));
  checks.push(check(!hasSparseNullPaddedWrite(calls), "no sparse/null-padded broad range write"));

  return { ok: checks.every((item) => item.ok), checks };
}

function check(ok, label, details) {
  return { ok, label, details };
}

function failureCategoriesForChecks(checks) {
  return [...new Set(checks.filter((item) => !item.ok).map(categorizeCheck))].sort();
}

function countFailureCategories(results) {
  const counts = {};
  for (const result of results) {
    for (const category of result.failureCategories ?? []) {
      counts[category] = (counts[category] ?? 0) + 1;
    }
  }
  return counts;
}

function categorizeCheck(checkResult) {
  const label = checkResult.label;
  const details = checkResult.details ?? "";
  const text = `${label} ${details}`;
  if (label === "codex process completed") {
    if (/before Open Workbook backend|Timed out waiting for http|Operation not permitted|Codex exited/.test(details)) {
      return "startup_error";
    }
    return "process_error";
  }
  if (/discovery/.test(label)) {
    return "missing_discovery";
  }
  if (/forbidden behavior/.test(label)) {
    if (/cell|loop/.test(text)) {
      return "per_cell_loop";
    }
    if (/read|rewrite|append|clear|delete|rollback|conversion/.test(text)) {
      return "unsafe_tool";
    }
    return "forbidden_tool";
  }
  if (/no repeated single-cell write loop/.test(label)) {
    return "per_cell_loop";
  }
  if (/formula\.repair|repair\.formula|fill_down|write_formulas/.test(text)) {
    return "missing_repair";
  }
  if (/validate|no_formula_errors|no_unintended_changes/.test(text)) {
    return "missing_validation";
  }
  if (/snapshot|diff/.test(text)) {
    return "missing_snapshot_diff";
  }
  if (/rollback.*preview|preview_rollback/.test(text)) {
    return "missing_rollback_preview";
  }
  if (/refresh|update_data_source/.test(text)) {
    return "missing_refresh";
  }
  if (/required tool|one of required tools/.test(label)) {
    return "missing_required_tool";
  }
  return "quality_gap";
}

function extractMcpCalls(events) {
  const calls = [];
  for (const [eventIndex, event] of events.entries()) {
    collectMcpCalls(event, calls, eventIndex, []);
  }
  return calls.filter((call, index) => {
    const previous = calls[index - 1];
    return !(previous && previous.name === call.name && JSON.stringify(previous.arguments) === JSON.stringify(call.arguments));
  });
}

function collectMcpCalls(value, calls, eventIndex, pathParts) {
  if (!value || typeof value !== "object") {
    return;
  }
  const call = candidateMcpCall(value);
  if (call) {
    calls.push({ ...call, eventIndex, path: pathParts.join(".") });
  }
  if (Array.isArray(value)) {
    value.forEach((item, index) => collectMcpCalls(item, calls, eventIndex, [...pathParts, String(index)]));
    return;
  }
  for (const [key, child] of Object.entries(value)) {
    collectMcpCalls(child, calls, eventIndex, [...pathParts, key]);
  }
}

function candidateMcpCall(value) {
  const type = typeof value.type === "string" ? value.type : "";
  const name = normalizeToolName(value.name ?? value.tool ?? value.tool_name ?? value.toolName ?? value.op ?? type);
  if (typeof name !== "string" || !name.startsWith("excel.")) {
    return undefined;
  }
  const looksLikeMcp = type.includes("mcp") || value.server === "open-workbook" || value.server_name === "open-workbook" || value.provider === "mcp" || "arguments" in value || "args" in value || "op" in value;
  if (!looksLikeMcp) {
    return undefined;
  }
  return {
    name,
    arguments: parseMaybeJson(value.arguments ?? value.args ?? value.input ?? plannedOperationArgs(value)),
    resultStatus: value.result?.isError === true ? "error" : value.error ? "error" : "ok"
  };
}

function normalizeToolName(value) {
  if (typeof value !== "string") {
    return undefined;
  }
  if (value.startsWith("excel.")) {
    return value;
  }
  const aliases = {
    "range.write": "excel.range.write_values",
    "range.format": "excel.range.write_number_formats",
    "style.apply": "excel.style.apply_style",
    "validate.formulas": "excel.formula.validate"
  };
  if (aliases[value]) {
    return aliases[value];
  }
  if (/^(runtime|workbook|backup|sheet|range|batch|workflow|plan|template|style|formula|table|filter|sort|pivot|chart|names|region|task|collab|lock|conflict|transaction|permissions|clean|validate|repair|snapshot|diff|events)\./.test(value)) {
    return `excel.${value}`;
  }
  return value;
}

function plannedOperationArgs(value) {
  const args = {};
  for (const [key, child] of Object.entries(value)) {
    if (["name", "tool", "tool_name", "toolName", "op", "type", "server", "server_name", "provider", "result", "error", "status"].includes(key)) {
      continue;
    }
    args[key] = child;
  }
  return args;
}

function parseJsonLines(text) {
  const events = [];
  for (const line of text.split(/\r?\n/)) {
    if (!line.trim()) {
      continue;
    }
    try {
      events.push(JSON.parse(line));
    } catch {
      events.push({ type: "unparsed", text: line });
    }
  }
  return events;
}

function parseMaybeJson(value) {
  if (typeof value !== "string") {
    return value ?? {};
  }
  try {
    return JSON.parse(value);
  } catch {
    return { raw: value };
  }
}

function hasPerCellWriteLoop(call, calls = []) {
  const writeCalls = calls.filter((item) => isTopLevelCall(item) && ["excel.range.write_values", "excel.range.write_formulas", "excel.range.write_styles"].includes(item.name));
  const singleCellWrites = writeCalls.filter((item) => cellCountFromArgs(item.arguments) === 1);
  if (call) {
    return isTopLevelCall(call) && ["excel.range.write_values", "excel.range.write_formulas", "excel.range.write_styles"].includes(call.name) && cellCountFromArgs(call.arguments) === 1 && singleCellWrites.length > 3;
  }
  return singleCellWrites.length > 3;
}

function hasPerCellStyleLoop(call, calls = []) {
  const styleCalls = calls.filter((item) => isTopLevelCall(item) && ["excel.range.write_styles", "excel.style.apply_style"].includes(item.name));
  const singleCellStyles = styleCalls.filter((item) => cellCountFromArgs(item.arguments) === 1);
  if (call) {
    return isTopLevelCall(call) && ["excel.range.write_styles", "excel.style.apply_style"].includes(call.name) && cellCountFromArgs(call.arguments) === 1 && singleCellStyles.length > 3;
  }
  return singleCellStyles.length > 3;
}

function isTopLevelCall(call) {
  return !call.path || call.path === "item";
}

function hasSparseNullPaddedWrite(calls = []) {
  return calls.some((call) => {
    if (call.name !== "excel.range.write_values") {
      return false;
    }
    const values = call.arguments?.values;
    if (!Array.isArray(values)) {
      return false;
    }
    const matrixCells = values.reduce((sum, row) => sum + (Array.isArray(row) ? row.length : 0), 0);
    const nonEmptyCells = values.reduce(
      (sum, row) => sum + (Array.isArray(row) ? row.filter((value) => value !== null && value !== undefined && value !== "").length : 0),
      0
    );
    const addressCells = typeof call.arguments?.address === "string" ? cellCountFromAddress(call.arguments.address) : undefined;
    const touchedCells = Math.max(matrixCells, addressCells ?? 0);
    return touchedCells >= 8 && nonEmptyCells > 0 && nonEmptyCells / touchedCells <= 0.25 && touchedCells - nonEmptyCells >= 4;
  });
}

function isLargeFullRead(call) {
  if (!["excel.range.read_full", "excel.range.read_values", "excel.table.read", "excel.table.read_compact"].includes(call.name)) {
    return false;
  }
  const args = call.arguments ?? {};
  if (call.name === "excel.table.read" && args.tableName === "TransactionsLarge" && !args.rowLimit) {
    return true;
  }
  if (call.name === "excel.table.read_compact" && args.tableName === "TransactionsLarge" && !args.maxRows && !args.budget?.maxRows) {
    return true;
  }
  return args.sheetName === "Large" && (!args.address || cellCountFromAddress(args.address) > 1000);
}

function isMutationTool(name) {
  if (name === "excel.workflow.preview_risky_edit") {
    return true;
  }
  return [
    ".write_",
    ".create",
    ".restore",
    ".append",
    ".update",
    ".clear",
    ".delete",
    ".insert",
    ".move",
    ".copy",
    ".fill",
    ".repair",
    ".set_",
    ".apply_",
    ".reorder",
    ".sort",
    ".refresh"
  ].some((marker) => name.includes(marker)) && !name.includes(".get_") && !name.includes(".list") && !name.includes(".read_") && !name.includes(".detect_") && !name.includes(".find_") && !name.includes(".compare");
}

function cellCountFromArgs(args) {
  if (!args || typeof args !== "object") {
    return undefined;
  }
  if (typeof args.address === "string") {
    return cellCountFromAddress(args.address);
  }
  if (Array.isArray(args.values)) {
    return args.values.reduce((sum, row) => sum + (Array.isArray(row) ? row.length : 1), 0);
  }
  if (Array.isArray(args.formulas)) {
    return args.formulas.reduce((sum, row) => sum + (Array.isArray(row) ? row.length : 1), 0);
  }
  return undefined;
}

function cellCountFromAddress(address) {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(address);
  if (!match) {
    return /^[A-Z]+\d+$/i.test(address) ? 1 : undefined;
  }
  const startColumn = columnToNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnToNumber(match[3]);
  const endRow = Number(match[4]);
  return Math.max(1, endColumn - startColumn + 1) * Math.max(1, endRow - startRow + 1);
}

function columnToNumber(column) {
  return column.toUpperCase().split("").reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

class MinimalFakeAddin {
  static async connect(url, scenarioId) {
    const { WebSocket } = await import("../apps/backend/node_modules/ws/wrapper.mjs");
    const socket = new WebSocket(url);
    const addin = new MinimalFakeAddin(socket, scenarioId);
    await new Promise((resolve, reject) => {
      socket.once("open", resolve);
      socket.once("error", reject);
    });
    socket.on("message", (raw) => addin.onMessage(JSON.parse(String(raw))));
    await addin.waitForConnectionId();
    addin.sendNotification("addin.hello", {
      capabilities: {
        platform: "mac",
        officeVersion: "fake-codex-e2e",
        apiSets: { ExcelApi: "1.16" },
        features: {
          ranges: "supported",
          tables: "supported",
          formulas: "supported",
          templates: "supported",
          styles: "supported",
          pivots: "partial",
          charts: "partial"
        }
      },
      activeWorkbook: addin.workbook
    });
    return addin;
  }

  constructor(socket, scenarioId) {
    this.socket = socket;
    this.scenarioId = scenarioId;
    this.workbook = { workbookId: "workbook_codex_e2e", name: "Agent Decision E2E.xlsx", platform: "mac" };
    this.connectionId = undefined;
    this.calls = [];
    this.connectedResolvers = [];
    this.tableColumns = ["Date", "Account", "Amount", "Region", "Status"];
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
    const result = this.handleRequest(message.method, message.params ?? {});
    this.socket.send(JSON.stringify({ jsonrpc: "2.0", id: message.id, result }));
  }

  handleRequest(method, params) {
    this.calls.push({ method, params });
    switch (method) {
      case "runtime.ping":
        return { ok: true, at: params.at };
      case "runtime.get_active_context":
        return this.workbook;
      case "workbook.get_info":
        return { ...this.workbook, sheetCount: 5 };
      case "workbook.get_map":
        return this.workbookMap();
      case "workbook.snapshot_ranges":
        return this.snapshotRanges(params.ranges ?? []);
      case "operation.execute_batch":
        return this.executeBatch(params.request, params.compiled);
      case "table.list":
        return { ok: true, tables: [this.tableInfo("TransactionsLarge")] };
      case "table.get_info":
        return { ok: true, info: this.tableInfo(params.tableName ?? "TransactionsLarge") };
      case "table.read":
        return this.tableRead(params);
      case "table.reorder_columns":
        this.tableColumns = params.columnOrder ?? this.tableColumns;
        return { ok: true, info: this.tableInfo(params.tableName ?? "TransactionsLarge") };
      case "table.append_rows":
      case "table.update_rows":
      case "table.apply_filters":
      case "table.clear_filters":
      case "table.sort":
      case "table.clear_sort":
      case "table.set_total_row":
      case "table.set_style":
      case "table.copy_structure":
      case "table.validate_against_template":
        return { ok: true, info: this.tableInfo(params.tableName ?? "TransactionsLarge"), warnings: [] };
      case "template.capture":
      case "template.capture_sheet":
        return {
          workbookId: this.workbook.workbookId,
          sourceSheetName: params.sourceSheetName ?? params.sheetName ?? "Template",
          dataRegions: params.dataRegions ?? ["A2:C12"],
          fingerprintPayload: { structureHash: stableHash(params), formulasHash: "formula_hash", stylesHash: "style_hash" }
        };
      case "template.repair":
        return { ok: true, repaired: ["styles", "formulas", "dataRegions"], warnings: [] };
      case "style.capture_fingerprint":
        return { ok: true, dimensions: params.dimensions ?? ["fills", "fonts", "numberFormats"], hash: stableHash(params), cellCount: cellCountFromAddress(params.address ?? "A1:C3") ?? 9 };
      case "style.copy_dimensions":
        return { ok: true, copied: params.dimensions ?? [], warnings: [] };
      case "formula.read_patterns":
        return {
          ok: true,
          patterns: {
            cells: [{ address: params.address ?? "D2:D20", formula: "=B2+C2", pattern: "=RC[-2]+RC[-1]" }]
          }
        };
      case "formula.copy_patterns":
      case "formula.fill_pattern":
      case "formula.recalculate":
        return { ok: true, warnings: [] };
      case "names.list":
        return { ok: true, names: [{ name: "Inputs_CurrentMonth", reference: "Report_Jan!B2:C12" }] };
      case "names.get":
      case "names.create":
      case "names.update":
      case "names.delete":
        return { ok: true, name: params.name, reference: params.reference ?? "Report_Jan!B2:C12" };
      case "pivot.list":
        return { ok: true, pivots: [{ pivotTableName: "PivotSales", sheetName: "Report_Jan", sourceTableName: "TransactionsLarge" }] };
      case "pivot.get_info":
      case "pivot.create":
      case "pivot.refresh":
      case "pivot.refresh_all":
      case "pivot.validate_source":
      case "pivot.get_fingerprint":
        return { ok: true, pivotTableName: params.pivotTableName ?? "PivotSales", capabilityStatus: "partial", warnings: [{ code: "FAKE_HOST_PARTIAL", message: "Fake host validates metadata paths only." }] };
      case "chart.list":
        return { ok: true, charts: [{ chartName: "SalesChart", sheetName: "Report_Jan" }] };
      case "chart.get_info":
      case "chart.create":
      case "chart.update_data_source":
      case "chart.refresh":
      case "chart.validate_against_template":
        return { ok: true, chartName: params.chartName ?? "SalesChart", warnings: [] };
      case "workbook.calculate":
      case "workbook.save":
        return { ok: true, workbookId: this.workbook.workbookId };
      default:
        return { ok: true, warnings: [] };
    }
  }

  workbookMap() {
    return {
      workbook: this.workbook,
      sheets: [
        { workbookId: this.workbook.workbookId, worksheetId: "sheet_Data", name: "Data", usedRange: range("Data", "A1:E20"), tables: [] },
        { workbookId: this.workbook.workbookId, worksheetId: "sheet_Large", name: "Large", usedRange: range("Large", "A1:E5001"), tables: [this.tableInfo("TransactionsLarge")] },
        { workbookId: this.workbook.workbookId, worksheetId: "sheet_Template", name: "Template", usedRange: range("Template", "A1:D20"), tables: [] },
        { workbookId: this.workbook.workbookId, worksheetId: "sheet_Report_Jan", name: "Report_Jan", usedRange: range("Report_Jan", "A1:H40"), tables: [] },
        { workbookId: this.workbook.workbookId, worksheetId: "sheet_ImportRaw", name: "ImportRaw", usedRange: range("ImportRaw", "A1:F250"), tables: [] }
      ]
    };
  }

  tableInfo(tableName) {
    return {
      workbookId: this.workbook.workbookId,
      tableName,
      sheetName: "Large",
      address: "A1:E5001",
      rowCount: 5000,
      columns: this.tableColumns.map((name, index) => ({ id: index + 1, index, name })),
      showHeaders: true,
      showTotals: false,
      showFilterButton: true,
      style: "TableStyleMedium2"
    };
  }

  tableRead(params) {
    const rowLimit = Math.min(params.rowLimit ?? 50, 50);
    const offset = params.rowOffset ?? 0;
    const columns = params.columns ?? this.tableColumns;
    return {
      ok: true,
      table: this.tableInfo(params.tableName ?? "TransactionsLarge"),
      values: Array.from({ length: rowLimit }, (_, row) => columns.map((column) => sampleTableValue(column, offset + row))),
      rowOffset: offset,
      rowLimit,
      truncated: offset + rowLimit < 5000
    };
  }

  executeBatch(request = {}, compiled = {}) {
    const operations = request.operations ?? [];
    const readData = [];
    let cellsRead = 0;
    let cellsWritten = 0;
    let sheetsChanged = 0;
    for (const operation of operations) {
      if (operation.kind?.startsWith("range.read")) {
        const snapshot = this.snapshotRanges([operation.target]).rangeSnapshots[0];
        readData.push({ operationId: operation.operationId, snapshot });
        cellsRead += snapshot.fingerprint.cellCount;
      }
      if (operation.kind?.startsWith("range.write") || operation.kind?.startsWith("range.clear")) {
        cellsWritten += cellCountFromAddress(operation.target?.address ?? "A1") ?? 1;
      }
      if (operation.kind?.startsWith("sheet.") || operation.kind?.startsWith("template.")) {
        sheetsChanged += 1;
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
        durationMs: 1,
        syncCount: 1,
        cellsRead,
        cellsWritten,
        rangeCount: operations.length,
        chunkCount: cellsWritten > 1000 ? Math.ceil(cellsWritten / 1000) : 1,
        engineName: "fake-codex-agent-e2e",
        warningCount: 0
      }
    };
  }

  snapshotRanges(ranges) {
    return {
      workbookId: this.workbook.workbookId,
      capturedAt: fixedNow(),
      workbookFingerprint: {
        workbookId: this.workbook.workbookId,
        workbookHash: stableHash(this.workbookMap()),
        structureHash: stableHash(this.tableColumns),
        capturedAt: fixedNow()
      },
      rangeSnapshots: ranges.map((target) => ({
        fingerprint: { range: target, hash: stableHash(target), cellCount: cellCountFromAddress(target.address ?? "A1") ?? 1, capturedAt: fixedNow() },
        values: [[null]],
        formulas: [[null]],
        numberFormat: [["General"]],
        text: [[""]]
      }))
    };
  }

  sendNotification(method, params) {
    this.socket.send(JSON.stringify({ jsonrpc: "2.0", method, params }));
  }

  close() {
    this.socket.close();
  }

  summary() {
    return { connectionId: this.connectionId, scenarioId: this.scenarioId, calls: this.calls };
  }
}

function sampleTableValue(column, index) {
  switch (column) {
    case "Date":
      return "2026-06-01";
    case "Account":
      return `A-${String(1000 + index).padStart(4, "0")}`;
    case "Amount":
      return 100 + index;
    case "Region":
      return ["North", "South", "East", "West"][index % 4];
    case "Status":
      return index % 2 === 0 ? "Open" : "Closed";
    default:
      return null;
  }
}

function range(sheetName, address) {
  return { workbookId: "workbook_codex_e2e", sheetName, address };
}

function stableHash(value) {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function fixedNow() {
  return "2026-01-01T00:00:00.000Z";
}

async function waitForHttp(url, waitMs) {
  const deadline = Date.now() + waitMs;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url);
      if (response.ok) {
        return;
      }
    } catch {
      // keep polling
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`Timed out waiting for ${url}`);
}

function waitForHttpOrEarlyExit(url, waitMs, exitPromise) {
  const deadline = Date.now() + waitMs;
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (callback) => {
      if (settled) {
        return;
      }
      settled = true;
      callback();
    };
    exitPromise.then((status) => finish(() => resolve(status)), (error) => finish(() => reject(error)));
    const poll = async () => {
      while (!settled && Date.now() < deadline) {
        try {
          const response = await fetch(url);
          if (response.ok) {
            finish(() => resolve(undefined));
            return;
          }
        } catch {
          // keep polling
        }
        if (!settled) {
          await new Promise((resolveDelay) => setTimeout(resolveDelay, 100));
        }
      }
      finish(() => reject(new Error(`Timed out waiting for ${url}`)));
    };
    void poll();
  });
}

function waitForExit(child) {
  return new Promise((resolve, reject) => {
    child.on("error", reject);
    child.on("exit", (status) => {
      resolve(status);
    });
  });
}

function waitForExitWithTimeout(exitPromise, waitMs) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error("Timed out waiting for Codex E2E process.")), waitMs);
    exitPromise.then(
      (status) => {
        clearTimeout(timeout);
        resolve(status);
      },
      (error) => {
        clearTimeout(timeout);
        reject(error);
      }
    );
  });
}

function parseScenarioSelection(value) {
  if (!value || value.trim() === "all") {
    return [];
  }
  if (value.trim() === "smoke" || value.trim() === "core") {
    return coreScenarioIds;
  }
  if (value.trim() === "quality") {
    return qualityScenarioIds;
  }
  return value.split(",").map((item) => item.trim()).filter(Boolean);
}

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function hasArg(name) {
  return process.argv.includes(name);
}

function loadSkillGuidance() {
  const files = [
    ["SKILL.md", "skills/open-workbook-excel/SKILL.md"],
    ["Tool Selection", "skills/open-workbook-excel/references/tool-selection.md"],
    ["Performance", "skills/open-workbook-excel/references/performance.md"],
    ["Multi-Agent", "skills/open-workbook-excel/references/multi-agent.md"],
    ["Reliability", "skills/open-workbook-excel/references/reliability.md"]
  ];
  return files
    .map(([label, relativePath]) => {
      const absolutePath = path.join(repoRoot, relativePath);
      const text = existsSync(absolutePath) ? stripFrontmatter(readFileSync(absolutePath, "utf8")).trim() : "";
      return `## ${label}\n${text}`;
    })
    .join("\n\n");
}

function stripFrontmatter(text) {
  return text.replace(/^---\n[\s\S]*?\n---\n?/, "");
}

function renderSuiteMarkdown(suite) {
  const passed = suite.results.filter((result) => result.ok).length;
  const lines = [
    "# Codex Agent Decision E2E Report",
    "",
    `- Model: ${suite.model}`,
    `- Reasoning: ${suite.reasoning}`,
    `- Repeat: ${suite.repeat}`,
    `- Report only: ${suite.reportOnly ? "yes" : "no"}`,
    `- Passed: ${passed}/${suite.results.length}`,
    `- Elapsed: ${suite.elapsedMs} ms`,
    `- Artifacts: ${suite.artifactDir}`,
    "",
    "## Failure Categories",
    ...(Object.keys(suite.failureCategoryCounts ?? {}).length > 0
      ? Object.entries(suite.failureCategoryCounts).sort(([left], [right]) => left.localeCompare(right)).map(([category, count]) => `- ${category}: ${count}`)
      : ["- none"]),
    "",
    "## Scenarios"
  ];
  for (const result of suite.results) {
    lines.push("");
    lines.push(`### ${result.ok ? "PASS" : "FAIL"} ${result.scenarioId}`);
    lines.push("");
    lines.push(`- Title: ${result.title}`);
    lines.push(`- Iteration: ${result.iteration}`);
    lines.push(`- Elapsed: ${result.elapsedMs} ms`);
    lines.push(`- Tools: ${result.uniqueTools.join(", ") || "none"}`);
    const failed = result.checks.filter((item) => !item.ok);
    if (result.failureCategories?.length) {
      lines.push(`- Failure categories: ${result.failureCategories.join(", ")}`);
    }
    if (failed.length > 0) {
      lines.push("- Failed checks:");
      for (const item of failed) {
        lines.push(`  - ${item.label}${item.details ? ` (${item.details})` : ""}`);
      }
    }
  }
  return lines.join("\n");
}

function renderScenarioMarkdown(result) {
  const lines = [
    `# ${result.scenarioId}`,
    "",
    `- Status: ${result.ok ? "PASS" : "FAIL"}`,
    `- Title: ${result.title}`,
    `- Iteration: ${result.iteration}`,
    `- Elapsed: ${result.elapsedMs} ms`,
    `- Exit status: ${result.exitStatus}`,
    "",
    "## Tool Calls",
    "",
    ...(result.toolCalls.length > 0 ? result.toolCalls.map((call, index) => `${index + 1}. ${call.name}`) : ["No MCP tool calls found."]),
    "",
    "## Checks"
  ];
  for (const item of result.checks) {
    lines.push(`- ${item.ok ? "PASS" : "FAIL"} ${item.label}${item.details ? `: ${item.details}` : ""}`);
  }
  if (result.failureCategories?.length) {
    lines.push("");
    lines.push("## Failure Categories");
    for (const category of result.failureCategories) {
      lines.push(`- ${category}`);
    }
  }
  return lines.join("\n");
}

function fail(message, result) {
  if (result?.stdout) {
    console.error(result.stdout);
  }
  if (result?.stderr) {
    console.error(result.stderr);
  }
  console.error(message);
  process.exit(1);
}

await main();
