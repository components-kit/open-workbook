#!/usr/bin/env node
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-office-agent-behavior-"));
const backendStateDir = path.join(tempRoot, "state");
const runId = new Date().toISOString().replace(/[:.]/g, "-");
const defaultRepoArtifactsDir = path.join(repoRoot, ".open-workbook", "office-agent-behavior", runId, "artifacts");
const artifactsDir = readArg("--artifact-dir") ?? process.env.OPEN_WORKBOOK_OFFICE_AGENT_BEHAVIOR_DIR ?? defaultRepoArtifactsDir;
mkdirSync(artifactsDir, { recursive: true });
const productionScenariosPath = path.join(repoRoot, "tests", "e2e", "fixtures", "office-agent-production-scenarios.json");

const workbookId = "workbook_office_behavior";
const backendPort = Number(readArg("--port") ?? process.env.OPEN_WORKBOOK_OFFICE_AGENT_BEHAVIOR_PORT ?? 39880 + Math.floor(Math.random() * 300));
const backendUrl = `http://127.0.0.1:${backendPort}`;
const backendWsUrl = `ws://127.0.0.1:${backendPort}/addin`;
const transcript = [];
const strictMode = hasArg("--strict") || process.env.OPEN_WORKBOOK_OFFICE_AGENT_BEHAVIOR_STRICT === "1";
const estimatedUsdPerMillionTokens = Number(process.env.OPEN_WORKBOOK_AGENT_E2E_USD_PER_MILLION_TOKENS ?? "0.06");
const hardMaxCompletedTaskCalls = Number(process.env.OPEN_WORKBOOK_AGENT_E2E_HARD_MAX_CALLS ?? "5");

async function main() {
  const scenarioCatalog = loadScenarioCatalog();
  const selection = selectScenarios(scenarioCatalog);
  const selected = selection.scenarios;
  const allowDestructiveActions = selected.some((scenario) => scenario.expected?.allowDestructiveActions === true);
  const server = spawn(process.execPath, ["apps/mcp-server/dist/index.js", "--standalone", "--agent-name", "office-agent-behavior"], {
    cwd: repoRoot,
    stdio: ["pipe", "pipe", "pipe"],
    env: {
      ...process.env,
      OPEN_WORKBOOK_HOST: "127.0.0.1",
      OPEN_WORKBOOK_PORT: String(backendPort),
      OPEN_WORKBOOK_ADDIN_PATH: "/addin",
      OPEN_WORKBOOK_STATE_DIR: backendStateDir,
      OPEN_WORKBOOK_BACKUP_DIR: path.join(tempRoot, "backups"),
      OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1",
      ...(allowDestructiveActions ? { OPEN_WORKBOOK_E2E_ALLOW_DESTRUCTIVE_ACTIONS: "1" } : {})
    }
  });
  const mcp = new McpClient(server);
  let addin;
  let serverStderr = "";
  server.stderr.on("data", (chunk) => {
    serverStderr += String(chunk);
  });

  try {
    await waitForHttp(`${backendUrl}/status`, 15_000);
    addin = await FakeAddin.connect(backendWsUrl, createWorkbookFixture(workbookId));
    await mcp.request("initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "open-workbook-office-agent-behavior", version: "0.0.0" }
    });
    mcp.notify("notifications/initialized", {});

    const listed = await mcp.request("tools/list", {});
    const toolNames = listed.tools.map((tool) => tool.name);
    const agentOutputSchema = listed.tools.find((tool) => tool.name === "excel.agent.run")?.outputSchema;
    const results = [];
    for (const scenario of selected) {
      results.push(await runScenario({ mcp, addin, scenario, agentOutputSchema }));
    }

    const report = buildReport({ toolNames, results, artifactDir: artifactsDir, scenarioSource: scenarioCatalog.source, selection });
    writeFileSync(path.join(artifactsDir, "office-agent-behavior-report.json"), JSON.stringify(report, null, 2));
    writeFileSync(path.join(artifactsDir, "office-agent-behavior-report.md"), renderReport(report));
    writeFileSync(path.join(artifactsDir, "mcp-transcript.jsonl"), transcript.map((event) => JSON.stringify(event)).join("\n"));
    console.log(renderReport(report));
    console.log(`\nSaved office agent behavior artifacts: ${artifactsDir}`);
    if (strictMode) {
      const failures = strictFailures(report);
      if (failures.length > 0) {
        writeFileSync(path.join(artifactsDir, "strict-failures.json"), JSON.stringify(failures, null, 2));
        throw new Error(`Strict MCP scenario runner failed ${failures.length}/${report.scenarioCount} scenario(s). See ${path.join(artifactsDir, "strict-failures.json")}`);
      }
    }
  } catch (error) {
    writeFileSync(path.join(artifactsDir, "office-agent-behavior-failure.json"), JSON.stringify({
      error: error instanceof Error ? { message: error.message, stack: error.stack } : String(error),
      serverStderr,
      transcript,
      fakeAddin: addin?.summary()
    }, null, 2));
    throw error;
  } finally {
    addin?.close();
    mcp.close();
    server.kill();
  }
}

function strictFailures(report) {
  return report.results
    .filter((result) => result.expectationIssues.length > 0 || result.notes.some((note) => note.startsWith("Tool call threw:")))
    .map((result) => ({
      scenarioId: result.scenarioId,
      category: result.category,
      status: result.status,
      nextAction: result.nextAction,
      expectationIssues: result.expectationIssues,
      notes: result.notes,
      artifactDir: result.artifactDir
    }));
}

async function runScenario({ mcp, addin, scenario, agentOutputSchema }) {
  const scenarioDir = path.join(artifactsDir, scenario.id);
  mkdirSync(scenarioDir, { recursive: true });
  addin.resetWorkbook(createWorkbookFixture(workbookId));
  const restoreWorkbookAfterScenario = scenario.fixture?.excelState === "noActiveWorkbook" || scenario.id === "closed-workbook-status-error";
  if (restoreWorkbookAfterScenario) {
    addin.setActiveWorkbook(undefined);
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  const before = addin.workbook.summary();
  addin.setSelection(scenario.selection ?? { sheetName: "Sales", address: "A1" });
  const started = performance.now();
  const chainStart = transcript.length;
  const hostCallStart = addin.calls.length;
  const stepResults = [];
  const stepContext = {};
  let error;
  try {
    for (const step of scenarioSteps(scenario)) {
      const input = normalizeAgentInput(step.input, stepContext);
      const stepResult = await callTool(mcp, "excel.agent.run", input, agentOutputSchema);
      const record = { label: step.label, input, expected: step.expected, result: stepResult };
      rememberStepContext(stepContext, step.label, stepResult);
      if (stepResult.status === "PREVIEW_READY" && stepResult.nextAction === "call_apply_update" && stepResult.operationId && stepResult.confirmationToken && shouldAutoApplyForStudy(scenario)) {
        record.applied = await callTool(mcp, "excel.agent.run", {
          request: `Apply preview for ${scenario.prompt}`,
          mode: "apply_update",
          operationId: stepResult.operationId,
          confirmationToken: stepResult.confirmationToken
        }, agentOutputSchema);
        rememberStepContext(stepContext, `${step.label}:applied`, record.applied);
      }
      stepResults.push(record);
    }
  } catch (caught) {
    error = caught;
  }
  const after = addin.workbook.summary();
  const hostCalls = addin.calls.slice(hostCallStart);
  const toolChain = transcript.slice(chainStart);
  const result = { steps: stepResults };
  const observation = observeScenario({ scenario, result, error, before, after, hostCalls, workbook: addin.workbook, toolChain, elapsedMs: Math.round(performance.now() - started) });
  if (scenario.expected?.xlsxAssertions === true) {
    const artifact = writeAndAssertWorkbookArtifact(addin.workbook, scenario, scenarioDir);
    observation.workbookArtifact = artifact.path;
    observation.workbookArtifactAssertions = artifact.assertions;
    for (const issue of artifact.issues) {
      observation.expectationIssues.push(issue);
      observation.notes.push(`Expectation: ${issue}`);
    }
  }
  writeFileSync(path.join(scenarioDir, "prompt.txt"), scenario.prompt);
  writeFileSync(path.join(scenarioDir, "tool-input.json"), JSON.stringify(scenario.steps ?? scenario.input, null, 2));
  writeFileSync(path.join(scenarioDir, "tool-calls.json"), JSON.stringify(toolChain.map((call) => ({ tool: call.tool, args: call.args, status: call.status, nextAction: call.nextAction, summary: call.summary })), null, 2));
  writeFileSync(path.join(scenarioDir, "tool-results.json"), JSON.stringify(stepResults, null, 2));
  writeFileSync(path.join(scenarioDir, "tool-result.json"), JSON.stringify(result ?? null, null, 2));
  writeFileSync(path.join(scenarioDir, "workbook-before.json"), JSON.stringify(before, null, 2));
  writeFileSync(path.join(scenarioDir, "workbook-after.json"), JSON.stringify(after, null, 2));
  writeFileSync(path.join(scenarioDir, "behavior-notes.md"), renderObservation(observation));
  if (restoreWorkbookAfterScenario) {
    addin.setActiveWorkbook(addin.workbook.ref);
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  return observation;
}

function shouldAutoApplyForStudy(scenario) {
  return scenario.expected?.autoApply === true || ["mock-data-blank-sheet", "expense-tracker", "summary-sheet", "add-notes-block", "office-multi-step-summary"].includes(scenario.id);
}

function scenarioSteps(scenario) {
  if (Array.isArray(scenario.steps)) {
    return scenario.steps;
  }
  return [{ label: "request", input: scenario.input }];
}

function loadScenarioCatalog() {
  const scenarioFile = readArg("--scenario-file") ?? process.env.OPEN_WORKBOOK_OFFICE_AGENT_SCENARIO_FILE ?? productionScenariosPath;
  const loaded = JSON.parse(readFileSync(scenarioFile, "utf8"));
  if (!Array.isArray(loaded)) {
    throw new Error(`Scenario file must contain a JSON array: ${scenarioFile}`);
  }
  loaded.source = scenarioFile;
  return loaded;
}

function selectScenarios(scenarioCatalog) {
  const categorySelectors = splitSelector(readArg("--category") ?? readArg("--categories"));
  const scenarioSelectors = splitSelector(readArg("--scenarios") ?? (categorySelectors.length > 0 ? "" : "all"));
  const fullSuite = scenarioSelectors.includes("all") && categorySelectors.length === 0;
  const scenarios = fullSuite
    ? scenarioCatalog
    : scenarioCatalog.filter((scenario) =>
      scenarioSelectors.includes(scenario.id) ||
      scenarioSelectors.includes(scenario.category) ||
      categorySelectors.includes(scenario.category)
    );
  if (scenarios.length === 0) {
    const availableCategories = [...new Set(scenarioCatalog.map((scenario) => scenario.category))].sort();
    throw new Error([
      "No office-agent behavior scenarios matched the requested selector.",
      `--scenarios: ${scenarioSelectors.join(", ") || "none"}`,
      `--category: ${categorySelectors.join(", ") || "none"}`,
      `Available categories: ${availableCategories.join(", ")}`
    ].join("\n"));
  }
  return {
    mode: fullSuite ? "full-suite" : categorySelectors.length > 0 ? "category" : "scenario",
    scenarioSelectors,
    categorySelectors,
    scenarioCount: scenarios.length,
    scenarios
  };
}

function splitSelector(raw) {
  if (typeof raw !== "string") return [];
  return raw.split(",").map((item) => item.trim()).filter(Boolean);
}

function normalizeAgentInput(rawInput = {}, stepContext = {}) {
  const input = JSON.parse(JSON.stringify(rawInput));
  if (input.contextId && !input.workbookContextId) {
    input.workbookContextId = resolveToken(input.contextId, stepContext);
    delete input.contextId;
  }
  if (input.previewId) {
    input.operationId = resolveToken(input.previewId, stepContext);
    delete input.previewId;
  }
  if (input.changeId) {
    input.transactionId = resolveToken(input.changeId, stepContext);
    delete input.changeId;
  }
  if (input.mode === "apply") {
    input.mode = "apply_update";
    const preview = stepContext.lastPreview;
    if (!input.operationId && preview?.operationId) input.operationId = preview.operationId;
    if (!input.confirmationToken && preview?.confirmationToken) input.confirmationToken = preview.confirmationToken;
  }
  if (input.mode === "rollback") {
    const apply = stepContext.lastApply;
    if (!input.operationId && apply?.operationId) input.operationId = apply.operationId;
    if (!input.transactionId && apply?.answer?.transactionId) input.transactionId = apply.answer.transactionId;
  }
  if (!input.request) {
    input.request = "Run Open Workbook agent workflow.";
  }
  return resolveTokensDeep(input, stepContext);
}

function resolveToken(value, stepContext) {
  if (typeof value !== "string" || !value.startsWith("$steps.")) {
    return value;
  }
  const [, label, field] = /^\$steps\.([^.]+)\.(.+)$/.exec(value) ?? [];
  const result = label ? stepContext.byLabel?.[label] : undefined;
  if (!result) {
    return undefined;
  }
  if (field === "contextId") {
    return result.workbookContextId ?? result.contextId;
  }
  if (field === "previewId") {
    return result.operationId;
  }
  if (field === "changeId") {
    return result.answer?.transactionId ?? result.operationId;
  }
  return valueAtPath(result, field);
}

function resolveTokensDeep(value, stepContext) {
  if (typeof value === "string") {
    return resolveToken(value, stepContext);
  }
  if (Array.isArray(value)) {
    return value.map((item) => resolveTokensDeep(item, stepContext));
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, resolveTokensDeep(item, stepContext)]));
  }
  return value;
}

function valueAtPath(value, pathExpression) {
  return String(pathExpression)
    .split(".")
    .reduce((current, part) => current?.[part], value);
}

function rememberStepContext(stepContext, label, result) {
  stepContext.byLabel ??= {};
  const latestTemplate = latestRuntimeTemplate();
  const enriched = latestTemplate?.templateId && result && typeof result === "object"
    ? { ...result, templateId: latestTemplate.templateId }
    : result;
  stepContext.byLabel[label] = enriched;
  if (latestTemplate?.templateId) {
    stepContext.lastTemplateId = latestTemplate.templateId;
  }
  if (result?.workbookContextId) {
    stepContext.workbookContextId = result.workbookContextId;
  }
  if (result?.status === "PREVIEW_READY") {
    stepContext.lastPreview = result;
  }
  if (result?.mode === "apply_update" || result?.answer?.kind === "apply_update_result") {
    stepContext.lastApply = result;
  }
}

function latestRuntimeTemplate() {
  try {
    const state = JSON.parse(readFileSync(path.join(backendStateDir, "collaboration-state.json"), "utf8"));
    return Array.isArray(state.templates) ? state.templates.at(-1) : undefined;
  } catch {
    return undefined;
  }
}

async function callTool(client, name, args, outputSchema) {
  const started = performance.now();
  const raw = await client.request("tools/call", { name, arguments: args });
  const wallMs = Math.round(performance.now() - started);
  const text = raw.content?.find((item) => item.type === "text")?.text;
  const parsed = raw.structuredContent;
  const event = {
    at: new Date().toISOString(),
    tool: name,
    args,
    wallMs,
    isError: raw.isError === true,
    text,
    telemetry: parsed?.telemetry,
    status: parsed?.status,
    nextAction: parsed?.nextAction,
    taskOutcome: parsed?.taskOutcome,
    maxRecommendedFollowupCalls: parsed?.maxRecommendedFollowupCalls,
    summary: parsed?.summary
  };
  if (outputSchema?.properties?.telemetry?.properties) {
    const schemaKeys = outputSchema.properties.telemetry.properties;
    event.undeclaredTelemetryKeys = Object.keys(parsed?.telemetry ?? {}).filter((key) => !(key in schemaKeys));
  }
  transcript.push(event);
  return parsed;
}

function observeScenario({ scenario, result, error, before, after, hostCalls, workbook, toolChain, elapsedMs }) {
  const effective = effectiveScenarioResult(result);
  const preview = result?.steps?.some((step) => step.result?.status === "PREVIEW_READY");
  const workbookChanged = stableHash(before) !== stableHash(after);
  const usage = summarizeToolChainUsage(toolChain, elapsedMs);
  const signals = [
    effective?.status ? `status:${effective.status}` : undefined,
    effective?.nextAction ? `nextAction:${effective.nextAction}` : undefined,
    effective?.telemetry?.internalReadCount !== undefined ? `reads:${effective.telemetry.internalReadCount}` : undefined,
    effective?.telemetry?.fullReadCellCount !== undefined ? `cellsRead:${effective.telemetry.fullReadCellCount}` : undefined,
    effective?.telemetry?.metadataCacheStatus ? `metadata:${effective.telemetry.metadataCacheStatus}` : undefined,
    preview ? "previewed" : undefined,
    workbookChanged ? "workbookChanged" : "workbookUnchanged"
  ].filter(Boolean);
  const notes = [];
  if (error) notes.push(`Tool call threw: ${error instanceof Error ? error.message : String(error)}`);
  if (effective?.status === "AMBIGUOUS_TARGET") notes.push("The agent surface returned candidates instead of choosing a target.");
  if (effective?.status === "NEEDS_INPUT") notes.push("The agent surface needed more input or manual review.");
  if (effective?.status === "VALIDATION_FAILED") notes.push("The agent surface blocked the request during validation.");
  if (effective?.nextAction === "manual_review") notes.push("A human or higher-level workflow would need to decide the next step.");
  if (scenario.category.includes("comparison") && effective?.status === "SUCCESS" && !/compare|difference|higher|lower|month/i.test(JSON.stringify(effective.answer ?? {}))) {
    notes.push("Comparison prompt did not produce an obviously comparative structured answer.");
  }
  if (scenario.category.includes("formula") && effective?.status === "VALIDATION_FAILED") {
    notes.push("Formula task reached a safety boundary in the generic value workflow.");
  }
  if (scenario.category.includes("simple edit") && !workbookChanged && effective?.status === "SUCCESS") {
    notes.push("Edit scenario reported success but workbook summary did not change.");
  }
  const expectationIssues = evaluateScenarioExpectations({ scenario, result, effective, workbookChanged, usage, hostCalls, workbook });
  for (const issue of expectationIssues) {
    notes.push(`Expectation: ${issue}`);
  }
  return {
    scenarioId: scenario.id,
    category: scenario.category,
    prompt: scenario.prompt,
    elapsedMs,
    status: effective?.status ?? "ERROR",
    nextAction: effective?.nextAction,
    workbookChanged,
    signals,
    notes,
    expectationIssues,
    expected: scenario.expected,
    budgets: scenario.budgets,
    toolChain: toolChain.map((call, index) => ({
      index: index + 1,
      tool: call.tool,
      mode: call.args?.mode ?? "auto",
      request: call.args?.request,
      wallMs: call.wallMs,
      status: call.status,
      nextAction: call.nextAction,
      taskOutcome: call.taskOutcome,
      maxRecommendedFollowupCalls: call.maxRecommendedFollowupCalls,
      summary: call.summary,
      telemetry: call.telemetry
    })),
    hostCalls: hostCalls.map((call) => ({
      method: call.method,
      operationKinds: call.params?.request?.operations?.map((operation) => operation.kind),
      tableName: call.params?.tableName
    })),
    usage,
    resultSummary: effective?.summary,
    artifactDir: path.join(artifactsDir, scenario.id)
  };
}

function summarizeToolChainUsage(toolChain, elapsedMs) {
  const telemetryItems = toolChain.map((call) => call.telemetry ?? {});
  const sum = (field) => telemetryItems.reduce((total, telemetry) => total + (Number(telemetry[field]) || 0), 0);
  const estimatedTokens = sum("estimatedTokens");
  const metadataCacheStatuses = countBy(telemetryItems, (telemetry) => telemetry.metadataCacheStatus ?? "unknown");
  const followupCounts = toolChain
    .map((call) => Number(call.maxRecommendedFollowupCalls))
    .filter((value) => Number.isFinite(value));
  return {
    toolCallCount: toolChain.length,
    modelCallCount: toolChain.length,
    scenarioElapsedMs: elapsedMs,
    toolWallMs: toolChain.reduce((total, call) => total + (Number(call.wallMs) || 0), 0),
    backendElapsedMs: sum("elapsedMs"),
    payloadBytes: sum("payloadBytes"),
    estimatedTokens,
    estimatedCostUsd: Number(((estimatedTokens / 1_000_000) * estimatedUsdPerMillionTokens).toFixed(8)),
    costModel: {
      source: "estimated_tokens_proxy",
      usdPerMillionTokens: estimatedUsdPerMillionTokens
    },
    estimatedTokensSaved: sum("estimatedTokensSaved"),
    internalCallCount: sum("internalCallCount"),
    internalReadCount: sum("internalReadCount"),
    fullReadCellCount: sum("fullReadCellCount"),
    cacheHits: telemetryItems.filter((telemetry) => telemetry.cacheHit === true).length,
    autoAppliedCount: telemetryItems.filter((telemetry) => telemetry.autoApplied === true).length,
    taskOutcomes: countBy(toolChain, (call) => call.taskOutcome ?? "unknown"),
    maxRecommendedFollowupCalls: followupCounts.length > 0 ? Math.max(...followupCounts) : undefined,
    metadataCacheStatuses
  };
}

function effectiveScenarioResult(result) {
  const last = result?.steps?.at(-1);
  return last?.applied ?? last?.result ?? result;
}

function writeAndAssertWorkbookArtifact(workbook, scenario, scenarioDir) {
  const filePath = path.join(scenarioDir, "workbook-after.xlsx");
  writeFileSync(filePath, createZip(createWorkbookArtifactFiles(workbook)));
  const xlsx = readXlsx(filePath);
  const issues = assertWorkbookArtifact(xlsx, scenario.expected ?? {});
  const assertions = [
    "xlsx-zip-central-directory",
    "xlsx-workbook-parts",
    ...(scenario.expected?.cellValues ? ["xlsx-cell-values"] : []),
    ...(scenario.expected?.cellFormulas ? ["xlsx-cell-formulas"] : []),
    ...(scenario.expected?.cellNumberFormats ? ["xlsx-cell-number-formats"] : []),
    ...(scenario.expected?.cellStyles ? ["xlsx-cell-styles"] : []),
    ...(scenario.expected?.insertedColumns ? ["xlsx-inserted-columns"] : []),
    ...(scenario.expected?.validationRanges ? ["xlsx-data-validations"] : []),
    ...(scenario.expected?.conditionalFormatRanges ? ["xlsx-conditional-formatting"] : []),
    ...(scenario.expected?.tableColumnOrder ? ["xlsx-table-columns"] : []),
    ...(scenario.expected?.sheetProtections ? ["xlsx-sheet-protection"] : [])
  ];
  writeFileSync(path.join(scenarioDir, "workbook-artifact-report.json"), JSON.stringify({ path: filePath, assertions, issues, entries: xlsx.entries().sort() }, null, 2));
  return { path: filePath, assertions, issues };
}

function assertWorkbookArtifact(workbook, expected) {
  const issues = [];
  const entrySet = new Set(workbook.entries());
  for (const entry of ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/styles.xml"]) {
    if (!entrySet.has(entry)) issues.push(`expected workbook artifact entry ${entry}`);
  }
  let sheet = "";
  try {
    sheet = workbookSheetText(workbook, "Sales");
  } catch (error) {
    issues.push(error instanceof Error ? error.message : String(error));
  }
  if (expected.cellValues) {
    for (const expectedCell of expected.cellValues) {
      const cellSheet = workbookSheetText(workbook, expectedCell.sheetName);
      if (!cellHasValue(cellSheet, expectedCell.cell, expectedCell.value)) {
        issues.push(`expected workbook artifact ${expectedCell.sheetName}!${expectedCell.cell} to equal ${JSON.stringify(expectedCell.value)}`);
      }
    }
  }
  if (expected.cellFormulas) {
    for (const expectedCell of expected.cellFormulas) {
      const cellSheet = workbookSheetText(workbook, expectedCell.sheetName);
      if (!cellHasFormula(cellSheet, expectedCell.cell, expectedCell.formula)) {
        issues.push(`expected workbook artifact ${expectedCell.sheetName}!${expectedCell.cell} formula to equal ${JSON.stringify(expectedCell.formula)}`);
      }
    }
  }
  if (expected.cellNumberFormats) {
    const styles = workbook.entryText("xl/styles.xml");
    for (const expectedCell of expected.cellNumberFormats) {
      const cellSheet = workbookSheetText(workbook, expectedCell.sheetName);
      const styleId = cellStyleId(cellSheet, expectedCell.cell);
      if (styleId === undefined || styleId === "0") {
        issues.push(`expected workbook artifact ${expectedCell.sheetName}!${expectedCell.cell} to have a number-format style`);
      }
      if (!styles.includes(`formatCode="${escapeXml(expectedCell.numberFormat)}"`)) {
        issues.push(`expected workbook artifact styles to include number format ${expectedCell.numberFormat}`);
      }
    }
  }
  if (expected.cellStyles) {
    const styles = workbook.entryText("xl/styles.xml");
    for (const expectedStyle of expected.cellStyles) {
      const styleSheet = workbookSheetText(workbook, expectedStyle.sheetName);
      const styleId = cellStyleId(styleSheet, expectedStyle.cell);
      if (styleId === undefined || styleId === "0") {
        issues.push(`expected workbook artifact ${expectedStyle.sheetName}!${expectedStyle.cell} to have a non-default style`);
      }
      const style = expectedStyle.style ?? {};
      if (style.fillColor && !styles.includes(`<fgColor rgb="${argb(style.fillColor)}"`)) {
        issues.push(`expected workbook artifact styles to include fill ${style.fillColor}`);
      }
      if (style.fontColor && !styles.includes(`<color rgb="${argb(style.fontColor)}"`)) {
        issues.push(`expected workbook artifact styles to include font color ${style.fontColor}`);
      }
    }
  }
  if (expected.validationRanges) {
    for (const expectedValidation of expected.validationRanges) {
      const validationSheet = workbookSheetText(workbook, expectedValidation.sheetName);
      const source = `"${(expectedValidation.source ?? []).join(",")}"`;
      const escapedFormula = `<formula1>${escapeXml(source)}</formula1>`;
      const literalFormula = `<formula1>${source}</formula1>`;
      if (!validationSheet.includes(`sqref="${expectedValidation.range}"`) || (!validationSheet.includes(escapedFormula) && !validationSheet.includes(literalFormula))) {
        issues.push(`expected workbook artifact validation ${expectedValidation.sheetName}!${expectedValidation.range} source ${JSON.stringify(expectedValidation.source)}`);
      }
    }
  }
  if (expected.insertedColumns) {
    for (const expectedInsert of expected.insertedColumns) {
      const insertSheet = workbookSheetText(workbook, expectedInsert.sheetName);
      const dimension = /<dimension ref="([^"]+)"/.exec(insertSheet)?.[1];
      const endCell = dimension?.split(":").at(-1);
      const endColumn = endCell ? /^[A-Z]+/.exec(endCell)?.[0] : undefined;
      if (!endColumn || columnIndex(endColumn) < columnIndex(expectedInsert.minDimensionEndColumn ?? expectedInsert.address.replace(/:.*/, ""))) {
        issues.push(`expected workbook artifact dimension to include inserted column ${expectedInsert.sheetName}!${expectedInsert.address}, got ${dimension ?? "none"}`);
      }
    }
  }
  if (expected.conditionalFormatRanges) {
    for (const expectedFormat of expected.conditionalFormatRanges) {
      const formatSheet = workbookSheetText(workbook, expectedFormat.sheetName);
      if (!formatSheet.includes(`<conditionalFormatting sqref="${expectedFormat.range}"`) || !formatSheet.includes(`<formula>${escapeXml(expectedFormat.formula)}</formula>`)) {
        issues.push(`expected workbook artifact conditional format ${expectedFormat.sheetName}!${expectedFormat.range} formula ${expectedFormat.formula}`);
      }
    }
  }
  if (expected.tableColumnOrder) {
    try {
      const table = workbookTableText(workbook, expected.tableColumnOrder.tableName);
      for (const [index, column] of expected.tableColumnOrder.columns.entries()) {
        if (!table.includes(`<tableColumn id="${index + 1}" name="${escapeXml(column)}"`)) {
          issues.push(`expected workbook artifact table column ${index + 1} to be ${column}`);
        }
      }
    } catch (error) {
      issues.push(error instanceof Error ? error.message : String(error));
    }
  }
  if (expected.sheetProtections) {
    for (const expectedProtection of expected.sheetProtections) {
      const protectedSheet = workbookSheetText(workbook, expectedProtection.sheetName);
      const hasProtection = protectedSheet.includes("<sheetProtection");
      if (hasProtection !== expectedProtection.protected) {
        issues.push(`expected workbook artifact ${expectedProtection.sheetName} protected=${expectedProtection.protected}, got ${hasProtection}`);
      }
      for (const [key, value] of Object.entries(expectedProtection.options ?? {})) {
        const attribute = sheetProtectionXmlAttribute(key, value);
        if (hasProtection && attribute && !protectedSheet.includes(attribute)) {
          issues.push(`expected workbook artifact ${expectedProtection.sheetName} protection option ${attribute}`);
        }
      }
    }
  }
  return issues;
}

function createWorkbookArtifactFiles(workbook) {
  const sheets = [...workbook.sheets.values()];
  const tables = [...workbook.tables.values()];
  const styleRegistry = collectWorkbookStyles(sheets);
  const sheetOverrides = sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const tableOverrides = tables.map((_, index) => `<Override PartName="/xl/tables/table${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join("");
  const workbookSheets = sheets.map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("");
  const workbookRelationships = [
    ...sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`),
    `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  ].join("");
  const files = {
    "[Content_Types].xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${tableOverrides}
</Types>`),
    "_rels/.rels": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    "xl/workbook.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${workbookSheets}</sheets>
</workbook>`),
    "xl/_rels/workbook.xml.rels": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${workbookRelationships}
</Relationships>`),
    "xl/styles.xml": xml(renderWorkbookStyles(styleRegistry))
  };
  const tableIndexes = new Map(tables.map((table, index) => [table.tableName, index + 1]));
  for (const [sheetIndex, sheet] of sheets.entries()) {
    const sheetTables = tables
      .filter((table) => table.sheetName === sheet.name)
      .map((table) => ({ table, artifactIndex: tableIndexes.get(table.tableName) }));
    files[`xl/worksheets/sheet${sheetIndex + 1}.xml`] = xml(renderWorksheetXml(sheet, styleRegistry, sheetTables));
    if (sheetTables.length > 0) {
      const relationships = sheetTables.map((entry) => `<Relationship Id="rIdTable${entry.artifactIndex}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${entry.artifactIndex}.xml"/>`).join("");
      files[`xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`] = xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relationships}
</Relationships>`);
    }
  }
  for (const [index, table] of tables.entries()) {
    files[`xl/tables/table${index + 1}.xml`] = xml(renderTableXml(table));
  }
  return files;
}

function renderWorksheetXml(sheet, styleRegistry, tables) {
  const used = sheet.usedRange();
  const rows = [];
  const range = parseRange(used.address);
  for (let row = 1; row <= range.endRow; row += 1) {
    const cells = [];
    for (let col = 1; col <= range.endCol; col += 1) {
      const cell = sheet.cell(row, col);
      const style = sheet.styles.get(`${row}:${col}`);
      const styleId = style ? styleRegistry.ids.get(stableJson(style)) : undefined;
      if ((cell.value === null || cell.value === undefined) && (cell.formula === null || cell.formula === undefined) && styleId === undefined) continue;
      cells.push(renderCell(row, col, cell, styleId));
    }
    if (cells.length > 0) rows.push(`<row r="${row}">${cells.join("")}</row>`);
  }
  const validations = [...sheet.validations.entries()].map(([address, validation]) => {
    const source = Array.isArray(validation.source) ? validation.source.join(",") : validation.source;
    return `<dataValidation type="list" allowBlank="${validation.ignoreBlanks === false ? "0" : "1"}" showDropDown="${validation.inCellDropDown === false ? "1" : "0"}" sqref="${escapeXml(address)}"><formula1>"${escapeXml(source)}"</formula1></dataValidation>`;
  });
  const conditionalFormats = [...sheet.conditionalFormats.entries()].map(([address, rule], index) =>
    `<conditionalFormatting sqref="${escapeXml(address)}"><cfRule type="expression" dxfId="${index}" priority="${index + 1}"><formula>${escapeXml(rule.formula)}</formula></cfRule></conditionalFormatting>`
  );
  const tableParts = tables.map((entry) => `<tablePart r:id="rIdTable${entry.artifactIndex}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${used.address}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  ${sheet.protected ? sheetProtectionXml(sheet.protectionOptions) : ""}
  <sheetData>${rows.join("")}</sheetData>
  ${validations.length > 0 ? `<dataValidations count="${validations.length}">${validations.join("")}</dataValidations>` : ""}
  ${conditionalFormats.join("")}
  ${tables.length > 0 ? `<tableParts count="${tables.length}">${tableParts}</tableParts>` : ""}
</worksheet>`;
}

function sheetProtectionXml(options = {}) {
  const attributes = [
    'sheet="1"',
    options.protectDrawingObjects === undefined ? sheetProtectionXmlAttribute("protectDrawingObjects", true) : undefined,
    options.protectScenarios === undefined ? sheetProtectionXmlAttribute("protectScenarios", true) : undefined,
    ...Object.entries(options).map(([key, value]) => sheetProtectionXmlAttribute(key, value))
  ].filter(Boolean);
  return `<sheetProtection ${[...new Set(attributes)].join(" ")}/>`;
}

function sheetProtectionXmlAttribute(key, value) {
  const attributeMap = {
    allowFormatCells: "formatCells",
    allowFormatColumns: "formatColumns",
    allowFormatRows: "formatRows",
    allowInsertColumns: "insertColumns",
    allowInsertRows: "insertRows",
    allowDeleteColumns: "deleteColumns",
    allowDeleteRows: "deleteRows",
    allowSort: "sort",
    allowAutoFilter: "autoFilter",
    allowPivotTables: "pivotTables",
    protectDrawingObjects: "objects",
    protectScenarios: "scenarios",
    selectionMode: "selectLockedCells"
  };
  const attribute = attributeMap[key];
  if (!attribute) return undefined;
  if (key === "selectionMode") {
    return value === "normal" || value === "unlocked" ? `${attribute}="1"` : `${attribute}="0"`;
  }
  return `${attribute}="${value ? "1" : "0"}"`;
}

function renderCell(row, col, cell, styleId) {
  const address = `${columnName(col)}${row}`;
  const styleAttr = styleId !== undefined ? ` s="${styleId}"` : "";
  if (cell.formula !== null && cell.formula !== undefined) {
    return `<c r="${address}"${styleAttr}><f>${escapeXml(String(cell.formula).replace(/^=/, ""))}</f></c>`;
  }
  if (typeof cell.value === "number") {
    return `<c r="${address}"${styleAttr}><v>${cell.value}</v></c>`;
  }
  return `<c r="${address}" t="inlineStr"${styleAttr}><is><t>${escapeXml(cell.value ?? "")}</t></is></c>`;
}

function renderWorkbookStyles(styleRegistry) {
  const styles = styleRegistry.styles;
  const numberFormatIds = new Map();
  for (const style of styles) {
    if (style.numberFormat && !numberFormatIds.has(style.numberFormat)) {
      numberFormatIds.set(style.numberFormat, 164 + numberFormatIds.size);
    }
  }
  const fonts = [{}, ...styles];
  const fills = [{}, {}, ...styles];
  const cellXfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  for (let index = 0; index < styles.length; index += 1) {
    const style = styles[index];
    const numFmtId = style.numberFormat ? numberFormatIds.get(style.numberFormat) : 0;
    cellXfs.push(`<xf numFmtId="${numFmtId}" fontId="${index + 1}" fillId="${index + 2}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"${style.numberFormat ? ' applyNumberFormat="1"' : ""}>${style.horizontalAlignment ? `<alignment horizontal="${escapeXml(style.horizontalAlignment)}"/>` : ""}</xf>`);
  }
  const numFmts = [...numberFormatIds.entries()].map(([formatCode, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(formatCode)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${numberFormatIds.size > 0 ? `<numFmts count="${numberFormatIds.size}">${numFmts}</numFmts>` : ""}
  <fonts count="${fonts.length}">${fonts.map((style) => `<font>${style.fontBold ? "<b/>" : ""}${style.fontItalic ? "<i/>" : ""}${style.fontColor ? `<color rgb="${argb(style.fontColor)}"/>` : ""}<sz val="${style.fontSize ?? 11}"/><name val="${escapeXml(style.fontName ?? "Calibri")}"/></font>`).join("")}</fonts>
  <fills count="${fills.length}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${styles.map((style) => `<fill><patternFill patternType="solid"><fgColor rgb="${argb(style.fillColor ?? "#FFFFFF")}"/><bgColor indexed="64"/></patternFill></fill>`).join("")}</fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${cellXfs.length}">${cellXfs.join("")}</cellXfs>
  <dxfs count="0"/>
</styleSheet>`;
}

function renderTableXml(table) {
  const info = table.info();
  const columns = info.columns.map((column, index) => `<tableColumn id="${index + 1}" name="${escapeXml(column.name)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${escapeXml(table.tableName)}" displayName="${escapeXml(table.tableName)}" ref="${escapeXml(table.address)}" totalsRowShown="0">
  <autoFilter ref="${escapeXml(table.address)}"/>
  <tableColumns count="${info.columns.length}">${columns}</tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`;
}

function collectWorkbookStyles(sheets) {
  const styles = [];
  const ids = new Map();
  for (const sheet of sheets) {
    for (const style of sheet.styles.values()) {
      const key = stableJson(style);
      if (!ids.has(key)) {
        styles.push(style);
        ids.set(key, styles.length);
      }
    }
  }
  return { styles, ids };
}

function readXlsx(filePath) {
  const buffer = readFileSync(filePath);
  const zip = readZip(buffer);
  return {
    entries: () => [...zip.keys()],
    entryText: (name) => {
      const entry = zip.get(name);
      if (!entry) throw new Error(`Missing workbook part: ${name}`);
      return entry.toString("utf8").replace(/\s+/g, " ");
    }
  };
}

function workbookSheetText(workbook, sheetName) {
  const workbookXml = workbook.entryText("xl/workbook.xml");
  const sheetMatches = [...workbookXml.matchAll(/<sheet name="([^"]+)" sheetId="\d+" r:id="([^"]+)"/g)];
  const match = sheetMatches.find((item) => unescapeXml(item[1]) === sheetName);
  if (!match) throw new Error(`Missing workbook sheet ${sheetName}`);
  const relationships = workbook.entryText("xl/_rels/workbook.xml.rels");
  const relationship = new RegExp(`<Relationship Id="${escapeRegExp(match[2])}"[^>]*Target="([^"]+)"`).exec(relationships);
  if (!relationship) throw new Error(`Missing workbook relationship for sheet ${sheetName}`);
  return workbook.entryText(`xl/${relationship[1]}`);
}

function workbookTableText(workbook, tableName) {
  for (const entry of workbook.entries().filter((name) => name.startsWith("xl/tables/") && name.endsWith(".xml"))) {
    const table = workbook.entryText(entry);
    if (table.includes(`name="${escapeXml(tableName)}"`) || table.includes(`displayName="${escapeXml(tableName)}"`)) {
      return table;
    }
  }
  throw new Error(`Missing workbook table ${tableName}`);
}

function buildReport({ toolNames, results, artifactDir, scenarioSource, selection }) {
  const categories = {};
  for (const result of results) {
    categories[result.category] ??= [];
    categories[result.category].push(result);
  }
  const statusCounts = countBy(results, (result) => result.status);
  const nextActionCounts = countBy(results, (result) => result.nextAction ?? "none");
  const usageTotals = summarizeScenarioUsage(results);
  const themes = inferThemes(results);
  return {
    title: "Open Workbook Office Agent Behavior Report",
    generatedAt: new Date().toISOString(),
    artifactDir,
    scenarioSource,
    selection: selection ? {
      mode: selection.mode,
      scenarioSelectors: selection.scenarioSelectors,
      categorySelectors: selection.categorySelectors,
      scenarioCount: selection.scenarioCount
    } : undefined,
    toolNames,
    scenarioCount: results.length,
    statusCounts,
    nextActionCounts,
    usageTotals,
    themes,
    categories,
    results
  };
}

function summarizeScenarioUsage(results) {
  const sum = (field) => results.reduce((total, result) => total + (Number(result.usage?.[field]) || 0), 0);
  const slowest = [...results]
    .sort((left, right) => (right.usage?.scenarioElapsedMs ?? 0) - (left.usage?.scenarioElapsedMs ?? 0))
    .slice(0, 5)
    .map((result) => ({
      scenarioId: result.scenarioId,
      category: result.category,
      elapsedMs: result.usage?.scenarioElapsedMs ?? 0,
      toolCallCount: result.usage?.toolCallCount ?? 0,
      modelCallCount: result.usage?.modelCallCount ?? 0,
      estimatedTokens: result.usage?.estimatedTokens ?? 0,
      estimatedCostUsd: result.usage?.estimatedCostUsd ?? 0,
      payloadBytes: result.usage?.payloadBytes ?? 0
    }));
  return {
    toolCallCount: sum("toolCallCount"),
    modelCallCount: sum("modelCallCount"),
    scenarioElapsedMs: sum("scenarioElapsedMs"),
    toolWallMs: sum("toolWallMs"),
    backendElapsedMs: sum("backendElapsedMs"),
    payloadBytes: sum("payloadBytes"),
    estimatedTokens: sum("estimatedTokens"),
    estimatedCostUsd: Number(sum("estimatedCostUsd").toFixed(8)),
    costModel: {
      source: "estimated_tokens_proxy",
      usdPerMillionTokens: estimatedUsdPerMillionTokens
    },
    estimatedTokensSaved: sum("estimatedTokensSaved"),
    internalCallCount: sum("internalCallCount"),
    internalReadCount: sum("internalReadCount"),
    fullReadCellCount: sum("fullReadCellCount"),
    cacheHits: sum("cacheHits"),
    autoAppliedCount: sum("autoAppliedCount"),
    slowest
  };
}

function evaluateScenarioExpectations({ scenario, result, effective, workbookChanged, usage, hostCalls = [], workbook }) {
  const issues = [];
  const expected = scenario.expected ?? {};
  const budgets = scenario.budgets ?? {};
  const stepResults = result?.steps ?? [];
  if (expected.resultType === "preview" && effective?.status !== "PREVIEW_READY") {
    issues.push(`expected preview result, got ${effective?.status ?? "no result"}`);
  }
  if (expected.resultType === "apply" && effective?.status !== "SUCCESS") {
    issues.push(`expected apply success result, got ${effective?.status ?? "no result"}`);
  }
  if (expected.resultType === "error" && ["SUCCESS", "PREVIEW_READY"].includes(effective?.status)) {
    issues.push(`expected error or clarification result, got ${effective?.status}`);
  }
  if (expected.status && effective?.status !== expected.status) {
    issues.push(`expected status ${expected.status}, got ${effective?.status ?? "no result"}`);
  }
  if (expected.nextAction && effective?.nextAction !== expected.nextAction) {
    issues.push(`expected nextAction ${expected.nextAction}, got ${effective?.nextAction ?? "none"}`);
  }
  if (expected.answerKind && effective?.answer?.kind !== expected.answerKind) {
    issues.push(`expected answer kind ${expected.answerKind}, got ${effective?.answer?.kind ?? "none"}`);
  }
  if (expected.answerAction && effective?.answer?.action !== expected.answerAction) {
    issues.push(`expected answer action ${expected.answerAction}, got ${effective?.answer?.action ?? "none"}`);
  }
  if (expected.requiredCapabilities) {
    const actualCapabilities = Array.isArray(effective?.answer?.requiredCapabilities) ? effective.answer.requiredCapabilities : [];
    for (const capability of expected.requiredCapabilities) {
      if (!actualCapabilities.includes(capability)) {
        issues.push(`expected required capability ${capability}, got ${actualCapabilities.join(", ") || "none"}`);
      }
    }
  }
  if (expected.warningIncludes) {
    const warnings = (effective?.warnings ?? []).map((warning) => String(warning));
    for (const expectedWarning of expected.warningIncludes) {
      if (!warnings.some((warning) => warning.includes(expectedWarning))) {
        issues.push(`expected warning containing ${JSON.stringify(expectedWarning)}, got ${JSON.stringify(warnings)}`);
      }
    }
  }
  if (expected.previewAnswerKind) {
    const previewKinds = stepResults.map((step) => step.result?.answer?.kind).filter(Boolean);
    if (!previewKinds.includes(expected.previewAnswerKind)) {
      issues.push(`expected preview answer kind ${expected.previewAnswerKind}, got ${previewKinds.join(", ") || "none"}`);
    }
  }
  if (expected.hostMethods) {
    const methods = hostCalls.map((call) => call.method);
    for (const method of expected.hostMethods) {
      if (!methods.includes(method)) issues.push(`expected host method ${method}, got ${methods.join(", ") || "none"}`);
    }
  }
  if (expected.operationKinds) {
    const operationKinds = hostCalls.flatMap((call) => call.params?.request?.operations?.map((operation) => operation.kind) ?? []);
    for (const kind of expected.operationKinds) {
      if (!operationKinds.includes(kind)) issues.push(`expected operation kind ${kind}, got ${operationKinds.join(", ") || "none"}`);
    }
  }
  if (expected.shouldMutateWorkbook === true && !workbookChanged) {
    issues.push("expected workbook mutation, but workbook summary did not change");
  }
  if ((expected.shouldMutateWorkbook === false || expected.mustNotMutateWorkbook === true) && workbookChanged) {
    issues.push("expected read-only behavior, but workbook summary changed");
  }
  if (expected.mustNotReadFullWorkbook === true && usage.fullReadCellCount > 500) {
    issues.push(`expected no full workbook read, but read ${usage.fullReadCellCount} cells`);
  }
  if (expected.mustNotReadSheetValues === true && usage.internalReadCount > 0) {
    issues.push(`expected metadata-only path, but used ${usage.internalReadCount} internal read(s)`);
  }
  if (expected.resolvedTarget) {
    const proof = Array.isArray(effective?.proof) ? effective.proof : [];
    const matched = proof.some((entry) =>
      (!expected.resolvedTarget.sheetName || entry.sheetName === expected.resolvedTarget.sheetName) &&
      (!expected.resolvedTarget.range || entry.range === expected.resolvedTarget.range)
    );
    if (!matched) {
      issues.push(`expected proof to include ${expected.resolvedTarget.sheetName ?? "*"}!${expected.resolvedTarget.range ?? "*"}`);
    }
  }
  if (expected.mustReturnSparseRows === true && !Array.isArray(effective?.answer?.sparseRows)) {
    issues.push("expected sparseRows in answer");
  }
  if (expected.mustReturnEmptySummary === true && !effective?.answer?.emptySummary) {
    issues.push("expected emptySummary in answer");
  }
  if (expected.mustNotReturnDenseRows === true && Array.isArray(effective?.answer?.rows)) {
    issues.push("expected dense rows to be omitted");
  }
  if (expected.requiresPreview === true && effective?.status !== "PREVIEW_READY") {
    issues.push(`expected preview, got ${effective?.status ?? "no result"}`);
  }
  if (expected.cellValues) {
    for (const expectedCell of expected.cellValues) {
      const value = workbook?.sheet(expectedCell.sheetName).cellValue(expectedCell.cell);
      if (value !== expectedCell.value) {
        issues.push(`expected ${expectedCell.sheetName}!${expectedCell.cell} to equal ${JSON.stringify(expectedCell.value)}, got ${JSON.stringify(value)}`);
      }
    }
  }
  if (expected.cellFormulas) {
    for (const expectedCell of expected.cellFormulas) {
      const formula = workbook?.sheet(expectedCell.sheetName).cellFormula(expectedCell.cell);
      if (formula !== expectedCell.formula) {
        issues.push(`expected ${expectedCell.sheetName}!${expectedCell.cell} formula to equal ${JSON.stringify(expectedCell.formula)}, got ${JSON.stringify(formula)}`);
      }
    }
  }
  if (expected.cellNumberFormats) {
    for (const expectedCell of expected.cellNumberFormats) {
      const style = workbook?.sheet(expectedCell.sheetName).cellStyle(expectedCell.cell) ?? {};
      if (style.numberFormat !== expectedCell.numberFormat) {
        issues.push(`expected ${expectedCell.sheetName}!${expectedCell.cell} numberFormat=${JSON.stringify(expectedCell.numberFormat)}, got ${JSON.stringify(style.numberFormat)}`);
      }
    }
  }
  if (expected.cellStyles) {
    for (const expectedStyle of expected.cellStyles) {
      const style = workbook?.sheet(expectedStyle.sheetName).cellStyle(expectedStyle.cell) ?? {};
      for (const [key, value] of Object.entries(expectedStyle.style ?? {})) {
        if (style[key] !== value) {
          issues.push(`expected ${expectedStyle.sheetName}!${expectedStyle.cell} style ${key}=${JSON.stringify(value)}, got ${JSON.stringify(style[key])}`);
        }
      }
    }
  }
  if (expected.validationRanges) {
    for (const expectedValidation of expected.validationRanges) {
      const validation = workbook?.sheet(expectedValidation.sheetName).validation(expectedValidation.range);
      if (!validation || JSON.stringify(validation.source ?? []) !== JSON.stringify(expectedValidation.source ?? [])) {
        issues.push(`expected validation ${expectedValidation.sheetName}!${expectedValidation.range} source ${JSON.stringify(expectedValidation.source)}, got ${JSON.stringify(validation)}`);
      }
    }
  }
  if (expected.conditionalFormatRanges) {
    for (const expectedFormat of expected.conditionalFormatRanges) {
      const rule = workbook?.sheet(expectedFormat.sheetName).conditionalFormat(expectedFormat.range);
      if (!rule || rule.formula !== expectedFormat.formula) {
        issues.push(`expected conditional format ${expectedFormat.sheetName}!${expectedFormat.range} formula ${expectedFormat.formula}, got ${JSON.stringify(rule)}`);
      }
    }
  }
  if (expected.tableColumnOrder) {
    const table = workbook?.table(expected.tableColumnOrder.tableName);
    const columns = table?.info().columns.map((column) => column.name) ?? [];
    if (JSON.stringify(columns) !== JSON.stringify(expected.tableColumnOrder.columns)) {
      issues.push(`expected table ${expected.tableColumnOrder.tableName} columns ${expected.tableColumnOrder.columns.join(", ")}, got ${columns.join(", ")}`);
    }
  }
  if (expected.sheetProtections) {
    for (const expectedProtection of expected.sheetProtections) {
      const protectedState = workbook?.sheet(expectedProtection.sheetName).protected;
      if (protectedState !== expectedProtection.protected) {
        issues.push(`expected ${expectedProtection.sheetName} protected=${expectedProtection.protected}, got ${protectedState}`);
      }
      for (const [key, value] of Object.entries(expectedProtection.options ?? {})) {
        const actual = workbook?.sheet(expectedProtection.sheetName).protectionOptions?.[key];
        if (actual !== value) {
          issues.push(`expected ${expectedProtection.sheetName} protection option ${key}=${JSON.stringify(value)}, got ${JSON.stringify(actual)}`);
        }
      }
    }
  }
  if (expected.insertedColumns) {
    for (const expectedInsert of expected.insertedColumns) {
      const inserted = workbook?.sheet(expectedInsert.sheetName).insertedColumns ?? [];
      if (!inserted.some((entry) => entry.address === expectedInsert.address)) {
        issues.push(`expected inserted column ${expectedInsert.sheetName}!${expectedInsert.address}, got ${JSON.stringify(inserted)}`);
      }
    }
  }
  if (expected.shouldAskClarification === true && !["AMBIGUOUS_TARGET", "NEEDS_INPUT"].includes(effective?.status)) {
    issues.push(`expected clarification, got ${effective?.status ?? "no result"}`);
  }
  if (expected.shouldFailGracefully === true && effective?.status === "SUCCESS") {
    issues.push("expected graceful failure or clarification, got SUCCESS");
  }
  if (budgets.maxToolCalls !== undefined && usage.toolCallCount > budgets.maxToolCalls) {
    issues.push(`tool calls ${usage.toolCallCount} exceeded budget ${budgets.maxToolCalls}`);
  }
  if (Number.isFinite(hardMaxCompletedTaskCalls) && usage.toolCallCount > hardMaxCompletedTaskCalls) {
    issues.push(`completed-task tool calls ${usage.toolCallCount} exceeded hard regression ceiling ${hardMaxCompletedTaskCalls}`);
  }
  if (budgets.maxModelCalls !== undefined && usage.modelCallCount > budgets.maxModelCalls) {
    issues.push(`model calls ${usage.modelCallCount} exceeded budget ${budgets.maxModelCalls}`);
  }
  if (budgets.maxPayloadBytes !== undefined && usage.payloadBytes > budgets.maxPayloadBytes) {
    issues.push(`payload bytes ${usage.payloadBytes} exceeded budget ${budgets.maxPayloadBytes}`);
  }
  if (budgets.maxEstimatedTokens !== undefined && usage.estimatedTokens > budgets.maxEstimatedTokens) {
    issues.push(`estimated tokens ${usage.estimatedTokens} exceeded budget ${budgets.maxEstimatedTokens}`);
  }
  if (budgets.maxEstimatedCostUsd !== undefined && usage.estimatedCostUsd > budgets.maxEstimatedCostUsd) {
    issues.push(`estimated cost $${usage.estimatedCostUsd} exceeded budget $${budgets.maxEstimatedCostUsd}`);
  }
  if (budgets.maxLatencyMs !== undefined && usage.scenarioElapsedMs > budgets.maxLatencyMs) {
    issues.push(`elapsed ${usage.scenarioElapsedMs}ms exceeded budget ${budgets.maxLatencyMs}ms`);
  }
  return issues;
}

function inferThemes(results) {
  const themes = [];
  if (results.some((result) => result.category.includes("comparison") && result.notes.some((note) => note.includes("comparative")))) {
    themes.push({ area: "compare", observation: "Simple comparison prompts need explicit multi-target orchestration instead of resolving one best target." });
  }
  if (results.some((result) => result.category.includes("formula") && ["VALIDATION_FAILED", "NEEDS_INPUT"].includes(result.status))) {
    themes.push({ area: "formula", observation: "Formula tasks hit the generic value safety boundary; expose a formula-aware default-surface workflow." });
  }
  if (results.some((result) => result.nextAction === "manual_review")) {
    themes.push({ area: "nextAction", observation: "Some normal office prompts stop at manual review; decide whether agent.run should return a clearer follow-up recipe." });
  }
  if (results.some((result) => result.status === "AMBIGUOUS_TARGET")) {
    themes.push({ area: "targeting", observation: "Ambiguous prompts return candidates; OpenCode guidance should teach retrying with candidateId." });
  }
  if (themes.length === 0) {
    themes.push({ area: "baseline", observation: "No broad behavior theme was inferred; inspect scenario notes for local improvements." });
  }
  return themes;
}

function renderReport(report) {
  const lines = [
    `# ${report.title}`,
    "",
    `Generated: ${report.generatedAt}`,
    `Scenario source: ${report.scenarioSource ?? "inline"}`,
    `Selection: ${renderSelection(report.selection)}`,
    `Artifacts: ${report.artifactDir}`,
    `Scenarios: ${report.scenarioCount}`,
    `Tools exposed: ${report.toolNames.join(", ")}`,
    "",
    "## Status Counts",
    ...Object.entries(report.statusCounts).sort().map(([status, count]) => `- ${status}: ${count}`),
    "",
    "## Next Actions",
    ...Object.entries(report.nextActionCounts).sort().map(([action, count]) => `- ${action}: ${count}`),
    "",
    "## Usage Summary",
    `- Tool calls: ${report.usageTotals.toolCallCount}`,
    `- Model calls: ${report.usageTotals.modelCallCount}`,
    `- Scenario elapsed: ${report.usageTotals.scenarioElapsedMs} ms`,
    `- Tool wall time: ${report.usageTotals.toolWallMs} ms`,
    `- Backend elapsed: ${report.usageTotals.backendElapsedMs} ms`,
    `- Payload bytes: ${report.usageTotals.payloadBytes}`,
    `- Estimated tokens: ${report.usageTotals.estimatedTokens}`,
    `- Estimated cost: $${report.usageTotals.estimatedCostUsd} (${report.usageTotals.costModel.usdPerMillionTokens}/1M token proxy)`,
    `- Estimated tokens saved: ${report.usageTotals.estimatedTokensSaved}`,
    `- Internal calls: ${report.usageTotals.internalCallCount}`,
    `- Internal reads: ${report.usageTotals.internalReadCount}`,
    `- Full-read cells: ${report.usageTotals.fullReadCellCount}`,
    `- Cache hits: ${report.usageTotals.cacheHits}`,
    `- Auto-applied updates: ${report.usageTotals.autoAppliedCount}`,
    `- Expectation issues: ${report.results.reduce((total, result) => total + result.expectationIssues.length, 0)}`,
    "",
    "## Slowest Scenarios",
    ...(report.usageTotals.slowest.length > 0
      ? report.usageTotals.slowest.map((item) => `- ${item.scenarioId}: ${item.elapsedMs} ms, ${item.toolCallCount} tool calls, ${item.modelCallCount} model calls, ${item.estimatedTokens} estimated tokens, $${item.estimatedCostUsd}, ${item.payloadBytes} bytes`)
      : ["- none"]),
    "",
    "## Behavior Themes",
    ...report.themes.map((theme) => `- ${theme.area}: ${theme.observation}`),
    "",
    "## Scenario Notes"
  ];
  for (const result of report.results) {
    lines.push("");
    lines.push(`### ${result.scenarioId}`);
    lines.push("");
    lines.push(`- Category: ${result.category}`);
    lines.push(`- Prompt: ${result.prompt}`);
    lines.push(`- Status: ${result.status}`);
    lines.push(`- Next action: ${result.nextAction ?? "none"}`);
    lines.push(`- Workbook changed: ${result.workbookChanged ? "yes" : "no"}`);
    lines.push(`- Tool calls: ${result.usage.toolCallCount}`);
    lines.push(`- Model calls: ${result.usage.modelCallCount}`);
    lines.push(`- Scenario elapsed: ${result.usage.scenarioElapsedMs} ms`);
    lines.push(`- Tool wall time: ${result.usage.toolWallMs} ms`);
    lines.push(`- Backend elapsed: ${result.usage.backendElapsedMs} ms`);
    lines.push(`- Payload bytes: ${result.usage.payloadBytes}`);
    lines.push(`- Estimated tokens: ${result.usage.estimatedTokens}`);
    lines.push(`- Estimated cost: $${result.usage.estimatedCostUsd}`);
    lines.push(`- Estimated tokens saved: ${result.usage.estimatedTokensSaved}`);
    lines.push(`- Task outcomes: ${Object.entries(result.usage.taskOutcomes ?? {}).map(([key, count]) => `${key}=${count}`).join(", ") || "none"}`);
    lines.push(`- Max recommended follow-up calls: ${result.usage.maxRecommendedFollowupCalls ?? "none"}`);
    lines.push(`- Internal calls: ${result.usage.internalCallCount}`);
    lines.push(`- Internal reads: ${result.usage.internalReadCount}`);
    lines.push(`- Full-read cells: ${result.usage.fullReadCellCount}`);
    lines.push(`- Cache hits: ${result.usage.cacheHits}`);
    lines.push(`- Metadata cache statuses: ${Object.entries(result.usage.metadataCacheStatuses).map(([key, count]) => `${key}=${count}`).join(", ")}`);
    lines.push(`- Signals: ${result.signals.join(", ") || "none"}`);
    lines.push(`- Artifacts: ${result.artifactDir}`);
    lines.push("- Tool chain:");
    for (const call of result.toolChain) {
      lines.push(`  - ${call.index}. ${call.tool} mode=${call.mode} status=${call.status ?? "unknown"} next=${call.nextAction ?? "none"} wall=${call.wallMs ?? 0}ms backend=${call.telemetry?.elapsedMs ?? 0}ms tokens=${call.telemetry?.estimatedTokens ?? 0} bytes=${call.telemetry?.payloadBytes ?? 0} reads=${call.telemetry?.internalReadCount ?? 0} cells=${call.telemetry?.fullReadCellCount ?? 0} request="${call.request ?? ""}"`);
    }
    if (result.notes.length > 0) {
      lines.push("- Notes:");
      for (const note of result.notes) lines.push(`  - ${note}`);
    }
  }
  return lines.join("\n");
}

function renderSelection(selection) {
  if (!selection) return "unknown";
  if (selection.mode === "full-suite") return "full-suite";
  const categories = selection.categorySelectors?.length ? `categories=${selection.categorySelectors.join(",")}` : undefined;
  const scenarios = selection.scenarioSelectors?.length ? `scenarios=${selection.scenarioSelectors.join(",")}` : undefined;
  return [selection.mode, categories, scenarios].filter(Boolean).join(" ");
}

function renderObservation(observation) {
  return [
    `# ${observation.scenarioId}`,
    "",
    `Prompt: ${observation.prompt}`,
    `Category: ${observation.category}`,
    `Status: ${observation.status}`,
    `Next action: ${observation.nextAction ?? "none"}`,
    `Workbook changed: ${observation.workbookChanged ? "yes" : "no"}`,
    `Elapsed: ${observation.elapsedMs} ms`,
    `Tool calls: ${observation.usage.toolCallCount}`,
    `Model calls: ${observation.usage.modelCallCount}`,
    `Tool wall time: ${observation.usage.toolWallMs} ms`,
    `Backend elapsed: ${observation.usage.backendElapsedMs} ms`,
    `Payload bytes: ${observation.usage.payloadBytes}`,
    `Estimated tokens: ${observation.usage.estimatedTokens}`,
    `Estimated cost: $${observation.usage.estimatedCostUsd}`,
    `Task outcomes: ${Object.entries(observation.usage.taskOutcomes ?? {}).map(([key, count]) => `${key}=${count}`).join(", ") || "none"}`,
    `Max recommended follow-up calls: ${observation.usage.maxRecommendedFollowupCalls ?? "none"}`,
    `Internal reads: ${observation.usage.internalReadCount}`,
    `Full-read cells: ${observation.usage.fullReadCellCount}`,
    "",
    "## Signals",
    ...(observation.signals.length > 0 ? observation.signals.map((signal) => `- ${signal}`) : ["- none"]),
    "",
    "## Tool Chain",
    ...(observation.toolChain.length > 0
      ? observation.toolChain.map((call) => `- ${call.index}. ${call.tool} mode=${call.mode} status=${call.status ?? "unknown"} next=${call.nextAction ?? "none"} wall=${call.wallMs ?? 0}ms backend=${call.telemetry?.elapsedMs ?? 0}ms tokens=${call.telemetry?.estimatedTokens ?? 0} bytes=${call.telemetry?.payloadBytes ?? 0} reads=${call.telemetry?.internalReadCount ?? 0} cells=${call.telemetry?.fullReadCellCount ?? 0} request="${call.request ?? ""}"`)
      : ["- none"]),
    "",
    "## Notes",
    ...(observation.notes.length > 0 ? observation.notes.map((note) => `- ${note}`) : ["- none"]),
    "",
    "## Result Summary",
    observation.resultSummary ?? "No result summary."
  ].join("\n");
}

class McpClient {
  constructor(child) {
    this.child = child;
    this.nextId = 1;
    this.pending = new Map();
    this.buffer = Buffer.alloc(0);
    child.stdout.on("data", (chunk) => this.read(chunk));
    child.on("exit", (code, signal) => {
      for (const pending of this.pending.values()) pending.reject(new Error(`MCP server exited code=${code} signal=${signal}`));
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
      if (!parsed) return;
      this.handle(parsed);
    }
  }

  readFramed() {
    const marker = this.buffer.indexOf("\r\n\r\n");
    if (marker < 0) return undefined;
    const header = this.buffer.slice(0, marker).toString("utf8");
    const match = /^Content-Length:\s*(\d+)/im.exec(header);
    if (!match) return undefined;
    const length = Number(match[1]);
    const bodyStart = marker + 4;
    const bodyEnd = bodyStart + length;
    if (this.buffer.length < bodyEnd) return undefined;
    const body = this.buffer.slice(bodyStart, bodyEnd).toString("utf8");
    this.buffer = this.buffer.slice(bodyEnd);
    return JSON.parse(body);
  }

  readLineDelimited() {
    const newline = this.buffer.indexOf("\n");
    if (newline < 0) return undefined;
    const line = this.buffer.slice(0, newline).toString("utf8").trim();
    this.buffer = this.buffer.slice(newline + 1);
    return line ? JSON.parse(line) : undefined;
  }

  handle(message) {
    if (!("id" in message) || "method" in message) return;
    const pending = this.pending.get(message.id);
    if (!pending) return;
    clearTimeout(pending.timeout);
    this.pending.delete(message.id);
    if (message.error) pending.reject(new Error(message.error.message));
    else pending.resolve(message.result);
  }

  close() {
    this.child.stdin.destroy();
  }
}

class FakeAddin {
  static async connect(url, workbook) {
    const { WebSocket } = await import("../../apps/backend/node_modules/ws/wrapper.mjs");
    const socket = new WebSocket(url);
    const addin = new FakeAddin(socket, workbook);
    await new Promise((resolve, reject) => {
      socket.once("open", resolve);
      socket.once("error", reject);
    });
    socket.on("message", (raw) => addin.onMessage(JSON.parse(String(raw))));
    await addin.waitForConnectionId();
    addin.sendNotification("addin.hello", {
      capabilities: { platform: "mac", officeVersion: "fake-office-agent-behavior", apiSets: { ExcelApi: "1.16" }, features: { ranges: "supported", tables: "supported", formulas: "supported" } },
      activeWorkbook: workbook.ref
    });
    return addin;
  }

  constructor(socket, workbook) {
    this.socket = socket;
    this.workbook = workbook;
    this.activeWorkbook = workbook.ref;
    this.selection = makeSelection(workbook.ref.workbookId, "Sales", "A1");
    this.connectionId = undefined;
    this.calls = [];
    this.connectedResolvers = [];
  }

  waitForConnectionId() {
    return this.connectionId ? Promise.resolve() : new Promise((resolve) => this.connectedResolvers.push(resolve));
  }

  onMessage(message) {
    if (message.method === "backend.connected") {
      this.connectionId = message.params.connectionId;
      for (const resolve of this.connectedResolvers.splice(0)) resolve();
      return;
    }
    if (!("id" in message) || !message.method) return;
    Promise.resolve()
      .then(() => this.handleRequest(message.method, message.params ?? {}))
      .then((result) => this.socket.send(JSON.stringify({ jsonrpc: "2.0", id: message.id, result })))
      .catch((error) => this.socket.send(JSON.stringify({ jsonrpc: "2.0", id: message.id, error: { code: -32000, message: error instanceof Error ? error.message : String(error) } })));
  }

  handleRequest(method, params) {
    this.calls.push({ method, params });
    switch (method) {
      case "runtime.ping":
        return { ok: true, at: params.at };
      case "runtime.get_active_context":
        return this.activeWorkbook;
      case "runtime.get_selection":
        return {
          workbook: this.activeWorkbook,
          ...(this.selection ? { selection: this.selection } : {})
        };
      case "workbook.get_map":
        return this.workbook.getMap();
      case "workbook.export_copy":
        return { ok: true, operation: "workbook.export_copy", payload: { kind: "compressed", encoding: "base64", data: Buffer.from(JSON.stringify(this.workbook.summary())).toString("base64") } };
      case "workbook.snapshot_ranges":
        return this.workbook.snapshotRanges(params.ranges);
      case "template.capture":
        return this.workbook.captureTemplate(params);
      case "template.capture_sheet":
        return this.workbook.captureSheetFingerprint(params);
      case "template.repair":
        return this.workbook.repairTemplateConsistency(params);
      case "table.list":
        return { ok: true, tables: [...this.workbook.tables.values()].map((table) => table.info()) };
      case "table.get_info":
        return { ok: true, info: this.workbook.table(params.tableName).info() };
      case "table.append_rows":
        return this.workbook.table(params.tableName).appendRows(params.values);
      case "table.reorder_columns":
        return this.workbook.table(params.tableName).reorderColumns(params.columnOrder);
      case "table.sort":
        return this.workbook.table(params.tableName).sort(params.fields);
      case "names.list":
        return { ok: true, names: [...this.workbook.names.values()] };
      case "operation.execute_batch":
        return this.workbook.executeBatch(params.request);
      case "validate.workbook":
        return { ok: true, issues: [] };
      default:
        return { ok: true, warnings: [] };
    }
  }

  sendNotification(method, params) {
    this.socket.send(JSON.stringify({ jsonrpc: "2.0", method, params }));
  }

  setActiveWorkbook(workbookRef) {
    this.activeWorkbook = workbookRef;
    this.sendNotification("workbook.contextChanged", { activeWorkbook: workbookRef ?? null });
  }

  resetWorkbook(workbook) {
    this.workbook = workbook;
    this.activeWorkbook = workbook.ref;
  }

  setSelection(selection) {
    this.selection = selection ? makeSelection(this.workbook.ref.workbookId, selection.sheetName, selection.address) : undefined;
  }

  close() {
    this.socket.close();
  }

  summary() {
    return { connectionId: this.connectionId, calls: this.calls.map((call) => call.method), workbook: this.workbook.summary() };
  }
}

class FakeWorkbook {
  constructor(id) {
    this.ref = { workbookId: id, name: "Office Agent Behavior.xlsx", path: path.join(tempRoot, "Office Agent Behavior.xlsx"), platform: "mac" };
    this.sheets = new Map();
    this.tables = new Map();
    this.names = new Map();
  }

  addSheet(name) {
    const sheet = new FakeSheet(this.ref.workbookId, name);
    this.sheets.set(name, sheet);
    return sheet;
  }

  sheet(name) {
    const sheet = this.sheets.get(name);
    if (!sheet) throw new Error(`Unknown fake sheet: ${name}`);
    return sheet;
  }

  table(name) {
    const table = this.tables.get(name);
    if (!table) throw new Error(`Unknown fake table: ${name}`);
    return table;
  }

  createTable({ sheetName, address, tableName }) {
    const table = new FakeTable(this, tableName, sheetName, address);
    this.tables.set(tableName, table);
    return table;
  }

  getMap() {
    return {
      workbook: this.ref,
      activeSheet: "Sales",
      sheets: [...this.sheets.values()].map((sheet) => ({
        workbookId: this.ref.workbookId,
        worksheetId: `sheet_${sheet.name.replace(/\W+/g, "_")}`,
        name: sheet.name,
        usedRange: sheet.usedRange(),
        tables: [...this.tables.values()].filter((table) => table.sheetName === sheet.name).map((table) => table.info())
      }))
    };
  }

  executeBatch(request) {
    const readData = [];
    let cellsRead = 0;
    let cellsWritten = 0;
    let sheetsChanged = 0;
    for (const operation of request.operations) {
      if (operation.kind === "range.read_full") {
        const snapshot = this.sheet(operation.target.sheetName).snapshot(operation.target);
        cellsRead += snapshot.fingerprint.cellCount;
        readData.push({ operationId: operation.operationId, snapshot });
      }
      if (operation.kind === "range.write_values") {
        cellsWritten += this.sheet(operation.target.sheetName).writeValues(operation.target.address, operation.values);
      }
      if (operation.kind === "range.write_formulas") {
        cellsWritten += this.sheet(operation.target.sheetName).writeFormulas(operation.target.address, operation.formulas);
      }
      if (operation.kind === "range.write_number_formats") {
        cellsWritten += this.sheet(operation.target.sheetName).writeNumberFormats(operation.target.address, operation.numberFormat);
      }
      if (operation.kind === "range.write_styles") {
        cellsWritten += this.sheet(operation.target.sheetName).writeStyles(operation.target.address, operation.style);
      }
      if (operation.kind === "range.write_styles_many") {
        for (const entry of operation.entries ?? []) {
          cellsWritten += this.sheet(entry.target.sheetName).writeStyles(entry.target.address, entry.style);
        }
      }
      if (operation.kind === "range.insert_columns") {
        cellsWritten += this.sheet(operation.target.sheetName).insertColumns(operation.target.address);
      }
      if (operation.kind === "range.reorder_columns") {
        cellsWritten += this.sheet(operation.target.sheetName).reorderColumns(operation.target.address, operation.columnOrder);
      }
      if (operation.kind === "range.write_data_validation") {
        cellsWritten += this.sheet(operation.target.sheetName).writeDataValidation(operation.target.address, operation.validation);
      }
      if (operation.kind === "range.write_conditional_formatting") {
        cellsWritten += this.sheet(operation.target.sheetName).writeConditionalFormatting(operation.target.address, operation.rule);
      }
      if (operation.kind === "range.clear_values_keep_format" || operation.kind === "range.clear_values") {
        cellsWritten += this.sheet(operation.target.sheetName).clearValues(operation.target.address);
      }
      if (operation.kind === "range.restore_snapshot") {
        cellsWritten += this.sheet(operation.target.sheetName).restoreSnapshot(operation.snapshot);
      }
      if (operation.kind === "sheet.copy") {
        this.copySheet(operation.sourceSheetName, operation.newSheetName);
        sheetsChanged += 1;
      }
      if (operation.kind === "sheet.protect") {
        this.sheet(operation.sheetName).protect(operation.password, operation.options);
        sheetsChanged += 1;
      }
      if (operation.kind === "sheet.unprotect") {
        this.sheet(operation.sheetName).unprotect(operation.password);
        sheetsChanged += 1;
      }
    }
    return {
      ok: true,
      rollbackAvailable: request.mode === "apply",
      backups: request.mode === "apply" && (cellsWritten > 0 || sheetsChanged > 0) ? [`backup_${Date.now()}`] : [],
      warnings: [],
      readData,
      diffSummary: { title: "Office behavior fake batch", changedRanges: [], cellsChanged: cellsWritten, formulasChanged: 0, stylesChanged: 0, tablesChanged: 0, sheetsChanged, destructiveLevel: sheetsChanged > 0 ? "structure" : cellsWritten > 0 ? "values" : "none" },
      telemetry: { cellsRead, cellsWritten, syncCount: 1, rangeCount: request.operations.length, chunkCount: 1, warningCount: 0 }
    };
  }

  copySheet(sourceName, newName) {
    const copy = this.sheet(sourceName).clone(newName);
    this.sheets.set(newName, copy);
    return copy;
  }

  snapshotRanges(ranges) {
    return {
      workbookId: this.ref.workbookId,
      capturedAt: fixedNow(),
      workbookFingerprint: {
        workbookId: this.ref.workbookId,
        workbookHash: stableHash(this.summary()),
        structureHash: stableHash([...this.sheets.keys(), ...this.tables.keys()]),
        capturedAt: fixedNow()
      },
      rangeSnapshots: ranges.map((range) => this.sheet(range.sheetName).snapshot(range))
    };
  }

  captureTemplate(request) {
    const captured = this.captureSheetFingerprint({
      workbookId: request.workbookId,
      sheetName: request.sourceSheetName,
      dataRegions: request.dataRegions
    });
    return {
      sourceSheetName: captured.sourceSheetName,
      dataRegions: captured.dataRegions,
      fingerprintPayload: captured.fingerprintPayload
    };
  }

  captureSheetFingerprint(request) {
    const sheet = this.sheet(request.sheetName);
    const usedRange = sheet.usedRange();
    return {
      sheetName: sheet.name,
      sourceSheetName: sheet.name,
      dataRegions: request.dataRegions ?? [],
      fingerprintPayload: {
        structure: {
          sheetName: sheet.name,
          position: [...this.sheets.keys()].indexOf(sheet.name),
          visibility: "Visible",
          usedRange: {
            address: usedRange.address,
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount
          },
          dataRegions: request.dataRegions ?? []
        },
        formulas: sheet.formulaMatrix(usedRange.address),
        styles: {
          usedRange: sheet.styleSummary(usedRange.address),
          numberFormat: sheet.numberFormatMatrix(usedRange.address)
        },
        filters: {
          note: "Filter capture will be expanded with table/filter-specific APIs."
        },
        tables: [...this.tables.values()].filter((table) => table.sheetName === sheet.name).map((table) => ({ name: table.name })),
        printLayout: {
          note: "Print layout capture will be expanded with page layout APIs."
        }
      }
    };
  }

  repairTemplateConsistency(request) {
    const source = this.sheet(request.sourceSheetName);
    const target = this.sheet(request.targetSheetName);
    const sourceRange = source.usedRange().address;
    const repaired = [];
    if (request.repair?.includes("styles")) {
      target.copyStylesFrom(source, sourceRange);
      repaired.push("styles");
    }
    if (request.repair?.includes("formulas")) {
      target.writeFormulas(sourceRange, source.formulaMatrix(sourceRange));
      repaired.push("formulas");
    }
    if (request.repair?.includes("dataRegions")) {
      for (const dataRegion of request.dataRegions ?? []) {
        target.clearValues(stripSheetName(dataRegion));
      }
      repaired.push("dataRegions");
    }
    return { ok: true, repaired };
  }

  summary() {
    return {
      workbook: this.ref,
      sheets: [...this.sheets.values()].map((sheet) => ({ name: sheet.name, usedRange: sheet.usedRange(), nonEmptyCells: sheet.nonEmptyCellCount() })),
      tables: [...this.tables.values()].map((table) => table.info()),
      hash: stableHash([...this.sheets.values()].map((sheet) => [sheet.name, sheet.valuesForSummary()]))
    };
  }
}

class FakeSheet {
  constructor(workbookId, name) {
    this.workbookId = workbookId;
    this.name = name;
    this.cells = new Map();
    this.styles = new Map();
    this.validations = new Map();
    this.conditionalFormats = new Map();
    this.insertedColumns = [];
    this.protected = false;
    this.protectionPassword = undefined;
    this.protectionOptions = {};
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

  writeNumberFormats(address, numberFormat) {
    const range = parseRange(address);
    let changed = 0;
    forEachMatrix(numberFormat, (format, row, col) => {
      const key = `${range.startRow + row}:${range.startCol + col}`;
      this.styles.set(key, { ...(this.styles.get(key) ?? {}), numberFormat: format });
      changed += 1;
    });
    return changed;
  }

  writeStyles(address, style) {
    const range = parseRange(address);
    let changed = 0;
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        this.styles.set(`${row}:${col}`, { ...(this.styles.get(`${row}:${col}`) ?? {}), ...style });
        changed += 1;
      }
    }
    return changed;
  }

  writeDataValidation(address, validation) {
    this.validations.set(address, validation);
    return parseRange(address).rowCount * parseRange(address).columnCount;
  }

  writeConditionalFormatting(address, rule) {
    this.conditionalFormats.set(address, rule);
    return parseRange(address).rowCount * parseRange(address).columnCount;
  }

  insertColumns(address) {
    const match = /^([A-Z]+)(?::([A-Z]+))?$/.exec(address);
    const startCol = match ? columnIndex(match[1]) : parseRange(address).startCol;
    const endCol = match?.[2] ? columnIndex(match[2]) : startCol;
    const count = endCol - startCol + 1;
    const nextCells = new Map();
    const nextStyles = new Map();
    for (const [key, cell] of this.cells.entries()) {
      const [row, col] = key.split(":").map(Number);
      nextCells.set(`${row}:${col >= startCol ? col + count : col}`, cell);
    }
    for (const [key, style] of this.styles.entries()) {
      const [row, col] = key.split(":").map(Number);
      nextStyles.set(`${row}:${col >= startCol ? col + count : col}`, style);
    }
    this.cells = nextCells;
    this.styles = nextStyles;
    this.insertedColumns.push({ address, count });
    return this.usedRange().rowCount * count;
  }

  reorderColumns(address, columnOrder) {
    const range = parseRange(address);
    const rows = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const rowValues = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.cell(row, col);
        rowValues.push(cell.formula ?? cell.value);
      }
      rows.push(columnOrder.map((order) => rowValues[Number(order) - 1] ?? null));
    }
    this.writeValues(address, rows);
    return range.rowCount * range.columnCount;
  }

  clearValues(address) {
    const range = parseRange(address);
    let cleared = 0;
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.cell(row, col);
        if (cell.value !== null) cleared += 1;
        cell.value = null;
      }
    }
    return cleared;
  }

  restoreSnapshot(snapshot) {
    const range = parseRange(snapshot.fingerprint.range.address);
    let changed = 0;
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const rowIndex = row - range.startRow;
        const colIndex = col - range.startCol;
        const nextValue = snapshot.values?.[rowIndex]?.[colIndex] ?? null;
        const nextFormula = snapshot.formulas?.[rowIndex]?.[colIndex] ?? null;
        const cell = this.cell(row, col);
        if (cell.value !== nextValue || cell.formula !== nextFormula) changed += 1;
        cell.value = nextValue;
        cell.formula = nextFormula;
      }
    }
    return changed;
  }

  protect(password, options = {}) {
    this.protected = true;
    this.protectionPassword = password;
    this.protectionOptions = { ...options };
  }

  unprotect() {
    this.protected = false;
    this.protectionPassword = undefined;
    this.protectionOptions = {};
  }

  clone(name) {
    const next = new FakeSheet(this.workbookId, name);
    for (const [key, cell] of this.cells.entries()) {
      next.cells.set(key, { ...cell });
    }
    next.protected = this.protected;
    next.protectionPassword = this.protectionPassword;
    next.protectionOptions = { ...this.protectionOptions };
    return next;
  }

  formulaMatrix(address) {
    const range = parseRange(address);
    const formulas = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const formulaRow = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        formulaRow.push(this.cell(row, col).formula);
      }
      formulas.push(formulaRow);
    }
    return formulas;
  }

  numberFormatMatrix(address) {
    const range = parseRange(address);
    const formats = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const formatRow = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        formatRow.push(this.styles.get(`${row}:${col}`)?.numberFormat ?? null);
      }
      formats.push(formatRow);
    }
    return formats;
  }

  styleSummary(address) {
    const range = parseRange(address);
    const styles = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const style = this.styles.get(`${row}:${col}`);
        if (style) {
          styles.push({ cell: `${columnName(col)}${row}`, style });
        }
      }
    }
    return styles;
  }

  copyStylesFrom(source, address) {
    const range = parseRange(address);
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const style = source.styles.get(`${row}:${col}`);
        if (style) {
          this.styles.set(`${row}:${col}`, { ...style });
        } else {
          this.styles.delete(`${row}:${col}`);
        }
      }
    }
  }

  snapshot(rangeRef) {
    const range = parseRange(rangeRef.address);
    const values = [];
    const formulas = [];
    const text = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const valueRow = [];
      const formulaRow = [];
      const textRow = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.cell(row, col);
        valueRow.push(cell.value);
        formulaRow.push(cell.formula);
        textRow.push(cell.formula ?? (cell.value === null || cell.value === undefined ? "" : String(cell.value)));
      }
      values.push(valueRow);
      formulas.push(formulaRow);
      text.push(textRow);
    }
    return {
      fingerprint: { range: rangeRef, hash: stableHash({ values, formulas }), cellCount: range.rowCount * range.columnCount, capturedAt: fixedNow() },
      values,
      formulas,
      text
    };
  }

  usedRange() {
    let maxRow = 1;
    let maxCol = 1;
    for (const [key, cell] of this.cells.entries()) {
      if (cell.value === null && cell.formula === null) continue;
      const [row, col] = key.split(":").map(Number);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }
    for (const inserted of this.insertedColumns) {
      const range = parseRange(inserted.address);
      maxRow = Math.max(maxRow, range.endRow);
      maxCol = Math.max(maxCol, range.endCol);
    }
    return { workbookId: this.workbookId, sheetName: this.name, address: `A1:${columnName(maxCol)}${maxRow}`, rowCount: maxRow, columnCount: maxCol };
  }

  nonEmptyCellCount() {
    return [...this.cells.values()].filter((cell) => cell.value !== null || cell.formula !== null).length;
  }

  valuesForSummary() {
    return {
      cells: [...this.cells.entries()].filter(([, cell]) => cell.value !== null || cell.formula !== null).sort(),
      styles: [...this.styles.entries()].sort(),
      validations: [...this.validations.entries()].sort(),
      conditionalFormats: [...this.conditionalFormats.entries()].sort(),
      insertedColumns: this.insertedColumns,
      protected: this.protected,
      protectionOptions: this.protectionOptions
    };
  }

  cell(row, col) {
    const key = `${row}:${col}`;
    if (!this.cells.has(key)) this.cells.set(key, { value: null, formula: null });
    return this.cells.get(key);
  }

  cellValue(address) {
    const { row, col } = parseCell(address);
    const cell = this.cell(row, col);
    return cell.formula ?? cell.value;
  }

  cellFormula(address) {
    const { row, col } = parseCell(address);
    return this.cell(row, col).formula;
  }

  cellStyle(address) {
    const { row, col } = parseCell(address);
    return this.styles.get(`${row}:${col}`) ?? {};
  }

  validation(address) {
    return this.validations.get(address);
  }

  conditionalFormat(address) {
    return this.conditionalFormats.get(address);
  }
}

class FakeTable {
  constructor(workbook, tableName, sheetName, address) {
    this.workbook = workbook;
    this.tableName = tableName;
    this.sheetName = sheetName;
    this.address = address;
  }

  info() {
    const range = parseRange(this.address);
    const headers = this.workbook.sheet(this.sheetName).snapshot({ workbookId: this.workbook.ref.workbookId, sheetName: this.sheetName, address: `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.startRow}` }).values[0].map(String);
    return {
      workbookId: this.workbook.ref.workbookId,
      tableName: this.tableName,
      name: this.tableName,
      id: `table_${this.tableName}`,
      sheetName: this.sheetName,
      address: this.address,
      headerAddress: `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.startRow}`,
      rowCount: Math.max(0, range.rowCount - 1),
      columnCount: range.columnCount,
      columns: headers.map((name, index) => ({ id: index + 1, index, name }))
    };
  }

  appendRows(values) {
    const range = parseRange(this.address);
    const startAddress = `${columnName(range.startCol)}${range.endRow + 1}`;
    this.workbook.sheet(this.sheetName).writeValues(startAddress, values);
    this.address = `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.endRow + values.length}`;
    return { ok: true, rowCount: values.length, tableName: this.tableName, address: this.address };
  }

  reorderColumns(columnOrder) {
    const range = parseRange(this.address);
    const sheet = this.workbook.sheet(this.sheetName);
    const headers = this.info().columns.map((column) => column.name);
    const orderIndexes = columnOrder.map((column) => typeof column === "number" ? column : headers.indexOf(column)).filter((index) => index >= 0);
    const rows = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const rowValues = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = sheet.cell(row, col);
        rowValues.push(cell.formula ?? cell.value);
      }
      rows.push(orderIndexes.map((index) => rowValues[index] ?? null));
    }
    sheet.writeValues(this.address, rows);
    return { ok: true, info: this.info(), warnings: [] };
  }

  sort(fields = []) {
    const [field] = fields;
    if (!field) return { ok: true, info: this.info(), warnings: [] };
    const range = parseRange(this.address);
    const sheet = this.workbook.sheet(this.sheetName);
    const rows = [];
    for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
      rows.push(sheet.snapshot({ workbookId: this.workbook.ref.workbookId, sheetName: this.sheetName, address: `${columnName(range.startCol)}${row}:${columnName(range.endCol)}${row}` }).values[0]);
    }
    rows.sort((left, right) => {
      const leftValue = left[field.key];
      const rightValue = right[field.key];
      const direction = field.ascending === false ? -1 : 1;
      return String(leftValue).localeCompare(String(rightValue), undefined, { numeric: true }) * direction;
    });
    sheet.writeValues(`${columnName(range.startCol)}${range.startRow + 1}:${columnName(range.endCol)}${range.endRow}`, rows);
    return { ok: true, info: this.info(), warnings: [] };
  }
}

function createWorkbookFixture(id) {
  const workbook = new FakeWorkbook(id);
  workbook.addSheet("Sales").writeValues("A1:E6", [
    ["Date", "Customer", "Product", "Amount", "Status"],
    ["2026-01-03", "Acme Co", "Consulting", 1200, "Open"],
    ["2026-01-04", "Northwind", "Support", 450, "Closed"],
    ["2026-01-08", "Contoso", "Implementation", 3200, "Open"],
    ["2026-01-10", "Fabrikam", "Training", 800, "Open"],
    ["2026-01-12", "Tailspin", "Support", 650, "Closed"]
  ]);
  workbook.createTable({ sheetName: "Sales", address: "A1:E6", tableName: "Transactions" });
  workbook.addSheet("January").writeValues("A1:C5", [["Metric", "Value", "Notes"], ["Revenue", 6300, ""], ["Expense", 2400, ""], ["Profit", 3900, ""], ["Open deals", 3, ""]]);
  workbook.addSheet("February").writeValues("A1:C5", [["Metric", "Value", "Notes"], ["Revenue", 7100, ""], ["Expense", 2600, ""], ["Profit", 4500, ""], ["Open deals", 5, ""]]);
  writeMonthlyPerformanceSheet(workbook.addSheet("Mar 2026"), "March", { received: 280000, spent: 240000, gas: 58000, repair: 18000, salary: 64000 });
  writeMonthlyPerformanceSheet(workbook.addSheet("Apr 2026"), "April", { received: 333881.72, spent: 363263.96, gas: 71000, repair: 42000, salary: 65000 });
  workbook.addSheet("Blank");
  workbook.addSheet("Summary");
  workbook.addSheet("HR").writeValues("A1:E5", [
    ["Employee", "Department", "Start Date", "Status", "Review"],
    ["Ana Gomez", "Operations", "2025-01-15", "Active", ""],
    ["Ben Smith", "Finance", "2024-11-01", "Active", ""],
    ["Chai Lee", "Sales", "2026-02-10", "Onboarding", ""],
    ["Dana Wu", "Support", "2023-08-20", "Active", ""]
  ]);
  workbook.addSheet("Data Cleanup").writeValues("A1:F6", [
    ["Raw Name", "Email", "Amount Text", "Clean Status", "Raw Date", "Currency Text"],
    ["  acme co  ", "INFO@ACME.COM ", "1,200.00", "", "26/6/26", "$1,200.00"],
    ["northwind", " sales@northwind.example", "450", "", "2026-06-27", "€450"],
    ["Contoso  ", "OPS@CONTOSO.EXAMPLE", "3,200.00", "", "not a date", "($3,200.00)"],
    ["tailspin", "help@tailspin.example ", "650", "", "28.6.26", "¥650"],
    ["northwind", " sales@northwind.example", "450", "", "2026-06-27", "€450"]
  ]);
  workbook.addSheet("Sparse").writeValues("A1:J10", [
    ["Owner", "", "", "", "", "", "", "", "", "Status"],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", "Ready"]
  ]);
  workbook.addSheet("Operations").writeValues("A1:J17", [
    ["Operations Review", "", "", "", "", "", "", "", "", ""],
    ["Period", "Jun 2026", "Prepared by", "Ops", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["Metric", "Value", "Status", "", "", "", "", "Owner", "Action", "Due"],
    ["Revenue", 7200, "On Track", "", "", "", "", "A. Chen", "Review invoices", "2026-06-20"],
    ["Expense", 2800, "Watch", "", "", "", "", "M. Lee", "Check variance", "2026-06-21"],
    ["Profit", 4400, "On Track", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["Invoice No", "Customer", "Job", "Amount", "Status", "Owner", "", "", "", ""],
    ["INV-100", "Acme", "Lift", 1200, "Open", "A. Chen", "", "", "", ""],
    ["INV-101", "Northwind", "Haul", 900, "Paid", "M. Lee", "", "", "", ""],
    ["INV-102", "Contoso", "Storage", 450, "Open", "A. Chen", "", "", "", ""],
    ["", "", "", "", "", "", "", "", "", ""],
    ["Reconciliation", "", "", "", "", "", "", "", "", ""],
    ["Expected", "Actual", "Variance", "Check", "", "", "", "", "", ""],
    [2500, 2550, null, "Review", "", "", "", "", "", ""]
  ]);
  workbook.sheet("Operations").writeFormulas("C17:C17", [["=B17-A17"]]);
  workbook.addSheet("FormulaSheet").writeValues("A1:D10", [
    ["Item", "Amount", "Tax", "Total"],
    ["A", 100, 7, null],
    ["B", 200, 14, null],
    ["C", 300, 21, null],
    ["D", 400, 28, null],
    ["E", 500, 35, null],
    ["F", 600, 42, null],
    ["G", 700, 49, null],
    ["H", 800, 56, null],
    ["I", 900, 63, null]
  ]);
  workbook.sheet("FormulaSheet").writeFormulas("D2:D10", [["=B2+C2"], ["=B3+C3"], ["=B4+C4"], ["=B5+C5"], ["=B6+C6"], ["=B7+C7"], ["=B8+C8"], ["=B9+C9"], ["=B10+C10"]]);
  workbook.addSheet("FormulaTarget").writeValues("A1:D10", [
    ["Item", "Amount", "Tax", "Total"],
    ["A", 100, 7, null],
    ["B", 200, 14, null],
    ["C", 300, 21, null],
    ["D", 400, 28, null],
    ["E", 500, 35, null],
    ["F", 600, 42, null],
    ["G", 700, 49, null],
    ["H", 800, 56, null],
    ["I", 900, 63, null]
  ]);
  workbook.sheet("FormulaTarget").writeFormulas("D2:D10", [["=B2-C2"], ["=B3-C3"], ["=B4-C4"], ["=B5-C5"], ["=B6-C6"], ["=B7-C7"], ["=B8-C8"], ["=B9-C9"], ["=B10-C10"]]);
  workbook.names.set(":RevenueTotal", { workbookId: id, name: "RevenueTotal", scope: "workbook", address: "January!B2" });
  workbook.names.set("OperationsKpiSection", { workbookId: id, name: "OperationsKpiSection", scope: "workbook", sheetName: "Operations", address: "Operations!A4:C7" });
  workbook.names.set("OperationsInvoiceSection", { workbookId: id, name: "OperationsInvoiceSection", scope: "workbook", sheetName: "Operations", address: "Operations!A10:F13" });
  workbook.names.set("OperationsNotesStatusSection", { workbookId: id, name: "OperationsNotesStatusSection", scope: "workbook", sheetName: "Operations", address: "Operations!H4:J7" });
  workbook.names.set("OperationsReconciliationSection", { workbookId: id, name: "OperationsReconciliationSection", scope: "workbook", sheetName: "Operations", address: "Operations!A15:D17" });
  return workbook;
}

function writeMonthlyPerformanceSheet(sheet, month, metrics) {
  const profit = metrics.received - metrics.spent;
  sheet.writeValues("A1:M2", [
    ["Transaction Date", "Job ID", "Truck ID", "Description", "Transaction Type", "Direction", "Cash Amount", "Actual Amount", "Payment Variance", "Reconciliation Note", "Transfer From/To", "Proof File", "Detail Notes"],
    [`2026-${month === "March" ? "03" : "04"}-01`, "204", "71-4653", "Company gas top-up", "company_gas_topup", "Outflow", 2211.21, 2211.21, 0, "", "Bank", "proof.pdf", "text note"]
  ]);
  sheet.writeValues("O1:AE2", [
    ["Invoice No", "Job ID", "Invoice Date", "Billed To", "Booking No", "Customer", "Job", "Container No", "Container Size", "Job Price", "Lifting On", "Lifting Off", "Total Lifting", "Other Fees", "Gross Billed", "W/H Tax", "Net Collect"],
    ["INV-001", "204", `2026-${month === "March" ? "03" : "04"}-01`, "ACME", "BK-001", "Customer A", "Job 204", "CONT-1", "20GP", 10000, 1000, 1000, 2000, 0, 12000, 360, 11640]
  ]);
  sheet.writeValues("AG1:AJ20", [
    [`${month} 2026 Summary`, "", "", ""],
    ["", "", "", ""],
    ["Cash Summary", "", "", ""],
    ["Metric", "Amount (THB)", "View", "Notes"],
    [`Cash received in ${month}`, metrics.received, "Cash in", `Actual cash/bank inflows received in ${month}.`],
    [`Cash spent in ${month}`, metrics.spent, "Cash out", `Actual cash/bank outflows paid in ${month}.`],
    ["Net cash movement", profit, "Cash net", "Cash received less cash spent."],
    ["", "", "", ""],
    ["P&L / Credit Float", "", "", ""],
    ["Billed / earned transport revenue", metrics.received + 45000, "Revenue", "Invoice revenue for jobs in the month."],
    ["Operating spend", metrics.spent, "Expense", "Cash operating spend."],
    ["Profit / loss", profit, "Profit", "Revenue less spend."],
    ["", "", "", ""],
    ["Spend Breakdown", "", "", ""],
    ["Company gas top-up", metrics.gas, "Expense", "Fuel-related spend."],
    ["Truck repair", metrics.repair, "Expense", "Maintenance spend."],
    ["Driver salary", metrics.salary, "Expense", "Payroll spend."],
    ["", "", "", ""],
    ["Checks / Interpretation", "", "", ""],
    ["Management takeaway", profit >= 0 ? "Profitable" : "Loss", "Status", profit >= 0 ? "Month remained profitable." : "Cash out exceeded cash in."]
  ]);
}

async function waitForHttp(url, timeoutMs) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // retry
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`Timed out waiting for ${url}`);
}

function parseRange(address) {
  const [start, end = start] = stripSheetName(address).replace(/\$/g, "").split(":");
  const startCell = parseCell(start);
  const endCell = parseCell(end);
  return { startRow: startCell.row, startCol: startCell.col, endRow: endCell.row, endCol: endCell.col, rowCount: endCell.row - startCell.row + 1, columnCount: endCell.col - startCell.col + 1 };
}

function stripSheetName(address) {
  return String(address).replace(/^'[^']+'!/, "").replace(/^[^!]+!/, "");
}

function makeSelection(workbookId, sheetName, address) {
  const range = parseRange(address);
  return {
    workbookId,
    sheetName,
    address,
    startCell: {
      workbookId,
      sheetName,
      address: `${columnName(range.startCol)}${range.startRow}`,
      row: range.startRow,
      column: range.startCol,
      rowIndex: range.startRow - 1,
      columnIndex: range.startCol - 1
    },
    endCell: {
      workbookId,
      sheetName,
      address: `${columnName(range.endCol)}${range.endRow}`,
      row: range.endRow,
      column: range.endCol,
      rowIndex: range.endRow - 1,
      columnIndex: range.endCol - 1
    },
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    cellCount: range.rowCount * range.columnCount,
    isSingleCell: range.rowCount === 1 && range.columnCount === 1
  };
}

function parseCell(cell) {
  const match = /^([A-Z]+)(\d+)$/i.exec(cell);
  if (!match) throw new Error(`Unsupported A1 cell: ${cell}`);
  return { col: columnIndex(match[1].toUpperCase()), row: Number(match[2]) };
}

function columnIndex(name) {
  let value = 0;
  for (const char of name) value = value * 26 + char.charCodeAt(0) - 64;
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
  return value;
}

function forEachMatrix(values, fn) {
  values.forEach((row, rowIndex) => row.forEach((value, columnIndex) => fn(value, rowIndex, columnIndex)));
}

function matrixCellCount(values) {
  return values.reduce((total, row) => total + row.length, 0);
}

function createZip(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const [name, data] of Object.entries(files)) {
    const nameBuffer = Buffer.from(name);
    const content = Buffer.isBuffer(data) ? data : Buffer.from(data);
    const crc = crc32(content);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(content.length, 18);
    local.writeUInt32LE(content.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    local.writeUInt16LE(0, 28);
    localParts.push(local, nameBuffer, content);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(content.length, 20);
    central.writeUInt32LE(content.length, 24);
    central.writeUInt16LE(nameBuffer.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuffer);
    offset += local.length + nameBuffer.length + content.length;
  }
  const centralDirectory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(centralDirectory.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, centralDirectory, end]);
}

function readZip(buffer) {
  const eocdOffset = findSignature(buffer, 0x06054b50);
  if (eocdOffset < 0) throw new Error("Invalid ZIP: missing end of central directory");
  const entryCount = buffer.readUInt16LE(eocdOffset + 10);
  const centralOffset = buffer.readUInt32LE(eocdOffset + 16);
  const entries = new Map();
  let cursor = centralOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (buffer.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error(`Invalid ZIP: bad central directory entry ${index}`);
    }
    const method = buffer.readUInt16LE(cursor + 10);
    if (method !== 0) throw new Error(`Unsupported ZIP compression method ${method}`);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const fileNameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLength = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.slice(cursor + 46, cursor + 46 + fileNameLength).toString("utf8");
    if (buffer.readUInt32LE(localOffset) !== 0x04034b50) {
      throw new Error(`Invalid ZIP: missing local header for ${name}`);
    }
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    entries.set(name, buffer.slice(dataOffset, dataOffset + compressedSize));
    cursor += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function findSignature(buffer, signature) {
  for (let index = buffer.length - 4; index >= 0; index -= 1) {
    if (buffer.readUInt32LE(index) === signature) return index;
  }
  return -1;
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff];
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const crcTable = Array.from({ length: 256 }, (_, index) => {
  let crc = index;
  for (let bit = 0; bit < 8; bit += 1) {
    crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
  }
  return crc >>> 0;
});

function xml(value) {
  return Buffer.from(value.replace(/\n\s*/g, ""), "utf8");
}

function cellHasValue(sheetXml, cell, value) {
  const escaped = escapeRegExp(escapeXml(value ?? ""));
  const inlinePattern = new RegExp(`<c r="${cell}"[^>]*><is><t>${escaped}</t></is></c>`);
  const valuePattern = new RegExp(`<c r="${cell}"[^>]*><v>${escapeRegExp(value)}</v></c>`);
  return inlinePattern.test(sheetXml) || valuePattern.test(sheetXml);
}

function cellHasFormula(sheetXml, cell, formula) {
  const normalized = String(formula ?? "").replace(/^=/, "");
  const formulaPattern = new RegExp(`<c r="${cell}"[^>]*><f>${escapeRegExp(escapeXml(normalized))}</f></c>`);
  return formulaPattern.test(sheetXml);
}

function cellStyleId(sheetXml, cell) {
  const match = new RegExp(`<c r="${cell}"[^>]* s="([^"]+)"`).exec(sheetXml);
  return match?.[1];
}

function argb(color) {
  const normalized = String(color).replace(/^#/, "").toUpperCase();
  return normalized.length === 8 ? normalized : `FF${normalized}`;
}

function escapeXml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("\"", "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function unescapeXml(value) {
  return String(value).replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&quot;", "\"").replaceAll("&apos;", "'").replaceAll("&amp;", "&");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function stableJson(value) {
  return JSON.stringify(value, Object.keys(value).sort());
}

function countBy(items, fn) {
  return items.reduce((counts, item) => {
    const key = fn(item);
    counts[key] = (counts[key] ?? 0) + 1;
    return counts;
  }, {});
}

function stableHash(value) {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex").slice(0, 16);
}

function fixedNow() {
  return "2026-06-16T00:00:00.000Z";
}

function readArg(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function hasArg(name) {
  return process.argv.includes(name);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
