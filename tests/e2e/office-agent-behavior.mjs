#!/usr/bin/env node
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-office-agent-behavior-"));
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

const scenarios = [
  {
    id: "workbook-about",
    category: "workbook overview",
    prompt: "Can you look into Office Agent Behavior.xlsx file, what is it about?",
    input: { request: "Can you look into Office Agent Behavior.xlsx file, what is it about?" }
  },
  {
    id: "xlsx-file-overview",
    category: "workbook overview",
    prompt: "Can you look into transactions.xlsx?",
    input: { request: "Can you look into transactions.xlsx?" }
  },
  {
    id: "sheet-count",
    category: "workbook overview",
    prompt: "How many sheets are in this Excel file?",
    input: { request: "How many sheets are in this Excel file?" }
  },
  {
    id: "list-sheet-purpose",
    category: "workbook overview",
    prompt: "List each sheet and what it seems to contain.",
    input: { request: "List each sheet and what it seems to contain." }
  },
  {
    id: "blank-sheets",
    category: "workbook overview",
    prompt: "Which sheets are blank or mostly empty?",
    input: { request: "Which sheets are blank or mostly empty?", mode: "find" }
  },
  {
    id: "active-sheet-analysis",
    category: "sheet reading",
    prompt: "Analyze the active sheet.",
    input: { request: "Analyze the active sheet." }
  },
  {
    id: "sales-sheet-summary",
    category: "sheet reading",
    prompt: "Can you look into Sales sheet, what is it about?",
    input: { request: "Can you look into Sales sheet, what is it about?", target: { sheetName: "Sales" } }
  },
  {
    id: "transactions-columns",
    category: "sheet reading",
    prompt: "What columns does the Transactions table have?",
    input: { request: "What columns does the Transactions table have?", target: { tableName: "Transactions" } }
  },
  {
    id: "sales-example-rows",
    category: "sheet reading",
    prompt: "Show me a few example rows from the Sales sheet.",
    input: { request: "Show me a few example rows from Sales", target: { sheetName: "Sales", range: "A1:E6" } }
  },
  {
    id: "sparse-range-token-saver",
    category: "token efficiency",
    prompt: "Read the sparse planning area without sending empty cells.",
    input: { request: "Read the sparse planning area without sending empty cells.", target: { sheetName: "Sparse", range: "A1:J10" } },
    expected: {
      resultType: "answer",
      shouldMutateWorkbook: false,
      resolvedTarget: { sheetName: "Sparse", range: "A1:J10" },
      mustReturnSparseRows: true,
      mustReturnEmptySummary: true,
      mustNotReturnDenseRows: true
    },
    budgets: {
      maxToolCalls: 2,
      maxPayloadBytes: 12000,
      maxLatencyMs: 4000
    }
  },
  {
    id: "compare-jan-feb",
    category: "sheet comparison",
    prompt: "Can you compare both sheets, January and February, how is it going?",
    input: { request: "Can you compare both sheets, January and February, how is it going?" }
  },
  {
    id: "compare-report-totals",
    category: "sheet comparison",
    prompt: "Which month has higher total revenue, January or February?",
    input: { request: "Which month has higher total revenue, January or February?" }
  },
  {
    id: "mock-data-blank-sheet",
    category: "simple edit",
    prompt: "Can you add mock data into the Blank sheet?",
    input: {
      request: "Add mock sales data to the blank sheet.",
      mode: "preview_update",
      target: { sheetName: "Blank", range: "A1:E6" },
      values: {
        values: [
          ["Date", "Customer", "Product", "Amount", "Status"],
          ["2026-06-01", "Acme Co", "Consulting", 1200, "Open"],
          ["2026-06-02", "Northwind", "Support", 450, "Closed"],
          ["2026-06-03", "Contoso", "Implementation", 3200, "Open"],
          ["2026-06-04", "Fabrikam", "Training", 800, "Open"],
          ["2026-06-05", "Tailspin", "Support", 650, "Closed"]
        ]
      }
    }
  },
  {
    id: "expense-tracker",
    category: "simple edit",
    prompt: "Create a small office expense tracker on the blank sheet.",
    input: {
      request: "Create a small office expense tracker on the blank sheet.",
      mode: "preview_update",
      target: { sheetName: "Blank", range: "A8:D13" },
      values: {
        values: [
          ["Date", "Category", "Vendor", "Amount"],
          ["2026-06-01", "Meals", "Cafe One", 42.5],
          ["2026-06-02", "Travel", "Metro", 16],
          ["2026-06-03", "Software", "SaaS Co", 99],
          ["2026-06-04", "Office", "Supply Store", 61.25],
          ["Total", "", "", 218.75]
        ]
      }
    }
  },
  {
    id: "header-format",
    category: "simple edit",
    prompt: "Format the header row on the Sales sheet.",
    input: { request: "Format the header row on the Sales sheet.", target: { sheetName: "Sales", range: "A1:E1" } }
  },
  {
    id: "find-formulas",
    category: "formula basics",
    prompt: "Find formulas in the FormulaSheet sheet.",
    input: { request: "Find formulas in the FormulaSheet sheet.", target: { sheetName: "FormulaSheet" } }
  },
  {
    id: "duplicate-formula",
    category: "formula basics",
    prompt: "Duplicate the formula from D2 down to D10.",
    input: {
      request: "Duplicate the formula from D2 down to D10.",
      mode: "preview_update",
      target: { sheetName: "FormulaSheet", range: "D3:D10" },
      values: { values: Array.from({ length: 8 }, (_item, index) => [`=B${index + 3}+C${index + 3}`]) }
    }
  },
  {
    id: "formula-explain",
    category: "formula basics",
    prompt: "Explain what the formula area is doing.",
    input: { request: "Explain what the formula area is doing.", target: { sheetName: "FormulaSheet", range: "A1:D10" } }
  },
  {
    id: "summary-sheet",
    category: "normal office workflow",
    prompt: "Create a simple summary sheet for this workbook.",
    input: {
      request: "Create a simple summary sheet for this workbook.",
      mode: "preview_update",
      target: { sheetName: "Summary", range: "A1:C6" },
      values: {
        values: [
          ["Metric", "Value", "Notes"],
          ["Sheet count", 7, "From workbook map"],
          ["Sales rows", 5, "Sales sample"],
          ["Transactions rows", 4, "Transactions table"],
          ["Formula sheet", "Yes", "Contains totals"],
          ["Next review", "2026-06-30", "Mock office workflow"]
        ]
      }
    }
  },
  {
    id: "workbook-status",
    category: "workbook overview",
    prompt: "Can you connect to Excel and tell me what workbook is active?",
    input: { request: "Can you connect to Excel and tell me what workbook is active?", mode: "status" }
  },
  {
    id: "find-tables",
    category: "workbook overview",
    prompt: "Which tables are in this workbook?",
    input: { request: "Which tables are in this workbook?" }
  },
  {
    id: "named-ranges",
    category: "workbook overview",
    prompt: "Are there named ranges in this workbook?",
    input: { request: "Are there named ranges in this workbook?" }
  },
  {
    id: "sales-totals",
    category: "sheet reading",
    prompt: "What is the total amount in the Sales sheet sample?",
    input: { request: "What is the total amount in the Sales sheet sample?", target: { sheetName: "Sales", range: "A1:E6" } }
  },
  {
    id: "sales-open-items",
    category: "sheet reading",
    prompt: "Which Sales rows are still open?",
    input: { request: "Which Sales rows are still open?", target: { sheetName: "Sales", range: "A1:E6" } }
  },
  {
    id: "analyze-sales-report",
    category: "sheet report",
    prompt: "Can you analyze sheet Sales and show me a report?",
    input: { request: "Can you analyze sheet Sales and show me a report?", target: { sheetName: "Sales", range: "A1:E6" } }
  },
  {
    id: "analyze-january-report",
    category: "sheet report",
    prompt: "Can you analyze sheet January and show me report?",
    input: { request: "Can you analyze sheet January and show me report?", target: { sheetName: "January", range: "A1:C5" } }
  },
  {
    id: "report-active-sheet",
    category: "sheet report",
    prompt: "Can you analyze this sheet and show me report?",
    input: { request: "Can you analyze this sheet and show me report?" }
  },
  {
    id: "january-summary",
    category: "sheet reading",
    prompt: "Summarize the January sheet.",
    input: { request: "Summarize the January sheet.", target: { sheetName: "January" } }
  },
  {
    id: "february-summary",
    category: "sheet reading",
    prompt: "Summarize the February sheet.",
    input: { request: "Summarize the February sheet.", target: { sheetName: "February" } }
  },
  {
    id: "compare-expense",
    category: "sheet comparison",
    prompt: "Compare January and February expenses.",
    input: { request: "Compare January and February expenses." }
  },
  {
    id: "compare-profit",
    category: "sheet comparison",
    prompt: "Compare January and February profit.",
    input: { request: "Compare January and February profit." }
  },
  {
    id: "compare-open-deals",
    category: "sheet comparison",
    prompt: "Which month has more open deals, January or February?",
    input: { request: "Which month has more open deals, January or February?" }
  },
  {
    id: "compare-both-sheets-explicit-chain",
    category: "sheet comparison",
    prompt: "Can you compare both sheets, January and February, and tell me how it is going?",
    input: { request: "Can you compare both sheets, January and February, and tell me how it is going?" }
  },
  {
    id: "sections-in-sales",
    category: "sheet structure",
    prompt: "How many sections do we have in the Sales sheet?",
    input: { request: "How many sections do we have in the Sales sheet?", target: { sheetName: "Sales" } }
  },
  {
    id: "sections-in-formula-sheet",
    category: "sheet structure",
    prompt: "How many sections do we have in FormulaSheet?",
    input: { request: "How many sections do we have in FormulaSheet?", target: { sheetName: "FormulaSheet" } }
  },
  {
    id: "sections-active-sheet",
    category: "sheet structure",
    prompt: "How many sections do we have in this sheet?",
    input: { request: "How many sections do we have in this sheet?" }
  },
  {
    id: "operations-section-inventory",
    category: "sheet structure",
    prompt: "What sections are on the Operations sheet?",
    input: { request: "What sections are on the Operations sheet?", target: { sheetName: "Operations" } }
  },
  {
    id: "operations-invoice-section",
    category: "sheet reading",
    prompt: "Show example rows from the invoice section on Operations.",
    input: { request: "Show example rows from the invoice section on Operations" }
  },
  {
    id: "operations-status-section-update",
    category: "simple edit",
    prompt: "Update the status section owner note on Operations.",
    input: {
      request: "Update the notes/status section on Operations.",
      mode: "preview_update",
      target: { sheetName: "Operations", range: "H5:J5" },
      values: { values: [["A. Chen", "Review complete", "2026-06-20"]] }
    }
  },
  {
    id: "duplicate-latest-sheet-template",
    category: "template cleanup",
    prompt: "Can you duplicate latest sheet, remove data and keep only template?",
    input: { request: "Can you duplicate latest sheet, remove data and keep only template?" }
  },
  {
    id: "duplicate-february-template",
    category: "template cleanup",
    prompt: "Can you duplicate February sheet, remove data and keep only template?",
    input: { request: "Can you duplicate February sheet, remove data and keep only template?", target: { sheetName: "February" } }
  },
  {
    id: "change-one-sales-cell",
    category: "simple edit",
    prompt: "Can you update the Status column, row 2, according to my input: Reviewed?",
    input: {
      request: "Change the first Sales status to Reviewed.",
      target: { sheetName: "Sales", range: "E2" },
      values: { values: [["Reviewed"]] }
    }
  },
  {
    id: "update-amount-cell",
    category: "simple edit",
    prompt: "Can you update Amount column, row 3, according to my input: 525?",
    input: {
      request: "Can you update Amount column, row 3, according to my input: 525?",
      target: { sheetName: "Sales", range: "D3" },
      values: { values: [[525]] }
    }
  },
  {
    id: "update-row-values",
    category: "simple edit",
    prompt: "Can you update row 4 according to my input?",
    input: {
      request: "Can you update row 4 according to my input?",
      target: { sheetName: "Sales", range: "A4:E4" },
      values: { values: [["2026-01-08", "Contoso", "Implementation", 3500, "Reviewed"]] }
    }
  },
  {
    id: "append-sales-row",
    category: "simple edit",
    prompt: "Can you add one new row according to my input?",
    steps: [
      {
        label: "preview",
        input: {
          request: "Append one new sales row to the Transactions table.",
          mode: "preview_update",
          target: { tableName: "Transactions" },
          values: { rows: [["2026-01-15", "Globex", "Support", 975, "Open"]] }
        }
      }
    ]
  },
  {
    id: "add-notes-block",
    category: "simple edit",
    prompt: "Add a small notes block to Summary.",
    input: {
      request: "Add a small notes block to Summary.",
      target: { sheetName: "Summary", range: "E1:F4" },
      values: { values: [["Notes", "Owner"], ["Review sales", "Finance"], ["Check formulas", "Ops"], ["Send update", "Admin"]] }
    }
  },
  {
    id: "formula-read-specific",
    category: "formula basics",
    prompt: "Can you look into FormulaSheet D2:D5 and explain the formulas?",
    input: { request: "Can you look into FormulaSheet D2:D5 and explain the formulas?", target: { sheetName: "FormulaSheet", range: "D2:D5" } }
  },
  {
    id: "formula-check-errors",
    category: "formula basics",
    prompt: "Check whether formulas look broken.",
    input: { request: "Check whether formulas look broken.", target: { sheetName: "FormulaSheet", range: "A1:D10" } }
  },
  {
    id: "formula-add-total-row",
    category: "formula basics",
    prompt: "Add a total row under the formula table.",
    input: {
      request: "Add a total row under the formula table.",
      mode: "preview_update",
      target: { sheetName: "FormulaSheet", range: "A11:D11" },
      values: { values: [["Total", "=SUM(B2:B10)", "=SUM(C2:C10)", "=SUM(D2:D10)"]] }
    }
  },
  {
    id: "office-multi-step-summary",
    category: "normal office workflow",
    prompt: "Can you analyze sheet Sales and show me a report?",
    steps: [
      { label: "inspect-sales", input: { request: "Show me a few example rows from Sales", target: { sheetName: "Sales", range: "A1:E6" } } },
      {
        label: "write-summary",
        input: {
          request: "Create a basic summary block.",
          target: { sheetName: "Summary", range: "A8:C11" },
          values: { values: [["Summary", "Value", "Source"], ["Sales rows", 5, "Sales"], ["Open items", 3, "Sales"], ["Formula rows", 9, "FormulaSheet"]] }
        }
      }
    ]
  },
  {
    id: "office-clarify-ambiguous-sheet",
    category: "targeting",
    prompt: "Analyze the month sheet.",
    input: { request: "Analyze the month sheet." }
  },
  {
    id: "office-explicit-range",
    category: "targeting",
    prompt: "Can you look into January!A1:C5?",
    input: { request: "Can you look into January!A1:C5 actual values" }
  },
  {
    id: "office-narrow-target-after-ambiguous",
    category: "targeting",
    prompt: "Analyze January after a broad monthly request.",
    steps: [
      { label: "broad", input: { request: "Analyze the month sheet." } },
      { label: "narrow", input: { request: "Analyze January", target: { sheetName: "January" } } }
    ]
  }
];

async function main() {
  const scenarioCatalog = loadScenarioCatalog();
  const selectedIds = new Set((readArg("--scenarios") ?? "all").split(",").map((item) => item.trim()).filter(Boolean));
  const selected = selectedIds.has("all") ? scenarioCatalog : scenarioCatalog.filter((scenario) => selectedIds.has(scenario.id) || selectedIds.has(scenario.category));
  const server = spawn(process.execPath, ["apps/mcp-server/dist/index.js", "--standalone", "--agent-name", "office-agent-behavior"], {
    cwd: repoRoot,
    stdio: ["pipe", "pipe", "pipe"],
    env: {
      ...process.env,
      OPEN_WORKBOOK_HOST: "127.0.0.1",
      OPEN_WORKBOOK_PORT: String(backendPort),
      OPEN_WORKBOOK_ADDIN_PATH: "/addin",
      OPEN_WORKBOOK_STATE_DIR: path.join(tempRoot, "state"),
      OPEN_WORKBOOK_BACKUP_DIR: path.join(tempRoot, "backups"),
      OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1"
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

    const report = buildReport({ toolNames, results, artifactDir: artifactsDir, scenarioSource: scenarioCatalog.source });
    writeFileSync(path.join(artifactsDir, "office-agent-behavior-report.json"), JSON.stringify(report, null, 2));
    writeFileSync(path.join(artifactsDir, "office-agent-behavior-report.md"), renderReport(report));
    writeFileSync(path.join(artifactsDir, "mcp-transcript.jsonl"), transcript.map((event) => JSON.stringify(event)).join("\n"));
    console.log(renderReport(report));
    console.log(`\nSaved office agent behavior artifacts: ${artifactsDir}`);
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

async function runScenario({ mcp, addin, scenario, agentOutputSchema }) {
  const scenarioDir = path.join(artifactsDir, scenario.id);
  mkdirSync(scenarioDir, { recursive: true });
  const restoreWorkbookAfterScenario = scenario.id === "closed-workbook-status-error";
  if (restoreWorkbookAfterScenario) {
    addin.setActiveWorkbook(undefined);
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  const before = addin.workbook.summary();
  addin.setSelection(scenario.selection ?? { sheetName: "Sales", address: "A1" });
  const started = performance.now();
  const chainStart = transcript.length;
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
  const toolChain = transcript.slice(chainStart);
  const result = { steps: stepResults };
  const observation = observeScenario({ scenario, result, error, before, after, toolChain, elapsedMs: Math.round(performance.now() - started) });
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
  return ["mock-data-blank-sheet", "expense-tracker", "summary-sheet", "add-notes-block", "office-multi-step-summary"].includes(scenario.id);
}

function scenarioSteps(scenario) {
  if (Array.isArray(scenario.steps)) {
    return scenario.steps;
  }
  return [{ label: "request", input: scenario.input }];
}

function loadScenarioCatalog() {
  if (hasArg("--legacy-scenarios") || process.env.OPEN_WORKBOOK_OFFICE_AGENT_LEGACY_SCENARIOS === "1") {
    scenarios.source = "legacy-inline";
    return scenarios;
  }
  const scenarioFile = readArg("--scenario-file") ?? process.env.OPEN_WORKBOOK_OFFICE_AGENT_SCENARIO_FILE ?? productionScenariosPath;
  const loaded = JSON.parse(readFileSync(scenarioFile, "utf8"));
  if (!Array.isArray(loaded)) {
    throw new Error(`Scenario file must contain a JSON array: ${scenarioFile}`);
  }
  loaded.source = scenarioFile;
  return loaded;
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
  return input;
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
  return result[field];
}

function rememberStepContext(stepContext, label, result) {
  stepContext.byLabel ??= {};
  stepContext.byLabel[label] = result;
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

async function callTool(client, name, args, outputSchema) {
  const started = performance.now();
  const raw = await client.request("tools/call", { name, arguments: args });
  const wallMs = Math.round(performance.now() - started);
  const text = raw.content?.find((item) => item.type === "text")?.text;
  const parsed = text ? JSON.parse(text) : raw.structuredContent;
  const event = {
    at: new Date().toISOString(),
    tool: name,
    args,
    wallMs,
    isError: raw.isError === true,
    telemetry: parsed?.telemetry,
    status: parsed?.status,
    nextAction: parsed?.nextAction,
    summary: parsed?.summary
  };
  if (outputSchema?.properties?.telemetry?.properties) {
    const schemaKeys = outputSchema.properties.telemetry.properties;
    event.undeclaredTelemetryKeys = Object.keys(parsed?.telemetry ?? {}).filter((key) => !(key in schemaKeys));
  }
  transcript.push(event);
  return parsed;
}

function observeScenario({ scenario, result, error, before, after, toolChain, elapsedMs }) {
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
  const expectationIssues = evaluateScenarioExpectations({ scenario, effective, workbookChanged, usage });
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
      summary: call.summary,
      telemetry: call.telemetry
    })),
    usage,
    resultSummary: effective?.summary,
    artifactDir: path.join(artifactsDir, scenario.id)
  };
}

function summarizeToolChainUsage(toolChain, elapsedMs) {
  const telemetryItems = toolChain.map((call) => call.telemetry ?? {});
  const sum = (field) => telemetryItems.reduce((total, telemetry) => total + (Number(telemetry[field]) || 0), 0);
  const metadataCacheStatuses = countBy(telemetryItems, (telemetry) => telemetry.metadataCacheStatus ?? "unknown");
  return {
    toolCallCount: toolChain.length,
    scenarioElapsedMs: elapsedMs,
    toolWallMs: toolChain.reduce((total, call) => total + (Number(call.wallMs) || 0), 0),
    backendElapsedMs: sum("elapsedMs"),
    payloadBytes: sum("payloadBytes"),
    estimatedTokens: sum("estimatedTokens"),
    estimatedTokensSaved: sum("estimatedTokensSaved"),
    internalCallCount: sum("internalCallCount"),
    internalReadCount: sum("internalReadCount"),
    fullReadCellCount: sum("fullReadCellCount"),
    cacheHits: telemetryItems.filter((telemetry) => telemetry.cacheHit === true).length,
    autoAppliedCount: telemetryItems.filter((telemetry) => telemetry.autoApplied === true).length,
    metadataCacheStatuses
  };
}

function effectiveScenarioResult(result) {
  const last = result?.steps?.at(-1);
  return last?.applied ?? last?.result ?? result;
}

function buildReport({ toolNames, results, artifactDir, scenarioSource }) {
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
      estimatedTokens: result.usage?.estimatedTokens ?? 0,
      payloadBytes: result.usage?.payloadBytes ?? 0
    }));
  return {
    toolCallCount: sum("toolCallCount"),
    scenarioElapsedMs: sum("scenarioElapsedMs"),
    toolWallMs: sum("toolWallMs"),
    backendElapsedMs: sum("backendElapsedMs"),
    payloadBytes: sum("payloadBytes"),
    estimatedTokens: sum("estimatedTokens"),
    estimatedTokensSaved: sum("estimatedTokensSaved"),
    internalCallCount: sum("internalCallCount"),
    internalReadCount: sum("internalReadCount"),
    fullReadCellCount: sum("fullReadCellCount"),
    cacheHits: sum("cacheHits"),
    autoAppliedCount: sum("autoAppliedCount"),
    slowest
  };
}

function evaluateScenarioExpectations({ scenario, effective, workbookChanged, usage }) {
  const issues = [];
  const expected = scenario.expected ?? {};
  const budgets = scenario.budgets ?? {};
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
  if (expected.shouldAskClarification === true && !["AMBIGUOUS_TARGET", "NEEDS_INPUT"].includes(effective?.status)) {
    issues.push(`expected clarification, got ${effective?.status ?? "no result"}`);
  }
  if (expected.shouldFailGracefully === true && effective?.status === "SUCCESS") {
    issues.push("expected graceful failure or clarification, got SUCCESS");
  }
  if (budgets.maxToolCalls !== undefined && usage.toolCallCount > budgets.maxToolCalls) {
    issues.push(`tool calls ${usage.toolCallCount} exceeded budget ${budgets.maxToolCalls}`);
  }
  if (budgets.maxPayloadBytes !== undefined && usage.payloadBytes > budgets.maxPayloadBytes) {
    issues.push(`payload bytes ${usage.payloadBytes} exceeded budget ${budgets.maxPayloadBytes}`);
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
    `- Scenario elapsed: ${report.usageTotals.scenarioElapsedMs} ms`,
    `- Tool wall time: ${report.usageTotals.toolWallMs} ms`,
    `- Backend elapsed: ${report.usageTotals.backendElapsedMs} ms`,
    `- Payload bytes: ${report.usageTotals.payloadBytes}`,
    `- Estimated tokens: ${report.usageTotals.estimatedTokens}`,
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
      ? report.usageTotals.slowest.map((item) => `- ${item.scenarioId}: ${item.elapsedMs} ms, ${item.toolCallCount} calls, ${item.estimatedTokens} estimated tokens, ${item.payloadBytes} bytes`)
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
    lines.push(`- Scenario elapsed: ${result.usage.scenarioElapsedMs} ms`);
    lines.push(`- Tool wall time: ${result.usage.toolWallMs} ms`);
    lines.push(`- Backend elapsed: ${result.usage.backendElapsedMs} ms`);
    lines.push(`- Payload bytes: ${result.usage.payloadBytes}`);
    lines.push(`- Estimated tokens: ${result.usage.estimatedTokens}`);
    lines.push(`- Estimated tokens saved: ${result.usage.estimatedTokensSaved}`);
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
    `Tool wall time: ${observation.usage.toolWallMs} ms`,
    `Backend elapsed: ${observation.usage.backendElapsedMs} ms`,
    `Payload bytes: ${observation.usage.payloadBytes}`,
    `Estimated tokens: ${observation.usage.estimatedTokens}`,
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
      case "workbook.snapshot_ranges":
        return this.workbook.snapshotRanges(params.ranges);
      case "table.list":
        return { ok: true, tables: [...this.workbook.tables.values()].map((table) => table.info()) };
      case "table.get_info":
        return { ok: true, info: this.workbook.table(params.tableName).info() };
      case "table.append_rows":
        return this.workbook.table(params.tableName).appendRows(params.values);
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

  clone(name) {
    const next = new FakeSheet(this.workbookId, name);
    for (const [key, cell] of this.cells.entries()) {
      next.cells.set(key, { ...cell });
    }
    return next;
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
    return { workbookId: this.workbookId, sheetName: this.name, address: `A1:${columnName(maxCol)}${maxRow}`, rowCount: maxRow, columnCount: maxCol };
  }

  nonEmptyCellCount() {
    return [...this.cells.values()].filter((cell) => cell.value !== null || cell.formula !== null).length;
  }

  valuesForSummary() {
    return [...this.cells.entries()].filter(([, cell]) => cell.value !== null || cell.formula !== null).sort();
  }

  cell(row, col) {
    const key = `${row}:${col}`;
    if (!this.cells.has(key)) this.cells.set(key, { value: null, formula: null });
    return this.cells.get(key);
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
  const [start, end = start] = address.replace(/\$/g, "").split(":");
  const startCell = parseCell(start);
  const endCell = parseCell(end);
  return { startRow: startCell.row, startCol: startCell.col, endRow: endCell.row, endCol: endCell.col, rowCount: endCell.row - startCell.row + 1, columnCount: endCell.col - startCell.col + 1 };
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
