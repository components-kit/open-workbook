#!/usr/bin/env node
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  assertOperationCoverageReport,
  buildOperationCoverageReport,
  renderOperationCoverageMarkdown
} from "./lib/operation-coverage.mjs";

const repoRoot = path.resolve(import.meta.dirname, "../..");
const artifactDir = process.env.OPEN_WORKBOOK_E2E_REPORT_DIR ?? path.join(tmpdir(), "open-workbook-e2e-reports");
mkdirSync(artifactDir, { recursive: true });
const scenarioFile = path.join(repoRoot, "tests", "e2e", "fixtures", "office-agent-production-scenarios.json");
const operationCoverage = await buildOperationCoverageReport({ repoRoot, scenarioFile });
assertOperationCoverageReport(operationCoverage);

const report = {
  title: "Open Workbook MCP E2E Test Report",
  generatedAt: new Date().toISOString(),
  operationCoverage,
  releaseTarget: {
    codexModel: process.env.OPEN_WORKBOOK_E2E_CODEX_MODEL ?? "gpt-5.4-mini",
    codexReasoning: process.env.OPEN_WORKBOOK_E2E_CODEX_REASONING ?? "low",
    apiKeyRequired: false,
    liveExcelReleaseHosts: ["macOS desktop Excel", "Windows desktop Excel"]
  },
  prerequisites: [
    {
      lane: "agent-surface",
      required: ["Built MCP server at apps/mcp-server/dist/index.js", "Node.js >=20.11"],
      notRequired: ["Excel", "Codex", "OpenAI API key"]
    },
    {
      lane: "codex-agent",
      required: ["Built CLI and MCP assets", "Signed-in Codex CLI session", "Codex access to gpt-5.4-mini"],
      notRequired: ["OpenAI API key", "Excel"]
    },
    {
      lane: "live-excel",
      required: ["Desktop Excel", "Sideloaded Open Workbook add-in", "Connected local MCP/backend runtime"],
      notRequired: ["OpenAI API key"]
    }
  ],
  lanes: [
    {
      name: "test:e2e:agent-surface",
      command: "pnpm run test:e2e:agent-surface",
      status: scriptExists("tests/e2e/agent-surface.mjs") ? "implemented" : "missing",
      startsServices: ["Open Workbook MCP stdio server"],
      coverage: [
        "Default MCP tools/list exposes only excel.agent.run",
        "excel.agent.run accepts the public request schema",
        "Status mode returns a structured status, nextAction, and telemetry without requiring Excel",
        "Default resources/list remains available"
      ],
      reliabilityAssertions: [
        "Primitive operation tools stay hidden from the public surface",
        "Agent output includes payload and estimated token telemetry"
      ],
      gaps: [
        "Workbook prepare/find/update agent behavior is covered by unit and fake-host tests; this lane only guards default MCP exposure."
      ]
    },
    {
      name: "test:e2e:agent-workflow",
      command: "pnpm run test:e2e:agent-workflow",
      status: scriptExists("tests/e2e/agent-workflow.mjs") ? "implemented" : "missing",
      startsServices: ["Open Workbook MCP stdio server", "Fake Excel add-in WebSocket client"],
      coverage: [
        "Default MCP tools/list exposes only excel.agent.run while connected to a workbook",
        "Agent prepare builds workbook metadata and repeated prepare reuses the cache",
        "Agent find obeys maxExamples and compact payload budgets",
        "Agent answer uses one targeted internal read and avoids broad workbook reads",
        "Agent preview/apply requires confirmationToken before mutation",
        "Agent auto mode applies clearly scoped low-risk value edits after preview checks",
        "Agent auto mode blocks formula-sensitive requests from plain value writes",
        "Agent validate succeeds after an applied value-only update"
      ],
      reliabilityAssertions: [
        "Repeated prepare reports metadataCacheStatus=hit",
        "Answer reports internalReadCount=1 and bounded fullReadCellCount",
        "Find payload stays under the deterministic fixture budget",
        "Apply without confirmationToken does not mutate",
        "Auto-applied edits report telemetry.autoApplied and safetyDecision",
        "Formula-sensitive auto requests return manual-review safety telemetry"
      ],
      gaps: [
        "This lane uses a deterministic fake host; live Excel fidelity remains covered by live gates."
      ]
    },
    {
      name: "test:e2e:agent:core",
      command: "pnpm run test:e2e:agent:core",
      status: scriptExists("tests/e2e/codex-agent.mjs") ? "implemented" : "missing",
      startsServices: ["Codex exec", "Open Workbook MCP via temporary Codex config", "Scenario fake Excel add-in"],
      coverage: [
        "Signed-in Codex CLI can initialize Open Workbook MCP",
        "Codex uses gpt-5.4-mini with low reasoning unless overridden",
        "Bundled open-workbook-skills skill guidance is loaded into every scenario prompt",
        "Every scenario expects workbook identity plus runtime discovery before mutation",
        "Large table reorder/filter/sort/append/update choice quality",
        "Multi-agent lock and conflict decisions",
        "Forbidden behavior checks for per-cell loops, broad large-table reads, manual large-table rewrites, destructive column operations, and writes during lock-only scenarios",
        "JSONL transcript plus parsed decision analysis artifacts per scenario"
      ],
      reliabilityAssertions: [
        "No OpenAI API key is required",
        "Global Codex config is not modified",
        "MCP server is required; missing server fails the run",
        "Assertions check tool families and safety invariants, not exact model wording",
        "OPEN_WORKBOOK_E2E_AGENT_SCENARIOS can run all, core, quality, smoke, or a comma-separated subset",
        "OPEN_WORKBOOK_E2E_AGENT_REPEAT can repeat scenarios for consistency checks"
      ],
      gaps: [
        "Agent decisions are validated against a deterministic fake host; live Excel fidelity remains covered by live gates.",
        "Transcript parser is intentionally tolerant, but may need updates if Codex JSONL event shapes change."
      ]
    },
    {
      name: "test:e2e:agent:quality",
      command: "pnpm run test:e2e:agent:quality",
      status: scriptExists("tests/e2e/codex-agent.mjs") ? "implemented" : "missing",
      startsServices: ["Codex exec", "Open Workbook MCP via temporary Codex config", "Scenario fake Excel add-in"],
      coverage: [
        "Natural-language agent decisions for sheet/range/formula creation",
        "Formula repair without formula-to-values conversion",
        "Template creation/fill/validation and style repair decisions",
        "Pivot/chart creation and partial capability warning behavior",
        "Snapshot/diff/backup/rollback-preview decisions",
        "Report-only quality diagnostics for cheap-model workflow gaps",
        "JSONL transcript plus parsed decision analysis artifacts per scenario"
      ],
      reliabilityAssertions: [
        "Runs in report-only mode by default so quality regressions are visible without blocking the core release gate.",
        "Uses the same fake host and transcript analysis as the core lane."
      ],
      gaps: [
        "Cheap-model quality scenarios may fail while skill and prompt guidance improves; failures should be reviewed before customer-facing workflow claims."
      ]
    },
    {
      name: "test:e2e:agent:quality:compare",
      command: "pnpm run test:e2e:agent:quality:compare",
      status: scriptExists("tests/e2e/agent-model-matrix.mjs") ? "implemented" : "missing",
      startsServices: ["Sequential Codex agent quality runs per configured model profile"],
      coverage: [
        "Report-only comparison between cheap and frontier model profiles",
        "Pass counts and failure-category counts per profile",
        "Artifacts for each underlying Codex agent quality run",
        "OPEN_WORKBOOK_E2E_CHEAP_MODEL and OPEN_WORKBOOK_E2E_FRONTIER_MODEL overrides"
      ],
      reliabilityAssertions: [
        "The comparison is diagnostic by default and does not replace the core safety gate.",
        "Use --strict or OPEN_WORKBOOK_E2E_AGENT_MATRIX_STRICT=1 only when a team intentionally wants quality comparisons to fail CI."
      ],
      gaps: [
        "Model matrix results are sensitive to Codex service latency and model availability."
      ]
    },
    {
      name: "test:e2e:live:mac",
      command: "pnpm run test:e2e:live:mac",
      status: scriptExists("tests/e2e/live-smoke.mjs") ? "opt-in connectivity smoke" : "missing",
      startsServices: ["Expected: desktop Excel with add-in and local backend"],
      coverage: [
        "Backend /status reachability",
        "Connected Excel add-in session",
        "Active workbook metadata",
        "Optional --deep scratch-sheet range write/formula/snapshot/diff/rollback-preview smoke"
      ],
      gaps: ["Deep host smoke must still be run on real macOS Excel before host-readiness claims; broader table/template/pivot/chart live automation remains future coverage."]
    },
    {
      name: "test:e2e:live:windows",
      command: "pnpm run test:e2e:live:windows",
      status: scriptExists("tests/e2e/live-smoke.mjs") ? "opt-in connectivity smoke" : "missing",
      startsServices: ["Expected: desktop Excel with add-in and local backend"],
      coverage: [
        "Backend /status reachability",
        "Connected Excel add-in session",
        "Active workbook metadata",
        "Optional --deep scratch-sheet range write/formula/snapshot/diff/rollback-preview smoke"
      ],
      gaps: ["Deep host smoke must still be run on real Windows Excel before host-readiness claims; broader table/template/pivot/chart live automation remains future coverage."]
    }
  ],
  recommendedRunOrder: [
    "pnpm run build",
    "pnpm run test:e2e:report",
    "pnpm run test:e2e:agent-surface",
    "pnpm run test:e2e:agent-workflow",
    "pnpm run test:e2e:agent:core",
    "pnpm run test:e2e:agent:quality",
    "pnpm run test:e2e:agent:quality:compare",
    "pnpm run test:e2e:live:mac",
    "pnpm run test:e2e:live:windows"
  ],
  releaseGatePolicy: [
    "Default one-tool public-surface E2E must pass before every release candidate.",
    "Codex core signed-in E2E must pass before customer-facing large-workbook and multi-agent safety claims.",
    "Codex quality E2E is report-only by default and should be reviewed before claiming complex workflow quality.",
    "Both live Excel gates must pass before declaring desktop host production readiness.",
    "Known unsupported Office.js paths must return capability warnings instead of fake success."
  ]
};

const markdown = renderMarkdown(report);
const operationCoverageMarkdown = renderOperationCoverageMarkdown(operationCoverage);
const markdownPath = path.join(artifactDir, "e2e-report.md");
const jsonPath = path.join(artifactDir, "e2e-report.json");
const operationCoverageMarkdownPath = path.join(artifactDir, "operation-coverage.md");
const operationCoverageJsonPath = path.join(artifactDir, "operation-coverage.json");
writeFileSync(markdownPath, markdown);
writeFileSync(jsonPath, JSON.stringify(report, null, 2));
writeFileSync(operationCoverageMarkdownPath, operationCoverageMarkdown);
writeFileSync(operationCoverageJsonPath, JSON.stringify(operationCoverage, null, 2));
console.log(markdown);
console.log(`\nSaved E2E report artifacts:\n- ${markdownPath}\n- ${jsonPath}\n- ${operationCoverageMarkdownPath}\n- ${operationCoverageJsonPath}`);

function scriptExists(file) {
  return existsSync(file);
}

function renderMarkdown(data) {
  const lines = [];
  lines.push(`# ${data.title}`);
  lines.push("");
  lines.push(`Generated: ${data.generatedAt}`);
  lines.push("");
  lines.push("## Release Target");
  lines.push("");
  lines.push("- Default E2E gate: one-tool public agent surface");
  lines.push(`- Codex agent model: ${data.releaseTarget.codexModel}`);
  lines.push(`- Codex reasoning: ${data.releaseTarget.codexReasoning}`);
  lines.push(`- OpenAI API key required: ${data.releaseTarget.apiKeyRequired ? "yes" : "no"}`);
  lines.push(`- Live Excel release hosts: ${data.releaseTarget.liveExcelReleaseHosts.join(", ")}`);
  lines.push("");
  lines.push("## Operation Coverage");
  lines.push("");
  lines.push(`- Scenario fixture: \`${data.operationCoverage.scenarioFile}\``);
  lines.push(`- Scenario count: ${data.operationCoverage.scenarioCount}`);
  lines.push(`- Capabilities represented: ${data.operationCoverage.covered.capabilities}/${data.operationCoverage.totals.capabilities}`);
  lines.push(`- Host methods represented: ${data.operationCoverage.covered.hostMethods}/${data.operationCoverage.totals.hostMethods}`);
  lines.push(`- Batch operation kinds represented: ${data.operationCoverage.covered.operationKinds}/${data.operationCoverage.totals.operationKinds}`);
  lines.push(`- Capability gaps: ${data.operationCoverage.uncoveredCapabilities.length}`);
  lines.push(`- Host method gaps: ${data.operationCoverage.uncoveredHostMethods.length}`);
  lines.push(`- Operation kind gaps: ${data.operationCoverage.uncoveredOperationKinds.length}`);
  lines.push("");
  lines.push("## Lanes");
  for (const lane of data.lanes) {
    lines.push("");
    lines.push(`### ${lane.name}`);
    lines.push("");
    lines.push(`- Command: \`${lane.command}\``);
    lines.push(`- Status: ${lane.status}`);
    lines.push(`- Starts: ${lane.startsServices.join(", ")}`);
    if (lane.coverage?.length) {
      lines.push("- Coverage:");
      for (const item of lane.coverage) {
        lines.push(`  - ${item}`);
      }
    }
    if (lane.fixtureSizes?.length) {
      lines.push("- Fixtures:");
      for (const fixture of lane.fixtureSizes) {
        lines.push(`  - ${fixture.name}: ${JSON.stringify(fixture)}`);
      }
    }
    if (lane.reliabilityAssertions?.length) {
      lines.push("- Reliability assertions:");
      for (const item of lane.reliabilityAssertions) {
        lines.push(`  - ${item}`);
      }
    }
    if (lane.gaps?.length) {
      lines.push("- Gaps:");
      for (const gap of lane.gaps) {
        lines.push(`  - ${gap}`);
      }
    }
  }
  lines.push("");
  lines.push("## Recommended Run Order");
  for (const command of data.recommendedRunOrder) {
    lines.push(`- \`${command}\``);
  }
  lines.push("");
  lines.push("## Release Gate Policy");
  for (const policy of data.releaseGatePolicy) {
    lines.push(`- ${policy}`);
  }
  return lines.join("\n");
}
