#!/usr/bin/env node
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const artifactDir = process.env.OPEN_WORKBOOK_E2E_REPORT_DIR ?? path.join(tmpdir(), "open-workbook-e2e-reports");
mkdirSync(artifactDir, { recursive: true });

const report = {
  title: "Open Workbook MCP E2E Test Report",
  generatedAt: new Date().toISOString(),
  releaseTarget: {
    fastGateBudgetMs: Number(process.env.OPEN_WORKBOOK_E2E_FAST_BUDGET_MS ?? 300_000),
    codexModel: process.env.OPEN_WORKBOOK_E2E_CODEX_MODEL ?? "gpt-5.4-mini",
    codexReasoning: process.env.OPEN_WORKBOOK_E2E_CODEX_REASONING ?? "low",
    apiKeyRequired: false,
    liveExcelReleaseHosts: ["macOS desktop Excel", "Windows desktop Excel"]
  },
  prerequisites: [
    {
      lane: "fast",
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
      name: "test:e2e:fast",
      command: "pnpm run test:e2e:fast",
      status: scriptExists("scripts/e2e-fast.mjs") ? "implemented" : "missing",
      startsServices: ["Open Workbook MCP stdio server", "Fake Excel add-in WebSocket client"],
      coverage: [
        "MCP initialize/tools/resources/prompts protocol sweep",
        "Runtime status, active context, capabilities, workbook map",
        "Sheet create",
        "Range values/formulas/number formats/copy/clear/read",
        "Large table create/read window/reorder/filter/sort",
        "Workbook backup and restore",
        "Transaction listing",
        "Manual lock conflict, guidance, release, retry",
        "Templates: detect/register/get/list/infer/create/fill/validate",
        "Styles: fingerprint/compare/apply/copy/repair",
        "Formulas: read patterns/dependency graph/trace/validate/explain/copy/fill/repair",
        "Names and regions: create/list/get/update/register/detect/fill",
        "Cleaning: header detect, trim, parse numbers, fill missing, dedupe, split, merge, outliers, fuzzy match",
        "Pivots/charts: create/list/get/validate/fingerprint/refresh/update",
        "Snapshots/diffs/events/permissions"
      ],
      fixtureSizes: [
        { name: "small", sheets: 3, tables: 1, rows: 5 },
        { name: "large", sheets: 1, tables: 1, rows: 5000, boundedReads: true }
      ],
      reliabilityAssertions: [
        "Every mutating range path returns transaction metadata",
        "Table mutation paths create backup metadata",
        "Large table reads use row windows",
        "Overlapping write is blocked while a manual lock is active",
        "Retry succeeds after lock release",
        "Suite fails if total runtime exceeds configured budget"
      ],
      gaps: [
        "The sweep is representative by stable tool group, not one assertion per individual stable tool.",
        "Live Office.js fidelity still requires macOS and Windows Excel gates.",
        "Pivot/chart fake-host assertions validate deterministic metadata paths; deep host-specific rendering remains live-gated."
      ]
    },
    {
      name: "test:e2e:agent:core",
      command: "pnpm run test:e2e:agent:core",
      status: scriptExists("scripts/e2e-codex-agent.mjs") ? "implemented" : "missing",
      startsServices: ["Codex exec", "Open Workbook MCP via temporary Codex config", "Scenario fake Excel add-in"],
      coverage: [
        "Signed-in Codex CLI can initialize Open Workbook MCP",
        "Codex uses gpt-5.4-mini with low reasoning unless overridden",
        "Bundled open-workbook-excel skill guidance is loaded into every scenario prompt",
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
      status: scriptExists("scripts/e2e-codex-agent.mjs") ? "implemented" : "missing",
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
      status: scriptExists("scripts/e2e-agent-model-matrix.mjs") ? "implemented" : "missing",
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
      status: scriptExists("scripts/e2e-live-smoke.mjs") ? "opt-in connectivity smoke" : "missing",
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
      status: scriptExists("scripts/e2e-live-smoke.mjs") ? "opt-in connectivity smoke" : "missing",
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
    "pnpm run test:e2e:fast",
    "pnpm run test:e2e:agent:core",
    "pnpm run test:e2e:agent:quality",
    "pnpm run test:e2e:agent:quality:compare",
    "pnpm run test:e2e:live:mac",
    "pnpm run test:e2e:live:windows"
  ],
  releaseGatePolicy: [
    "Fast E2E must pass before every release candidate.",
    "Codex core signed-in E2E must pass before customer-facing large-workbook and multi-agent safety claims.",
    "Codex quality E2E is report-only by default and should be reviewed before claiming complex workflow quality.",
    "Both live Excel gates must pass before declaring desktop host production readiness.",
    "Known unsupported Office.js paths must return capability warnings instead of fake success."
  ]
};

const markdown = renderMarkdown(report);
const markdownPath = path.join(artifactDir, "e2e-report.md");
const jsonPath = path.join(artifactDir, "e2e-report.json");
writeFileSync(markdownPath, markdown);
writeFileSync(jsonPath, JSON.stringify(report, null, 2));
console.log(markdown);
console.log(`\nSaved E2E report artifacts:\n- ${markdownPath}\n- ${jsonPath}`);

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
  lines.push(`- Fast gate budget: ${data.releaseTarget.fastGateBudgetMs} ms`);
  lines.push(`- Codex agent model: ${data.releaseTarget.codexModel}`);
  lines.push(`- Codex reasoning: ${data.releaseTarget.codexReasoning}`);
  lines.push(`- OpenAI API key required: ${data.releaseTarget.apiKeyRequired ? "yes" : "no"}`);
  lines.push(`- Live Excel release hosts: ${data.releaseTarget.liveExcelReleaseHosts.join(", ")}`);
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
