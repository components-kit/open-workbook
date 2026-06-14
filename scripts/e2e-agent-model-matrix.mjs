#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const baseArtifactDir = readArg("--artifact-dir") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_MATRIX_DIR ?? path.join(tmpdir(), "open-workbook-e2e-agent-matrix");
const scenarios = readArg("--scenarios") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_MATRIX_SCENARIOS ?? "quality";
const reasoning = readArg("--reasoning") ?? process.env.OPEN_WORKBOOK_E2E_CODEX_REASONING ?? "low";
const timeoutMs = readArg("--timeout-ms") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_TIMEOUT_MS;
const strict = hasArg("--strict") || process.env.OPEN_WORKBOOK_E2E_AGENT_MATRIX_STRICT === "1";
const profileNames = (readArg("--profiles") ?? process.env.OPEN_WORKBOOK_E2E_AGENT_MATRIX_PROFILES ?? "cheap,frontier")
  .split(",")
  .map((item) => item.trim())
  .filter(Boolean);

const knownProfiles = {
  cheap: {
    label: "cheap",
    model: process.env.OPEN_WORKBOOK_E2E_CHEAP_MODEL ?? "gpt-5.4-mini"
  },
  frontier: {
    label: "frontier",
    model: process.env.OPEN_WORKBOOK_E2E_FRONTIER_MODEL ?? "gpt-5.4"
  }
};

mkdirSync(baseArtifactDir, { recursive: true });

const started = performance.now();
const results = [];
for (const profileName of profileNames) {
  const profile = knownProfiles[profileName] ?? { label: profileName, model: profileName };
  results.push(runProfile(profile));
}

const matrix = {
  ok: results.every((result) => result.ok || result.reportOnly),
  strictOk: results.every((result) => result.ok),
  scenarios,
  reasoning,
  strict,
  elapsedMs: Math.round(performance.now() - started),
  artifactDir: baseArtifactDir,
  results
};

writeFileSync(path.join(baseArtifactDir, "model-matrix.json"), JSON.stringify(matrix, null, 2));
writeFileSync(path.join(baseArtifactDir, "model-matrix.md"), renderMarkdown(matrix));
console.log(renderMarkdown(matrix));
console.log(`\nSaved model matrix artifacts:\n- ${path.join(baseArtifactDir, "model-matrix.md")}\n- ${path.join(baseArtifactDir, "model-matrix.json")}`);

if (strict && !matrix.strictOk) {
  process.exit(1);
}

function runProfile(profile) {
  const artifactDir = path.join(baseArtifactDir, profile.label);
  mkdirSync(artifactDir, { recursive: true });
  const args = [
    "scripts/e2e-codex-agent.mjs",
    "--scenarios",
    scenarios,
    "--report-only",
    "--model",
    profile.model,
    "--reasoning",
    reasoning
  ];
  if (timeoutMs !== undefined) {
    args.push("--timeout-ms", timeoutMs);
  }
  const startedProfile = performance.now();
  const child = spawnSync(process.execPath, args, {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      OPEN_WORKBOOK_E2E_AGENT_ARTIFACT_DIR: artifactDir
    },
    stdio: ["ignore", "pipe", "pipe"]
  });
  writeFileSync(path.join(artifactDir, "matrix-run.stdout.log"), child.stdout ?? "");
  writeFileSync(path.join(artifactDir, "matrix-run.stderr.log"), child.stderr ?? "");

  const suitePath = path.join(artifactDir, "codex-agent-suite.json");
  const suite = existsSync(suitePath) ? JSON.parse(readFileSync(suitePath, "utf8")) : undefined;
  const passed = suite?.results?.filter((result) => result.ok).length ?? 0;
  const total = suite?.results?.length ?? 0;
  return {
    label: profile.label,
    model: profile.model,
    ok: Boolean(suite?.ok),
    reportOnly: Boolean(suite?.reportOnly),
    exitStatus: child.status,
    elapsedMs: Math.round(performance.now() - startedProfile),
    passed,
    total,
    failureCategoryCounts: suite?.failureCategoryCounts ?? {},
    artifactDir,
    error: suite ? undefined : "codex-agent-suite.json was not produced"
  };
}

function renderMarkdown(matrix) {
  const lines = [
    "# Codex Agent Model Matrix",
    "",
    `- Scenarios: ${matrix.scenarios}`,
    `- Reasoning: ${matrix.reasoning}`,
    `- Strict: ${matrix.strict ? "yes" : "no"}`,
    `- Elapsed: ${matrix.elapsedMs} ms`,
    `- Artifacts: ${matrix.artifactDir}`,
    "",
    "## Results"
  ];
  for (const result of matrix.results) {
    lines.push("");
    lines.push(`### ${result.label}`);
    lines.push("");
    lines.push(`- Model: ${result.model}`);
    lines.push(`- Passed: ${result.passed}/${result.total}`);
    lines.push(`- Exit status: ${result.exitStatus}`);
    lines.push(`- Elapsed: ${result.elapsedMs} ms`);
    lines.push(`- Artifacts: ${result.artifactDir}`);
    if (result.error) {
      lines.push(`- Error: ${result.error}`);
    }
    const categories = Object.entries(result.failureCategoryCounts);
    lines.push("- Failure categories:");
    if (categories.length === 0) {
      lines.push("  - none");
    } else {
      for (const [category, count] of categories.sort(([left], [right]) => left.localeCompare(right))) {
        lines.push(`  - ${category}: ${count}`);
      }
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
