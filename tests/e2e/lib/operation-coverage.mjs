import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const scenarioCoverageRules = [
  {
    match: ({ category, intent, tags }) => category === "connection" || intent === "status" || intent === "prepare" || tags.includes("session"),
    capabilities: [
      "excel.agent.run",
      "excel.runtime.get_status",
      "excel.runtime.get_active_context",
      "excel.runtime.get_capabilities",
      "excel.workbook.get_workbook_map",
      "excel.workflow.prepare_session"
    ],
    hostMethods: ["runtime.get_active_context", "workbook.get_map"]
  },
  {
    match: ({ category, tags }) => category === "workbook overview" || tags.includes("overview"),
    capabilities: ["excel.workbook.get_summary", "excel.workbook.get_workbook_map", "excel.sheet.list", "excel.table.list", "excel.names.list"],
    hostMethods: ["workbook.get_map", "table.list", "names.list"]
  },
  {
    match: ({ category, tags }) => ["sheet reading", "sheet report", "sheet comparison", "token efficiency"].includes(category) || tags.includes("compact"),
    capabilities: ["excel.range.read_compact", "excel.range.get_summary", "excel.lookup.resolve_range", "excel.workflow.inspect_analyze"],
    hostMethods: ["workbook.snapshot_ranges", "operation.execute_batch"],
    operationKinds: ["range.read_full"]
  },
  {
    match: ({ category }) => category === "sheet structure",
    capabilities: ["excel.region.detect", "excel.region.list", "excel.names.list", "excel.lookup.inspect_match"],
    hostMethods: ["workbook.get_map", "names.list"]
  },
  {
    match: ({ category }) => category === "simple edit" || category === "smart edit",
    capabilities: ["excel.range.write_values", "excel.workflow.preview_risky_edit", "excel.validate.no_unintended_changes"],
    hostMethods: ["operation.execute_batch"],
    operationKinds: ["range.write_values"]
  },
  {
    match: ({ category }) => category === "table reading",
    capabilities: ["excel.table.get_info", "excel.table.get_schema", "excel.table.read_compact"],
    hostMethods: ["table.list", "table.get_info", "workbook.snapshot_ranges"]
  },
  {
    match: ({ category }) => category === "table edit",
    capabilities: ["excel.table.append_rows", "excel.table.update_rows", "excel.table.apply_filters", "excel.table.sort", "excel.validate.tables"],
    hostMethods: ["table.get_info", "table.append_rows", "table.sort"]
  },
  {
    match: ({ category }) => category.includes("formula"),
    capabilities: [
      "excel.formula.read_patterns",
      "excel.formula.validate",
      "excel.formula.find_errors",
      "excel.formula.recalculate",
      "excel.validate.no_formula_errors"
    ],
    hostMethods: ["formula.read_patterns", "workbook.calculate", "range.find_errors"],
    operationKinds: ["range.write_formulas", "workbook.calculate"]
  },
  {
    match: ({ category }) => category === "formatting",
    capabilities: ["excel.style.apply_style", "excel.style.get_fingerprint", "excel.style.validate_consistency"],
    hostMethods: ["style.capture_fingerprint", "operation.execute_batch"],
    operationKinds: ["range.write_styles", "range.write_number_formats"]
  },
  {
    match: ({ category }) => category === "template cleanup",
    capabilities: ["excel.template.detect_templates", "excel.template.clear_data_regions", "excel.template.repair_sheet_from_template"],
    hostMethods: ["template.capture_sheet", "template.repair", "operation.execute_batch"],
    operationKinds: ["sheet.copy", "range.clear_values_keep_format", "template.create_sheet_from_template"]
  },
  {
    match: ({ category }) => category === "cleanup",
    capabilities: [
      "excel.clean.detect_header_row",
      "excel.clean.normalize_headers",
      "excel.clean.trim_whitespace",
      "excel.clean.remove_duplicates",
      "excel.clean.parse_numbers"
    ]
  },
  {
    match: ({ category }) => category === "safety",
    capabilities: [
      "excel.snapshot.create",
      "excel.snapshot.compare_compact",
      "excel.backup.create_file",
      "excel.transaction.preview_rollback",
      "excel.workflow.rollback_validate"
    ],
    hostMethods: ["workbook.snapshot_ranges", "operation.execute_batch"],
    operationKinds: ["range.restore_snapshot"]
  },
  {
    match: ({ category }) => category === "calculation",
    capabilities: ["excel.workbook.calculate", "excel.formula.recalculate", "excel.validate.workbook"],
    hostMethods: ["workbook.calculate"],
    operationKinds: ["workbook.calculate"]
  },
  {
    match: ({ category }) => category === "save",
    capabilities: ["excel.workbook.save", "excel.backup.create_file"],
    hostMethods: ["workbook.save"],
    operationKinds: ["workbook.save"]
  },
  {
    match: ({ category }) => category === "targeting" || category === "error handling",
    capabilities: ["excel.lookup.search_workbook", "excel.lookup.resolve_range", "excel.lookup.inspect_match"],
    hostMethods: ["workbook.get_map", "runtime.get_selection"]
  },
  {
    match: ({ category }) => category === "multilingual",
    capabilities: ["excel.agent.run", "excel.lookup.resolve_range"],
    hostMethods: ["workbook.get_map"]
  },
  {
    match: ({ category }) => category === "normal office workflow",
    capabilities: ["excel.workflow.prepare_session", "excel.workflow.inspect_analyze", "excel.workflow.preview_risky_edit", "excel.validate.workbook"],
    hostMethods: ["workbook.get_map", "workbook.snapshot_ranges", "operation.execute_batch"],
    operationKinds: ["range.read_full", "range.write_values"]
  }
];

const explicitScenarioCoverage = {
  "pivot-chart": {
    capabilities: [
      "excel.workflow.create_pivot_chart_summary",
      "excel.pivot.create",
      "excel.pivot.refresh",
      "excel.pivot.validate_source",
      "excel.chart.create",
      "excel.chart.refresh"
    ],
    hostMethods: ["pivot.create", "pivot.refresh", "pivot.validate_source", "chart.create", "chart.refresh"]
  },
  "multi-agent-locking": {
    capabilities: ["excel.lock.acquire", "excel.lock.release", "excel.collab.get_status", "excel.conflict.get_guidance"]
  },
  "closed-workbook-status-error": {
    capabilities: ["excel.runtime.get_status", "excel.workbook.list_open_workbooks"]
  }
};

export async function buildOperationCoverageReport({ repoRoot, scenarioFile }) {
  const scenarios = readScenarioFile(scenarioFile);
  const registries = await loadRegistries(repoRoot);
  const scenarioCoverage = scenarios.map((scenario) => coverageForScenario(scenario));
  const coveredCapabilities = unique(scenarioCoverage.flatMap((item) => item.capabilities));
  const coveredHostMethods = unique(scenarioCoverage.flatMap((item) => item.hostMethods));
  const coveredOperationKinds = unique(scenarioCoverage.flatMap((item) => item.operationKinds));

  const knownCapabilities = new Set(registries.capabilities.map((entry) => entry.name));
  const knownHostMethods = new Set(registries.hostMethods.map((entry) => entry.method));
  const knownOperationKinds = new Set(registries.operationKinds);

  const unknown = {
    capabilities: coveredCapabilities.filter((name) => !knownCapabilities.has(name)),
    hostMethods: coveredHostMethods.filter((name) => !knownHostMethods.has(name)),
    operationKinds: coveredOperationKinds.filter((name) => !knownOperationKinds.has(name))
  };

  const capabilitiesByGroup = summarizeCapabilitiesByGroup(registries.capabilities, coveredCapabilities);
  const capabilitiesByPlanningStatus = summarizeCapabilitiesByPlanningStatus(registries.capabilities, coveredCapabilities);
  const scenarioCategories = countBy(scenarios, (scenario) => scenario.category ?? "uncategorized");

  return {
    title: "Open Workbook Operation Coverage Report",
    scenarioFile,
    scenarioCount: scenarios.length,
    scenarioCategories,
    covered: {
      capabilities: coveredCapabilities.length,
      hostMethods: coveredHostMethods.length,
      operationKinds: coveredOperationKinds.length
    },
    totals: {
      capabilities: registries.capabilities.length,
      hostMethods: registries.hostMethods.length,
      operationKinds: registries.operationKinds.length
    },
    capabilitiesByGroup,
    capabilitiesByPlanningStatus,
    uncoveredCapabilities: registries.capabilities
      .filter((entry) => !coveredCapabilities.includes(entry.name))
      .map((entry) => entry.name),
    uncoveredStableCapabilities: registries.capabilities
      .filter((entry) => entry.coverageStatus === "covered" && !coveredCapabilities.includes(entry.name))
      .map((entry) => entry.name),
    uncoveredHostMethods: registries.hostMethods.filter((entry) => !coveredHostMethods.includes(entry.method)).map((entry) => entry.method),
    uncoveredOperationKinds: registries.operationKinds.filter((kind) => !coveredOperationKinds.includes(kind)),
    unknown,
    scenarioCoverage
  };
}

export function assertOperationCoverageReport(report) {
  const unknown = [
    ...report.unknown.capabilities.map((name) => `unknown capability ${name}`),
    ...report.unknown.hostMethods.map((name) => `unknown host method ${name}`),
    ...report.unknown.operationKinds.map((name) => `unknown operation kind ${name}`)
  ];
  if (unknown.length > 0) {
    throw new Error(`Operation coverage fixture references invalid registry names:\n${unknown.join("\n")}`);
  }
  const uncovered = [
    ...report.uncoveredCapabilities.map((name) => `uncovered capability ${name}`),
    ...report.uncoveredHostMethods.map((name) => `uncovered host method ${name}`),
    ...report.uncoveredOperationKinds.map((name) => `uncovered operation kind ${name}`)
  ];
  if (uncovered.length > 0) {
    throw new Error(`Operation coverage fixture is incomplete:\n${uncovered.join("\n")}`);
  }
}

export function renderOperationCoverageMarkdown(report) {
  const lines = [
    `# ${report.title}`,
    "",
    `Scenario file: ${report.scenarioFile}`,
    `Scenarios: ${report.scenarioCount}`,
    `Capabilities represented: ${report.covered.capabilities}/${report.totals.capabilities}`,
    `Host methods represented: ${report.covered.hostMethods}/${report.totals.hostMethods}`,
    `Batch operation kinds represented: ${report.covered.operationKinds}/${report.totals.operationKinds}`,
    "",
    "## Capability Groups",
    ...report.capabilitiesByGroup.map((group) => `- ${group.group}: ${group.covered}/${group.total} represented, stable ${group.coveredCoveredStatus}/${group.totalCoveredStatus}`),
    "",
    "## Planning Status",
    ...report.capabilitiesByPlanningStatus.map((item) => `- ${item.status}: ${item.covered}/${item.total} represented`),
    "",
    "## Scenario Categories",
    ...Object.entries(report.scenarioCategories).sort().map(([category, count]) => `- ${category}: ${count}`),
    "",
    "## Remaining Gaps",
    `- Capabilities without representative E2E fixture: ${report.uncoveredCapabilities.length}`,
    `- Host methods without representative E2E fixture: ${report.uncoveredHostMethods.length}`,
    `- Batch operation kinds without representative E2E fixture: ${report.uncoveredOperationKinds.length}`
  ];
  if (report.unknown.capabilities.length || report.unknown.hostMethods.length || report.unknown.operationKinds.length) {
    lines.push("", "## Invalid References");
    for (const name of report.unknown.capabilities) lines.push(`- capability: ${name}`);
    for (const name of report.unknown.hostMethods) lines.push(`- host method: ${name}`);
    for (const name of report.unknown.operationKinds) lines.push(`- operation kind: ${name}`);
  }
  return lines.join("\n");
}

function readScenarioFile(scenarioFile) {
  const scenarios = JSON.parse(readFileSync(scenarioFile, "utf8"));
  if (!Array.isArray(scenarios)) {
    throw new Error(`Scenario file must contain a JSON array: ${scenarioFile}`);
  }
  return scenarios;
}

async function loadRegistries(repoRoot) {
  const backendRegistryPath = path.join(repoRoot, "apps", "backend", "dist", "capabilities", "registry.js");
  const hostRegistryPath = path.join(repoRoot, "apps", "excel-addin", "dist", "host", "registry.js");
  if (!existsSync(backendRegistryPath) || !existsSync(hostRegistryPath)) {
    throw new Error("Operation coverage report requires built backend and add-in registry files. Run `corepack pnpm build` first.");
  }
  const backendRegistry = await import(pathToFileURL(backendRegistryPath).href);
  const hostRegistry = await import(pathToFileURL(hostRegistryPath).href);
  return {
    capabilities: backendRegistry.listBackendCapabilityRegistry({ includeInternal: true }),
    hostMethods: hostRegistry.HOST_METHOD_REGISTRY,
    operationKinds: hostRegistry.BATCH_OPERATION_KINDS
  };
}

function coverageForScenario(scenario) {
  const normalized = {
    id: scenario.id,
    category: scenario.category ?? "uncategorized",
    intent: typeof scenario.intent === "string" ? scenario.intent : scenario.input?.intent?.action,
    tags: Array.isArray(scenario.tags) ? scenario.tags : [],
    capabilities: new Set(scenario.coverage?.capabilities ?? []),
    hostMethods: new Set(scenario.coverage?.hostMethods ?? []),
    operationKinds: new Set(scenario.coverage?.operationKinds ?? [])
  };
  for (const rule of scenarioCoverageRules) {
    if (rule.match(normalized)) {
      for (const capability of rule.capabilities ?? []) normalized.capabilities.add(capability);
      for (const method of rule.hostMethods ?? []) normalized.hostMethods.add(method);
      for (const kind of rule.operationKinds ?? []) normalized.operationKinds.add(kind);
    }
  }
  const explicit = explicitScenarioCoverage[scenario.id];
  if (explicit) {
    for (const capability of explicit.capabilities ?? []) normalized.capabilities.add(capability);
    for (const method of explicit.hostMethods ?? []) normalized.hostMethods.add(method);
    for (const kind of explicit.operationKinds ?? []) normalized.operationKinds.add(kind);
  }
  return {
    id: scenario.id,
    category: normalized.category,
    capabilities: [...normalized.capabilities].sort(),
    hostMethods: [...normalized.hostMethods].sort(),
    operationKinds: [...normalized.operationKinds].sort()
  };
}

function summarizeCapabilitiesByGroup(capabilities, coveredCapabilities) {
  const covered = new Set(coveredCapabilities);
  const groups = new Map();
  for (const capability of capabilities) {
    const group = groups.get(capability.group) ?? { group: capability.group, total: 0, covered: 0, totalCoveredStatus: 0, coveredCoveredStatus: 0 };
    group.total += 1;
    if (covered.has(capability.name)) group.covered += 1;
    if (capability.coverageStatus === "covered") {
      group.totalCoveredStatus += 1;
      if (covered.has(capability.name)) group.coveredCoveredStatus += 1;
    }
    groups.set(capability.group, group);
  }
  return [...groups.values()].sort((left, right) => left.group.localeCompare(right.group));
}

function summarizeCapabilitiesByPlanningStatus(capabilities, coveredCapabilities) {
  const covered = new Set(coveredCapabilities);
  const statuses = new Map();
  for (const capability of capabilities) {
    const status = capability.coverageStatus ?? "unknown";
    const item = statuses.get(status) ?? { status, total: 0, covered: 0 };
    item.total += 1;
    if (covered.has(capability.name)) item.covered += 1;
    statuses.set(status, item);
  }
  return [...statuses.values()].sort((left, right) => left.status.localeCompare(right.status));
}

function countBy(values, selector) {
  const counts = {};
  for (const value of values) {
    const key = selector(value);
    counts[key] = (counts[key] ?? 0) + 1;
  }
  return counts;
}

function unique(values) {
  return [...new Set(values)].sort();
}
