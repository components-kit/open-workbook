#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

const regressionScenarios = readScenarioFile("tests/e2e/fixtures/office-agent-regression-scenarios.json");
const departmentScenarios = readScenarioFile("tests/e2e/fixtures/office-agent-department-scenarios.json");

const finalStateAssertionFields = [
  "cellValues",
  "cellFormulas",
  "cellNumberFormats",
  "cellStyles",
  "conditionalFormatRanges",
  "insertedColumns",
  "sheetProtections",
  "tableColumnOrder",
  "validationRanges"
];

main();

function main() {
  validateCommonScenarioContract(regressionScenarios, "regression");
  validateCommonScenarioContract(departmentScenarios, "department");
  validateRegressionPack();
  validateDepartmentPack();
  console.log("Scenario fixture contract E2E passed.");
}

function readScenarioFile(relativePath) {
  const absolutePath = path.join(repoRoot, relativePath);
  const parsed = JSON.parse(readFileSync(absolutePath, "utf8"));
  assert(Array.isArray(parsed), `${relativePath} must contain a scenario array`);
  return parsed.map((scenario) => ({ ...scenario, __file: relativePath }));
}

function validateCommonScenarioContract(scenarios, label) {
  const seenIds = new Set();
  assert(scenarios.length > 0, `${label} scenario pack must not be empty`);

  for (const scenario of scenarios) {
    assertNonEmptyString(scenario.id, `${scenario.__file}: scenario id is required`);
    assert(!seenIds.has(scenario.id), `${scenario.__file}: duplicate scenario id ${scenario.id}`);
    seenIds.add(scenario.id);

    assertNonEmptyString(scenario.category, `${scenario.id}: category is required`);
    assertNonEmptyString(scenario.intent, `${scenario.id}: intent is required`);
    assertNonEmptyString(scenario.prompt, `${scenario.id}: prompt is required`);
    validateScenarioRequests(scenario);
    assert(scenario.expected && typeof scenario.expected === "object", `${scenario.id}: expected contract is required`);

    assertPositiveNumber(scenario.budgets?.maxToolCalls, `${scenario.id}: budgets.maxToolCalls must be positive`);
    assertPositiveNumber(scenario.budgets?.maxPayloadBytes, `${scenario.id}: budgets.maxPayloadBytes must be positive`);
    assertPositiveNumber(scenario.budgets?.maxLatencyMs, `${scenario.id}: budgets.maxLatencyMs must be positive`);
    assert(scenario.expected.mustNotReadFullWorkbook === true, `${scenario.id}: mustNotReadFullWorkbook must be explicit`);
    assert(typeof scenario.expected.shouldMutateWorkbook === "boolean", `${scenario.id}: shouldMutateWorkbook must be explicit`);

    if (scenario.expected.shouldMutateWorkbook) {
      validateMutatingScenario(scenario);
    } else {
      validateNonMutatingScenario(scenario);
    }
  }
}

function validateMutatingScenario(scenario) {
  const expected = scenario.expected;
  assert(expected.resultType === "apply", `${scenario.id}: mutating scenarios must apply through preview/apply`);
  assert(expected.autoApply === true, `${scenario.id}: mutating scenarios must exercise apply after preview`);
  assert(expected.xlsxAssertions === true, `${scenario.id}: mutating scenarios must reload and assert workbook artifacts`);
  assert(Array.isArray(expected.hostMethods) && expected.hostMethods.length > 0, `${scenario.id}: mutating scenarios must assert host route`);
  assert(hasFinalStateAssertion(expected), `${scenario.id}: mutating scenarios must assert final workbook state`);
}

function validateScenarioRequests(scenario) {
  if (scenario.input) {
    assertNonEmptyString(scenario.input.request, `${scenario.id}: input.request is required`);
    return;
  }

  assert(Array.isArray(scenario.steps) && scenario.steps.length > 0, `${scenario.id}: input or steps are required`);
  for (const [index, step] of scenario.steps.entries()) {
    assertNonEmptyString(step.label, `${scenario.id}: steps[${index}].label is required`);
    assertNonEmptyString(step.input?.request, `${scenario.id}: steps[${index}].input.request is required`);
  }
}

function validateNonMutatingScenario(scenario) {
  const expected = scenario.expected;
  assert(expected.xlsxAssertions !== true, `${scenario.id}: non-mutating scenarios should not claim xlsx mutation assertions`);
  assert(expected.mustNotMutateWorkbook === true || expected.shouldMutateWorkbook === false, `${scenario.id}: non-mutating scenario must state safety expectation`);

  const isBoundaryScenario =
    expected.resultType === "error" ||
    expected.status === "VALIDATION_FAILED" ||
    expected.nextAction === "manual_review" ||
    expected.shouldFailGracefully === true ||
    expected.answerKind === "workflow_plan";

  if (isBoundaryScenario) {
    assert(
      expected.nextAction === "manual_review" || expected.shouldFailGracefully === true || expected.answerKind === "workflow_plan",
      `${scenario.id}: boundary scenarios need manual-review, workflow-plan, or graceful-failure guidance`
    );
  }
}

function validateRegressionPack() {
  const requiredIntents = new Set([
    "connection",
    "format_range",
    "insert_columns",
    "reorder_range_columns",
    "reorder_table_columns",
    "write_conditional_formatting",
    "write_data_validation"
  ]);
  const actualIntents = new Set(regressionScenarios.map((scenario) => scenario.intent));
  for (const intent of requiredIntents) {
    assert(actualIntents.has(intent), `regression pack missing ${intent}`);
  }

  for (const scenario of regressionScenarios) {
    assert(scenario.category === "production regression", `${scenario.id}: regression category must be production regression`);
    if (scenario.expected.shouldMutateWorkbook) {
      assert(
        scenario.expected.hostMethods.includes("operation.execute_batch") || scenario.expected.hostMethods.includes("table.reorder_columns"),
        `${scenario.id}: regression scenario must assert the broken layer route`
      );
    } else {
      assert(scenario.fixture?.excelState === "noActiveWorkbook", `${scenario.id}: non-mutating regression must declare its disconnected workbook fixture`);
      assert(scenario.expected.shouldFailGracefully === true, `${scenario.id}: disconnected regression must assert graceful failure`);
      assert(scenario.expected.mustNotMutateWorkbook === true, `${scenario.id}: disconnected regression must forbid mutation`);
    }
  }
}

function validateDepartmentPack() {
  const requiredDepartments = [
    "finance",
    "sales/ops",
    "logistics",
    "HR",
    "executive reporting",
    "data cleanup"
  ];

  for (const department of requiredDepartments) {
    const scenarios = departmentScenarios.filter((scenario) => scenario.category.includes(`department workflow: ${department}`));
    assert(scenarios.length > 0, `department pack missing ${department}`);
    assert(scenarios.some((scenario) => scenario.expected.shouldMutateWorkbook === false), `${department}: missing read-only/safety scenario`);
    assert(scenarios.some((scenario) => scenario.expected.shouldMutateWorkbook === true), `${department}: missing mutating scenario`);
  }

  const requiredOperationOrHostCapabilities = new Set([
    "range.write_values",
    "range.write_formulas",
    "range.write_number_formats",
    "range.write_styles",
    "range.write_styles_many",
    "range.insert_columns",
    "range.reorder_columns",
    "range.write_data_validation",
    "range.write_conditional_formatting",
    "sheet.protect",
    "template.capture",
    "template.capture_sheet",
    "template.repair"
  ]);
  const actualOperationOrHostCapabilities = new Set(departmentScenarios.flatMap((scenario) => [
    ...(scenario.expected.operationKinds ?? []),
    ...(scenario.expected.hostMethods ?? [])
  ]));
  for (const capability of requiredOperationOrHostCapabilities) {
    assert(actualOperationOrHostCapabilities.has(capability), `department pack missing operation or host capability ${capability}`);
  }

  const pivotWorkflow = departmentScenarios.find((scenario) => scenario.intent === "create_pivot_chart_summary");
  assert(pivotWorkflow, "department pack missing pivot/chart workflow boundary scenario");
  assert(pivotWorkflow.expected.answerKind === "workflow_plan", "pivot/chart scenario must return workflow_plan");
  assert(pivotWorkflow.expected.nextAction === "manual_review", "pivot/chart scenario must require manual review");
  assert(Array.isArray(pivotWorkflow.expected.requiredCapabilities) && pivotWorkflow.expected.requiredCapabilities.length >= 4, "pivot/chart scenario must list required capabilities");

  const formulaBoundary = departmentScenarios.find((scenario) => scenario.intent === "repair_formula_errors");
  assert(formulaBoundary, "department pack missing unsupported formula-repair boundary scenario");
  assert(formulaBoundary.expected.status === "VALIDATION_FAILED", "formula-repair boundary must return VALIDATION_FAILED");
  assert(formulaBoundary.expected.nextAction === "manual_review", "formula-repair boundary must require manual review");
}

function hasFinalStateAssertion(expected) {
  return finalStateAssertionFields.some((field) => {
    const value = expected[field];
    return Array.isArray(value) ? value.length > 0 : Boolean(value);
  });
}

function assertNonEmptyString(value, message) {
  assert(typeof value === "string" && value.trim().length > 0, message);
}

function assertPositiveNumber(value, message) {
  assert(typeof value === "number" && Number.isFinite(value) && value > 0, message);
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}
