import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import * as z from "zod/v4";
import { AGENT_DETAIL_LEVELS, AGENT_RUN_MODES, AGENT_RUN_STATUSES, type AgentRunExecutionContext, type AgentRunInput } from "@components-kit/open-workbook-protocol";
import type { RuntimeFacade } from "../runtime-facade.js";
import { agentJsonResult } from "../results.js";

export function registerAgentTools(mcp: McpServer, runtime: RuntimeFacade, context?: AgentRunExecutionContext): void {
  (mcp.registerTool as any)(
    "excel.agent.run",
    {
      title: "Run Open Workbook agent workflow",
      description:
        "Single default Open Workbook interface. Send workbook intent; the backend handles discovery, cached metadata, target resolution, preview/apply, validation, rollback, and compact proof without exposing low-level Excel tools. In default auto mode, safe exact small edits may auto-apply and return taskOutcome apply_complete with maxRecommendedFollowupCalls 0; set autoApply false when you need preview-only behavior. This tool can read the current live Excel selection; when the user says this, here, selected, current, this row/cell/range/column, or asks a vague question while Excel is connected, call excel.agent.run before asking for row or column numbers.",
      inputSchema: agentRunInputSchema(),
      outputSchema: agentRunOutputSchema(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false
      }
    },
    async (args: AgentRunInput) => agentJsonResult(await runtime.runAgent(args, context))
  );
}

export function agentRunInputSchema() {
  const targetObjectSchema = z.object({
    workbookId: z.string().optional(),
    workbookName: z.string().optional(),
    candidateId: z.string().optional(),
    sheetName: z.string().optional(),
    tableName: z.string().optional(),
    range: z.string().optional(),
    row: z.number().int().optional(),
    column: z.string().optional(),
    entity: z.string().optional()
  });
  const targetSchema = jsonObjectString(targetObjectSchema, "target");
  const continuationSchema = jsonObjectString(z.object({
    workbookContextId: z.string().optional(),
    operationId: z.string().optional(),
    transactionId: z.string().optional(),
    resultUri: z.string().optional(),
    fullResultUri: z.string().optional(),
    freshness: z.object({
      workbookId: z.string().optional(),
      workbookContentVersion: z.number().optional(),
      workbookStructureHash: z.string().optional(),
      contextUpdatedAt: z.number().optional()
    }).optional(),
    nextRequest: z.string().optional(),
    responseMode: z.enum(["brief", "standard", "verbose"]).optional()
  }), "continuation");
  const intentSchema = jsonObjectString(z.object({
    action: z.string().optional().describe("Optional internal action hint, such as read_values, write_values, validate_workbook, or create_pivot_chart_summary. Invalid hints are ignored by the backend."),
    confidence: z.number().min(0).max(1).optional(),
    reason: z.string().optional(),
    targetHints: z.array(z.string()).optional()
  }), "intent");
  const colorSchema = z.string().regex(/^#[0-9a-fA-F]{6}$|^[A-Za-z]+$/, "colors must be CSS-style names or #RRGGBB hex strings");
  const styleSchema = z.object({
    fillColor: colorSchema.optional(),
    fill: colorSchema.optional(),
    backgroundColor: colorSchema.optional(),
    background: colorSchema.optional(),
    fontColor: colorSchema.optional(),
    textColor: colorSchema.optional(),
    fontBold: z.boolean().optional(),
    bold: z.boolean().optional(),
    fontItalic: z.boolean().optional(),
    italic: z.boolean().optional(),
    fontName: z.string().optional(),
    fontSize: z.number().positive().optional(),
    horizontalAlignment: z.string().optional(),
    align: z.string().optional(),
    verticalAlignment: z.string().optional(),
    rowHeight: z.number().positive().optional(),
    columnWidth: z.number().positive().optional(),
    borders: z.any().optional()
  }).catchall(z.any());
  const validationSchema = z.object({
    type: z.enum(["list"]).optional(),
    source: z.union([z.string(), z.array(z.string())]).optional(),
    formula1: z.string().optional(),
    inCellDropDown: z.boolean().optional(),
    ignoreBlanks: z.boolean().optional()
  }).catchall(z.any());
  const conditionalRuleSchema = z.object({
    type: z.enum(["custom"]).optional(),
    formula: z.string().optional(),
    style: styleSchema.optional()
  }).catchall(z.any());
  const numberFormatSchema = z.union([z.string(), z.array(z.array(z.string()))]);
  const tableRowUpdateSchema = z.object({
    index: z.number().int(),
    values: z.array(z.any())
  }).catchall(z.any());
  const valuesSchema = jsonObjectString(z.object({
    values: z.array(z.array(z.any())).optional(),
    rows: z.array(z.union([z.array(z.any()), tableRowUpdateSchema])).optional(),
    formulas: z.array(z.array(z.union([z.string(), z.null()]))).optional(),
    numberFormat: numberFormatSchema.optional(),
    numberFormats: numberFormatSchema.optional(),
    formats: numberFormatSchema.optional(),
    style: styleSchema.optional(),
    entries: z.array(z.object({
      sheetName: z.string().optional(),
      range: z.string().optional(),
      address: z.string().optional(),
      target: targetSchema.optional(),
      style: styleSchema.optional()
    }).catchall(z.any())).optional(),
    options: z.array(z.string()).optional(),
    allowedValues: z.array(z.string()).optional(),
    source: z.union([z.string(), z.array(z.string())]).optional(),
    validation: validationSchema.optional(),
    formula: z.string().optional(),
    rule: conditionalRuleSchema.optional(),
    columnOrder: z.array(z.union([z.string(), z.number()])).optional(),
    columns: z.array(z.union([z.string(), z.number()])).optional(),
    order: z.array(z.union([z.string(), z.number()])).optional(),
    patches: z.array(z.object({
      target: targetSchema,
      values: z.array(z.array(z.any())).optional(),
      rows: z.array(z.array(z.any())).optional(),
      reason: z.string().optional()
    })).optional()
  }).catchall(z.any()), "values");
  return {
    request: z.string(),
    mode: z.enum(AGENT_RUN_MODES).optional(),
    workbookContextId: z.string().optional(),
    operationId: z.string().optional(),
    transactionId: z.string().optional(),
    confirmationToken: z.string().optional(),
    continuation: continuationSchema.optional(),
    intent: intentSchema.optional(),
    target: targetSchema.optional(),
    values: valuesSchema.optional(),
    autoApply: z.boolean().optional(),
    detailLevel: z.enum(AGENT_DETAIL_LEVELS).optional(),
    responseMode: z.enum(["brief", "standard", "verbose"]).optional(),
    budget: z.object({
      maxPayloadBytes: z.number().int().positive().optional(),
      maxEstimatedTokens: z.number().int().positive().optional(),
      maxExamples: z.number().int().positive().optional()
    }).optional()
  };
}

function agentRunOutputSchema() {
  return {
    status: z.enum(AGENT_RUN_STATUSES),
    mode: z.string(),
    workbookContextId: z.string().optional(),
    operationId: z.string().optional(),
    transactionId: z.string().optional(),
    confirmationToken: z.string().optional(),
    summary: z.string(),
    answer: z.any().optional(),
    metrics: z.record(z.string(), z.any()).optional(),
    changes: z.array(z.any()).optional(),
    candidates: z.array(z.any()).optional(),
    proof: z.array(z.any()),
    resourceLinks: z.array(z.any()),
    invalidatedContextIds: z.array(z.string()).optional(),
    invalidatedResourceUris: z.array(z.string()).optional(),
    continuation: z.object({
      workbookContextId: z.string().optional(),
      operationId: z.string().optional(),
      transactionId: z.string().optional(),
      resultUri: z.string().optional(),
      fullResultUri: z.string().optional(),
      freshness: z.object({
        workbookId: z.string().optional(),
        workbookContentVersion: z.number().optional(),
        workbookStructureHash: z.string().optional(),
        contextUpdatedAt: z.number().optional()
      }).optional(),
      nextRequest: z.string().optional(),
      responseMode: z.enum(["brief", "standard", "verbose"]).optional()
    }).optional(),
    nextAction: z.string(),
    taskOutcome: z.enum(["final_answer", "preview_ready", "apply_complete", "needs_user_input", "cannot_complete"]).optional(),
    finalAnswer: z.string().optional(),
    agentInstruction: z.string().optional(),
    maxRecommendedFollowupCalls: z.number().int().nonnegative().optional(),
    requiredFollowup: z.object({
      mode: z.enum(AGENT_RUN_MODES).optional(),
      nextAction: z.string().optional(),
      operationId: z.string().optional(),
      confirmationToken: z.string().optional(),
      instruction: z.string()
    }).optional(),
    warnings: z.array(z.string()),
    telemetry: z.object({
      internalCallCount: z.number(),
      payloadBytes: z.number(),
      estimatedTokens: z.number(),
      elapsedMs: z.number(),
      cacheHit: z.boolean(),
      autoApplied: z.boolean().optional(),
      safetyDecision: z.string().optional(),
      previewOperationId: z.string().optional(),
      validationStatus: z.enum(["passed", "failed", "not_run"]).optional(),
      metadataCacheStatus: z.enum(["hit", "miss", "not_applicable"]).optional(),
      metadataFreshnessReason: z.string().optional(),
      metadataDetailLevel: z.enum(["structure", "sampled"]).optional(),
      internalReadCount: z.number().optional(),
      fullReadCellCount: z.number().optional(),
      fullReadUsed: z.boolean().optional(),
      safetyFingerprintOnly: z.boolean().optional(),
      workflowRoute: z.string().optional(),
      workflowConfidence: z.number().optional(),
      workflowReasons: z.array(z.string()).optional(),
      semanticIndexStatus: z.enum(["built", "not_applicable"]).optional(),
      semanticEntryCount: z.number().optional(),
      semanticCandidateUsed: z.boolean().optional(),
      metadataPolicy: z.enum(["structure_only", "sampled_allowed", "sampled_required"]).optional(),
      readPolicy: z.enum(["metadata_only", "targeted_read", "preview_only", "apply_only", "not_applicable"]).optional(),
      candidateCount: z.number().optional(),
      resourceLinkCount: z.number().optional(),
      estimatedTokensSaved: z.number().optional(),
      routeMode: z.string().optional(),
      routeMatchedRule: z.string().optional(),
      routeConfidence: z.number().optional(),
      routeReasons: z.array(z.string()).optional(),
      operationRisk: z.string().optional(),
      actionHandlerId: z.string().optional(),
      autoApplyBlockedReason: z.string().optional(),
      workflowKind: z.string().optional(),
      groupedOperationCount: z.number().optional(),
      styleCopyCount: z.number().optional(),
      clearFormatCount: z.number().optional(),
      fragmentationRedirectCount: z.number().optional(),
      detectedFamily: z.string().optional(),
      suggestedWorkflowKind: z.string().optional(),
      targetFingerprintStatus: z.enum(["matched", "changed", "not_applicable"]).optional(),
      targetHintCount: z.number().optional(),
      targetHintUsed: z.boolean().optional(),
      intentSource: z.enum(["caller_structured", "deterministic_fallback", "mixed"]).optional(),
      intentAction: z.string().optional(),
      intentAccepted: z.boolean().optional(),
      intentRejectedReason: z.string().optional()
    })
  };
}

function jsonObjectString<T extends z.ZodType>(schema: T, fieldName: string): T {
  return z.preprocess((value) => {
    if (typeof value !== "string") {
      return value;
    }
    try {
      const parsed = JSON.parse(value) as unknown;
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : value;
    } catch {
      return value;
    }
  }, schema.describe(`${fieldName} must be an object. If an agent has a JSON string, it should send the parsed object instead.`)) as unknown as T;
}
