import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import * as z from "zod/v4";
import { type AgentRunExecutionContext, type AgentRunInput } from "@components-kit/open-workbook-protocol";
import type { RuntimeFacade } from "../runtime-facade.js";
import { agentJsonResult } from "../results.js";

export function registerAgentTools(mcp: McpServer, runtime: RuntimeFacade, context?: AgentRunExecutionContext): void {
  (mcp.registerTool as any)(
    "excel.agent.run",
    {
      title: "Run Open Workbook agent workflow",
      description:
        "Single default Open Workbook interface. Send workbook intent; the backend handles discovery, cached metadata, target resolution, preview/apply, validation, rollback, and compact proof without exposing low-level Excel tools.",
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

function agentRunInputSchema() {
  const targetSchema = z.object({
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
  return {
    request: z.string(),
    mode: z.enum(["auto", "status", "prepare", "find", "answer", "preview_update", "apply_update", "validate", "rollback"]).optional(),
    workbookContextId: z.string().optional(),
    operationId: z.string().optional(),
    transactionId: z.string().optional(),
    confirmationToken: z.string().optional(),
    continuation: z.object({
      workbookContextId: z.string().optional(),
      operationId: z.string().optional(),
      transactionId: z.string().optional(),
      resultUri: z.string().optional(),
      fullResultUri: z.string().optional(),
      nextRequest: z.string().optional(),
      responseMode: z.enum(["brief", "standard", "verbose"]).optional()
    }).optional(),
    intent: z.object({
      action: z.string().optional().describe("Optional internal action hint, such as read_values, write_values, validate_workbook, or create_pivot_chart_summary. Invalid hints are ignored by the backend."),
      confidence: z.number().min(0).max(1).optional(),
      reason: z.string().optional(),
      targetHints: z.array(z.string()).optional()
    }).optional(),
    target: targetSchema.optional(),
    values: z.record(z.string(), z.any()).and(z.object({
      patches: z.array(z.object({
        target: targetSchema,
        values: z.array(z.array(z.any())).optional(),
        rows: z.array(z.array(z.any())).optional(),
        reason: z.string().optional()
      })).optional()
    })).optional(),
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
    status: z.enum(["SUCCESS", "PREVIEW_READY", "NEEDS_INPUT", "AMBIGUOUS_TARGET", "NOT_FOUND", "STALE_CONTEXT", "VALIDATION_FAILED", "CONFLICT", "ERROR"]),
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
    continuation: z.object({
      workbookContextId: z.string().optional(),
      operationId: z.string().optional(),
      transactionId: z.string().optional(),
      resultUri: z.string().optional(),
      fullResultUri: z.string().optional(),
      nextRequest: z.string().optional(),
      responseMode: z.enum(["brief", "standard", "verbose"]).optional()
    }).optional(),
    nextAction: z.string(),
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
      internalReadCount: z.number().optional(),
      fullReadCellCount: z.number().optional(),
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
