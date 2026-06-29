import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import * as z from "zod/v4";
import {
  AGENT_CONTEXT_FACETS,
  AGENT_CONTEXT_LEVELS,
  AGENT_CONTEXT_SCOPES,
  AGENT_CONTEXT_STRATEGIES,
  AGENT_DETAIL_LEVELS,
  AGENT_RUN_MODES,
  AGENT_RUN_STATUSES,
  type AgentRunExecutionContext,
  type AgentRunInput
} from "@components-kit/open-workbook-protocol";
import type { RuntimeFacade } from "../runtime-facade.js";
import { agentJsonResult } from "../results.js";

export function registerAgentTools(mcp: McpServer, runtime: RuntimeFacade, context?: AgentRunExecutionContext): void {
  (mcp.registerTool as any)(
    "excel.agent.run",
    {
      title: "Run Open Workbook agent workflow",
      description:
        "Single default Open Workbook interface. Send workbook intent; the backend handles discovery, cached metadata, target resolution, session-scoped write permission, preview/apply, validation, rollback, and compact proof without exposing low-level Excel tools. For 'what is this workbook/file', 'summarize this workbook', or broad overview requests, call mode answer with detailLevel workbook_summary or sheet_summary and answer when nextAction is answer_now; do not chase fullResultUri, chunk-read sheets, list MCP resources, or call low-level resource reads unless the user explicitly asks for all raw rows or exact cell values. In default auto mode, safe exact small edits may auto-apply after workbook write access is allowed for the session and return taskOutcome apply_complete with maxRecommendedFollowupCalls 0; use mode auto with intent.action write_values and values.patches for edits like setting one price cell. Even one cell is one patch, and patch.target is the mutation destination; top-level target is only context/default scope. Set autoApply false or use preview_update only when the user asks to review first, the edit is broad/risky/ambiguous, or the workflow is formula/style/table/structural. Do not ask the user to confirm every small exact edit once session write access exists. For read-only row lookup such as show rows where Status = Unpaid, list transactions in June, or which rows are blank, use intent.action query_rows with values.where predicates and optional values.return/limit/format; use filter_range only when the user explicitly asks to change the visible Excel filter. For visible filter/sort work, use filter_range to add autofilter controls to all columns in a resolved range, clear_table_filters to clear table filters and show full data, filter_range with values.filters or values.column plus values.value for structured table criteria, sort_table for one table sort column, and apply_table_view when one request combines clearFilters, filters, and sort fields. Ordinary ranges support enabling/clearing autofilter controls here; criteria-based filtering is table-scoped. For freeze/frozen pane status questions such as which columns/rows are frozen, call mode answer with the natural request text or put the freeze question in intent.reason; do not answer from cached workbook context, workbook_design_overview, or read_style_summary if they do not show it. For grouped-header summary questions such as summarizing row 1 grouped headers, call one mode answer request targeting the sheet/header range; the backend returns grouped_header_summary with spans, labels, and merged/unmerged status. Do not chase workbook_design_overview, semantic_index, fullResultUri, or broad row reads for the same summary. For explicit worksheet row or column structure changes, use mode preview_update with intent.action delete_rows/insert_rows/delete_columns/insert_columns and an exact row/column target such as target {sheetName:\"Invoices\", range:\"1:1\"} or target.row 1; apply the returned preview once. Do not emulate row deletion with row-height changes, blank writes, table conversion, or delete_columns. For multi-range merge plus style requests such as merging grouped header spans and centering them, send one preview_update with values.merges or values.entries containing sheetName/range/style entries; the preview must include range.merge operations before range.write_styles_many. Do not describe a merge if the returned preview only shows style updates. For column-by-column workbook design review requests such as deciding free text vs date vs money vs ID/text code vs dropdown vs lookup/reference, call intent.action workbook_design_overview once in mode answer; it returns column recommendations, related-sheet hints, and next workflows from cached metadata without reading Customer/Bookings/Drivers or empty data rows manually. For structural workbook styling such as grouped_header that inserts rows or merges group labels, if the user authorizes the formatting permission or an apply returns DESTRUCTIVE_ACTION_BLOCKED/PERMISSION_DENIED, call intent.action set_permissions with values.permissions {allowWrites:true, allowDestructiveActions:true, scopeToWorkbook:true, requireConfirmationFor:[]} and then create a fresh preview; do not tell the user you cannot grant it or ask for a manual Excel click unless set_permissions itself fails. For multiple explicit value edits from the same user request, use one mode:auto call with values.patches; independent row/range edits should still be grouped when targets and values are known. Do not issue parallel or sequential excel.agent.run update calls for related prompt work unless one grouped call fails with actionable details. For broad column/range changes like add prefix/suffix, replace text, fill blanks, normalize, or map values, call intent.action transform_values so the backend scans and previews bounded examples; do not read full columns into model context and generate a giant write matrix. For row-aware updates like fill Column X from Column Y, copy-if-blank, extract patterns, conditional maps, lookup-style derivations, or formula_like calculations such as Payment Variance = Actual Amount - Cash Amount, call intent.action derive_values so Open Workbook resolves source/target columns and compiles changed cells server-side. For exact formula inspection such as 'is this a formula?', 'raw formula', 'show formula', or 'formula in I165', call intent.action read_formulas; never infer formula existence from displayed values or numbers alone. Formula mutations, formula repairs, and formula-like broad derivations are preview/apply workflows with validation, not blind auto-applies. For full-range formula repair from one repeated same-sheet A1 pattern, send intent.action write_formulas with one values.patches entry targeting the full range and formulas [[\"=H2-G2\"]]; Open Workbook expands relative references, so do not build large formula arrays or add dummy values.values. For transaction settlement consistency involving Payment Variance, Reconciliation Note, and Detail Notes, call intent.action settle_reconciliation so Open Workbook inspects reference-month formula convention, compiles variance formula/value updates plus note updates as one grouped preview, and keeps note columns distinct by header/role. For workbook structure batches like adding a prefix/suffix to many sheet names, call intent.action transform_sheets so Open Workbook previews one bounded rename plan instead of issuing sheet-by-sheet calls. This tool can read the current live Excel selection; when the user says this, here, selected, current cell/range/row/column, or asks for values from the selected area, call excel.agent.run before asking for row or column numbers. A normal selected cell is incidental for broad workbook/worksheet overview requests. For sheet sections, use sheet_summary/semantic_index anchors; when editing by row label and column header, send values.semanticPatches with sectionId, rowMatch, columnMatch, and value instead of reading whole sections or guessing coordinates. For cross-sheet labels, historical examples, or 'how did we classify this before?' use intent.action find_similar_rows with the source row/range and any named reference sheet instead of broad-reading sheets. For formula/reference month work such as 'look at Apr 2026 for reference', use read_formulas/read_formula_patterns or validate_formula_against_template by matching headers and roles, then preview grouped formula/note repairs. For style/template references, use intent.action find_style_references before reading values. For dropdown values, call intent.action read_data_validation once on the selected/current column or exact target and answer from data_validation_summary; do not fetch fullResultUri, chunk-read sheets, list MCP resources, or read raw rows for dropdown options unless the user explicitly asks for raw audit metadata. If the user asks to read values from a source-list sheet such as Dropdown Lists, read the actual cell values with read_values or a targeted range read; do not treat the sheet name itself as validation intent. To add, rename, delete, replace, or reorder dropdown options, use preview_update with intent.action update_dropdown_options and values.operation; if the dropdown rule/source itself must change, use one preview_update with intent.action write_data_validation and one apply_update. Do not retry source formula variants, fetch resources, or test-write cells; if a fresh apply fails, report the exact error and stop. If Open Workbook is connected but a live read fails or returns a diagnostic, report that Open Workbook failure; do not fall back to Python/openpyxl/offline `.xlsx` parsing unless the user explicitly asks for offline file analysis. excel:// resultUri/fullResultUri values are internal Open Workbook handles, not web URLs; never use Webfetch/browser for them. To read stored detail, call excel.agent.run again with continuation.fullResultUri or paste the excel:// handle in request, but only when exact raw detail is necessary.",
      inputSchema: agentRunInputSchema(),
      outputSchema: agentRunOutputSchema(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false
      }
    },
    async (args: AgentRunInput) => agentJsonResult(await runtime.runAgent(normalizeAgentRunArgs(args), context))
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
    action: z.string().optional().describe("Optional internal action hint, such as read_values, query_rows, style_overview, workbook_design_overview, grouped_header, improve_visual_readability, set_permissions, read_formulas, write_values, derive_values, delete_rows, insert_rows, delete_columns, insert_columns, validate_workbook, or create_pivot_chart_summary. Invalid hints are ignored by the backend."),
    confidence: z.number().min(0).max(1).optional(),
    reason: z.string().optional(),
    targetHints: z.array(z.string()).optional()
  }), "intent");
  const contextSchema = jsonObjectString(z.object({
    level: z.enum(AGENT_CONTEXT_LEVELS.map(String) as [string, ...string[]]).transform((value) => Number(value)).or(z.number().int().min(0).max(5)).optional(),
    strategy: z.enum(AGENT_CONTEXT_STRATEGIES).optional(),
    scope: z.enum(AGENT_CONTEXT_SCOPES).optional(),
    include: z.array(z.enum(AGENT_CONTEXT_FACETS)).optional()
  }), "context");
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
  const queryPredicateSchema = z.object({
    column: z.string(),
    op: z.enum(["=", "!=", "contains", "starts_with", "ends_with", "in", "not_in", "blank", "not_blank", ">", ">=", "<", "<=", "between"]),
    value: z.any().optional()
  }).catchall(z.any());
  const valuesSchema = jsonObjectString(z.object({
    where: z.array(queryPredicateSchema).optional(),
    return: z.array(z.string()).optional(),
    limit: z.number().int().positive().optional(),
    format: z.enum(["json_rows", "csv", "summary"]).optional(),
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
      formulas: z.array(z.array(z.union([z.string(), z.null()]))).optional(),
      style: styleSchema.optional(),
      numberFormat: numberFormatSchema.optional(),
      numberFormats: numberFormatSchema.optional(),
      validation: validationSchema.optional(),
      conditionalFormatting: conditionalRuleSchema.optional(),
      note: z.string().optional(),
      comment: z.string().optional(),
      options: z.union([z.array(z.string()), z.record(z.string(), z.any())]).optional(),
      allowedValues: z.array(z.string()).optional(),
      reason: z.string().optional()
    })).optional()
  }).catchall(z.any()), "values");
  return {
    request: z.string().optional(),
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
    context: contextSchema.optional(),
    detailLevel: z.enum(AGENT_DETAIL_LEVELS).optional(),
    responseMode: z.union([z.enum(["brief", "standard", "verbose"]), z.literal("apply_update")]).optional(),
    budget: z.object({
      maxPayloadBytes: z.number().int().positive().optional(),
      maxEstimatedTokens: z.number().int().positive().optional(),
      maxExamples: z.number().int().positive().optional()
    }).optional()
  };
}

export function normalizeAgentRunArgs(args: AgentRunInput): AgentRunInput {
  const current = args as AgentRunInput & { request?: string };
  const rawResponseMode = (args as { responseMode?: unknown }).responseMode;
  if (current.mode === undefined && current.operationId && current.confirmationToken) {
    if (rawResponseMode === "apply_update") {
      const { responseMode: _responseMode, ...rest } = current;
      return normalizeAgentRunArgs({ ...rest, mode: "apply_update" } as AgentRunInput);
    }
    return normalizeAgentRunArgs({ ...current, mode: "apply_update" } as AgentRunInput);
  }
  if (typeof current.request === "string" && current.request.trim()) {
    return args;
  }
  const continuation = current.continuation;
  const handle = continuation?.fullResultUri ?? continuation?.resultUri;
  return {
    ...current,
    request: handle
      ? `Read stored Open Workbook result detail from ${handle}`
      : "Continue Open Workbook agent workflow"
  } as AgentRunInput;
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
    suggestedOperation: z.record(z.string(), z.any()).optional(),
    operationJournalRef: z.object({
      workbookContextId: z.string(),
      operationId: z.string(),
      contextVersion: z.number(),
      appliedAt: z.number()
    }).optional(),
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
    contextUsed: z.object({
      strategy: z.enum(AGENT_CONTEXT_STRATEGIES),
      scope: z.enum(AGENT_CONTEXT_SCOPES),
      levelRequested: z.number().int().min(0).max(5).optional(),
      levelUsed: z.number().int().min(0).max(5),
      levelReason: z.string().optional(),
      stagesPlanned: z.array(z.string()).optional(),
      stagesUsed: z.array(z.string()),
      skippedStages: z.array(z.string()).optional(),
      stopReason: z.string().optional(),
      included: z.array(z.string()),
      rangesRead: z.array(z.string()).optional(),
      requiredFacets: z.array(z.string()).optional(),
      cachedFacetsUsed: z.array(z.string()).optional(),
      missingFacets: z.array(z.string()).optional(),
      staleFacets: z.array(z.string()).optional(),
      facetsToRefresh: z.array(z.string()).optional(),
      refreshReason: z.string().optional(),
      freshnessRequiresRead: z.boolean().optional(),
      rowsRead: z.number().optional(),
      estimatedTokens: z.number().optional(),
      truncated: z.boolean().optional(),
      confidence: z.number().optional(),
      source: z.enum(["cache", "live", "mixed", "none"]).optional(),
      continuation: z.object({
        available: z.boolean(),
        suggestedNext: z.array(z.string()).optional()
      }).optional()
    }).optional(),
    contextFreshness: z.object({
      status: z.enum(["fresh", "mostly_fresh", "partially_stale", "stale"]),
      freshFacets: z.array(z.string()),
      staleFacets: z.array(z.string()),
      staleRanges: z.array(z.string()).optional(),
      confidence: z.number(),
      updatedAt: z.number()
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
      contextDecision: z.object({
        strategy: z.enum(AGENT_CONTEXT_STRATEGIES),
        scope: z.enum(AGENT_CONTEXT_SCOPES),
        include: z.array(z.enum(AGENT_CONTEXT_FACETS)),
        source: z.enum(["caller", "inferred"]),
        reason: z.string()
      }).optional(),
      contextRefresh: z.object({
        requiredFacets: z.array(z.string()),
        cachedFacets: z.array(z.string()),
        facetsToRefresh: z.array(z.string()),
        missingFacets: z.array(z.string()).optional(),
        staleFacets: z.array(z.string()).optional(),
        readStrategy: z.string(),
        reason: z.string(),
        requiresRead: z.boolean(),
        confidence: z.number()
      }).optional(),
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
