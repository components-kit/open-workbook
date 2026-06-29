import {
  AGENT_QUERY_ROW_OPERATORS,
  makeId,
  runtimeError,
  type AgentCandidate,
  type AgentOperationId,
  type AgentProofReference,
  type AgentQueryRowsPredicate,
  type AgentRequiredFollowup,
  type AgentRunInput,
  type AgentRunExecutionContext,
  type AgentRunMode,
  type AgentRunOutput,
  type AgentRunTarget,
  type AddinTemplateRepairRequest,
  type A1Range,
  type BatchRequest,
  type CellMatrix,
  type ExcelOperation,
  type FormulaCopyPatternsRequest,
  type FormulaFillRequest,
  type FormulaPatternRequest,
  type NameCreateRequest,
  type NameSelector,
  type NameUpdateRequest,
  type OperationId,
  type RangeMetadataRequest,
  type RangeSnapshot,
  type RangeSearchRequest,
  type RegionRegisterRequest,
  type RegionSelector,
  type StyleCopyRequest,
  type StyleDimension,
  type TableApplyFiltersRequest,
  type TableApplyViewRequest,
  type TableAppendRowsRequest,
  type TableCopyStructureRequest,
  type TableCreateRequest,
  type TableReadRequest,
  type TableReorderColumnsRequest,
  type TableResizeRequest,
  type TableSelector,
  type TableSetStyleRequest,
  type TableSetTotalRowRequest,
  type TableSortField,
  type TableSortRequest,
  type TableUpdateRowsRequest,
  type TemplateCaptureRequest,
  type TemplateId,
  type TransactionId,
  type WorkbookBackupRetentionRequest,
  type WorkbookCreateFileBackupRequest,
  type WorkbookLocalConfigImportRequest,
  type WorkbookRestoreFileBackupRequest,
  type BackupId,
  type PermissionState,
  type SnapshotId,
  type WorkbookId
} from "@components-kit/open-workbook-protocol";
import {
  cellCountFromAddress,
  columnNameToNumber as columnToNumber,
  matrixCellCount,
  numberToColumnName as numberToColumn,
  rangesOverlap as rangesOverlapAddresses,
  hashStable,
  stripSheetName,
  tryParseA1Address
} from "@components-kit/open-workbook-excel-core";
import { columnLetter, normalizeHeaderName, type CacheImpactSummary, type ColumnMetadata, type ContextFacet, type HeaderMetadata, type TableMetadata, type WorkbookMetadata, WorkbookMetadataCache } from "./workbook-metadata-cache.js";
import { AgentOperationStore, type AgentCleanMutationAction, type AgentCleanRequest, type PendingAgentOperation } from "./agent-operation-store.js";
import { WorkbookMetadataBuilder } from "./workbook-metadata-builder.js";
import { findAgentCandidates, resolveAgentReadTarget, resolveAgentUpdateTarget, type AgentTargetResolution } from "./agent-target-resolver.js";
import type { RuntimeService } from "./runtime-service.js";
import { assessAgentUpdateRisk, classifyAgentActionRisk, type AgentOperationRisk } from "./agent-action-policy.js";
import { routeAgentRequest, type IntentRoute } from "./agent-routing.js";
import { isAgentIntentAction, modeForIntentAction, normalizeAgentIntent, type AgentIntentAction, type NormalizedAgentIntent } from "./agent-intent.js";
import { findAgentActionHandler, type AgentActionHandlerDefinition, type AgentActionHandlerId } from "./agent-action-handlers.js";
import { buildSemanticWorkbookIndex } from "./semantic-workbook-index.js";
import { resolveSemanticFields } from "./semantic-field-resolver.js";

const AGENT_RESULT_TTL_MS = 60 * 60 * 1000;
const AGENT_LARGE_RANGE_CELL_LIMIT = 10_000;
const AGENT_DETAILED_PREVIEW_CELL_LIMIT = 200;
const AGENT_FRAGMENT_WINDOW_MS = 2 * 60 * 1000;
const AGENT_FRAGMENT_REDIRECT_THRESHOLD = 2;

type AgentResponseMode = NonNullable<AgentRunInput["responseMode"]>;
type RangeStructuralOperationKind =
  | "range.clear_values"
  | "range.insert_rows"
  | "range.delete_rows"
  | "range.insert_columns"
  | "range.delete_columns"
  | "range.hide_columns"
  | "range.unhide_columns"
  | "range.merge"
  | "range.unmerge";

interface StoredAgentResult {
  resultId: string;
  resourceUri: string;
  fullResourceUri: string;
  workbookContextId?: string | undefined;
  freshness?: AgentResultFreshness | undefined;
  kind?: string | undefined;
  summary: string;
  hash: string;
  answer: unknown;
  createdAt: number;
  expiresAt: number;
}

interface AgentResultFreshness {
  workbookId?: WorkbookId | string;
  workbookContentVersion?: number;
  workbookStructureHash?: string;
  contextUpdatedAt?: number;
}

class AgentResultStore {
  private readonly results = new Map<string, StoredAgentResult>();

  create(input: { workbookContextId?: string | undefined; freshness?: AgentResultFreshness | undefined; summary: string; answer: unknown }): StoredAgentResult {
    this.clearExpired();
    const resultId = makeId("agentres");
    const now = Date.now();
    const kind = typeof input.answer === "object" && input.answer !== null && typeof (input.answer as { kind?: unknown }).kind === "string"
      ? (input.answer as { kind: string }).kind
      : undefined;
    const result: StoredAgentResult = {
      resultId,
      resourceUri: resultResource(resultId).uri,
      fullResourceUri: resultResource(resultId, "full").uri,
      ...(input.workbookContextId !== undefined ? { workbookContextId: input.workbookContextId } : {}),
      ...(input.freshness !== undefined ? { freshness: input.freshness } : {}),
      ...(kind !== undefined ? { kind } : {}),
      summary: input.summary,
      hash: hashStable(input.answer),
      answer: input.answer,
      createdAt: now,
      expiresAt: now + AGENT_RESULT_TTL_MS
    };
    this.results.set(resultId, result);
    return result;
  }

  get(resultId: string, options?: { view?: "summary" | "full"; maxBytes?: number }): StoredAgentResult | Record<string, unknown> | undefined {
    const result = this.results.get(resultId);
    if (!result) {
      return undefined;
    }
    if (Date.now() > result.expiresAt) {
      this.results.delete(resultId);
      return undefined;
    }
    if (options?.view === "summary") {
      return compactStoredResult(result);
    }
    if (options?.maxBytes !== undefined && options.maxBytes > 0) {
      return enforceStoredResultByteBudget(result, options.maxBytes);
    }
    return result;
  }

  invalidateByWorkbookContextId(workbookContextId: string): string[] {
    const invalidated: string[] = [];
    for (const [resultId, result] of this.results.entries()) {
      if (result.workbookContextId === workbookContextId) {
        invalidated.push(result.resourceUri, result.fullResourceUri, compactResultResource(resultId).uri);
        this.results.delete(resultId);
      }
    }
    return invalidated;
  }

  private clearExpired(now = Date.now()): void {
    for (const [resultId, result] of this.results.entries()) {
      if (now > result.expiresAt) {
        this.results.delete(resultId);
      }
    }
  }
}

export class AgentOrchestrator {
  readonly metadataCache = new WorkbookMetadataCache();
  private readonly operations: AgentOperationStore;
  private readonly results = new AgentResultStore();
  private readonly metadataBuilder: WorkbookMetadataBuilder;
  private readonly previewFragments: AgentPreviewFragment[] = [];

  constructor(private readonly runtime: RuntimeService, options: { onOperationsChanged?: () => void } = {}) {
    this.operations = new AgentOperationStore({
      ...(options.onOperationsChanged !== undefined ? { onChange: options.onOperationsChanged } : {})
    });
    this.metadataBuilder = new WorkbookMetadataBuilder(runtime, this.metadataCache);
  }

  dumpOperations(): PendingAgentOperation[] {
    return this.operations.dump();
  }

  loadOperations(operations: PendingAgentOperation[]): void {
    this.operations.load(operations);
  }

  async run(rawInput: AgentRunInput, context?: AgentRunExecutionContext): Promise<AgentRunOutput> {
    const input = applyContinuationInput(rawInput);
    const startedAt = Date.now();
    let internalCallCount = 0;
    const intent = normalizeAgentIntent(input);
    const route = routeAgentRequest(input.request, input.mode ?? "auto", intent, input.context, input.target);
    const runMetrics: AgentRunMetrics = { internalReadCount: 0, fullReadCellCount: 0, validationStatus: "not_run", route, intent };
    const mode = input.mode ?? "auto";
    const finish = (output: Omit<AgentRunOutput, "telemetry">, cacheHit = false): AgentRunOutput => {
      const outputMetrics = output.metrics as Record<string, unknown> | undefined;
      const operationRisk = typeof outputMetrics?.operationRisk === "string" ? outputMetrics.operationRisk : runMetrics.operationRisk;
      const actionHandlerId = typeof outputMetrics?.actionHandlerId === "string" ? outputMetrics.actionHandlerId : runMetrics.actionHandlerId;
      const autoApplyBlockedReason = typeof outputMetrics?.autoApplyBlockedReason === "string" ? outputMetrics.autoApplyBlockedReason : runMetrics.autoApplyBlockedReason;
      const workflowKind = typeof outputMetrics?.workflowKind === "string" ? outputMetrics.workflowKind : undefined;
      const groupedOperationCount = typeof outputMetrics?.groupedOperationCount === "number" ? outputMetrics.groupedOperationCount : undefined;
      const styleCopyCount = typeof outputMetrics?.styleCopyCount === "number" ? outputMetrics.styleCopyCount : undefined;
      const clearFormatCount = typeof outputMetrics?.clearFormatCount === "number" ? outputMetrics.clearFormatCount : undefined;
      const fragmentationRedirectCount = typeof outputMetrics?.fragmentationRedirectCount === "number" ? outputMetrics.fragmentationRedirectCount : undefined;
      const detectedFamily = typeof outputMetrics?.detectedFamily === "string" ? outputMetrics.detectedFamily : undefined;
      const suggestedWorkflowKind = typeof outputMetrics?.suggestedWorkflowKind === "string" ? outputMetrics.suggestedWorkflowKind : undefined;
      const metadataFreshnessReason = typeof outputMetrics?.metadataFreshnessReason === "string" ? outputMetrics.metadataFreshnessReason : runMetrics.metadataFreshnessReason;
      const metadataDetailLevel = outputMetrics?.metadataDetailLevel === "sampled" || outputMetrics?.metadataDetailLevel === "structure" ? outputMetrics.metadataDetailLevel : runMetrics.metadataDetailLevel;
      const targetFingerprintStatus = isTargetFingerprintStatus(outputMetrics?.targetFingerprintStatus) ? outputMetrics.targetFingerprintStatus : runMetrics.targetFingerprintStatus;
      const safetyFingerprintOnly = typeof outputMetrics?.safetyFingerprintOnly === "boolean" ? outputMetrics.safetyFingerprintOnly : runMetrics.safetyFingerprintOnly;
      const preBudgetPayloadBytes = Buffer.byteLength(JSON.stringify(output));
      const budgeted = withTaskOutcomeContract(enforceFullDataFollowupAfterBudget(applyOutputBudget(output, input, this.results, this.metadataCache), input));
      const targetHintCount = runMetrics.intent.targetHints?.length ?? 0;
      const targetHintUsed = targetHintCount > 0 && outputUsedCallerTargetHint(output);
      const preliminaryPayloadBytes = Buffer.byteLength(JSON.stringify(budgeted));
      const fullContextUsed = buildContextUsed(input, budgeted, runMetrics, cacheHit, preliminaryPayloadBytes, this.metadataCache);
      const contextUsed = compactContextUsedForBudget(fullContextUsed, input.budget?.maxPayloadBytes);
      const contextFreshness = compactContextFreshnessForBudget(budgeted.contextFreshness ?? contextFreshnessForOutput(budgeted, this.metadataCache).contextFreshness, input.budget?.maxPayloadBytes);
      const contextRefresh = buildContextRefreshTelemetry(budgeted, runMetrics, this.metadataCache);
      const budgetedWithContext = stripUndefinedOptionals({
        ...budgeted,
        ...(contextUsed !== undefined ? { contextUsed } : {}),
        ...(contextFreshness !== undefined ? { contextFreshness } : {})
      });
      const finalBudgeted = compactPostContextForTinyBudget(budgetedWithContext, input.budget?.maxPayloadBytes);
      const payloadBytes = Buffer.byteLength(JSON.stringify(finalBudgeted));
      return {
        ...finalBudgeted,
        telemetry: {
          internalCallCount,
          payloadBytes,
          estimatedTokens: Math.ceil(payloadBytes / 4),
          elapsedMs: Date.now() - startedAt,
          cacheHit,
          ...(runMetrics.autoApplied !== undefined ? { autoApplied: runMetrics.autoApplied } : {}),
          ...(runMetrics.safetyDecision !== undefined ? { safetyDecision: runMetrics.safetyDecision } : {}),
          ...(runMetrics.previewOperationId !== undefined ? { previewOperationId: runMetrics.previewOperationId } : {}),
          validationStatus: runMetrics.validationStatus,
          metadataCacheStatus: mode === "status" || mode === "apply_update" || mode === "rollback" ? "not_applicable" : cacheHit ? "hit" : "miss",
          internalReadCount: runMetrics.internalReadCount,
          fullReadCellCount: runMetrics.fullReadCellCount,
          fullReadUsed: runMetrics.fullReadUsed === true,
          ...(safetyFingerprintOnly !== undefined ? { safetyFingerprintOnly } : {}),
          candidateCount: finalBudgeted.candidates?.length ?? 0,
          resourceLinkCount: finalBudgeted.resourceLinks.length,
          estimatedTokensSaved: Math.max(0, Math.ceil((preBudgetPayloadBytes - payloadBytes) / 4)),
          routeMode: runMetrics.route.mode,
          routeMatchedRule: runMetrics.route.matchedRule,
          routeConfidence: runMetrics.route.confidence,
          routeReasons: runMetrics.route.reasons,
          workflowRoute: runMetrics.route.workflowRoute,
          workflowConfidence: runMetrics.route.workflowConfidence,
          workflowReasons: runMetrics.route.workflowReasons,
          semanticIndexStatus: runMetrics.semanticEntryCount !== undefined ? "built" : "not_applicable",
          ...(runMetrics.semanticEntryCount !== undefined ? { semanticEntryCount: runMetrics.semanticEntryCount } : {}),
          semanticCandidateUsed: outputUsedSemanticCandidate(output),
          metadataPolicy: runMetrics.route.metadataPolicy,
          readPolicy: runMetrics.route.readPolicy,
          contextDecision: runMetrics.route.contextDecision,
          ...(contextRefresh !== undefined ? { contextRefresh } : {}),
          ...(operationRisk !== undefined ? { operationRisk } : {}),
          ...(actionHandlerId !== undefined ? { actionHandlerId } : {}),
          ...(autoApplyBlockedReason !== undefined ? { autoApplyBlockedReason } : {}),
          ...(workflowKind !== undefined ? { workflowKind } : {}),
          ...(groupedOperationCount !== undefined ? { groupedOperationCount } : {}),
          ...(styleCopyCount !== undefined ? { styleCopyCount } : {}),
          ...(clearFormatCount !== undefined ? { clearFormatCount } : {}),
          ...(fragmentationRedirectCount !== undefined ? { fragmentationRedirectCount } : {}),
          ...(detectedFamily !== undefined ? { detectedFamily } : {}),
          ...(suggestedWorkflowKind !== undefined ? { suggestedWorkflowKind } : {}),
          ...(metadataFreshnessReason !== undefined ? { metadataFreshnessReason } : {}),
          ...(metadataDetailLevel !== undefined ? { metadataDetailLevel } : {}),
          ...(targetFingerprintStatus !== undefined ? { targetFingerprintStatus } : {}),
          ...(targetHintCount > 0 ? { targetHintCount, targetHintUsed } : {}),
          intentSource: runMetrics.intent.source,
          ...(runMetrics.intent.action !== undefined ? { intentAction: runMetrics.intent.action } : {}),
          intentAccepted: runMetrics.intent.accepted,
          ...(runMetrics.intent.rejectedReason !== undefined ? { intentRejectedReason: runMetrics.intent.rejectedReason } : {})
        }
      };
    };

    try {
      if (mode === "status") {
        internalCallCount += 1;
        const readiness = await this.runtime.getConnectionReadiness();
        const status = readiness.status;
        const workbookId = activeWorkbookIdFromStatus(status);
        const collaboration = compactAgentCollaborationSummary(this.runtime.getCollaborationStatus(workbookId));
        if (!readiness.ok) {
          return finish({
            status: "NEEDS_INPUT",
            mode,
            summary: connectionReadinessSummary(readiness.connectionState),
            answer: { ...status, connectionState: readiness.connectionState, collaboration },
            proof: [],
            resourceLinks: [{ uri: "excel://runtime/status", name: "runtime status", description: "Runtime connection and capability status.", mimeType: "application/json" }],
            nextAction: "ask_user",
            warnings: connectionReadinessWarnings(readiness.connectionState)
          });
        }
        return finish({
          status: "SUCCESS",
          mode,
          summary: "Open Workbook is connected to Excel and an active workbook is ready.",
          answer: { ...status, collaboration },
          proof: [],
          resourceLinks: [{ uri: "excel://runtime/status", name: "runtime status", description: "Runtime connection and capability status.", mimeType: "application/json" }],
          nextAction: "answer_now",
          warnings: []
        });
      }

      if (mode === "apply_update") {
        return finish(await this.applyUpdate(input));
      }

      if (mode === "operation_status") {
        return finish(this.operationStatus(input));
      }

      if (mode === "cancel_operation") {
        return finish(this.cancelOperation(input));
      }

      if (mode === "rollback") {
        return finish(await this.rollback(input));
      }

      if (mode === "preview_update" && input.operationId !== undefined) {
        return finish(invalidPreviewOperationReuseOutput(input));
      }

      const handleOutput = this.resourceHandleOutput(input, mode);
      if (handleOutput) {
        return finish(handleOutput, true);
      }

      const requestedOverviewIntent = workbookOverviewIntent(input);
      const effectiveMode = mode === "auto" && shouldRouteStructuredWritePayloadToPreview(input)
        ? "preview_update"
        : route.mode;
      if (effectiveMode !== route.mode) {
        runMetrics.route = {
          ...route,
          mode: effectiveMode,
          matchedRule: "structured_write_payload",
          confidence: Math.max(route.confidence, 0.9),
          reasons: ["Structured write values with an explicit target were routed to preview/update."]
        };
      }
      const includeSamples = shouldBuildSampledMetadata(input, effectiveMode, requestedOverviewIntent, route);
      internalCallCount += 1;
      const readiness = await this.runtime.getConnectionReadiness();
      if (!readiness.ok) {
        return finish({
          status: "NEEDS_INPUT",
          mode,
          summary: connectionReadinessSummary(readiness.connectionState),
          answer: { ...readiness.status, connectionState: readiness.connectionState },
          proof: [],
          resourceLinks: [{ uri: "excel://runtime/status", name: "runtime status", description: "Runtime connection and capability status.", mimeType: "application/json" }],
          nextAction: "ask_user",
          warnings: connectionReadinessWarnings(readiness.connectionState)
        });
      }
      const targetFreshnessRanges = targetFreshnessRangesFromInput(input, readiness.activeWorkbook?.workbookId);
      const { metadata, cacheHit, freshnessReason } = await this.metadataBuilder.getOrBuild({
        ...(input.workbookContextId ? { workbookContextId: String(input.workbookContextId) } : {}),
        ...(input.target?.workbookId !== undefined ? { workbookId: input.target.workbookId } : {}),
        ...(input.target?.workbookName !== undefined ? { workbookName: input.target.workbookName } : {}),
        includeSamples,
        ...(targetFreshnessRanges !== undefined ? { targetFreshnessRanges } : {})
      });
      runMetrics.metadataFreshnessReason = freshnessReason;
      runMetrics.metadataDetailLevel = metadata.detailLevel;
      runMetrics.semanticEntryCount = buildSemanticWorkbookIndex(metadata).entryCount;
      internalCallCount += cacheHit ? 1 : 4;

      const detailLevelOutput = (mode === "answer" || mode === "find" || mode === "prepare" || mode === "auto") ? detailLevelAnswerOutput(metadata, input, mode) : undefined;
      if (detailLevelOutput) {
        return finish(detailLevelOutput, cacheHit);
      }

      if (mode === "validate") {
        internalCallCount += 1;
        const validationOutput = await this.validationAnswerOutput(metadata, input, mode, runMetrics);
        if (validationOutput) {
          return finish(validationOutput, cacheHit);
        }
        const validation = await this.runtime.validateWorkbook({ workbookId: metadata.workbook.workbookId as WorkbookId });
        runMetrics.validationStatus = validation.ok === false ? "failed" : "passed";
        return finish({
          status: validation.ok === false ? "VALIDATION_FAILED" : "SUCCESS",
          mode,
          workbookContextId: metadata.workbookContextId,
          summary: validation.ok === false ? "Workbook validation reported issues." : "Workbook validation completed.",
          answer: validation,
          proof: [],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: validation.ok === false ? "manual_review" : "answer_now",
          warnings: []
        }, cacheHit);
      }

      if (!intent.accepted) {
        return finish({
          status: "VALIDATION_FAILED",
          mode,
          workbookContextId: metadata.workbookContextId,
          summary: "The caller-provided structured intent is not supported by this agent surface.",
          answer: { kind: "intent_rejected", rejectedReason: intent.rejectedReason },
          proof: [],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "ask_user",
          warnings: ["Retry without intent.action or use one of the supported agent intent actions."]
        }, cacheHit);
      }

      if (mode === "auto" && effectiveMode === "preview_update") {
        const advancedMutation = advancedMutationDecision(input);
        if (advancedMutation) {
          const advancedPreview = await this.previewAdvancedMutation(metadata, input, mode, advancedMutation.kind);
          if (advancedPreview) {
            internalCallCount += 1;
            return finish(advancedPreview, cacheHit);
          }
          runMetrics.safetyDecision = advancedMutation.safetyDecision;
          return finish({
            status: "NEEDS_INPUT",
            mode,
            workbookContextId: metadata.workbookContextId,
            summary: advancedMutation.summary,
            candidates: findAgentCandidates(metadata, input).slice(0, 5),
            proof: [],
            resourceLinks: [contextResource(metadata.workbookContextId)],
            nextAction: "manual_review",
            warnings: [advancedMutation.warning]
          }, cacheHit);
        }
      }
      if (effectiveMode === "prepare") {
        return finish(this.prepareOutput(metadata, mode), cacheHit);
      }
      if (effectiveMode === "find" && shouldInspectFormulaInline(input)) {
        internalCallCount += 1;
        const answerInput: AgentRunInput = {
          ...input,
          mode: "answer",
          intent: { ...(input.intent ?? {}), action: "read_formulas", reason: "Exact formula inspection request was routed from find to answer." }
        };
        return finish(await this.answerOutput(metadata, answerInput, "answer", runMetrics), cacheHit);
      }
      if (effectiveMode === "find") {
        const sectionAnswer = sectionAnswerOutput(metadata, input, mode);
        return finish(sectionAnswer ?? this.findOutput(metadata, input, mode), cacheHit);
      }
      if (effectiveMode === "preview_update" && isReadOnlyInspectionRequest(input.request) && !hasStructuredMutationPayload(input) && (intent.action === undefined || modeForIntentAction(intent.action) !== "preview_update")) {
        internalCallCount += 1;
        const answerInput: AgentRunInput = { ...input, mode: "answer" };
        return finish(await this.answerOutput(metadata, answerInput, "answer", runMetrics), cacheHit);
      }
      if (effectiveMode === "preview_update") {
        internalCallCount += 1;
      const preview = await this.previewUpdate(metadata, input, mode);
        if (mode !== "auto" || preview.status !== "PREVIEW_READY") {
          return finish(preview, cacheHit);
        }
        if (input.autoApply === false || process.env.OPEN_WORKBOOK_AGENT_AUTO_APPLY === "0") {
          runMetrics.safetyDecision = "manual_review:auto_apply_disabled";
          return finish({
            ...preview,
            summary: `${preview.summary} Preview is required before applying this workbook change.`,
            warnings: [...preview.warnings, "Auto-apply is disabled for agent safety. Call apply_update with the returned confirmationToken to apply."]
          }, cacheHit);
        }
        const autoDecision = autoApplyDecision(input, preview);
        runMetrics.safetyDecision = autoDecision.safetyDecision;
        if (!autoDecision.allow) {
          runMetrics.autoApplyBlockedReason = autoDecision.reason;
          return finish({
            ...preview,
            summary: `${preview.summary} Auto-apply was not used: ${autoDecision.reason}.`,
            warnings: [...preview.warnings, autoDecision.reason]
          }, cacheHit);
        }
        const operationId = preview.operationId ? String(preview.operationId) : "";
        runMetrics.previewOperationId = operationId;
        const confirmationToken = preview.confirmationToken;
        if (!confirmationToken) {
          runMetrics.safetyDecision = "manual_review:missing_confirmation_token";
          return finish({
            ...preview,
            summary: `${preview.summary} Auto-apply was not used: preview did not return a confirmation token.`,
            warnings: [...preview.warnings, "Preview did not return a confirmation token."]
          }, cacheHit);
        }
        const applied = await this.applyUpdate({ request: input.request, mode: "apply_update", operationId, confirmationToken });
        runMetrics.autoApplied = applied.status === "SUCCESS";
        runMetrics.validationStatus = applied.status === "SUCCESS" ? "passed" : applied.status === "VALIDATION_FAILED" ? "failed" : "not_run";
        const autoOutput: Omit<AgentRunOutput, "telemetry"> = {
          ...applied,
          mode,
          ...(applied.status === "SUCCESS" ? { taskOutcome: "apply_complete" as const, maxRecommendedFollowupCalls: 0 } : {}),
          ...(applied.status === "SUCCESS" && isMatchUpdatePreviewAnswer(preview.answer)
            ? {
                answer: {
                  ...preview.answer,
                  kind: "match_update_result",
                  applied: true,
                  transactionId: typeof (applied.answer as { transactionId?: unknown } | undefined)?.transactionId === "string" ? (applied.answer as { transactionId: string }).transactionId : undefined
                },
                agentInstruction: "Use matchedRows and changes from this match/update result. Answer now. Do not fetch full rows, chunk-read, or split this task into single-cell writes."
              }
            : {}),
          summary: applied.status === "SUCCESS"
            ? `${applied.summary} Auto-applied after safe preview.`
            : `${preview.summary} Auto-apply attempted after safe preview but did not complete.`,
          metrics: {
            ...(applied.metrics ?? {}),
            autoApplied: applied.status === "SUCCESS",
            previewOperationId: operationId,
            safetyDecision: autoDecision.safetyDecision,
            validationStatus: runMetrics.validationStatus
          },
          resourceLinks: applied.resourceLinks.length > 0 ? applied.resourceLinks : preview.resourceLinks
        };
        if (applied.status === "SUCCESS") {
          delete autoOutput.confirmationToken;
        }
        return finish(autoOutput, cacheHit);
      }
      internalCallCount += 1;
      return finish(await this.answerOutput(metadata, input, mode, runMetrics), cacheHit);
    } catch (error) {
      return finish({
        status: "ERROR",
        mode,
        summary: error instanceof Error ? error.message : String(error),
        proof: [],
        resourceLinks: [],
        nextAction: "manual_review",
        warnings: []
      });
    }
  }

  getContextResource(workbookContextId: string) {
    const metadata = this.metadataCache.getByContextId(workbookContextId);
    if (!metadata) {
      return { ok: false, error: runtimeError("NOT_FOUND", "Workbook context metadata was not found or expired.", { retryable: true }) };
    }
    return {
      ok: true,
      workbookContextId,
      workbook: metadata.workbook,
      semanticIndex: buildSemanticWorkbookIndex(metadata, { maxEntries: 25 }),
      ...(metadata.selection ? { selection: metadata.selection } : {}),
      sheets: metadata.sheets.map((sheet) => ({
        name: sheet.name,
        kind: sheet.kind,
        usedRange: sheet.usedRange,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        headers: sheet.headers.map((header) => ({
          range: header.range,
          confidence: header.confidence,
          columns: header.columns.map((column) => ({ name: column.name, normalizedName: column.normalizedName, inferredType: column.inferredType, role: column.role, importance: column.importance }))
        }))
      })),
      sections: metadata.sections.map((section) => ({
        id: section.id,
        sheetName: section.sheetName,
        label: section.label,
        kind: section.kind,
        range: section.range,
        headerRange: section.headerRange,
        headerRow: section.headerRow,
        columns: section.columns.map((column) => ({ name: column.name, normalizedName: column.normalizedName, inferredType: column.inferredType, role: column.role, importance: column.importance })),
        labels: section.labels,
        confidence: section.confidence
      })),
      tables: metadata.tables.map((table) => ({
        name: table.name,
        sheetName: table.sheetName,
        range: table.range,
        columns: table.columns.map((column) => ({ name: column.name, normalizedName: column.normalizedName, inferredType: column.inferredType, role: column.role, importance: column.importance }))
      })),
      namedRanges: metadata.namedRanges,
      summaryBlocks: metadata.summaryBlocks,
      formulaRegions: metadata.formulaRegions,
      contentVersion: metadata.contentVersion,
      fingerprint: metadata.fingerprint,
      freshness: {
        workbookContentVersion: metadata.contentVersion,
        workbookStructureHash: metadata.fingerprint.structureHash,
        contextUpdatedAt: metadata.updatedAt
      },
      updatedAt: metadata.updatedAt,
      expiresAt: metadata.expiresAt
    };
  }

  getSemanticIndexResource(workbookContextId: string) {
    const metadata = this.metadataCache.getByContextId(workbookContextId);
    if (!metadata) {
      return { ok: false, error: runtimeError("NOT_FOUND", "Workbook context metadata was not found or expired.", { retryable: true }) };
    }
    return {
      ok: true,
      workbookContextId,
      semanticIndex: buildSemanticWorkbookIndex(metadata)
    };
  }

  getOperationResource(operationId: string) {
    const pending = this.operations.get(operationId);
    if (!pending) {
      return { ok: false, error: runtimeError("NOT_FOUND", "Pending agent operation was not found or has already been applied.", { retryable: false }) };
    }
    return {
      ok: true,
      operationId,
      workbookContextId: pending.workbookContextId,
      workbookId: pending.workbookId,
      ...(pending.workflowKind !== undefined ? { workflowKind: pending.workflowKind } : {}),
      summary: pending.summary,
      changes: pending.changes,
      createdAt: pending.createdAt,
      expiresAt: pending.expiresAt,
      applyStatus: pending.applyStatus,
      ...(pending.applyStartedAt !== undefined ? { applyStartedAt: pending.applyStartedAt } : {}),
      ...(pending.completedAt !== undefined ? { completedAt: pending.completedAt } : {}),
      ...(pending.terminalOutput !== undefined ? { terminalOutput: pending.terminalOutput } : {})
    };
  }

  getResultResource(resultId: string, options?: { view?: "summary" | "full"; maxBytes?: number }) {
    const result = this.results.get(resultId, options);
    if (!result) {
      return { ok: false, error: runtimeError("NOT_FOUND", "Agent result was not found or expired.", { retryable: true }) };
    }
    return { ok: true, ...result };
  }

  getCompactResource(resourceId: string, options?: { view?: "summary" | "full"; maxBytes?: number }) {
    return this.getResultResource(resourceId, options);
  }

  invalidateWorkbook(workbookId: WorkbookId | string): void {
    this.metadataCache.deleteByWorkbookId(workbookId);
  }

  private recordApplyCacheImpact(pending: PendingAgentOperation): CacheImpactSummary | undefined {
    const updateRisk = pending.updateRisk;
    if (!updateRisk) {
      return undefined;
    }
    const optimisticallyUpdatedValues = updateRisk.cacheAction === "update_cache"
      ? this.metadataCache.applyOptimisticValueChanges(pending.workbookContextId, String(pending.operationId), pending.changes)
      : [];
    const updatedFacets: ContextFacet[] = optimisticallyUpdatedValues.length > 0 ? ["values"] : [];
    const staleFacets = updateRisk.invalidatedFacets
      .filter(isContextFacet)
      .filter((facet) => !(facet === "values" && optimisticallyUpdatedValues.length > 0));
    const affectedRanges = uniqueChangeRanges(pending.changes);
    const freshness = this.metadataCache.markFacetsStale(pending.workbookContextId, staleFacets, affectedRanges);
    const state = freshness ?? this.metadataCache.getContextState(pending.workbookContextId);
    if (!state) {
      return undefined;
    }
    const journalEntry = this.metadataCache.appendJournalEntry(pending.workbookContextId, {
      operationId: String(pending.operationId),
      affectedRanges,
      affectedFacets: [...new Set<ContextFacet>([...staleFacets, ...updatedFacets])],
      invalidatedFacets: staleFacets,
      preservedFacets: updateRisk.preservedFacets.filter(isContextFacet),
      changes: pending.changes,
      cacheAction: updateRisk.cacheAction === "update_cache" ? "updated_from_patch" : updateRisk.cacheAction === "none" ? "recorded" : "invalidated"
    });
    if (!journalEntry) {
      return undefined;
    }
    const latestState = this.metadataCache.getContextState(pending.workbookContextId) ?? state;
    return {
      cacheAction: journalEntry.cacheAction,
      contextVersion: latestState.contextVersion,
      freshness: latestState.freshness,
      journalEntry,
      ...(updatedFacets.length > 0 ? { updatedFacets } : {})
    };
  }

  private invalidateWorkbookContext(workbookContextId: string): { invalidatedContextIds: string[]; invalidatedResourceUris: string[] } {
    const invalidatedContextIds = this.metadataCache.deleteByContextId(workbookContextId) ? [workbookContextId] : [];
    const invalidatedResourceUris = this.results.invalidateByWorkbookContextId(workbookContextId);
    return { invalidatedContextIds, invalidatedResourceUris };
  }

  private applyContextInvalidation(pending: PendingAgentOperation): { invalidatedContextIds: string[]; invalidatedResourceUris: string[] } {
    const invalidatedResourceUris = this.results.invalidateByWorkbookContextId(pending.workbookContextId);
    if (pending.updateRisk?.cacheAction === "rebuild_context") {
      const invalidatedContextIds = this.metadataCache.deleteByContextId(pending.workbookContextId) ? [pending.workbookContextId] : [];
      return { invalidatedContextIds, invalidatedResourceUris };
    }
    return { invalidatedContextIds: [], invalidatedResourceUris };
  }

  private resourceHandleOutput(input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
    const handle = detectAgentResourceHandle(input);
    if (!handle) {
      return undefined;
    }
    if (handle.kind === "operation") {
      return this.operationStatus({ ...input, operationId: handle.id });
    }
    if (handle.kind === "result" || handle.kind === "compact") {
      const view = shouldReturnFullResource(input, handle.view) ? "full" : "summary";
      const result = this.results.get(handle.id, { view, maxBytes: view === "full" ? 24_000 : 8_000 });
      if (!result) {
        return {
          status: "NOT_FOUND",
          mode: requestedMode,
          summary: "Stored agent result was not found or expired.",
          answer: { kind: "agent_result_not_found", resultId: handle.id, retryable: true },
          proof: [],
          resourceLinks: [],
          nextAction: "retry_after_refresh",
          warnings: ["Retry the original workbook request to create a fresh result handle."]
        };
      }
      const resource = handle.kind === "compact" ? compactResultResource(handle.id) : resultResource(handle.id);
      const workbookContextId = typeof (result as { workbookContextId?: unknown }).workbookContextId === "string" ? (result as { workbookContextId: string }).workbookContextId : undefined;
      return {
        status: "SUCCESS",
        mode: requestedMode,
        ...(workbookContextId !== undefined ? { workbookContextId } : {}),
        summary: view === "full" ? "Returned stored agent result detail from a resource handle." : "Returned stored agent result summary from a resource handle.",
        answer: { kind: "agent_result_resource", view, result },
        proof: [],
        resourceLinks: [resource],
        nextAction: "answer_now",
        warnings: view === "summary" ? ["Full result detail remains behind fullResultUri. excel:// handles are internal Open Workbook handles, not web URLs; never use Webfetch/browser. Call excel.agent.run with continuation.fullResultUri when exact rows, raw values, audit detail, or transformation input is needed."] : []
      };
    }
    if (handle.kind === "semantic_index") {
      const resource = this.getSemanticIndexResource(handle.id);
      const ok = (resource as { ok?: unknown }).ok === true;
      return {
        status: ok ? "SUCCESS" : "NOT_FOUND",
        mode: requestedMode,
        workbookContextId: handle.id,
        summary: ok ? "Returned semantic workbook index from a context handle." : "Workbook context metadata was not found or expired.",
        answer: resource,
        proof: [],
        resourceLinks: [semanticIndexResource(handle.id)],
        nextAction: ok ? "answer_now" : "retry_after_refresh",
        warnings: ok ? [] : ["Retry prepare to create a fresh workbook context handle."]
      };
    }
    if (handle.kind === "context") {
      const resource = this.getContextResource(handle.id);
      const ok = (resource as { ok?: unknown }).ok === true;
      return {
        status: ok ? "SUCCESS" : "NOT_FOUND",
        mode: requestedMode,
        workbookContextId: handle.id,
        summary: ok ? "Returned cached workbook context from a handle." : "Workbook context metadata was not found or expired.",
        answer: resource,
        proof: [],
        resourceLinks: [contextResource(handle.id)],
        nextAction: ok ? "answer_now" : "retry_after_refresh",
        warnings: ok ? [] : ["Retry prepare to create a fresh workbook context handle."]
      };
    }
    return undefined;
  }

  private prepareOutput(metadata: WorkbookMetadata, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Prepared workbook context for ${metadata.workbook.name}: ${metadata.sheets.length} sheets, ${metadata.tables.length} tables, ${metadata.namedRanges.length} named ranges.`,
      answer: {
        workbook: metadata.workbook,
        ...(metadata.selection ? { selection: metadata.selection } : {}),
        sheetCount: metadata.sheets.length,
        tableCount: metadata.tables.length,
        namedRangeCount: metadata.namedRanges.length,
        sectionCount: metadata.sections.length,
        semanticIndex: buildSemanticWorkbookIndex(metadata, { maxEntries: 12 }),
        collaboration: compactAgentCollaborationSummary(this.runtime.getCollaborationStatus(metadata.workbook.workbookId as WorkbookId)),
        sheets: metadata.sheets.map((sheet) => ({ name: sheet.name, kind: sheet.kind, usedRange: sheet.usedRange, sectionCount: sheet.sectionIds.length }))
      },
      proof: metadata.sheets.slice(0, 5).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "used range" }] : []),
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: []
    };
  }

  private findOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const candidates = findAgentCandidates(metadata, input).slice(0, input.budget?.maxExamples ?? 10);
    return {
      status: candidates.length > 0 ? "SUCCESS" : "NOT_FOUND",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: candidates.length > 0 ? `Found ${candidates.length} candidate workbook target(s).` : "No matching workbook targets were found in cached metadata.",
      candidates,
      proof: candidates.flatMap((candidate) => candidate.sheetName && candidate.range ? [{ sheetName: candidate.sheetName, range: candidate.range, label: candidate.label }] : []).slice(0, 5),
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: candidates.length > 0 ? "answer_now" : "ask_user",
      warnings: []
    };
  }

  private async answerOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, runMetrics: AgentRunMetrics): Promise<Omit<AgentRunOutput, "telemetry">> {
    const validationAnswer = await this.validationAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (validationAnswer) {
      return validationAnswer;
    }
    const safetyArtifactAnswer = await this.safetyArtifactAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (safetyArtifactAnswer) {
      return safetyArtifactAnswer;
    }
    const workbookActionAnswer = await this.workbookActionAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (workbookActionAnswer) {
      return workbookActionAnswer;
    }
    const workflowAnswer = await this.workflowAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (workflowAnswer) {
      return workflowAnswer;
    }
    const freezeStatusAnswer = await freezePaneStatusAnswerOutput(this.runtime, metadata, input, requestedMode, runMetrics);
    if (freezeStatusAnswer) {
      return freezeStatusAnswer;
    }
    const workbookDumpGuard = workbookDumpGuardOutput(metadata, input, requestedMode);
    if (workbookDumpGuard) {
      return workbookDumpGuard;
    }
    const designOverviewAnswer = workbookDesignOverviewAnswerOutput(metadata, input, requestedMode);
    if (designOverviewAnswer) {
      return designOverviewAnswer;
    }
    const workbookAnswer = workbookOverviewAnswer(metadata, input, requestedMode);
    if (workbookAnswer) {
      return workbookAnswer;
    }
    const sheetOverviewAnswer = sheetOverviewAnswerOutput(metadata, input, requestedMode);
    if (sheetOverviewAnswer) {
      return sheetOverviewAnswer;
    }
    const sectionAnswer = sectionAnswerOutput(metadata, input, requestedMode);
    if (sectionAnswer) {
      return sectionAnswer;
    }
    if (intentAction(input) === "read_named_item") {
      return this.namedItemAnswerOutput(metadata, input, requestedMode, runMetrics);
    }
    if (intentAction(input) === "read_region") {
      return this.regionAnswerOutput(metadata, input, requestedMode, runMetrics);
    }
    if (intentAction(input) === "query_rows") {
      return this.queryRowsAnswerOutput(metadata, input, requestedMode, runMetrics);
    }
    const templateAnswer = await this.templateAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (templateAnswer) {
      return templateAnswer;
    }
    const referenceAnalysisAnswer = await this.referenceSheetAnalysisAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (referenceAnalysisAnswer) {
      return referenceAnalysisAnswer;
    }
    const styleAnswer = await this.styleAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (styleAnswer) {
      return styleAnswer;
    }
    const exactSearchAnswer = await this.exactRowSearchAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (exactSearchAnswer) {
      return exactSearchAnswer;
    }
    const similarRowsAnswer = await this.similarRowsAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (similarRowsAnswer) {
      return similarRowsAnswer;
    }
    const cleaningAnswer = await this.cleaningAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (cleaningAnswer) {
      return cleaningAnswer;
    }
    const repairAnswer = this.repairAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (repairAnswer) {
      return repairAnswer;
    }
    const comparisonTargets = resolveComparisonTargets(metadata, input.request);
    if (comparisonTargets.length >= 2) {
      const profiles = await this.readAndProfileRanges(metadata.workbook.workbookId as WorkbookId, comparisonTargets, runMetrics);
      return comparisonAnswerOutput(metadata, input, requestedMode, profiles);
    }
    const resolved = resolveAgentReadTarget(metadata, input);
    if (!resolved.ok) {
      return {
        status: resolved.status,
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: resolved.summary,
        ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: resolved.nextAction,
        warnings: resolved.warnings
      };
    }
    const candidates = findAgentCandidates(metadata, input);
    if (answerIntent(input) === "schema") {
      return this.schemaAnswerOutput(metadata, input, requestedMode, resolved, candidates);
    }
    const adjustedTarget = adjustReadRangeForSemanticColumn(metadata, input, resolved.sheetName, resolved.range);
    const normalizedRange = normalizeOperationRange(metadata, resolved.sheetName, adjustedTarget.range);
    if (shouldSummarizeGroupedHeader(input)) {
      return this.groupedHeaderSummaryOutput(metadata, input, requestedMode, resolved.sheetName, normalizedRange, runMetrics);
    }
    if (hasRangeMetadataReadIntent(input)) {
      const rangeMetadataAnswer = await this.rangeMetadataAnswerOutput(metadata, input, requestedMode, normalizedRange, resolved, runMetrics);
      if (rangeMetadataAnswer) {
        return rangeMetadataAnswer;
      }
    }
    if (isFormulaReadIntentAction(intentAction(input)) || shouldInspectFormulaInline(input)) {
      return this.formulaAnswerOutput(metadata, input, requestedMode, normalizedRange, resolved, runMetrics);
    }
    const largeRangeGuard = largeRangeGuardOutput(metadata, input, requestedMode, resolved.sheetName, normalizedRange, resolved.candidate.label);
    if (largeRangeGuard) {
      return largeRangeGuard;
    }
    const rangeMetadataAnswer = await this.rangeMetadataAnswerOutput(metadata, input, requestedMode, normalizedRange, resolved, runMetrics);
    if (rangeMetadataAnswer) {
      return rangeMetadataAnswer;
    }
    const table = tableFromResolution(metadata, resolved) ?? explicitlyRequestedTable(metadata, input);
    if (table && (resolved.candidate.kind === "table" || input.target?.tableName || requestExplicitlyNamesTable(input, table))) {
      return this.tableCompactAnswerOutput(metadata, input, requestedMode, resolved, table, runMetrics);
    }
    const profile = await this.readAndProfileRange(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, normalizedRange, runMetrics);
    const emptyLiveReadDiagnostic = emptyLiveReadDiagnosticOutput(metadata, input, requestedMode, resolved.sheetName, normalizedRange, resolved.candidate.label, profile);
    if (emptyLiveReadDiagnostic) {
      return emptyLiveReadDiagnostic;
    }
    const inlinePreview = inlinePreviewForProfile(input, profile, normalizedRange);
    const rangeInlineCompleteness = inlineCompletenessMetadata(inlinePreview, (profile.rows ?? []) as CellMatrix, profile.shape.rows);
    const emptyResultKind = emptyResultKindForProfile(metadata, resolved.sheetName, normalizedRange, profile);
    const aggregate = aggregateProfileForRequest(input, profile, adjustedTarget.column);
    if (aggregate) {
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Answered ${aggregate.uniqueCount} unique value(s) from ${adjustedTarget.column?.name ?? resolved.candidate.label} on ${resolved.sheetName}.`,
        answer: aggregate,
        metrics: { ...profile.metrics, uniqueCount: aggregate.uniqueCount },
        candidates: candidates.slice(0, 5),
        proof: selectionAwareProof(metadata, resolved, normalizedRange, adjustedTarget.column?.name ?? resolved.candidate.label),
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: [...adjustedTarget.warnings, ...(profile.warning ? [profile.warning] : [])]
      };
    }
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Answered from ${resolved.candidate.label} on ${resolved.sheetName} using a targeted compact read.`,
      answer: { ...profile, ...rangeInlineCompleteness, ...(emptyResultKind ? { emptyResultKind } : {}), ...inlinePreview },
      metrics: profile.metrics,
      candidates: candidates.slice(0, 5),
      proof: selectionAwareProof(metadata, resolved, normalizedRange, adjustedTarget.column?.name ?? resolved.candidate.label),
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: [...adjustedTarget.warnings, ...(profile.warning ? [profile.warning] : [])]
    };
  }

  private async queryRowsAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const values = input.values as Record<string, unknown> | undefined;
    if (!Array.isArray(values?.where) || values.where.length === 0) {
      return queryRowsContractOutput(metadata, input, requestedMode);
    }
    const table = resolveQueryRowsTable(metadata, input.target);
    if (!table) {
      return {
        status: "AMBIGUOUS_TARGET",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "query_rows needs a table target or a sheet with one known table.",
        candidates: findAgentCandidates(metadata, input).filter((candidate) => candidate.kind === "table" || candidate.kind === "sheet").slice(0, 5),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "call_with_target",
        warnings: ["Retry query_rows with target.tableName or a narrower target.sheetName."]
      };
    }
    if (!table.name) {
      return queryRowsContractOutput(metadata, input, requestedMode);
    }
    const predicates = normalizeQueryRowsPredicates(values.where);
    const returnTerms = Array.isArray(values.return) ? values.return.filter((value): value is string => typeof value === "string") : [];
    const updateColumn = stringValue(values.updateColumn ?? values.setColumn ?? values.targetColumn);
    const updateValue = values.updateValue ?? values.setValue ?? values.value;
    const fieldTerms = [...new Set([...predicates.map((predicate) => predicate.column), ...returnTerms, ...(updateColumn ? [updateColumn] : [])])];
    const fieldResolutions = resolveSemanticFields(metadata, fieldTerms, {
      ...(input.target ?? {}),
      sheetName: table.sheetName,
      tableName: table.name
    });
    const unresolved = fieldResolutions.filter((resolution) => !resolution.best || resolution.ambiguous);
    if (unresolved.length > 0) {
      return {
        status: "AMBIGUOUS_TARGET",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `query_rows could not resolve ${unresolved.map((resolution) => `"${resolution.term}"`).join(", ")} unambiguously.`,
        answer: {
          kind: "query_rows_field_candidates",
          fieldCandidates: fieldResolutions.map((resolution) => ({
            term: resolution.term,
            ambiguous: resolution.ambiguous,
            candidates: resolution.candidates
          }))
        },
        proof: [{ sheetName: table.sheetName, range: table.range, label: "query table" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "call_with_target",
        warnings: ["Retry query_rows with exact column names from fieldCandidates."]
      };
    }
    const fieldByTerm = new Map(fieldResolutions.map((resolution) => [resolution.term, resolution.best!]));
    const returnFields = returnTerms.length > 0
      ? returnTerms.map((term) => fieldByTerm.get(term)!).filter(Boolean)
      : table.columns.map((column) => ({
          field: column.name,
          sheetName: table.sheetName,
          ...(table.name ? { tableName: table.name } : {}),
          range: table.range,
          columnIndex: column.index,
          columnLetter: column.letter,
          score: 1,
          evidence: ["table column"]
        }));
    const updateField = updateColumn ? fieldByTerm.get(updateColumn) : undefined;
    const predicateFields = predicates.map((predicate) => fieldByTerm.get(predicate.column)!);
    const readColumns = [...new Set([...predicateFields, ...returnFields].map((field) => field.field))];
    const limit = clampQueryRowsLimit(values.limit);
    const scanLimit = clampQueryRowsScanLimit(values.scanLimit, limit);
    runMetrics.internalReadCount += 1;
    const result = await this.runtime.readTable({
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName: table.name,
      columns: readColumns,
      rowLimit: scanLimit
    });
    if ((result as { ok?: boolean }).ok === false) {
      return formulaRuntimeErrorOutput(metadata, requestedMode, `query_rows failed to read ${table.name ?? table.range}.`, result);
    }
    const tableResult = (result as { table?: { headers?: string[]; values?: CellMatrix } }).table;
    const headers = tableResult?.headers ?? readColumns;
    const rows = tableResult?.values ?? [];
    const matched = rows
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => predicates.every((predicate, predicateIndex) => queryPredicateMatches(row[headers.indexOf(predicateFields[predicateIndex]!.field)], predicate)));
    const returned = matched.slice(0, limit);
    const outputColumns = returnFields.map((field) => field.field);
    const rowObjects = returned.map(({ row }) => Object.fromEntries(outputColumns.map((column) => [column, row[headers.indexOf(column)]])));
    const rowAddresses = returned.map(({ index }) => queryRowAddress(table, index));
    const format = values.format === "csv" || values.format === "summary" ? values.format : "json_rows";
    const suggestedOperation = updateField && updateValue !== undefined
      ? queryRowsSuggestedPatchOperation(input, table, returned.map(({ index }) => index), updateField.columnLetter, updateValue)
      : undefined;
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `query_rows matched ${matched.length} row(s) in ${table.name ?? table.range}; returned ${returned.length}.`,
      answer: {
        kind: "query_rows_result",
        matchedRows: matched.length,
        returnedRows: returned.length,
        format,
        columns: outputColumns,
        truncated: matched.length > returned.length || rows.length >= scanLimit,
        rows: format === "csv" ? rowsToCsv(outputColumns, rowObjects) : format === "summary" ? undefined : rowObjects,
        rowAddresses,
        predicates,
        fieldCandidates: fieldResolutions.map((resolution) => ({
          term: resolution.term,
          ambiguous: resolution.ambiguous,
          candidates: resolution.candidates
        }))
      },
      proof: [{ sheetName: table.sheetName, range: table.dataRange ?? table.range, label: "queried rows" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      ...(suggestedOperation ? { suggestedOperation } : {}),
      nextAction: "answer_now",
      agentInstruction: suggestedOperation
        ? "Answer from query_rows_result. If the user asked to update matched rows, call the returned suggestedOperation once; do not apply visible Excel filters for read-only lookup."
        : "Answer from query_rows_result. Use rowAddresses for any follow-up preview patches; do not apply visible Excel filters for read-only lookup.",
      maxRecommendedFollowupCalls: 0,
      warnings: rows.length >= scanLimit ? [`Scanned row limit ${scanLimit} was reached; narrow the query or increase values.scanLimit if more rows are needed.`] : []
    };
  }

  private async workbookActionAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    let result: unknown;
    let kind = "";
    let summary = "";
    switch (action) {
      case "list_open_workbooks":
        runMetrics.internalReadCount += 1;
        result = this.runtime.listOpenWorkbooks();
        kind = "open_workbooks";
        summary = "Listed open workbook references from connected Excel sessions.";
        break;
      case "get_workbook_info":
        runMetrics.internalReadCount += 1;
        result = await this.runtime.getWorkbookInfo();
        kind = "workbook_info";
        summary = "Read workbook info from the connected Excel runtime.";
        break;
      case "refresh_workbook_snapshot": {
        const snapshotId = snapshotIdFromInput(input);
        if (!snapshotId) {
          return safetyArtifactNeedsInput(metadata, requestedMode, "Workbook snapshot refresh needs values.snapshotId, target.entity, or operationId.");
        }
        runMetrics.internalReadCount += 1;
        result = await this.runtime.refreshWorkbookSnapshot({ snapshotId: snapshotId as SnapshotId, reason: input.request });
        kind = "workbook_snapshot_refresh";
        summary = `Refreshed workbook snapshot ${snapshotId}.`;
        break;
      }
      case "get_workbook_snapshot": {
        const snapshotId = snapshotIdFromInput(input);
        if (!snapshotId) {
          return safetyArtifactNeedsInput(metadata, requestedMode, "Workbook snapshot read needs values.snapshotId, target.entity, or operationId.");
        }
        runMetrics.internalReadCount += 1;
        result = this.runtime.getWorkbookSnapshot(snapshotId as SnapshotId);
        kind = "workbook_snapshot";
        summary = `Read workbook snapshot ${snapshotId}.`;
        break;
      }
      case "detect_external_changes": {
        const snapshotId = snapshotIdFromInput(input);
        if (!snapshotId) {
          return safetyArtifactNeedsInput(metadata, requestedMode, "External change detection needs values.snapshotId, target.entity, or operationId.");
        }
        runMetrics.internalReadCount += 1;
        result = await this.runtime.detectExternalChanges({ workbookId, snapshotId: snapshotId as SnapshotId });
        kind = "workbook_external_changes";
        summary = `Compared current workbook state against snapshot ${snapshotId}.`;
        break;
      }
      case "export_local_config":
        runMetrics.internalReadCount += 1;
        result = this.runtime.exportWorkbookLocalConfig(workbookId, workbookLocalConfigOptionsFromInput(input));
        kind = "workbook_local_config_export";
        summary = "Exported local Open Workbook config for this workbook.";
        break;
      case "read_embedded_local_config":
        runMetrics.internalReadCount += 1;
        result = await this.runtime.readWorkbookEmbeddedLocalConfig(workbookId);
        kind = "workbook_embedded_local_config";
        summary = "Read embedded Open Workbook config from this workbook.";
        break;
      case "get_permissions":
        runMetrics.internalReadCount += 1;
        result = this.runtime.getPermissions();
        kind = "permissions";
        summary = "Read current Open Workbook permission policy.";
        break;
      case "set_permissions": {
        const update = permissionUpdateFromInput(input, workbookId);
        if (Object.keys(update).length === 0) {
          return safetyArtifactNeedsInput(metadata, requestedMode, "Permission update needs values.permissions or explicit permission fields such as allowWrites or allowDestructiveActions.");
        }
        result = this.runtime.setPermissions(update);
        kind = "permissions_update";
        summary = "Updated Open Workbook permission policy.";
        break;
      }
      case "allow_destructive_actions": {
        const allow = booleanValue(input.values?.allow ?? input.values?.enabled ?? input.values?.allowDestructiveActions) ?? true;
        result = this.runtime.allowDestructiveActions(allow);
        kind = "permissions_destructive_actions";
        summary = `${allow ? "Allowed" : "Blocked"} Open Workbook structure/workbook actions.`;
        break;
      }
      default:
        return undefined;
    }
    return {
      status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary,
      answer: { kind, result },
      metrics: { source: "runtime_workbook_action" },
      proof: metadata.sheets.slice(0, 1).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "workbook context" }] : []),
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
      warnings: []
    };
  }

  private async workflowAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const plan = workflowPlanForAction(action);
    if (!plan) {
      return undefined;
    }
    const target = workflowTargetFromInput(metadata, input);
    let analysis: unknown;
    if (action === "inspect_analyze" && target) {
      const profile = await this.readAndProfileRange(metadata.workbook.workbookId as WorkbookId, target.sheetName, target.range, runMetrics);
      analysis = {
        source: "range_compact_profile",
        sheetName: target.sheetName,
        range: target.range,
        profile
      };
    } else {
      runMetrics.internalReadCount += 1;
    }
    const proof = target ? [{ sheetName: target.sheetName, range: target.range, label: "workflow target" }] : [];
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `${plan.title} is available through excel.agent.run as a backend-composed workflow plan.`,
      answer: {
        kind: "workflow_plan",
        action,
        workflow: plan.workflow,
        title: plan.title,
        mutatesWorkbook: plan.mutatesWorkbook,
        steps: plan.steps,
        requiredCapabilities: plan.requiredCapabilities,
        continuation: plan.continuation,
        target,
        analysis
      },
      metrics: { source: "agent_workflow_plan", workflow: plan.workflow },
      proof,
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: plan.mutatesWorkbook ? "manual_review" : "answer_now",
      warnings: plan.warnings
    };
  }

  private schemaAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    candidates: AgentCandidate[]
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = resolveSchemaTable(metadata, input, resolved);
    if (table) {
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Answered schema for ${table.name ?? resolved.candidate.label} on ${table.sheetName} from cached workbook metadata.`,
        answer: {
          kind: "table_schema",
          source: "cached_metadata",
          tableName: table.name,
          sheetName: table.sheetName,
          range: table.range,
          headerRange: table.headerRange,
          dataRange: table.dataRange,
          columns: table.columns.map((column) => ({
            name: column.name,
            normalizedName: column.normalizedName,
            inferredType: column.inferredType,
            index: column.index,
            letter: column.letter
          })),
          contextHints: compactWorkbookContextHints(metadata, table.sheetName, table)
        },
        metrics: {
          columnCount: table.columns.length,
          source: "cached_metadata"
        },
        candidates: candidates.slice(0, 5),
        proof: [{ sheetName: table.sheetName, range: table.headerRange ?? table.range, label: table.name ?? "table schema" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: table.columns.length > 0 ? [] : ["No table columns were found in cached metadata."]
      };
    }

    const headers = resolveSchemaHeaders(metadata, resolved);
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Answered header/schema context for ${resolved.candidate.label} on ${resolved.sheetName} from cached workbook metadata.`,
      answer: {
        kind: "range_schema",
        source: "cached_metadata",
        sheetName: resolved.sheetName,
        range: resolved.range,
        headers: headers.map(headerAnswer),
        contextHints: compactWorkbookContextHints(metadata, resolved.sheetName)
      },
      metrics: {
        headerCount: headers.length,
        columnCount: headers.reduce((total, header) => Math.max(total, header.columns.length), 0),
        source: "cached_metadata"
      },
      candidates: candidates.slice(0, 5),
      proof: [{ sheetName: resolved.sheetName, range: headers[0]?.range ?? resolved.range, label: "detected headers" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: headers.length > 0 ? [] : ["No headers were found in cached metadata for the selected target."]
    };
  }

  private async formulaAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    range: string,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const request = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      sheetName: resolved.sheetName,
      address: range
    };
    const action = intentAction(input);
    runMetrics.internalReadCount += 1;
    if (action === "read_formulas" || shouldInspectFormulaInline(input)) {
      const snapshot = await this.readRangeSnapshot(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, range, ["values", "text", "formulas", "numberFormat"], runMetrics);
      const patternsResult = await this.runtime.readFormulaPatterns(request);
      const patterns = (patternsResult as { ok?: boolean; patterns?: unknown }).ok === false ? undefined : (patternsResult as { patterns?: unknown }).patterns;
      const answer = formulaReadAnswerFromSnapshot(snapshot, patterns, resolved.sheetName, range);
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: formulaReadSummary(answer, resolved.sheetName, range),
        answer,
        metrics: { source: "runtime_formula_read", formulaCount: answer.formulaCount, hardcodedCount: answer.hardcodedCount, blankCount: answer.blankCount },
        proof: [{ sheetName: resolved.sheetName, range, label: "formula read" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: answer.formulaCount === 0 ? ["No formulas were found in the requested range; returned displayed values and hardcoded/blank status."] : []
      };
    }
    if (action === "get_formula_dependency_graph") {
      const result = await this.runtime.getFormulaDependencyGraph(request);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, "Formula dependency graph is unavailable.", result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Read formula dependency graph for ${resolved.sheetName}!${range}.`,
        answer: { kind: "formula_dependency_graph", ...(result as { graph?: unknown }) },
        metrics: { source: "runtime_formula_dependency_graph" },
        proof: [{ sheetName: resolved.sheetName, range, label: resolved.candidate.label }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: formulaWarnings((result as { graph?: { warnings?: unknown[] } }).graph?.warnings)
      };
    }
    if (action === "trace_formula_precedents" || action === "trace_formula_dependents") {
      const result = action === "trace_formula_precedents"
        ? await this.runtime.traceFormulaPrecedents(request)
        : await this.runtime.traceFormulaDependents(request);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, "Formula trace is unavailable.", result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Read ${action === "trace_formula_precedents" ? "precedents" : "dependents"} for ${resolved.sheetName}!${range}.`,
        answer: { kind: action, result },
        metrics: { source: action === "trace_formula_precedents" ? "runtime_formula_trace_precedents" : "runtime_formula_trace_dependents" },
        proof: [{ sheetName: resolved.sheetName, range, label: resolved.candidate.label }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: formulaWarnings((result as { warnings?: unknown[] }).warnings)
      };
    }
    if (action === "find_formula_errors") {
      const result = await this.runtime.validateFormulas(request);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, "Formula error scan is unavailable.", result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Scanned formula errors for ${resolved.sheetName}!${range}.`,
        answer: { kind: "formula_errors", result },
        metrics: { source: "runtime_formula_find_errors" },
        proof: [{ sheetName: resolved.sheetName, range, label: resolved.candidate.label }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: validationIssueMessages(result)
      };
    }
    if (action === "explain_formula") {
      const result = await this.explainFormulaFromInput(input, request);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, "Formula explanation is unavailable.", result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Explained formula for ${resolved.sheetName}!${range}.`,
        answer: { kind: "formula_explain", result },
        metrics: { source: "runtime_formula_explain" },
        proof: [{ sheetName: resolved.sheetName, range, label: resolved.candidate.label }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: []
      };
    }
    const result = await this.runtime.readFormulaPatterns(request);
    if ((result as { ok?: boolean }).ok === false) {
      return formulaRuntimeErrorOutput(metadata, requestedMode, "Formula patterns are unavailable.", result);
    }
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Read formula patterns for ${resolved.sheetName}!${range}.`,
      answer: { kind: "formula_patterns", ...(result as { patterns?: unknown }) },
      metrics: { source: "runtime_formula_patterns" },
      proof: [{ sheetName: resolved.sheetName, range, label: resolved.candidate.label }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: formulaWarnings((result as { patterns?: { warnings?: unknown[] } }).patterns?.warnings)
    };
  }

  private async tableCompactAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    table: TableMetadata,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const tableName = table.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const values = input.values as Record<string, unknown> | undefined;
    const rowOffset = tableReadRowOffset(input);
    const rowLimit = compactTableRowLimit(input, table.columns.length, rowOffset);
    const columns = tableReadColumnsFromInput(input);
    const request: TableReadRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName,
      includeValues: values?.includeValues === false ? false : true,
      includeFormulas: values?.includeFormulas === true,
      includeText: values?.includeText === true,
      includeNumberFormats: values?.includeNumberFormats === true,
      rowOffset,
      rowLimit,
      ...(columns.length > 0 ? { columns } : {})
    };
    runMetrics.internalReadCount += 1;
    const result = await this.runtime.readTable(request);
    if ((result as { ok?: boolean }).ok === false) {
      return formulaRuntimeErrorOutput(metadata, requestedMode, `Compact table read failed for ${tableName}.`, result);
    }
    const payload = compactTablePayloadFromResult(result);
    const matrix = payload.values ?? payload.text ?? payload.formulas ?? [];
    const profile = profileValues(matrix as CellMatrix, table.dataRange ?? table.range);
    const inlinePreview = inlinePreviewForMatrix(input, matrix as CellMatrix, table.dataRange ?? table.range)
      ?? tinyInlinePreviewForMatrix(matrix as CellMatrix, table.dataRange ?? table.range);
    const totalRows = dimensionsFromAddress(table.dataRange ?? table.range)?.rows ?? profile.shape.rows;
    const inlineCompleteness = inlineCompletenessMetadata(inlinePreview, matrix as CellMatrix, totalRows);
    const rowMetadata = tableReadRowMetadata(table.dataRange ?? table.range, rowOffset, matrix as CellMatrix);
    const compactColumn = (column: TableMetadata["columns"][number]) => ({
      name: column.name,
      index: column.index,
      letter: column.letter,
      inferredType: column.inferredType,
      role: column.role,
      importance: column.importance
    });
    const projectedColumns = columns.length > 0
      ? table.columns.filter((column) => columns.includes(column.name) || columns.includes(column.index)).map(compactColumn)
      : table.columns.map(compactColumn);
    const fieldContext = tableFieldContext(table, projectedColumns, {
      values: payload.values,
      text: payload.text,
      formulas: payload.formulas,
      numberFormat: payload.numberFormat
    });
    const truncated = rowOffset + rowLimit < totalRows;
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Read compact table data from ${tableName}.`,
      answer: {
        kind: "table_compact_read",
        source: "runtime_table_read",
        tableName,
        sheetName: table.sheetName,
        range: table.range,
        dataRange: table.dataRange,
        rowOffset,
        rowLimit,
        ...inlineCompleteness,
        projectedColumns,
        fieldContext,
        truncated,
        ...(truncated ? { nextPage: { rowOffset: rowOffset + rowLimit } } : {}),
        schema: table.columns.map((column) => ({
          name: column.name,
          index: column.index,
          letter: column.letter,
          inferredType: column.inferredType,
          role: column.role,
          importance: column.importance
        })),
        profile,
        ...(rowMetadata.length > 0 ? { rowMetadata } : {}),
        ...inlinePreview,
        ...(payload.headers ? { headers: payload.headers } : {}),
        ...(payload.values ? { values: payload.values } : {}),
        ...(payload.formulas ? { formulas: payload.formulas } : {}),
        ...(payload.text ? { text: payload.text } : {}),
        ...(payload.numberFormat ? { numberFormat: payload.numberFormat } : {})
      },
      metrics: {
        source: "runtime_table_read_compact",
        rowOffset,
        rowLimit,
        columnCount: projectedColumns.length,
        truncated
      },
      proof: [{ sheetName: table.sheetName, range: table.dataRange ?? table.range, label: tableName }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: profile.warning ? [profile.warning] : []
    };
  }

  private async explainFormulaFromInput(input: AgentRunInput, request: { workbookId: WorkbookId; sheetName: string; address: string }) {
    const values = input.values as Record<string, unknown> | undefined;
    const explicitFormula = stringValue(values?.formula);
    if (explicitFormula) {
      return explainFormulaString(explicitFormula);
    }
    const patterns = await this.runtime.readFormulaPatterns(request);
    if ((patterns as { ok?: boolean }).ok === false) {
      return patterns;
    }
    const formula = (patterns as { patterns?: { cells?: Array<{ formula?: unknown }> } }).patterns?.cells?.find((cell) => typeof cell.formula === "string")?.formula;
    return typeof formula === "string"
      ? explainFormulaString(formula)
      : {
          ok: false,
          error: runtimeError("INVALID_ARGUMENT", "Provide a formula string or target a cell containing at least one formula.", { retryable: false })
        };
  }

  private async namedItemAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const selector = nameSelectorFromInput(metadata, input);
    if (!selector) {
      return nameRegionNeedsInput(metadata, requestedMode, "Named-item reads need target.candidateId, target.entity, or values.name.");
    }
    runMetrics.internalReadCount += 1;
    const result = await this.runtime.getName({
      workbookId: metadata.workbook.workbookId as WorkbookId,
      name: selector.name,
      ...(selector.sheetName !== undefined ? { sheetName: selector.sheetName } : {})
    });
    if ((result as { ok?: boolean }).ok === false) {
      return formulaRuntimeErrorOutput(metadata, requestedMode, `Named item ${selector.name} is unavailable.`, result);
    }
    const name = (result as { name?: { sheetName?: string; address?: string } }).name;
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Read named item ${selector.name}.`,
      answer: { kind: "named_item", result },
      metrics: { source: "runtime_names_get" },
      proof: name?.sheetName && name.address ? [{ sheetName: name.sheetName, range: stripSheetName(name.address), label: selector.name }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: []
    };
  }

  private async regionAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const regionName = regionNameFromInput(input);
    if (!regionName) {
      return nameRegionNeedsInput(metadata, requestedMode, "Region reads need target.candidateId, target.entity, or values.name.");
    }
    runMetrics.internalReadCount += 1;
    const result = await this.runtime.getRegion({
      workbookId: metadata.workbook.workbookId as WorkbookId,
      regionName
    });
    if ((result as { ok?: boolean }).ok === false) {
      return formulaRuntimeErrorOutput(metadata, requestedMode, `Region ${regionName} is unavailable.`, result);
    }
    const region = (result as { region?: { sheetName?: string; address?: string } }).region;
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Read region ${regionName}.`,
      answer: { kind: "region", result },
      metrics: { source: "runtime_region_get" },
      proof: region?.sheetName && region.address ? [{ sheetName: region.sheetName, range: stripSheetName(region.address), label: regionName }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: []
    };
  }

  private async templateAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    if (action === "list_templates") {
      runMetrics.internalReadCount += 1;
      const templates = this.runtime.listTemplates(workbookId);
      return templateMetadataOutput(metadata, requestedMode, "template_list", "Listed registered templates.", { templates });
    }
    if (action === "read_template") {
      const templateId = templateIdFromInput(input);
      if (!templateId) {
        return nameRegionNeedsInput(metadata, requestedMode, "Template read needs values.templateId or target.entity.");
      }
      runMetrics.internalReadCount += 1;
      const result = this.runtime.getTemplate(templateId);
      return templateMetadataOutput(metadata, requestedMode, "template", `Read template ${templateId}.`, { result });
    }
    if (action === "detect_templates") {
      runMetrics.internalReadCount += 1;
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Detected template candidates from workbook structure.",
        answer: { kind: "template_candidates", result: await this.runtime.detectTemplates(workbookId) },
        metrics: { source: "runtime_template_detect" },
        proof: metadata.sheets.flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "template candidate" }] : []).slice(0, 5),
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "infer_template_regions") {
      const templateId = templateIdFromInput(input);
      if (!templateId) {
        return nameRegionNeedsInput(metadata, requestedMode, "Template region inference needs values.templateId or target.entity.");
      }
      runMetrics.internalReadCount += 1;
      const result = this.runtime.inferTemplateRegions(templateId);
      return templateMetadataOutput(metadata, requestedMode, "template_regions", `Inferred template regions for ${templateId}.`, { result });
    }
    return undefined;
  }

  private async styleAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input)
      ?? (input.detailLevel === "style_overview" ? "style_overview" : undefined)
      ?? (runMetrics.route.workflowRoute === "style.inspect" && isStyleOverviewRequest(input.request) ? "style_overview" : undefined)
      ?? (runMetrics.route.workflowRoute === "style.inspect" ? "read_style_summary" : undefined);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    if (action === "style_overview") {
      return this.styleOverviewOutput(metadata, input, requestedMode, runMetrics);
    }
    if (action === "find_style_references" || shouldRunStyleReferenceSearch(input)) {
      const candidates = styleReferenceCandidates(metadata, input).slice(0, Math.max(1, Math.min(input.budget?.maxExamples ?? 5, 8)));
      const enriched = [];
      for (const candidate of candidates.slice(0, 3)) {
        runMetrics.internalReadCount += 1;
        const result = await this.runtime.getStyleFingerprint({
          workbookId,
          sheetName: candidate.sheetName,
          address: candidate.range,
          maxCellSamples: 60
        });
        enriched.push({
          ...candidate,
          styleSummary: compactStyleReferenceFingerprint(result)
        });
      }
      enriched.push(...candidates.slice(enriched.length));
      return {
        status: enriched.length > 0 ? "SUCCESS" : "NOT_FOUND",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: enriched.length > 0
          ? `Found ${enriched.length} style reference candidate(s).`
          : "No style reference candidates were found in workbook metadata.",
        answer: {
          kind: "style_reference_candidates",
          source: "cached_metadata_and_style_fingerprint",
          candidates: enriched
        },
        metrics: { source: "runtime_style_reference_search", candidateCount: enriched.length },
        proof: enriched.slice(0, 5).map((candidate) => ({ sheetName: candidate.sheetName, range: candidate.range, label: candidate.label })),
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: enriched.length > 0 ? "answer_now" : "call_with_target",
        warnings: []
      };
    }
    if (action === "read_style_summary") {
      const resolved = resolveAgentReadTarget(metadata, input);
      if (!resolved.ok) {
        return {
          status: resolved.status,
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: resolved.summary,
          ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
          proof: [],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: resolved.nextAction,
          warnings: resolved.warnings
        };
      }
      runMetrics.internalReadCount += 1;
      const result = await this.runtime.getStyleFingerprint({
        workbookId,
        sheetName: resolved.sheetName,
        address: resolved.range,
        maxCellSamples: 200
      });
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, `Style summary is unavailable for ${resolved.sheetName}!${resolved.range}.`, result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Read style summary for ${resolved.sheetName}!${resolved.range}.`,
        answer: { kind: "style_summary", ...styleSummaryFromFingerprint((result as { fingerprint?: unknown }).fingerprint ?? result) },
        metrics: { source: "runtime_style_summary" },
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "style summary" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: styleWarnings(result)
      };
    }
    if (action === "format_diagnostics") {
      const resolved = resolveAgentReadTarget(metadata, input);
      if (!resolved.ok) {
        return {
          status: resolved.status,
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: resolved.summary,
          ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
          proof: [],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: resolved.nextAction,
          warnings: resolved.warnings
        };
      }
      const snapshot = await this.readRangeSnapshot(workbookId, resolved.sheetName, resolved.range, ["values", "text", "formulas", "numberFormat", "style"], runMetrics);
      const diagnostics = formatDiagnosticsFromSnapshot(snapshot, resolved.sheetName, resolved.range);
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Diagnosed formatting for ${resolved.sheetName}!${resolved.range}.`,
        answer: diagnostics,
        metrics: { source: "runtime_format_diagnostics", issueCount: diagnostics.issues.length },
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "format diagnostics" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: diagnostics.issues.slice(0, 5).map((issue) => issue.message)
      };
    }
    if (action === "read_style_fingerprint") {
      const request = styleFingerprintRequestFromInput(metadata, input);
      if (!request) {
        return nameRegionNeedsInput(metadata, requestedMode, "Style fingerprint reads need target.sheetName or values.sheetName.");
      }
      runMetrics.internalReadCount += 1;
      const result = await this.runtime.getStyleFingerprint(request);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, `Style fingerprint is unavailable for ${request.sheetName}.`, result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Read style fingerprint for ${request.sheetName}${request.address ? `!${request.address}` : ""}.`,
        answer: { kind: "style_fingerprint", result },
        metrics: { source: "runtime_style_fingerprint" },
        proof: [{ sheetName: request.sheetName, range: request.address ?? usedRangeForSheet(metadata, request.sheetName), label: "style fingerprint" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "compare_style_fingerprint") {
      const request = styleCompareRequestFromInput(metadata, input);
      if (!request) {
        return nameRegionNeedsInput(metadata, requestedMode, "Style comparison needs values.source and values.destination, or source/target sheet names.");
      }
      runMetrics.internalReadCount += 2;
      const result = await this.runtime.compareStyleFingerprints(request);
      if ((result as { ok?: boolean }).ok === false && (result as { error?: unknown }).error) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, "Style comparison failed.", result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Compared styles from ${request.sourceSheetName} to ${request.targetSheetName}.`,
        answer: { kind: "style_compare", result },
        metrics: { source: "runtime_style_compare" },
        proof: [
          { sheetName: request.sourceSheetName, range: request.sourceAddress ?? usedRangeForSheet(metadata, request.sourceSheetName), label: "source style" },
          { sheetName: request.targetSheetName, range: request.targetAddress ?? usedRangeForSheet(metadata, request.targetSheetName), label: "target style" }
        ],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "get_theme" || action === "apply_theme") {
      runMetrics.internalReadCount += 1;
      const result = action === "get_theme"
        ? this.runtime.getTheme(workbookId)
        : this.runtime.applyTheme({ workbookId, theme: (input.values as Record<string, unknown> | undefined)?.theme });
      return {
        status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: action === "get_theme" ? "Workbook theme inspection is not available in the connected Excel runtime." : "Workbook theme application is not available in the connected Excel runtime.",
        answer: { kind: "theme_capability_report", result },
        metrics: { source: "runtime_theme" },
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "manual_review",
        warnings: ((result as { warnings?: Array<{ message?: string }> }).warnings ?? [])
          .map((warning) => warning.message)
          .filter((message): message is string => typeof message === "string" && message.length > 0)
      };
    }
    return undefined;
  }

  private async styleOverviewOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const resolved = resolveAgentReadTarget(metadata, input);
    if (!resolved.ok) {
      return {
        status: resolved.status,
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: resolved.summary,
        ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: resolved.nextAction,
        warnings: resolved.warnings
      };
    }
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    runMetrics.internalReadCount += 1;
    const fingerprint = await this.runtime.getStyleFingerprint({
      workbookId,
      sheetName: resolved.sheetName,
      address: resolved.range,
      maxCellSamples: 240
    });
    const styleSummary = styleSummaryFromFingerprint((fingerprint as { fingerprint?: unknown }).fingerprint ?? fingerprint);
    const overview = styleOverviewFromMetadata(metadata, resolved.sheetName, resolved.range, styleSummary);
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Returned style overview for ${resolved.sheetName}!${resolved.range}.`,
      answer: overview,
      metrics: { source: "metadata_and_runtime_style_overview", fullReadCellCount: 0, internalReadCount: 1 },
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "style overview" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: styleWarnings(fingerprint)
    };
  }

  private async similarRowsAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    if (intentAction(input) !== "find_similar_rows" && !shouldRunReferenceRowSearch(input)) {
      return undefined;
    }
    const resolved = resolveAgentReadTarget(metadata, input);
    if (!resolved.ok) {
      return {
        status: resolved.status,
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: resolved.summary,
        ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: resolved.nextAction,
        warnings: resolved.warnings
      };
    }

    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const sourceRange = normalizeOperationRange(metadata, resolved.sheetName, resolved.range);
    const sourceParsed = tryParseA1Address(stripSheetName(sourceRange));
    const sourceRowCount = sourceParsed ? sourceParsed.endRow - sourceParsed.startRow + 1 : undefined;
    const sourceValues = sourceRowCount !== undefined && sourceRowCount <= 3
      ? await this.readRangeValues(workbookId, resolved.sheetName, sourceRange, runMetrics)
      : [];
    const sourceSignals = similarRowSignals(input, sourceValues);
    const candidates = similarRowCandidateRanges(metadata, resolved.sheetName, sourceRange, input);
    const searchedRanges = shouldSearchResolvedRange(input, metadata, resolved.sheetName, sourceRange)
      ? dedupeSimilarRanges([{ sheetName: resolved.sheetName, range: clampRangeForSimilarRows(sourceRange), reason: "requested reference range" }, ...candidates])
      : candidates;
    const rowMatches: SimilarRowMatch[] = [];

    for (const candidate of searchedRanges.slice(0, 8)) {
      const rows = await this.readRangeValues(workbookId, candidate.sheetName, candidate.range, runMetrics);
      rowMatches.push(...rankSimilarRows(metadata, candidate, rows, sourceSignals));
      if (rowMatches.length >= 25) {
        break;
      }
    }

    const maxRows = Math.max(1, Math.min(input.budget?.maxExamples ?? 5, 10));
    const examples = rowMatches
      .sort((left, right) => right.score - left.score || left.sheetName.localeCompare(right.sheetName) || left.sheetRowNumber - right.sheetRowNumber)
      .slice(0, maxRows);
    return {
      status: examples.length > 0 ? "SUCCESS" : "NOT_FOUND",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: examples.length > 0
        ? `Found ${examples.length} similar historical row(s) across related workbook sheets.`
        : "No similar rows were found in bounded related-sheet samples.",
      answer: {
        kind: "similar_rows",
        source: { sheetName: resolved.sheetName, range: sourceRange },
        sourceMode: sourceValues.length > 0 ? "exact_source_row" : "request_predicates",
        signals: sourceSignals.tokens.slice(0, 12),
        predicates: sourceSignals.predicates.map((predicate) => ({ label: predicate.label, value: predicate.value })),
        comparedRanges: searchedRanges.slice(0, 8),
        rows: examples
      },
      metrics: { source: "runtime_similar_rows", comparedRangeCount: searchedRanges.slice(0, 8).length, matchedRowCount: examples.length },
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      proof: examples.map((example) => ({ sheetName: example.sheetName, range: example.range, label: "similar row" })).slice(0, 5),
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: examples.length > 0 ? "answer_now" : "call_with_target",
      warnings: searchedRanges.length > 8 ? ["Similar-row search sampled the most relevant requested and related sheets/ranges first."] : []
    };
  }

  private async exactRowSearchAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const search = resolveExactRowSearchRequest(input);
    if (!search) {
      return undefined;
    }
    const resolved = resolveAgentReadTarget(metadata, input);
    if (!resolved.ok) {
      return undefined;
    }
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const range = normalizeOperationRange(metadata, resolved.sheetName, resolved.range);
    const parsed = tryParseA1Address(stripSheetName(range));
    if (!parsed) {
      return undefined;
    }
    const values = await this.readRangeValues(workbookId, resolved.sheetName, range, runMetrics);
    const matchedRows: number[] = [];
    const evidenceRows: Array<Record<string, unknown>> = [];
    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex] ?? [];
      const matchedColumnIndexes = search.columnOffset !== undefined
        ? [search.columnOffset]
        : row.map((_value, index) => index);
      const matchedValue = matchedColumnIndexes.map((index) => row[index]).find((value) => cellMatchesPredicate(value, search.predicate));
      if (matchedValue === undefined) {
        continue;
      }
      const sheetRow = parsed.startRow + rowIndex;
      matchedRows.push(sheetRow);
      if (evidenceRows.length < 10) {
        evidenceRows.push({ row: sheetRow, matchedValue });
      }
    }
    const answer = {
      kind: matchedRows.length > 0 ? "exact_search_rows" : "exact_search_no_match",
      sheetName: resolved.sheetName,
      searchedRange: range,
      predicate: search.predicate,
      matchedRowCount: matchedRows.length,
      matchedRows,
      evidenceRows
    };
    return {
      status: matchedRows.length > 0 ? "SUCCESS" : "NEEDS_INPUT",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: matchedRows.length > 0
        ? `Found ${matchedRows.length} exact row match(es) in ${resolved.sheetName}!${range}.`
        : `No exact row matches were found in ${resolved.sheetName}!${range}.`,
      answer,
      metrics: { source: "runtime_exact_row_search", matchedRowCount: matchedRows.length },
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      proof: [{ sheetName: resolved.sheetName, range, label: "searched range" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: matchedRows.length > 0 ? "answer_now" : "ask_user",
      agentInstruction: matchedRows.length > 0
        ? "Answer from matchedRows and evidenceRows. Do not fetch full rows or chunk-read this search result."
        : "Ask the user to refine the search term, sheet, or range; do not broad-read chunks.",
      maxRecommendedFollowupCalls: 0,
      warnings: []
    };
  }

  private async referenceSheetAnalysisAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    if (intentAction(input) !== "analyze_reference_sheet" && !shouldAnalyzeReferenceSheet(input)) {
      return undefined;
    }
    const reference = resolveReferenceAnalysisTarget(metadata, input);
    if (!reference) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Reference analysis needs a reference sheet, table, or range.",
        candidates: findAgentCandidates(metadata, input).slice(0, 8),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: ["Ask the user which reference sheet/range to learn from before broad-reading workbook data."]
      };
    }

    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const snapshot = await this.readRangeSnapshot(workbookId, reference.sheetName, reference.range, ["values", "text", "formulas", "numberFormat", "style"], runMetrics);
    const analysis = referenceAnalysisFromSnapshot(metadata, input, reference, snapshot);
    const styleTargets = referenceStyleTargets(metadata, reference.sheetName, reference.range);
    const styleSummaries = [];
    for (const target of styleTargets.slice(0, 2)) {
      runMetrics.internalReadCount += 1;
      const result = await this.runtime.getStyleFingerprint({
        workbookId,
        sheetName: reference.sheetName,
        address: target.range,
        maxCellSamples: target.label === "header_style" ? 60 : 120
      });
      styleSummaries.push({
        label: target.label,
        range: target.range,
        summary: compactStyleReferenceFingerprint(result)
      });
    }

    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Analyzed reference patterns from ${reference.sheetName}!${reference.range}.`,
      answer: {
        kind: "reference_sheet_analysis",
        source: "live_read_metadata_formula_style",
        reference,
        objectives: analysis.objectives,
        structure: analysis.structure,
        columnProfiles: analysis.columnProfiles,
        relationships: analysis.relationships,
        formulaPatterns: analysis.formulaPatterns,
        stylePatterns: styleSummaries,
        rowConditionPatterns: analysis.rowConditionPatterns,
        evidenceRows: analysis.evidenceRows,
        nextRequestHints: [
          "Use this reference analysis to answer convention questions; do not chunk-read the same reference sheet for the same objective.",
          "For updates, ask the user for target rows/facts or call preview_update with grouped structured values."
        ]
      },
      metrics: {
        source: "runtime_reference_sheet_analysis",
        rowCount: analysis.structure.rowCount,
        columnCount: analysis.structure.columnCount,
        profileCount: analysis.columnProfiles.length,
        relationshipCount: analysis.relationships.length
      },
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      proof: [
        { sheetName: reference.sheetName, range: reference.range, label: "reference analysis range" },
        ...styleTargets.slice(0, 2).map((target) => ({ sheetName: reference.sheetName, range: target.range, label: target.label }))
      ],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      agentInstruction: "Answer the user's reference-sheet questions from this derived analysis. Do not broad-read or chunk-read the same reference sheet again for this task.",
      maxRecommendedFollowupCalls: 0,
      warnings: analysis.warnings
    };
  }

  private async rangeMetadataAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    normalizedRange: string,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    if (action === "read_range_compact" || action === "get_range_summary") {
      const profile = await this.readAndProfileRange(workbookId, resolved.sheetName, normalizedRange, runMetrics);
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: action === "read_range_compact" ? `Read compact range ${resolved.sheetName}!${normalizedRange}.` : `Read range summary for ${resolved.sheetName}!${normalizedRange}.`,
        answer: { kind: action === "read_range_compact" ? "range_compact" : "range_summary", profile },
        metrics: profile.metrics,
        proof: [{ sheetName: resolved.sheetName, range: normalizedRange, label: "range" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: profile.warning ? [profile.warning] : []
      };
    }
    const method = rangeMetadataMethodForAction(action ?? inferredRangeMetadataReadAction(input));
    if (!method) {
      return undefined;
    }
    const values = input.values as Record<string, unknown> | undefined;
    const request: RangeMetadataRequest | RangeSearchRequest = method === "range.search"
      ? { workbookId, sheetName: resolved.sheetName, address: normalizedRange, text: stringValue(values?.text ?? values?.query ?? input.request) ?? "" }
      : { workbookId, sheetName: resolved.sheetName, address: normalizedRange };
    runMetrics.internalReadCount += 1;
    const result = await this.runtime.readRangeMetadata(method, request);
    if (method === "range.search" && searchResultHasNoMatches(result)) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `No exact match was found in ${resolved.sheetName}!${normalizedRange}. Ask the user for a more exact search term, a different sheet/range, or permission to broaden the search.`,
        answer: {
          kind: "range_search_no_match",
          sheetName: resolved.sheetName,
          range: normalizedRange,
          searchedText: (request as RangeSearchRequest).text,
          searchedRangeWasComplete: true,
          result
        },
        metrics: { source: "runtime_range_search", matchCount: 0 },
        candidates: findAgentCandidates(metadata, input).slice(0, 5),
        proof: [{ sheetName: resolved.sheetName, range: normalizedRange, label: "searched range" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: ["No exact match was found. Do not broad-read adjacent chunks; ask the user to refine the search or target."]
      };
    }
    const answer = method === "range.read_data_validation"
      ? dataValidationSummaryAnswer(result, metadata, resolved.sheetName, normalizedRange)
      : { kind: "range_metadata", method, result };
    return {
      status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Read ${method.replace("range.", "").replace(/_/g, " ")} for ${resolved.sheetName}!${normalizedRange}.`,
      answer,
      metrics: { source: "runtime_range_metadata" },
      proof: [{ sheetName: resolved.sheetName, range: normalizedRange, label: "range metadata" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
      ...((result as { ok?: boolean }).ok === false ? {} : method === "range.read_data_validation" ? {
        taskOutcome: "final_answer" as const,
        agentInstruction: "Answer from this data_validation_summary. Dropdown validation metadata/options are complete inline for the requested range; do not fetch fullResultUri, chunk-read sheets, list MCP resources, or read raw rows unless the user explicitly asks for raw audit metadata.",
        maxRecommendedFollowupCalls: 0
      } : {}),
      warnings: []
    };
  }

  private async groupedHeaderSummaryOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    sheetName: string,
    normalizedRange: string,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const headerRange = groupedHeaderSummaryRange(metadata, sheetName, normalizedRange, input);
    runMetrics.internalReadCount += 1;
    const mergedResult = await this.runtime.readRangeMetadata("range.read_merged_cells", { workbookId, sheetName, address: headerRange });
    const snapshot = await this.readRangeSnapshot(workbookId, sheetName, headerRange, ["values", "text", "style"], runMetrics);
    const mergedRanges = mergedRangesFromMetadataResult(mergedResult);
    const spans = groupedHeaderSpansFromSnapshot(headerRange, snapshot, mergedRanges);
    const unmergedLabels = groupedHeaderUnmergedLabels(headerRange, snapshot, mergedRanges);
    const warnings = [
      ...styleWarnings(mergedResult),
      ...(mergedRanges.length === 0 ? ["No merged areas were detected in the grouped header range."] : [])
    ];
    return {
      status: (mergedResult as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Summarized grouped header row for ${sheetName}!${headerRange}.`,
      answer: {
        kind: "grouped_header_summary",
        sheetName,
        range: headerRange,
        mergedRangeCount: mergedRanges.length,
        spans,
        unmergedLabels,
        mergeStatus: mergedRanges.length > 0 ? "merged_spans_detected" : "no_merged_spans_detected",
        rawMergedCellSummary: (mergedResult as { data?: unknown }).data
      },
      metrics: {
        source: "runtime_grouped_header_summary",
        internalReadCount: 2,
        mergedRangeCount: mergedRanges.length,
        labelCount: spans.length + unmergedLabels.length
      },
      proof: [{ sheetName, range: headerRange, label: "grouped header summary" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      agentInstruction: "Answer from this grouped_header_summary. Do not call workbook_design_overview, semantic_index, fullResultUri, or broad value reads for the same grouped-header summary question.",
      maxRecommendedFollowupCalls: 0,
      warnings
    };
  }

  private async cleaningAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    if (action !== "detect_header_row" && action !== "detect_outliers" && action !== "fuzzy_match") {
      return undefined;
    }
    const request = cleanRequestFromInput(metadata, input);
    if (!request) {
      return nameRegionNeedsInput(metadata, requestedMode, "Cleaning inspection needs target.sheetName and target.range or values.sheetName and values.address.");
    }
    runMetrics.internalReadCount += 1;
    let result: unknown;
    const values = input.values as Record<string, unknown> | undefined;
    if (action === "detect_header_row") {
      const maxRows = positiveIntegerValue(values?.maxRows);
      result = await this.runtime.cleanDetectHeaderRow({ ...request, ...(maxRows !== undefined ? { maxRows } : {}) });
    } else if (action === "detect_outliers") {
      const columnIndex = numberValue(values?.columnIndex);
      const threshold = numberValue(values?.threshold);
      result = await this.runtime.cleanDetectOutliers({
        ...request,
        ...(columnIndex !== undefined ? { columnIndex } : {}),
        ...(threshold !== undefined ? { threshold } : {})
      });
    } else {
      const threshold = numberValue(values?.threshold);
      result = await this.runtime.cleanFuzzyMatch({
        ...request,
        lookupValues: stringArrayValue(values?.lookupValues),
        ...(threshold !== undefined ? { threshold } : {})
      });
    }
    return {
      status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Ran cleaning inspection ${action} on ${request.sheetName}!${request.address}.`,
      answer: { kind: "cleaning_report", action, result },
      metrics: { source: `runtime_clean_${action}` },
      proof: [{ sheetName: request.sheetName, range: request.address, label: "cleaning target" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
      warnings: (result as { warnings?: Array<{ message?: string }> }).warnings?.map((warning) => warning.message ?? String(warning)).slice(0, 5) ?? []
    };
  }

  private repairAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Omit<AgentRunOutput, "telemetry"> | undefined {
    const action = intentAction(input);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const request = repairTemplateRequestFromInput(metadata, input);
    let result: unknown;
    if (action === "repair_filters_from_template") {
      result = this.runtime.repairFiltersFromTemplate({ workbookId, ...(request?.templateId ? { templateId: request.templateId } : {}), ...(request?.targetSheetName ? { targetSheetName: request.targetSheetName } : {}) });
    } else if (action === "repair_print_layout") {
      result = this.runtime.repairPrintLayout({ workbookId });
    } else if (action === "repair_named_ranges") {
      result = this.runtime.repairNamedRanges({ workbookId });
    } else if (action === "repair_formula_errors") {
      result = this.runtime.repairFormulaErrors({ workbookId });
    } else if (action === "repair_merged_cells") {
      result = this.runtime.repairMergedCells({ workbookId });
    } else {
      return undefined;
    }
    runMetrics.internalReadCount += 1;
    return {
      status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Ran repair capability check for ${action}.`,
      answer: { kind: "repair_report", action, result },
      metrics: { source: `runtime_${action}` },
      proof: request?.targetSheetName ? [{ sheetName: request.targetSheetName, range: usedRangeForSheet(metadata, request.targetSheetName), label: "repair target" }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
      warnings: (result as { warnings?: Array<{ message?: string }> }).warnings?.map((warning) => warning.message ?? String(warning)).slice(0, 5) ?? []
    };
  }

  private async safetyArtifactAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    if (action === "list_snapshots") {
      runMetrics.internalReadCount += 1;
      const result = this.runtime.listSnapshots(workbookId);
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Listed workbook snapshots.",
        answer: { kind: "snapshot_list", result: compactSnapshotListResult(result) },
        metrics: { source: "runtime_snapshot_list" },
        proof: [],
        resourceLinks: [{ uri: `excel://workbooks/${workbookId}/snapshots`, name: "workbook snapshots", description: "Workbook snapshot metadata.", mimeType: "application/json" }],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "read_snapshot") {
      const snapshotId = snapshotIdFromInput(input);
      if (!snapshotId) {
        return safetyArtifactNeedsInput(metadata, requestedMode, "Snapshot reads need target.entity, values.snapshotId, or operationId.");
      }
      runMetrics.internalReadCount += 1;
      const result = this.runtime.getSnapshot(snapshotId as SnapshotId);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, `Snapshot ${snapshotId} is unavailable.`, result);
      }
      const snapshot = compactSnapshot((result as { snapshot?: unknown }).snapshot);
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Read compact snapshot metadata for ${snapshotId}.`,
        answer: { kind: "snapshot", snapshot },
        metrics: { source: "runtime_snapshot_get_compact" },
        proof: snapshot.affectedRanges.slice(0, 5).map((range) => ({ sheetName: range.sheetName, range: range.address, label: "snapshot range" })),
        resourceLinks: [{ uri: `excel://workbooks/${workbookId}/snapshots/${snapshotId}`, name: "snapshot", description: "Compact workbook snapshot metadata.", mimeType: "application/json" }],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "compare_snapshots") {
      const pair = snapshotPairFromInput(input);
      if (!pair) {
        return safetyArtifactNeedsInput(metadata, requestedMode, "Snapshot comparison needs values.leftSnapshotId and values.rightSnapshotId.");
      }
      runMetrics.internalReadCount += 1;
      const result = this.runtime.compareSnapshots(pair.left as SnapshotId, pair.right as SnapshotId);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, "Snapshot comparison is unavailable.", result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Compared snapshots ${pair.left} and ${pair.right}.`,
        answer: { kind: "snapshot_compare", result },
        metrics: { source: "runtime_snapshot_compare_compact" },
        proof: ((result as { diff?: { changedRanges?: A1Range[] } }).diff?.changedRanges ?? []).slice(0, 5).map((range) => ({ sheetName: range.sheetName, range: range.address, label: "changed range" })),
        resourceLinks: [{ uri: `excel://workbooks/${workbookId}/snapshots/compare`, name: "snapshot comparison", description: "Compact snapshot comparison metadata.", mimeType: "application/json" }],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "list_backups") {
      runMetrics.internalReadCount += 1;
      const result = this.runtime.listFileBackups(workbookId);
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Listed persisted workbook backups.",
        answer: { kind: "backup_list", result },
        metrics: { source: "runtime_backup_list" },
        proof: [],
        resourceLinks: [{ uri: `excel://workbooks/${workbookId}/backups`, name: "workbook backups", description: "Persisted workbook backup metadata.", mimeType: "application/json" }],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "read_backup") {
      const backupId = backupIdFromInput(input);
      if (!backupId) {
        return safetyArtifactNeedsInput(metadata, requestedMode, "Backup reads need target.entity, values.backupId, or operationId.");
      }
      runMetrics.internalReadCount += 1;
      const result = this.runtime.getFileBackup(backupId as BackupId);
      if ((result as { ok?: boolean }).ok === false) {
        return formulaRuntimeErrorOutput(metadata, requestedMode, `Backup ${backupId} is unavailable.`, result);
      }
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Read persisted backup metadata for ${backupId}.`,
        answer: { kind: "backup", result },
        metrics: { source: "runtime_backup_get" },
        proof: [],
        resourceLinks: [{ uri: `excel://workbooks/${workbookId}/backups/${backupId}`, name: "backup", description: "Persisted workbook backup metadata.", mimeType: "application/json" }],
        nextAction: "answer_now",
        warnings: []
      };
    }
    if (action === "verify_backup") {
      const backupId = backupIdFromInput(input);
      if (!backupId) {
        return safetyArtifactNeedsInput(metadata, requestedMode, "Backup verification needs target.entity, values.backupId, or operationId.");
      }
      runMetrics.internalReadCount += 1;
      const result = await this.runtime.verifyFileBackup(backupId as BackupId);
      return {
        status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: (result as { ok?: boolean }).ok === false ? `Backup verification failed for ${backupId}.` : `Verified backup ${backupId}.`,
        answer: { kind: "backup_verify", result },
        metrics: { source: "runtime_backup_verify" },
        proof: [],
        resourceLinks: [{ uri: `excel://workbooks/${workbookId}/backups/${backupId}`, name: "backup", description: "Persisted workbook backup metadata.", mimeType: "application/json" }],
        nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
        warnings: []
      };
    }
    return undefined;
  }

  private async validationAnswerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    if (!isValidationIntentAction(action)) {
      return undefined;
    }
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const request = validationRequestFromInput(metadata, input);
    let result: unknown;
    let source = "";
    switch (action) {
      case "validate_formula_range":
        result = await this.runtime.validateFormulas({ workbookId, ...(request.sheetName ? { sheetName: request.sheetName } : {}), ...(request.address ? { address: request.address } : {}) });
        source = "runtime_formula_validate";
        break;
      case "validate_formula_against_template": {
        if (!request.templateId || !request.targetSheetName) return validationNeedsInput(metadata, requestedMode, "Formula template validation needs values.templateId and target.sheetName or values.targetSheetName.");
        const template = this.runtime.getTemplate(request.templateId);
        if ((template as { ok?: boolean }).ok === false) {
          result = template;
          source = "runtime_formula_validate_template_lookup";
          break;
        }
        const sourceSheetName = (template as { template?: { sourceSheetName?: unknown } }).template?.sourceSheetName;
        if (typeof sourceSheetName !== "string" || !sourceSheetName.trim()) {
          return validationNeedsInput(metadata, requestedMode, "Formula template validation needs a registered template with a source sheet.");
        }
        result = await this.runtime.compareFormulaPatterns({
          workbookId,
          sourceSheetName,
          targetSheetName: request.targetSheetName,
          ...(request.address ? { targetAddress: request.address } : {})
        });
        source = "runtime_formula_validate_against_template";
        break;
      }
      case "validate_compact":
        result = compactValidationReport(await this.runtime.validateWorkbook({ workbookId }));
        source = "runtime_validate_compact";
        break;
      case "validate_workbook":
        result = await this.runtime.validateWorkbook({ workbookId });
        source = "runtime_validate_workbook";
        break;
      case "validate_sheet":
        if (!request.sheetName) return validationNeedsInput(metadata, requestedMode, "Sheet validation needs target.sheetName or values.sheetName.");
        result = await this.runtime.validateSheet({ workbookId, sheetName: request.sheetName });
        source = "runtime_validate_sheet";
        break;
      case "validate_template_consistency":
        if (!request.templateId || !request.targetSheetName) return validationNeedsInput(metadata, requestedMode, "Template consistency validation needs values.templateId and target.sheetName or values.targetSheetName.");
        result = await this.runtime.validateTemplateConsistency({ workbookId, templateId: request.templateId, targetSheetName: request.targetSheetName });
        source = "runtime_validate_template_consistency";
        break;
      case "validate_sheet_against_template":
        if (!request.templateId || !request.targetSheetName) return validationNeedsInput(metadata, requestedMode, "Template sheet validation needs values.templateId and target.sheetName or values.targetSheetName.");
        result = await this.runtime.validateSheetAgainstTemplate({ workbookId, templateId: request.templateId, targetSheetName: request.targetSheetName });
        source = "runtime_template_validate_sheet";
        break;
      case "validate_formulas":
        result = await this.runtime.validateFormulas({ workbookId, ...(request.sheetName ? { sheetName: request.sheetName } : {}), ...(request.address ? { address: request.address } : {}) });
        source = "runtime_validate_formulas";
        break;
      case "validate_styles":
        result = await this.runtime.validateStyles({ workbookId, ...(request.sheetName ? { sheetName: request.sheetName } : {}), ...(request.templateId ? { templateId: request.templateId } : {}), ...(request.targetSheetName ? { targetSheetName: request.targetSheetName } : {}) });
        source = "runtime_validate_styles";
        break;
      case "validate_tables":
        result = await this.runtime.validateTables({ workbookId, ...(request.tableName ? { tableName: request.tableName } : {}), ...(request.templateId ? { templateId: request.templateId } : {}) });
        source = "runtime_validate_tables";
        break;
      case "validate_table_against_template":
        if (!request.tableName || !request.templateId) return validationNeedsInput(metadata, requestedMode, "Table template validation needs target.tableName or values.tableName plus values.templateId.");
        result = await this.runtime.validateTableAgainstTemplate({ workbookId, tableName: request.tableName, templateId: request.templateId });
        source = "runtime_table_validate_against_template";
        break;
      case "validate_filters":
        result = await this.runtime.validateFilters({ workbookId, ...(request.tableName ? { tableName: request.tableName } : {}) });
        source = "runtime_validate_filters";
        break;
      case "validate_print_layout":
        result = this.runtime.validatePrintLayout({ workbookId, ...(request.templateId ? { templateId: request.templateId } : {}), ...(request.targetSheetName ? { targetSheetName: request.targetSheetName } : {}) });
        source = "runtime_validate_print_layout";
        break;
      case "validate_no_broken_references":
        result = await this.runtime.validateNoBrokenReferences({ workbookId, ...(request.sheetName ? { sheetName: request.sheetName } : {}), ...(request.address ? { address: request.address } : {}) });
        source = "runtime_validate_no_broken_references";
        break;
      case "validate_no_formula_errors":
        result = await this.runtime.validateNoFormulaErrors({ workbookId, ...(request.sheetName ? { sheetName: request.sheetName } : {}), ...(request.address ? { address: request.address } : {}) });
        source = "runtime_validate_no_formula_errors";
        break;
      case "validate_no_unintended_changes":
        if (!request.snapshotId && (!request.leftSnapshotId || !request.rightSnapshotId)) {
          return validationNeedsInput(metadata, requestedMode, "Unintended-change validation needs values.snapshotId or values.leftSnapshotId and values.rightSnapshotId.");
        }
        result = await this.runtime.validateNoUnintendedChanges({
          workbookId,
          ...(request.snapshotId ? { snapshotId: request.snapshotId as SnapshotId } : {}),
          ...(request.leftSnapshotId ? { leftSnapshotId: request.leftSnapshotId as SnapshotId } : {}),
          ...(request.rightSnapshotId ? { rightSnapshotId: request.rightSnapshotId as SnapshotId } : {})
        });
        source = "runtime_validate_no_unintended_changes";
        break;
    }
    runMetrics.internalReadCount += 1;
    runMetrics.validationStatus = (result as { ok?: boolean }).ok === false ? "failed" : "passed";
    const issues = Array.isArray((result as { issues?: unknown }).issues) ? (result as { issues: Array<{ message?: unknown; target?: A1Range }> }).issues : [];
    return {
      status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: validationSummary(action, result),
      answer: { kind: action, result },
      metrics: { source, issueCount: issues.length },
      proof: issues.flatMap((issue) => issue.target ? [{ sheetName: issue.target.sheetName, range: issue.target.address, label: "validation issue" }] : []).slice(0, 5),
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
      warnings: issues.slice(0, 5).flatMap((issue) => typeof issue.message === "string" ? [issue.message] : [])
    };
  }

  private async previewUpdate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry">> {
    const matchUpdate = await this.previewMatchUpdate(metadata, input, requestedMode);
    if (matchUpdate) {
      return matchUpdate;
    }
    if (hasMergeBatchInput(input)) {
      const mergePreview = this.previewMergeRangesWithStyles(metadata, input, requestedMode);
      if (mergePreview) {
        return mergePreview;
      }
    }
    if (intentAction(input) !== "write_formulas" && hasStyleBatchInput(input) && styleEntriesFromInput(metadata, metadata.workbook.workbookId as WorkbookId, input).length > 0) {
      return this.previewWriteStylesMany(metadata, input, requestedMode);
    }
    const patchValidation = validateDirectPatchInput(input);
    if (!patchValidation.ok) {
      return canonicalPatchRequiredOutput(metadata, requestedMode, patchValidation.summary, patchValidation.warnings);
    }
    const patches = valuePatchesFromInput(input);
    if (patches.length > 0) {
      return this.previewPatchUpdate(metadata, input, requestedMode, patches);
    }
    const semanticPatches = await this.resolveSemanticValuePatches(metadata, input, requestedMode);
    if (!semanticPatches.ok) {
      return semanticPatches.output;
    }
    if (semanticPatches.patches.length > 0) {
      return this.previewPatchUpdate(metadata, input, requestedMode, semanticPatches.patches);
    }
    if (isPotentialBroadMutationRequest(input)) {
      const scoped = resolveUpdateTarget(metadata, input);
      if (scoped.ok) {
        const guard = broadMutationNeedsScopeOutput(metadata, input, requestedMode, scoped, objectToCellMatrix(input.values ?? {}));
        if (guard) {
          return guard;
        }
      } else {
        const guard = broadMutationExplicitTargetNeedsScopeOutput(metadata, input, requestedMode);
        if (guard) {
          return guard;
        }
      }
    }
    const directMutationIssue = legacyDirectMutationPatchIssue(input);
    if (directMutationIssue) {
      return canonicalPatchRequiredOutput(metadata, requestedMode, directMutationIssue.summary, directMutationIssue.warnings);
    }
    const operationPreview = await this.previewOperationIntent(metadata, input, requestedMode);
    if (operationPreview) {
      return operationPreview;
    }
    if (usesLegacyDirectValueWrite(input)) {
      return canonicalPatchRequiredOutput(
        metadata,
        requestedMode,
        "Direct range value updates must use values.patches. Top-level target plus values.values is no longer accepted for agent-facing direct writes.",
        ["Use values.patches for direct cell/range updates, including single-cell updates."]
      );
    }
    const resolved = resolveUpdateTarget(metadata, input);
    if (!resolved.ok) {
      const guard = broadMutationExplicitTargetNeedsScopeOutput(metadata, input, requestedMode);
      if (guard) {
        return guard;
      }
      return {
        status: resolved.status,
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: resolved.summary,
        ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: resolved.nextAction,
        warnings: resolved.warnings
      };
    }
    const rawValues = input.values as Record<string, unknown> | undefined;
    if (rawValues?.validation && intentAction(input) !== "write_data_validation") {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Data validation rules must use intent.action write_data_validation; they cannot be applied through generic value writes.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "validation target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "call_preview_update",
        warnings: ["Use values.validation.source or values.options with intent.action write_data_validation."]
      };
    }
    const matrix = objectToCellMatrix(input.values ?? {});
    const broadMutationGuard = broadMutationNeedsScopeOutput(metadata, input, requestedMode, resolved, matrix);
    if (broadMutationGuard) {
      return broadMutationGuard;
    }
    const dateNormalization = normalizeShortYearDatesForWrite(metadata.workbook.workbookId as WorkbookId, input, resolved.range, matrix);
    const writeMatrix = dateNormalization.matrix;
    const shapeIssue = matrixShapeIssue(resolved.range, writeMatrix);
    if (shapeIssue) {
      return {
        status: "VALIDATION_FAILED",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Value payload has a shape mismatch for ${resolved.sheetName}!${resolved.range}.`,
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "shape mismatch" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "manual_review",
        warnings: [shapeIssue]
      };
    }
    if (containsFormulaLikeValue(writeMatrix) && (intentAction(input) === "write_formulas" || (!isConditionalFormattingMutationRequest(input.request) && isFormulaMutationRequest(input.request)))) {
      return this.previewFormulaUpdate(metadata, input, requestedMode, resolved, writeMatrix);
    }
    if ((intentAction(input) === "append_table_rows" || isTableAppendIntent(input.request)) && resolved.candidate.kind === "table") {
      return this.previewTableAppend(metadata, input, requestedMode, resolved, writeMatrix);
    }
    if (isSparseBroadWrite(resolved.range, writeMatrix)) {
      return {
        status: "VALIDATION_FAILED",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Preview blocked a sparse broad overwrite. Provide a smaller target range or an explicit clear operation.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "blocked target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: ["Sparse/null-padded broad writes are blocked by the agent workflow."]
      };
    }
    if (containsFormulaLikeValue(writeMatrix) && intentAction(input) !== "write_conditional_formatting") {
      return {
        status: "VALIDATION_FAILED",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Preview blocked formula-like values in a generic value update.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula-like value target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "manual_review",
        warnings: ["Use a formula-aware workflow for values that start with '='."]
      };
    }
    const formulaWarnings = overlapsFormulaRegion(metadata, resolved.sheetName, resolved.range)
      ? ["Target overlaps detected formula regions. Review carefully before applying."]
      : [];
    if (formulaWarnings.length > 0 && (!input.target?.range || (cellCountFromAddress(resolved.range) ?? 0) > Math.max(1, matrixCellCount(writeMatrix)))) {
      return {
        status: "VALIDATION_FAILED",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Preview blocked an update that overlaps detected formula regions.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula-protected target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "manual_review",
        warnings: ["Use a formula-aware repair/update path for formula regions."]
      };
    }
    if (shouldTrackFragmentedValueWrite(input.request)) {
      const redirect = this.fragmentationRedirect(metadata, requestedMode, {
        family: "write_values",
        workbookContextId: metadata.workbookContextId,
        targetSheetName: resolved.sheetName,
        targetAddress: resolved.range,
        request: input.request,
        values: writeMatrix
      });
      if (redirect) return redirect;
    }
    const detailedPreview = shouldBuildDetailedPreviewChanges(writeMatrix);
    const before = detailedPreview ? await this.readRangeValues(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, resolved.range) : [];
    const operation: ExcelOperation = {
      kind: "range.write_values",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      values: writeMatrix,
      preserveFormats: true
    };
    const operations: ExcelOperation[] = [operation];
    if (dateNormalization.numberFormatEntries.length > 0) {
      operations.push({
        kind: "range.write_number_formats_many",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "format",
        reason: "Apply date number format for normalized short-year date values.",
        entries: dateNormalization.numberFormatEntries
      });
    }
    const changes = previewChangesForMatrix(resolved.sheetName, resolved.range, writeMatrix, before);
    const action = { kind: "batch" as const, operations };
    const pending = this.createPendingOperation(metadata, {
      action,
      changes,
      summary: `Prepared ${changes.length} cell update(s) on ${resolved.sheetName}!${resolved.range}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", safetyFingerprintOnly: true },
      changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "preview target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: [...formulaWarnings, ...dateNormalization.warnings]
    };
  }

  private async resolveSemanticValuePatches(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode
  ): Promise<{ ok: true; patches: AgentValuePatch[] } | { ok: false; output: Omit<AgentRunOutput, "telemetry"> }> {
    const requests = semanticValuePatchesFromInput(input);
    if (requests.length === 0) {
      return { ok: true, patches: [] };
    }

    const patches: AgentValuePatch[] = [];
    const dataRangeCache = new Map<string, CellMatrix>();
    for (const [index, request] of requests.entries()) {
      const section = findSectionForSemanticPatch(metadata, input, request);
      if (!section.ok) {
        return {
          ok: false,
          output: semanticPatchNeedsInput(metadata, requestedMode, `Semantic patch ${index + 1} could not resolve a workbook section. ${section.summary}`, section.warnings)
        };
      }
      const dataRange = dataRangeForSection(section.section);
      if (!dataRange) {
        return {
          ok: false,
          output: semanticPatchNeedsInput(metadata, requestedMode, `Semantic patch ${index + 1} resolved section ${section.section.id}, but that section has no data range below its header.`, [
            "Use an explicit target range or choose a section with header/data anchors."
          ])
        };
      }
      const targetColumn = findSectionColumn(section.section, request.columnMatch);
      if (!targetColumn) {
        return {
          ok: false,
          output: semanticPatchNeedsInput(metadata, requestedMode, `Semantic patch ${index + 1} could not match target column "${String(request.columnMatch)}" in section ${section.section.id}.`, [
            `Available columns: ${section.section.columns.map((column) => column.name || column.letter).filter(Boolean).join(", ")}`
          ])
        };
      }
      const dataRangeCacheKey = `${section.section.sheetName}!${dataRange}`;
      let dataValues = dataRangeCache.get(dataRangeCacheKey);
      if (!dataValues) {
        dataValues = await this.readRangeValues(metadata.workbook.workbookId as WorkbookId, section.section.sheetName, dataRange);
        dataRangeCache.set(dataRangeCacheKey, dataValues);
      }
      const rowMatch = findSemanticPatchRow(section.section, dataRange, dataValues, request.rowMatch);
      if (!rowMatch.ok) {
        return {
          ok: false,
          output: semanticPatchNeedsInput(metadata, requestedMode, `Semantic patch ${index + 1} could not match row "${String(request.rowMatch.value)}" in section ${section.section.id}.`, rowMatch.warnings)
        };
      }
      const targetAddress = `${columnLetter(targetColumn.index)}${rowMatch.sheetRow}`;
      const values: CellMatrix = [[request.value as CellMatrix[number][number]]];
      patches.push({
        target: { sheetName: section.section.sheetName, range: targetAddress },
        values,
        reason: request.reason ?? `Resolved ${section.section.id} row "${String(request.rowMatch.value)}" column "${targetColumn.name || targetColumn.letter}".`
      });
    }
    return { ok: true, patches };
  }

  private async previewPatchUpdate(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    patches: AgentValuePatch[]
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const entries: Array<{ target: A1Range; values: CellMatrix; preserveFormats?: true }> = [];
    const numberFormatEntries: Array<{ target: A1Range; numberFormat: string[][] }> = [];
    const changes: NonNullable<AgentRunOutput["changes"]> = [];
    const warnings: string[] = [];
    const validationChecks: Array<Record<string, unknown>> = [];
    let cellCount = 0;
    let detailedPreviewCells = 0;

    for (const [index, patch] of patches.entries()) {
      const patchTarget = inheritPatchTargetSheet(input, patch.target);
      const resolved = resolveUpdateTarget(metadata, {
        ...input,
        target: patchTarget,
        values: { values: patch.values }
      });
      if (!resolved.ok) {
        return {
          status: resolved.status,
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: `Patch ${index + 1} could not be resolved. ${resolved.summary}`,
          ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
          proof: [],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: resolved.nextAction,
          warnings: resolved.warnings
        };
      }
      const dateNormalization = normalizeShortYearDatesForWrite(metadata.workbook.workbookId as WorkbookId, {
        ...input,
        target: patchTarget
      }, resolved.range, patch.values);
      const writeMatrix = dateNormalization.matrix;
      numberFormatEntries.push(...dateNormalization.numberFormatEntries);
      warnings.push(...dateNormalization.warnings);
      const shapeIssue = matrixShapeIssue(resolved.range, writeMatrix);
      if (shapeIssue) {
        return {
          status: "VALIDATION_FAILED",
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: `Patch ${index + 1} has a shape mismatch for ${resolved.sheetName}!${resolved.range}.`,
          proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "shape mismatch" }],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "manual_review",
          warnings: [shapeIssue]
        };
      }
      if (isSparseBroadWrite(resolved.range, writeMatrix)) {
        return {
          status: "VALIDATION_FAILED",
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: `Patch ${index + 1} would sparsely overwrite ${resolved.sheetName}!${resolved.range}.`,
          proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "blocked patch target" }],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "ask_user",
          warnings: ["Sparse/null-padded broad writes are blocked by the agent workflow."]
        };
      }
      if (containsFormulaLikeValue(writeMatrix)) {
        return {
          status: "VALIDATION_FAILED",
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: `Patch ${index + 1} contains formula-like values.`,
          proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula-like patch target" }],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "manual_review",
          warnings: ["Use a formula-aware workflow for values that start with '='."]
        };
      }
      const formulaWarning = overlapsFormulaRegion(metadata, resolved.sheetName, resolved.range)
        ? `Patch ${index + 1} overlaps detected formula regions at ${resolved.sheetName}!${resolved.range}. Review carefully before applying.`
        : undefined;
      if (formulaWarning && (!patch.target.range || (cellCountFromAddress(resolved.range) ?? 0) > Math.max(1, matrixCellCount(writeMatrix)))) {
        return {
          status: "VALIDATION_FAILED",
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: `Patch ${index + 1} overlaps detected formula regions.`,
          proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula-protected patch target" }],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "manual_review",
          warnings: ["Use a formula-aware repair/update path for formula regions."]
        };
      }
      if (formulaWarning) {
        warnings.push(formulaWarning);
      }
      const validationCheck = await this.patchDataValidationCheck(metadata, index, resolved.sheetName, resolved.range, writeMatrix);
      if (!validationCheck.ok) {
        return {
          status: "VALIDATION_FAILED",
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: validationCheck.summary,
          answer: validationCheck.answer,
          proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "dropdown validation" }],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "manual_review",
          warnings: validationCheck.warnings
        };
      }
      warnings.push(...validationCheck.warnings);
      validationChecks.push(...validationCheck.checks);
      const patchCellCount = matrixCellCount(writeMatrix);
      const detailedPreview = patchCellCount <= Math.max(0, AGENT_DETAILED_PREVIEW_CELL_LIMIT - detailedPreviewCells);
      const before = detailedPreview ? await this.readRangeValues(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, resolved.range) : [];
      if (detailedPreview) {
        detailedPreviewCells += patchCellCount;
      }
      entries.push({
        target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
        values: writeMatrix,
        preserveFormats: true
      });
      cellCount += patchCellCount;
      changes.push(...previewChangesForMatrix(resolved.sheetName, resolved.range, writeMatrix, before, !detailedPreview));
    }

    const pending = this.createPendingOperation(metadata, {
      action: {
        kind: "batch",
        operations: [
          {
            kind: "range.write_values_many",
            operationId: makeId<OperationId>("op"),
            workbookId: metadata.workbook.workbookId as WorkbookId,
            destructiveLevel: "values",
            reason: input.request,
            entries
          },
          ...(numberFormatEntries.length > 0 ? [{
            kind: "range.write_number_formats_many" as const,
            operationId: makeId<OperationId>("op"),
            workbookId: metadata.workbook.workbookId as WorkbookId,
            destructiveLevel: "format" as const,
            reason: "Apply date number format for normalized short-year date values.",
            entries: numberFormatEntries
          }] : [])
        ]
      },
      changes,
      summary: `Prepared ${cellCount} cell update(s) across ${patches.length} grouped range patch(es).`
    });
    const proof = uniqueProofFromChanges(changes).slice(0, 8);
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: `${pending.summary} Apply this grouped preview once with apply_update; do not split these related patches unless apply returns a hard failure.`,
      answer: {
        kind: "multi_range_preview",
        patchCount: patches.length,
        cellCount,
        operationCount: 1,
        grouped: true,
        updateRisk: pending.updateRisk,
        ...(validationChecks.length > 0 ? { validationChecks } : {})
      },
      metrics: { operationRisk: pending.risk, updateRisk: pending.updateRisk, targetFingerprintStatus: "matched", safetyFingerprintOnly: true },
      changes,
      proof,
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings
    };
  }

  private async patchDataValidationCheck(
    metadata: WorkbookMetadata,
    patchIndex: number,
    sheetName: string,
    range: string,
    values: CellMatrix
  ): Promise<
    | { ok: true; warnings: string[]; checks: Array<Record<string, unknown>> }
    | { ok: false; summary: string; answer: Record<string, unknown>; warnings: string[] }
  > {
    const targetShape = rangeShape(range);
    if (targetShape && targetShape.columns !== 1) {
      return { ok: true, warnings: [], checks: [] };
    }
    const result = await this.runtime.readRangeMetadata("range.read_data_validation", {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      sheetName,
      address: range
    });
    const answer = dataValidationSummaryAnswer(result, metadata, sheetName, range);
    const fieldContext = Array.isArray(answer.fieldContext)
      ? answer.fieldContext.filter((field): field is Record<string, unknown> => Boolean(field && typeof field === "object"))
      : [];
    const fieldsWithValidation = fieldContext.filter((field) => fieldHasEnforceableDropdownContext(field));
    if (fieldsWithValidation.length === 0) {
      return { ok: true, warnings: [], checks: [] };
    }
    const proposedValues = uniqueProposedValidationValues(values);
    if (proposedValues.length === 0) {
      return { ok: true, warnings: [], checks: [] };
    }
    const warnings: string[] = [];
    const checks: Array<Record<string, unknown>> = [];
    for (const field of fieldsWithValidation) {
      const allowedValues = Array.isArray(field.allowedValues)
        ? field.allowedValues.filter((value): value is string => typeof value === "string")
        : [];
      const validation = field.validation && typeof field.validation === "object" ? field.validation as Record<string, unknown> : {};
      const fieldName = stringValue(field.field) ?? stringValue(field.range) ?? "validated field";
      if (allowedValues.length === 0) {
        checks.push(stripUndefinedRecord({
          patchIndex: patchIndex + 1,
          field: fieldName,
          range: field.range,
          validation,
          optionsResolved: false
        }));
        continue;
      }
      const allowed = new Map(allowedValues.map((value) => [normalizeValidationOption(value), value]));
      const invalidValues = proposedValues.filter((value) => !allowed.has(normalizeValidationOption(value)));
      const check = stripUndefinedRecord({
        patchIndex: patchIndex + 1,
        field: fieldName,
        range: field.range,
        allowedValues,
        allowedValueCount: allowedValues.length,
        proposedValues,
        invalidValues,
        validation
      });
      checks.push(check);
      if (invalidValues.length > 0) {
        const warning = `VALUE_NOT_IN_DROPDOWN_OPTIONS: Patch ${patchIndex + 1} proposes ${invalidValues.map((value) => `"${value}"`).join(", ")} for ${fieldName}, but allowed values are ${allowedValues.join(", ")}.`;
        return {
          ok: false,
          summary: `Patch ${patchIndex + 1} writes a value outside the dropdown options for ${fieldName}.`,
          answer: {
            kind: "patch_validation_failed",
            code: "VALUE_NOT_IN_DROPDOWN_OPTIONS",
            patchIndex: patchIndex + 1,
            field: fieldName,
            range: field.range,
            proposedValues,
            invalidValues,
            allowedValues,
            fieldContext: [field]
          },
          warnings: [warning]
        };
      }
      warnings.push(`Patch ${patchIndex + 1} validated ${fieldName} against ${allowedValues.length} dropdown option(s).`);
    }
    return { ok: true, warnings, checks };
  }

  private async previewMatchUpdate(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const spec = resolveMatchUpdateSpec(metadata, input);
    if (!spec) {
      return undefined;
    }
    if (!spec.ok) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: spec.summary,
        ...(spec.candidates !== undefined ? { candidates: spec.candidates } : {}),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: spec.warnings,
        taskOutcome: "needs_user_input",
        maxRecommendedFollowupCalls: 0
      };
    }

    const parsed = tryParseA1Address(stripSheetName(spec.searchedRange));
    if (!parsed) {
      return undefined;
    }

    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const matrix = await this.readRangeValues(workbookId, spec.sheetName, spec.searchedRange);
    const headerOffset = Math.max(0, spec.headerRow - parsed.startRow);
    const dataStartIndex = Math.max(headerOffset + 1, 0);
    const matchedRows: number[] = [];
    const evidenceRows: Array<Record<string, unknown>> = [];
    const changes: NonNullable<AgentRunOutput["changes"]> = [];
    const entries: Array<{ target: A1Range; values: CellMatrix; preserveFormats?: true }> = [];

    for (let rowIndex = dataStartIndex; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex] ?? [];
      const matchValue = row[spec.matchColumn.index];
      if (!cellMatchesPredicate(matchValue, spec.predicate)) {
        continue;
      }
      const sheetRow = parsed.startRow + rowIndex;
      const targetAddress = `${numberToColumn(spec.updateColumn.index + 1)}${sheetRow}`;
      matchedRows.push(sheetRow);
      entries.push({
        target: { workbookId, sheetName: spec.sheetName, address: targetAddress },
        values: [[spec.updateValue as CellMatrix[number][number]]],
        preserveFormats: true
      });
      changes.push({
        sheetName: spec.sheetName,
        cell: targetAddress,
        range: `${targetAddress}:${targetAddress}`,
        before: row[spec.updateColumn.index] ?? "",
        after: spec.updateValue
      });
      if (evidenceRows.length < 8) {
        evidenceRows.push(stripUndefinedRecord({
          row: sheetRow,
          matchValue,
          before: row[spec.updateColumn.index] ?? "",
          after: spec.updateValue,
          date: valueByHeader(spec.headers, row, /date/i),
          amount: valueByHeader(spec.headers, row, /amount|cash|actual/i)
        }));
      }
    }

    if (matchedRows.length === 0) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `No rows matched ${spec.matchColumn.name} ${spec.predicate.operator} "${spec.predicate.value}" in ${spec.sheetName}!${spec.searchedRange}.`,
        answer: {
          kind: "match_update_no_match",
          sheetName: spec.sheetName,
          searchedRange: spec.searchedRange,
          matchColumn: spec.matchColumn,
          predicate: spec.predicate,
          updateColumn: spec.updateColumn,
          updateValue: spec.updateValue,
          matchedRowCount: 0,
          matchedRows: []
        },
        proof: [{ sheetName: spec.sheetName, range: spec.searchedRange, label: "searched range" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: ["No exact match was found. Ask the user to refine the match text, sheet, or column; do not broad-read chunks."]
      };
    }

    const pending = this.createPendingOperation(metadata, {
      action: {
        kind: "batch",
        operations: [{
          kind: "range.write_values_many",
          operationId: makeId<OperationId>("op"),
          workbookId,
          destructiveLevel: "values",
          reason: input.request,
          entries
        }]
      },
      changes,
      summary: `Prepared match/update for ${matchedRows.length} row(s): set ${spec.updateColumn.name} to ${String(spec.updateValue)} where ${spec.matchColumn.name} ${spec.predicate.operator} "${spec.predicate.value}".`
    });

    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: `${pending.summary} Apply this grouped preview once; do not fetch full rows or split into single-cell writes.`,
      answer: {
        kind: "match_update_preview",
        sheetName: spec.sheetName,
        searchedRange: spec.searchedRange,
        matchColumn: spec.matchColumn,
        predicate: spec.predicate,
        updateColumn: spec.updateColumn,
        updateValue: spec.updateValue,
        matchedRowCount: matchedRows.length,
        matchedRows,
        evidenceRows,
        operationCount: 1,
        grouped: true
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", safetyFingerprintOnly: true, workflowKind: "match_update", groupedOperationCount: 1 },
      changes,
      proof: [
        { sheetName: spec.sheetName, range: spec.searchedRange, label: "searched range" },
        ...uniqueProofFromChanges(changes).slice(0, 5)
      ],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      agentInstruction: "Use matchedRows and changes from this match/update preview. Do not fetch full rows, chunk-read, or split this task into single-cell writes.",
      maxRecommendedFollowupCalls: 1,
      warnings: []
    };
  }

  private async previewOperationIntent(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const operationInput = isCanonicalDirectMutationAction(action)
      ? canonicalDirectMutationInput(input, action)
      : input;
    if (action === "transform_values") {
      return this.previewTransformValues(metadata, input, requestedMode);
    }
    if (action === "derive_values") {
      return this.previewDeriveValues(metadata, input, requestedMode);
    }
    if (action === "settle_reconciliation" || shouldPreviewSettlementBundle(input)) {
      return this.previewSettlementBundle(metadata, input, requestedMode);
    }
    if (action === "transform_sheets") {
      return this.previewTransformSheets(metadata, input, requestedMode);
    }
    if (action === "replace_range_with_styled_table" || shouldPreviewReplaceStyledTable(input)) {
      return this.previewReplaceStyledTable(metadata, input, requestedMode);
    }
    if (action === "grouped_header" || (!action && shouldPreviewGroupedHeader(input))) {
      return this.previewGroupedHeader(metadata, input, requestedMode);
    }
    if (action === "improve_visual_readability") {
      return this.previewVisualReadability(metadata, input, requestedMode);
    }
    const workbookLevelHandler = findAgentActionHandler(operationInput, action, false);
    if (workbookLevelHandler) {
      return this.previewActionHandler(metadata, operationInput, requestedMode, workbookLevelHandler);
    }
    const resolved = resolveAgentUpdateTarget(metadata, operationInput);
    if (!resolved.ok) {
      return undefined;
    }
    const normalizedRange = normalizeOperationRange(metadata, resolved.sheetName, resolved.range);
    const normalizedResolved = { ...resolved, range: normalizedRange };
    const targetLevelHandler = findAgentActionHandler(operationInput, action, true);
    if (targetLevelHandler) {
      return this.previewActionHandler(metadata, operationInput, requestedMode, targetLevelHandler, normalizedResolved);
    }
    return undefined;
  }

  private previewVisualReadability(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const options = visualReadabilityOptionsFromInput(input);
    const requestedSheetName = input.target?.sheetName;
    const fallbackSheetName = metadata.workbook.activeSheet ?? metadata.sheets[0]?.name;
    const sheetName = requestedSheetName && !sameText(requestedSheetName, "active") && !sameText(requestedSheetName, "active_sheet")
      ? requestedSheetName
      : fallbackSheetName;
    const sheet = metadata.sheets.find((candidate) => sameText(candidate.name, sheetName)) ?? metadata.sheets.find((candidate) => sameText(candidate.name, fallbackSheetName));
    if (!sheet) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Visual readability preview needs a target sheet.",
        candidates: findAgentCandidates(metadata, input).slice(0, 5),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: ["Provide target.sheetName or prepare a workbook context with an active sheet."]
      };
    }
    const range = input.target?.range ?? sheet.usedRange;
    const detected = detectVisualReadabilityStructure(metadata, sheet, range);
    const detectionFailure = visualReadabilityDetectionFailure(metadata, requestedMode, sheet.name, range, detected);
    if (detectionFailure) {
      return detectionFailure;
    }
    const columnRoles = inferVisualReadabilityColumns(metadata, sheet, detected);
    const sheetType = inferVisualReadabilitySheetType(sheet, columnRoles);
    const resolvedProfile = options.profile === "auto" ? profileForVisualSheetType(sheetType) : options.profile;
    const visualPlan = compileVisualReadabilityPlan(metadata, sheet, detected, columnRoles, { ...options, profile: resolvedProfile }, sheetType);
    const groupedHeaderSuggestion = groupedHeaderSuggestionFromColumns(detected, suggestedColumnGroups(columnRoles, detected));
    const compiledOperations = compileVisualReadabilityOperations(metadata.workbook.workbookId as WorkbookId, sheet.name, visualPlan.rules, visualPlan.validationSuggestions, {
      preserveExistingStyle: options.preserveExistingStyle,
      stylePreservationMode: options.stylePreservationMode,
      allowReplaceConditionalFormatting: options.allowReplaceConditionalFormatting,
      allowReplaceDataValidation: options.allowReplaceDataValidation,
      applySuggestionBuckets: options.applySuggestionBuckets,
      preservation: detected
    });
    const formulaRanges = visualReadabilityFormulaCheckRanges(detected);
    const pending = this.createPendingOperation(metadata, {
      action: {
        kind: "visual_readability.apply",
        operations: compiledOperations.operations,
        request: {
          workbookId: metadata.workbook.workbookId as WorkbookId,
          sheetName: sheet.name,
          formulaRanges,
          ruleCount: visualPlan.counts.totalRules,
          skippedRuleCount: compiledOperations.skipped.length
        }
      },
      changes: visualPlan.previewExamples.length > 0
        ? visualPlan.previewExamples.map((example) => ({ sheetName: sheet.name, range: example.range, before: example.before, after: example.after }))
        : [{ sheetName: sheet.name, ...(range ? { range } : {}), after: "visual readability compiled preview; no workbook styles will be changed yet" }],
      summary: `Prepared visual readability preview for ${sheet.name}.`,
      workflowKind: "visual_readability_preview"
    });
    const hasApplyReadyOperations = compiledOperations.operations.length > 0;
    const zeroOperationWarnings = !hasApplyReadyOperations
      ? visualReadabilityZeroOperationWarnings(compiledOperations.skipped)
      : [];
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: "visual_readability_preview",
        action: "improve_visual_readability",
        sheetName: sheet.name,
        ...(range ? { range } : {}),
        defaults: stripUndefinedRecord({ ...options, profile: resolvedProfile }),
        detected,
        columnRoles,
        sheetType,
        groupedHeaderSuggestion,
        visualPlan: {
          ...visualPlan,
          ruleIds: visualPlan.rules.map((rule) => rule.id),
          operationId: pending.operationId,
          risk: pending.risk,
          operationCount: compiledOperations.operations.length,
          skipped: compiledOperations.skipped,
          preservation: {
            formulas: options.preserveFormulas ? "preserved" : "not_checked",
            formulaRanges,
            existingStyle: options.preserveExistingStyle ? options.stylePreservationMode : "not_checked"
          }
        }
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", workflowKind: "visual_readability_preview", groupedOperationCount: visualPlan.counts.totalRules, operationCount: compiledOperations.operations.length, skippedRuleCount: compiledOperations.skipped.length },
      changes: pending.changes,
      proof: range ? [{ sheetName: sheet.name, range, label: "visual readability target" }] : [],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: hasApplyReadyOperations ? "call_apply_update" : "answer_now",
      ...(!hasApplyReadyOperations ? { agentInstruction: "Do not call apply_update for this visual readability preview because it compiled zero workbook operations. Explain the skipped reasons and ask for a narrower target or supported workflow." } : {}),
      warnings: compiledOperations.skipped.length > 0
        ? [
            `Compiled ${compiledOperations.operations.length} safe visual operation(s). ${compiledOperations.skipped.length} rule(s) are preview-only because the current operation schema does not support them, an opt-in bucket is required, or preservation settings skipped them.`,
            ...zeroOperationWarnings
          ]
        : [`Compiled ${compiledOperations.operations.length} safe visual operation(s) for apply.`]
    };
  }

  private previewGroupedHeader(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const requestedSheetName = input.target?.sheetName
      ?? stringValue((input.values as Record<string, unknown> | undefined)?.sheetName)
      ?? metadata.workbook.activeSheet
      ?? metadata.sheets[0]?.name;
    const sheet = requestedSheetName ? metadata.sheets.find((candidate) => sameText(candidate.name, requestedSheetName)) : undefined;
    if (!sheet) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Grouped header preview needs target.sheetName.");
    }
    const range = input.target?.range ?? sheet.usedRange;
    const detected = detectVisualReadabilityStructure(metadata, sheet, range);
    const detectionFailure = visualReadabilityDetectionFailure(metadata, requestedMode, sheet.name, range, detected);
    if (detectionFailure) {
      return detectionFailure;
    }
    const columns = inferVisualReadabilityColumns(metadata, sheet, detected);
    const groups = groupedHeaderGroupsFromInput(input, columns, detected);
    if (groups.length < 2) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Grouped header preview needs at least two column groups or a wide table/header range that can be grouped.");
    }
    const headerRow = detected.headerRow ?? 1;
    const groupRow = headerRow;
    const shiftedHeaderRow = headerRow + 1;
    const firstGroupColumn = groups[0]!.startColumn;
    const lastGroupColumn = groups[groups.length - 1]!.endColumn;
    const groupRowRange = `${firstGroupColumn}${groupRow}:${lastGroupColumn}${groupRow}`;
    const shiftedHeaderRowRange = `${firstGroupColumn}${shiftedHeaderRow}:${lastGroupColumn}${shiftedHeaderRow}`;
    const styleEntries: Extract<ExcelOperation, { kind: "range.write_styles_many" }>["entries"] = [
      { target: { workbookId, sheetName: sheet.name, address: groupRowRange }, style: { rowHeight: groupedHeaderRowHeight(input, 34) }, preserveValues: true },
      { target: { workbookId, sheetName: sheet.name, address: shiftedHeaderRowRange }, style: { rowHeight: groupedHeaderRowHeight(input, 26, "headerRowHeight") }, preserveValues: true },
      ...groups.flatMap((group): Extract<ExcelOperation, { kind: "range.write_styles_many" }>["entries"] => [
        {
          target: { workbookId, sheetName: sheet.name, address: `${group.startColumn}${groupRow}:${group.endColumn}${groupRow}` },
          style: {
            fillColor: group.fillColor,
            fontColor: "#FFFFFF",
            fontBold: true,
            horizontalAlignment: "center",
            verticalAlignment: "center"
          },
          preserveValues: true
        },
        {
          target: { workbookId, sheetName: sheet.name, address: `${group.startColumn}${shiftedHeaderRow}:${group.endColumn}${shiftedHeaderRow}` },
          style: {
            fillColor: group.headerFillColor,
            fontColor: "#1F2937",
            fontBold: true,
            horizontalAlignment: "center",
            verticalAlignment: "center",
            wrapText: true,
            borders: { edgeBottom: { style: "continuous", weight: "thin", color: group.fillColor } }
          },
          preserveValues: true
        }
      ])
    ];
    const operations: ExcelOperation[] = [
      {
        kind: "range.insert_rows",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "structure",
        reason: input.request,
        target: { workbookId, sheetName: sheet.name, address: groupRowRange }
      },
      {
        kind: "range.write_values_many",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "values",
        reason: input.request,
        entries: groups.map((group) => ({
          target: { workbookId, sheetName: sheet.name, address: `${group.startColumn}${groupRow}:${group.startColumn}${groupRow}` },
          values: [[group.label]],
          preserveFormats: true
        }))
      },
      ...groups
        .filter((group) => columnToNumber(group.endColumn) > columnToNumber(group.startColumn))
        .map((group): ExcelOperation => ({
          kind: "range.merge",
          operationId: makeId<OperationId>("op"),
          workbookId,
          destructiveLevel: "structure",
          reason: input.request,
          target: { workbookId, sheetName: sheet.name, address: `${group.startColumn}${groupRow}:${group.endColumn}${groupRow}` },
          across: false
        })),
      {
        kind: "range.write_styles_many",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "format",
        reason: input.request,
        entries: styleEntries
      }
    ];
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      operations,
      [
        { sheetName: sheet.name, range: groupRowRange, after: "insert grouped visual header row above existing table headers" },
        ...groups.map((group) => ({ sheetName: sheet.name, range: `${group.startColumn}${groupRow}:${group.endColumn}${groupRow}`, after: `group header ${group.label}` })),
        { sheetName: sheet.name, range: shiftedHeaderRowRange, after: "existing header row restyled with lighter group fills" }
      ],
      `Prepared grouped header preview for ${sheet.name}; existing headers shift from row ${headerRow} to row ${shiftedHeaderRow}.`,
      {
        kind: "grouped_header_preview",
        sheetName: sheet.name,
        headerRow,
        groupRow,
        shiftedHeaderRow,
        groups,
        operationCount: operations.length,
        preservesExistingHeaderLabels: true
      }
    );
  }

  private async previewTransformValues(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry">> {
    const values = input.values ?? {};
    const operation = transformOperationFromInput(input);
    if (!operation.ok) {
      return transformNeedsInput(metadata, requestedMode, operation.summary, operation.warnings);
    }
    const target = resolveValueColumnScope(metadata, input, "target");
    if (!target.ok) {
      return transformNeedsInput(metadata, requestedMode, target.summary, target.warnings, target.candidates);
    }
    const snapshot = await this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, target.scope.sheetName, target.scope.address);
    const plan = compileTransformPlan(target.scope, snapshot, operation.operation, values);
    if (!plan.ok) {
      return transformNeedsInput(metadata, requestedMode, plan.summary, plan.warnings);
    }
    return this.previewCompiledColumnPlan(metadata, input, requestedMode, {
      kind: "transform_values_preview",
      summary: `Prepared ${operation.operation} transform for ${target.scope.sheetName}!${target.scope.address}.`,
      target: target.scope,
      sources: [],
      rowAlignment: { type: "single_column", rows: target.scope.rowCount },
      plan
    });
  }

  private async previewDeriveValues(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry">> {
    const derivation = derivationOperationFromInput(input);
    if (!derivation.ok) {
      return transformNeedsInput(metadata, requestedMode, derivation.summary, derivation.warnings);
    }
    const target = resolveValueColumnScope(metadata, input, "target");
    if (!target.ok) {
      return transformNeedsInput(metadata, requestedMode, target.summary, target.warnings, target.candidates);
    }
    const formulaMode = derivation.operation === "formula_like" && shouldDeriveAsFormula(input);
    const deriveInput = formulaMode ? withFormulaDeriveSourceDefaults(input, target.scope) : input;
    const values = deriveInput.values ?? {};
    const sourceScopes = resolveDeriveSourceScopes(metadata, deriveInput);
    if (!sourceScopes.ok) {
      return transformNeedsInput(metadata, requestedMode, sourceScopes.summary, sourceScopes.warnings, sourceScopes.candidates);
    }
    if (derivation.operation === "lookup_map") {
      const lookup = resolveLookupMapScope(metadata, deriveInput);
      if (!lookup.ok) {
        return transformNeedsInput(metadata, requestedMode, lookup.summary, lookup.warnings, lookup.candidates);
      }
      const targetSnapshot = await this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, target.scope.sheetName, target.scope.address);
      const sourceSnapshots = await Promise.all(sourceScopes.scopes.map((source) =>
        this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, source.sheetName, source.address)
      ));
      const lookupKeySnapshot = await this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, lookup.keyScope.sheetName, lookup.keyScope.address);
      const lookupValueSnapshot = await this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, lookup.valueScope.sheetName, lookup.valueScope.address);
      const plan = compileLookupDerivePlan(target.scope, targetSnapshot, sourceScopes.scopes, sourceSnapshots, lookup, lookupKeySnapshot, lookupValueSnapshot);
      if (!plan.ok) {
        return transformNeedsInput(metadata, requestedMode, plan.summary, plan.warnings);
      }
      return this.previewCompiledColumnPlan(metadata, deriveInput, requestedMode, {
        kind: "derive_values_preview",
        summary: `Prepared lookup_map derivation for ${target.scope.sheetName}!${target.scope.address}.`,
        target: target.scope,
        sources: [...sourceScopes.scopes, lookup.keyScope, lookup.valueScope],
        rowAlignment: { type: "lookup_map", rows: target.scope.rowCount, lookupRows: lookup.keyScope.rowCount },
        plan
      });
    }
    const misaligned = sourceScopes.scopes.find((source) => source.rowCount !== target.scope.rowCount);
    if (misaligned) {
      return transformNeedsInput(metadata, requestedMode, `Source ${misaligned.headerName ?? misaligned.address} does not align with target ${target.scope.headerName ?? target.scope.address}.`, [
        `Target rows: ${target.scope.rowCount}; source rows: ${misaligned.rowCount}.`
      ]);
    }
    const targetSnapshot = await this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, target.scope.sheetName, target.scope.address);
    const sourceSnapshots = await Promise.all(sourceScopes.scopes.map((source) =>
      this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, source.sheetName, source.address)
    ));
    if (formulaMode) {
      const plan = compileFormulaDeriveMatrix(target.scope, sourceScopes.scopes, sourceSnapshots, values);
      if (!plan.ok) {
        return transformNeedsInput(metadata, requestedMode, plan.summary, plan.warnings);
      }
      const formulaInput: AgentRunInput = {
        ...deriveInput,
        target: { ...(deriveInput.target ?? {}), sheetName: target.scope.sheetName, range: target.scope.address },
        values: { ...(deriveInput.values ?? {}), formulas: plan.formulas }
      };
      const resolved = {
        ok: true,
        sheetName: target.scope.sheetName,
        range: target.scope.address,
        candidate: {
          id: `formula-derive:${target.scope.sheetName}:${target.scope.address}`,
          kind: "range" as const,
          label: target.scope.headerName ?? target.scope.address,
          sheetName: target.scope.sheetName,
          range: target.scope.address,
          semanticRole: "formula_region" as const,
          confidence: 1,
          reason: "Formula derivation target"
        }
      } satisfies Extract<AgentTargetResolution, { ok: true }>;
      const preview = this.previewFormulaUpdate(metadata, formulaInput, requestedMode, resolved, plan.formulas);
      if (isRecord(preview.answer)) {
        preview.answer = {
          ...preview.answer,
          target: compactValueColumnScope(target.scope),
          sources: sourceScopes.scopes.map(compactValueColumnScope),
          rowAlignment: { type: "same_row", rows: target.scope.rowCount },
          scannedRows: target.scope.rowCount,
          changedCells: plan.changedCount,
          examples: plan.examples
        };
      }
      preview.summary = `Prepared formula derivation for ${target.scope.sheetName}!${target.scope.address}: ${plan.changedCount} formula cell(s).`;
      return preview;
    }
    const plan = compileDerivePlan(target.scope, targetSnapshot, sourceScopes.scopes, sourceSnapshots, derivation.operation, values);
    if (!plan.ok) {
      return transformNeedsInput(metadata, requestedMode, plan.summary, plan.warnings);
    }
    return this.previewCompiledColumnPlan(metadata, input, requestedMode, {
      kind: "derive_values_preview",
      summary: `Prepared ${derivation.operation} derivation for ${target.scope.sheetName}!${target.scope.address}.`,
      target: target.scope,
      sources: sourceScopes.scopes,
      rowAlignment: { type: "same_row", rows: target.scope.rowCount },
      plan
    });
  }

  private async previewSettlementBundle(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry">> {
    const bundle = await compileSettlementBundle(metadata, input, async (scope) =>
      this.readColumnSnapshot(metadata.workbook.workbookId as WorkbookId, scope.sheetName, scope.address)
    );
    if (!bundle.ok) {
      return transformNeedsInput(metadata, requestedMode, bundle.summary, bundle.warnings, bundle.candidates);
    }
    if (bundle.operations.length === 0) {
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Settlement convention preview found no changes to apply.",
        answer: {
          kind: "settlement_bundle_preview",
          target: compactValueColumnScope(bundle.scopes.paymentVariance),
          sources: [bundle.scopes.cashAmount, bundle.scopes.actualAmount].map(compactValueColumnScope),
          noteTargets: [bundle.scopes.reconciliationNote, bundle.scopes.detailNotes].map(compactValueColumnScope),
          reference: bundle.reference,
          scannedCount: bundle.scannedCount,
          changedCount: 0,
          skipped: bundle.skipped,
          examples: []
        },
        metrics: { operationRisk: "read_only", targetFingerprintStatus: "matched", workflowKind: "settlement_bundle_preview" },
        proof: bundle.proof,
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: bundle.warnings
      };
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations: bundle.operations },
      changes: bundle.changes,
      summary: `Prepared settlement update for ${bundle.sheetName}: ${bundle.changedCount} changed cell(s) across ${bundle.operations.length} grouped operation(s).`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: `${pending.summary} Apply once with apply_update; do not split variance and note updates into separate calls.`,
      answer: {
        kind: "settlement_bundle_preview",
        sheetName: bundle.sheetName,
        tableName: bundle.tableName,
        target: compactValueColumnScope(bundle.scopes.paymentVariance),
        sources: [bundle.scopes.cashAmount, bundle.scopes.actualAmount].map(compactValueColumnScope),
        noteTargets: [bundle.scopes.reconciliationNote, bundle.scopes.detailNotes].map(compactValueColumnScope),
        reference: bundle.reference,
        scannedCount: bundle.scannedCount,
        changedCount: bundle.changedCount,
        skipped: bundle.skipped,
        examples: bundle.examples,
        operationCount: bundle.operations.length,
        grouped: true
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", workflowKind: "settlement_bundle_preview", groupedOperationCount: bundle.operations.length },
      changes: bundle.changes,
      proof: bundle.proof,
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: bundle.warnings
    };
  }

  private previewTransformSheets(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const plan = compileSheetTransformPlan(metadata, input);
    if (!plan.ok) {
      return transformNeedsInput(metadata, requestedMode, plan.summary, plan.warnings, plan.candidates);
    }
    const operations: ExcelOperation[] = plan.renames.map((rename) => ({
      kind: "sheet.rename",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "structure",
      reason: input.request,
      sheetName: rename.from,
      newSheetName: rename.to
    }));
    const changes: NonNullable<AgentRunOutput["changes"]> = plan.renames.slice(0, 12).map((rename) => ({
      sheetName: rename.from,
      before: rename.from,
      after: rename.to
    }));
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
      changes,
      summary: `Prepared ${plan.renames.length} sheet rename(s) as one workbook structure plan.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: "transform_sheets_preview",
        operation: plan.operation,
        changedCount: plan.renames.length,
        skippedCount: plan.skipped.length,
        examples: plan.renames.slice(0, 10),
        skipped: plan.skipped.slice(0, 10)
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", workflowKind: "transform_sheets_preview", groupedOperationCount: operations.length },
      changes,
      proof: changes.map((change) => ({ sheetName: change.sheetName, range: usedRangeForSheet(metadata, change.sheetName), label: "sheet rename target" })),
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: plan.warnings
    };
  }

  private previewCompiledColumnPlan(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    compiled: {
      kind: "transform_values_preview" | "derive_values_preview";
      summary: string;
      target: ValueColumnScope;
      sources: ValueColumnScope[];
      rowAlignment: Record<string, unknown>;
      plan: CompiledColumnPlan;
    }
  ): Omit<AgentRunOutput, "telemetry"> {
    const entries = changedColumnRuns(compiled.target, compiled.plan.afterValues, compiled.plan.changedRows);
    if (entries.length === 0) {
      return {
        status: "SUCCESS",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `${compiled.summary} No cell changes are needed.`,
        answer: {
          kind: compiled.kind,
          target: compactValueColumnScope(compiled.target),
          sources: compiled.sources.map(compactValueColumnScope),
          rowAlignment: compiled.rowAlignment,
          scannedCount: compiled.plan.scannedCount,
          changedCount: 0,
          skipped: compiled.plan.skipped,
          examples: []
        },
        metrics: { operationRisk: "read_only", targetFingerprintStatus: "matched", workflowKind: compiled.kind },
        proof: [{ sheetName: compiled.target.sheetName, range: compiled.target.address, label: "target column" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "answer_now",
        warnings: compiled.plan.warnings
      };
    }

    const operation: ExcelOperation = {
      kind: "range.write_values_many",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      entries: entries.map((entry) => ({
        target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: compiled.target.sheetName, address: entry.address },
        values: entry.values,
        preserveFormats: true
      }))
    };
    const changes: NonNullable<AgentRunOutput["changes"]> = [
      { sheetName: compiled.target.sheetName, range: compiled.target.address, before: `${compiled.plan.scannedCount} scanned`, after: `${compiled.plan.changedCount} changed` },
      ...compiled.plan.examples.map((example) => ({
        sheetName: compiled.target.sheetName,
        range: `${compiled.target.columnLetter}${example.row}`,
        before: example.before,
        after: example.after
      }))
    ];
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations: [operation] },
      changes,
      summary: `${compiled.summary} Previewed ${compiled.plan.changedCount} changed cell(s) across ${entries.length} run(s).`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: compiled.kind,
        target: compactValueColumnScope(compiled.target),
        sources: compiled.sources.map(compactValueColumnScope),
        rowAlignment: compiled.rowAlignment,
        scannedCount: compiled.plan.scannedCount,
        changedCount: compiled.plan.changedCount,
        skipped: compiled.plan.skipped,
        unmatchedCount: compiled.plan.unmatchedCount,
        examples: compiled.plan.examples,
        runCount: entries.length
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", workflowKind: compiled.kind, groupedOperationCount: entries.length },
      changes,
      proof: [
        { sheetName: compiled.target.sheetName, range: compiled.target.address, label: "target column" },
        ...compiled.sources.slice(0, 3).map((source) => ({ sheetName: source.sheetName, range: source.address, label: "source column" }))
      ],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: compiled.plan.warnings
    };
  }

  private async previewActionHandler(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    handler: AgentActionHandlerDefinition,
    resolved?: Extract<AgentTargetResolution, { ok: true }>
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const result = await this.previewActionHandlerOutput(metadata, input, requestedMode, handler, resolved);
    return result ? withActionHandlerMetric(result, handler.id) : result;
  }

  private previewActionHandlerOutput(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    handler: AgentActionHandlerDefinition,
    resolved?: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> | Promise<Omit<AgentRunOutput, "telemetry"> | undefined> | undefined {
    switch (handler.id) {
      case "save_workbook":
        return this.previewWorkbookOperation(metadata, input, requestedMode, "workbook.save");
      case "calculate_workbook":
        return this.previewWorkbookOperation(metadata, input, requestedMode, "workbook.calculate");
      case "recalculate_formulas":
        return this.previewWorkbookOperation(metadata, input, requestedMode, "workbook.calculate", { kind: "formula.recalculate_preview" });
      case "create_snapshot":
        return this.previewWorkbookLocalStateOperation(metadata, input, requestedMode, "workbook.snapshot");
      case "create_backup":
        return this.previewWorkbookLocalStateOperation(metadata, input, requestedMode, "workbook.create_backup");
      case "refresh_snapshot":
        return this.previewSafetyArtifactMutation(metadata, input, requestedMode, "snapshot.refresh");
      case "invalidate_snapshot":
        return this.previewSafetyArtifactMutation(metadata, input, requestedMode, "snapshot.invalidate");
      case "delete_snapshot":
        return this.previewSafetyArtifactMutation(metadata, input, requestedMode, "snapshot.delete");
      case "create_file_backup":
        return this.previewFileBackupCreate(metadata, input, requestedMode);
      case "restore_file_backup":
        return this.previewFileBackupRestore(metadata, input, requestedMode);
      case "prune_backups":
        return this.previewBackupPrune(metadata, input, requestedMode);
      case "pin_backup":
        return this.previewSafetyArtifactMutation(metadata, input, requestedMode, "backup.pin");
      case "unpin_backup":
        return this.previewSafetyArtifactMutation(metadata, input, requestedMode, "backup.unpin");
      case "delete_backup":
        return this.previewSafetyArtifactMutation(metadata, input, requestedMode, "backup.delete");
      case "restore_workbook_backup":
        return this.previewWorkbookRestoreBackup(metadata, input, requestedMode);
      case "import_local_config":
        return this.previewWorkbookImportLocalConfig(metadata, input, requestedMode);
      case "embed_local_config":
        return this.previewWorkbookEmbedLocalConfig(metadata, input, requestedMode);
      case "import_embedded_local_config":
        return this.previewWorkbookImportEmbeddedLocalConfig(metadata, input, requestedMode);
      case "close_workbook":
        return this.previewWorkbookClose(metadata, input, requestedMode);
      case "copy_template_sheet":
        return this.previewTemplateCleanup(metadata, input, requestedMode);
      case "register_template":
        return this.previewTemplateRegister(metadata, input, requestedMode);
      case "unregister_template":
        return this.previewTemplateUnregister(metadata, input, requestedMode);
      case "clear_template_data_regions":
        return this.previewTemplateClearDataRegions(metadata, input, requestedMode);
      case "fill_template_regions":
        return this.previewTemplateFillRegions(metadata, input, requestedMode);
      case "repair_sheet_from_template":
        return this.previewTemplateRepair(metadata, input, requestedMode);
      case "copy_style_from_template":
        return this.previewStyleCopy(metadata, input, requestedMode);
      case "repair_style_consistency":
        return this.previewStyleRepairConsistency(metadata, input, requestedMode);
      case "repair_style_from_template":
        return this.previewStyleRepairConsistency(metadata, input, requestedMode);
      case "repair_formulas_from_template":
        return this.previewFormulaRepairPatterns(metadata, input, requestedMode);
      case "create_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "create");
      case "copy_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "copy");
      case "rename_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "rename");
      case "delete_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "delete");
      case "hide_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "hide");
      case "unhide_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "unhide");
      case "protect_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "protect");
      case "unprotect_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "unprotect");
      case "clear_sheet":
        return this.previewSheetOperation(metadata, input, requestedMode, "clear");
      case "set_sheet_tab_color":
        return this.previewSheetOperation(metadata, input, requestedMode, "set_tab_color");
      case "copy_formula_patterns":
        return this.previewFormulaCopyPatterns(metadata, input, requestedMode);
      case "fill_formula_down":
        return this.previewFormulaFill(metadata, input, requestedMode, "down");
      case "fill_formula_right":
        return this.previewFormulaFill(metadata, input, requestedMode, "right");
      case "repair_formula_patterns":
        return this.previewFormulaRepairPatterns(metadata, input, requestedMode);
      case "create_name":
        return this.previewNameMutation(metadata, input, requestedMode, "create");
      case "update_name":
        return this.previewNameMutation(metadata, input, requestedMode, "update");
      case "delete_name":
        return this.previewNameMutation(metadata, input, requestedMode, "delete");
      case "register_region":
        return this.previewRegionMutation(metadata, input, requestedMode, "register");
      case "clear_region_values":
        return this.previewRegionMutation(metadata, input, requestedMode, "clear_values");
      case "write_region_values":
        return this.previewRegionMutation(metadata, input, requestedMode, "write_values");
      case "fill_region":
        return this.previewRegionMutation(metadata, input, requestedMode, "fill");
      case "first_open_reviewed":
        return this.previewFirstStatusMatchUpdate(metadata, input, requestedMode);
      case "apply_table_view":
        return resolved?.candidate.kind === "table" ? this.previewTableApplyView(metadata, input, requestedMode, resolved) : undefined;
      case "sort_table":
        return this.previewTableSort(metadata, input, requestedMode);
      case "append_table_rows":
        return resolved ? this.previewTableAppend(metadata, input, requestedMode, resolved, objectToCellMatrix(input.values ?? {})) : undefined;
      case "update_table_rows":
        return resolved ? this.previewTableUpdateRows(metadata, input, requestedMode, resolved) : undefined;
      case "create_table":
        return this.previewTableCreate(metadata, input, requestedMode);
      case "resize_table":
        return resolved ? this.previewTableResize(metadata, input, requestedMode, resolved) : undefined;
      case "reorder_table_columns":
        return resolved ? this.previewTableReorderColumns(metadata, input, requestedMode, resolved) : undefined;
      case "clear_table_data":
        return resolved ? this.previewTableSelectorMutation(metadata, input, requestedMode, resolved, "clear_data_keep_formulas") : undefined;
      case "clear_table_filters":
        return resolved
          ? resolved.candidate.kind === "table"
            ? this.previewTableSelectorMutation(metadata, input, requestedMode, resolved, "clear_filters")
            : this.previewClearAutoFilter(metadata, input, requestedMode, resolved)
          : undefined;
      case "filter_range":
        return resolved?.candidate.kind === "table"
          ? this.previewTableApplyFilters(metadata, input, requestedMode, resolved)
          : resolved ? this.previewAutoFilterMutation(metadata, input, requestedMode, resolved) : undefined;
      case "set_table_total_row":
        return resolved ? this.previewTableTotalRow(metadata, input, requestedMode, resolved) : undefined;
      case "set_table_style":
        return resolved ? this.previewTableStyle(metadata, input, requestedMode, resolved) : undefined;
      case "copy_table_structure":
        return resolved ? this.previewTableCopyStructure(metadata, input, requestedMode, resolved) : undefined;
      case "repair_table_structure":
        return resolved ? this.previewTableCopyStructure(metadata, input, requestedMode, resolved) : undefined;
      case "autofit_columns":
        return resolved ? this.previewAutofit(metadata, input, requestedMode, resolved, "columns") : undefined;
      case "autofit_rows":
        return resolved ? this.previewAutofit(metadata, input, requestedMode, resolved, "rows") : undefined;
      case "freeze_panes":
        return this.previewFreezePanes(metadata, input, requestedMode, resolved);
      case "clear_range":
        return resolved ? this.previewClearRange(metadata, input, requestedMode, resolved) : undefined;
      case "normalize_headers":
      case "trim_whitespace":
      case "remove_duplicates":
      case "parse_dates":
      case "parse_numbers":
      case "standardize_currency":
      case "fill_missing_values":
      case "split_column":
      case "merge_columns":
        return resolved ? this.previewCleanMutation(metadata, input, requestedMode, resolved, handler.id) : undefined;
      case "clear_values":
        return resolved ? this.previewClearValues(metadata, input, requestedMode, resolved) : undefined;
      case "clear_values_raw":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.clear_values") : undefined;
      case "clear_formats":
        return resolved ? this.previewClearFormats(metadata, input, requestedMode, resolved) : undefined;
      case "copy_range":
        return this.previewRangeCopyMove(metadata, input, requestedMode, "copy");
      case "move_range":
        return this.previewRangeCopyMove(metadata, input, requestedMode, "move");
      case "reorder_range_columns":
        return resolved ? this.previewRangeColumnReorder(metadata, input, requestedMode, resolved) : undefined;
      case "write_styles_many":
        return this.previewWriteStylesMany(metadata, input, requestedMode);
      case "write_data_validation":
        return resolved ? this.previewWriteDataValidation(metadata, canonicalDirectMutationInput(input, "write_data_validation"), requestedMode, resolved) : undefined;
      case "write_conditional_formatting":
        if (resolved && hasFormulaWriteInput(input) && !isConditionalFormattingMutationRequest(input.request)) {
          const canonicalInput = canonicalDirectMutationInput(input, "write_conditional_formatting");
          return this.previewFormulaUpdate(metadata, canonicalInput, requestedMode, resolved, objectToCellMatrix(canonicalInput.values ?? {}));
        }
        return resolved ? this.previewWriteConditionalFormatting(metadata, canonicalDirectMutationInput(input, "write_conditional_formatting"), requestedMode, resolved) : undefined;
      case "insert_rows":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, structuralRangeOperationKind(input, resolved, "range.insert_rows")) : undefined;
      case "delete_rows":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, structuralRangeOperationKind(input, resolved, "range.delete_rows")) : undefined;
      case "insert_columns":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, structuralRangeOperationKind(input, resolved, "range.insert_columns")) : undefined;
      case "delete_columns":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, structuralRangeOperationKind(input, resolved, "range.delete_columns")) : undefined;
      case "hide_columns":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.hide_columns") : undefined;
      case "unhide_columns":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.unhide_columns") : undefined;
      case "merge_range":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.merge") : undefined;
      case "unmerge_range":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.unmerge") : undefined;
      case "write_formulas":
        if (!resolved) return undefined;
        {
          const canonicalInput = canonicalDirectMutationInput(input, "write_formulas");
          return this.previewFormulaUpdate(metadata, canonicalInput, requestedMode, resolved, objectToCellMatrix(canonicalInput.values ?? {}));
        }
      case "convert_formulas_to_values":
        return resolved ? this.previewFormulaConvertToValues(metadata, input, requestedMode, resolved) : undefined;
      case "write_number_formats":
        return resolved ? this.previewNumberFormatUpdate(metadata, canonicalDirectMutationInput(input, "write_number_formats"), requestedMode, resolved) : undefined;
      case "clear_style_dimensions":
        return resolved ? this.previewClearStyleDimensions(metadata, input, requestedMode, resolved) : undefined;
      case "format_range":
        return resolved ? this.previewStyleUpdate(metadata, canonicalDirectMutationInput(input, "format_range"), requestedMode, resolved) : undefined;
    }
  }

  private previewWorkbookOperation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    kind: "workbook.calculate" | "workbook.save",
    answer: Record<string, unknown> = { kind: `${kind}_preview` }
  ): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind,
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: kind === "workbook.save" ? "workbook" : "none",
      reason: input.request,
      ...(kind === "workbook.calculate" ? { calculationType: "full" as const } : {})
    };
    const label = kind === "workbook.save" ? "save workbook" : "recalculate workbook";
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: label }], `Prepared ${label}.`, answer);
  }

  private previewWorkbookLocalStateOperation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    kind: "workbook.snapshot" | "workbook.create_backup"
  ): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const values = input.values as Record<string, unknown> | undefined;
    const ranges = backupRangesFromInput(workbookId, input);
    const reason = stringValue(values?.reason) ?? input.request;
    const request = {
      workbookId,
      reason,
      ...(ranges.length > 0 ? { ranges } : {})
    };
    const label = kind === "workbook.snapshot" ? "workbook snapshot" : "workbook backup";
    const changes = ranges.length > 0
      ? ranges.map((range) => ({ sheetName: range.sheetName, range: range.address, after: label }))
      : metadata.sheets.slice(0, 5).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, after: label }] : []);
    const pending = this.createPendingOperation(metadata, {
      action: { kind, request },
      changes,
      summary: `Prepared ${label} creation.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: `${kind}_preview`, workbookId, reason, rangeCount: ranges.length },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: changes.flatMap((change) => change.range ? [{ sheetName: change.sheetName, range: change.range, label }] : []).slice(0, 5),
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewSafetyArtifactMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    kind: "snapshot.refresh" | "snapshot.invalidate" | "snapshot.delete" | "backup.pin" | "backup.unpin" | "backup.delete"
  ): Omit<AgentRunOutput, "telemetry"> {
    const isSnapshot = kind.startsWith("snapshot.");
    const id = isSnapshot ? snapshotIdFromInput(input) : backupIdFromInput(input);
    if (!id) {
      return safetyArtifactNeedsInput(metadata, requestedMode, `${isSnapshot ? "Snapshot" : "Backup"} ${kind.split(".")[1]} needs target.entity, values.${isSnapshot ? "snapshotId" : "backupId"}, or operationId.`);
    }
    const action = safetyArtifactAction(kind, id);
    const label = kind.replace(".", " ");
    const pending = this.createPendingOperation(metadata, {
      action,
      changes: [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: `${label}: ${id}` }],
      summary: `Prepared ${label} for ${id}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: `${kind}_preview`, id },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "not_applicable" },
      changes: pending.changes,
      proof: [],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewWorkbookRestoreBackup(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const backupId = backupIdFromInput(input);
    if (!backupId) {
      return safetyArtifactNeedsInput(metadata, requestedMode, "Workbook backup restore needs values.backupId, target.entity, or operationId.");
    }
    return this.previewWorkbookPendingAction(metadata, requestedMode, {
      action: { kind: "workbook.restore_backup", backupId: backupId as BackupId },
      summary: `Prepared workbook backup restore for ${backupId}.`,
      answer: { kind: "workbook_restore_backup_preview", backupId },
      after: `restore workbook backup ${backupId}`
    });
  }

  private previewWorkbookImportLocalConfig(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = workbookLocalConfigImportRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Local config import needs values.config.");
    }
    return this.previewWorkbookPendingAction(metadata, requestedMode, {
      action: { kind: "workbook.import_local_config", request },
      summary: "Prepared workbook local config import.",
      answer: { kind: "workbook_import_local_config_preview", workbookId: request.workbookId, overwrite: request.overwrite ?? false },
      after: "imported local config"
    });
  }

  private previewWorkbookEmbedLocalConfig(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const options = workbookLocalConfigOptionsFromInput(input);
    return this.previewWorkbookPendingAction(metadata, requestedMode, {
      action: { kind: "workbook.embed_local_config", workbookId, ...options },
      summary: "Prepared embedding workbook local config.",
      answer: { kind: "workbook_embed_local_config_preview", workbookId, ...options },
      after: "embedded local config"
    });
  }

  private previewWorkbookImportEmbeddedLocalConfig(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const request = workbookEmbeddedLocalConfigImportRequestFromInput(workbookId, input);
    return this.previewWorkbookPendingAction(metadata, requestedMode, {
      action: { kind: "workbook.import_embedded_local_config", request },
      summary: "Prepared embedded workbook local config import.",
      answer: { kind: "workbook_import_embedded_local_config_preview", workbookId, overwrite: request.overwrite ?? false },
      after: "imported embedded local config"
    });
  }

  private previewWorkbookClose(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const closeBehavior = workbookCloseBehaviorFromInput(input);
    return this.previewWorkbookPendingAction(metadata, requestedMode, {
      action: { kind: "workbook.close", workbookId, ...(closeBehavior ? { closeBehavior } : {}) },
      summary: "Prepared workbook close.",
      answer: { kind: "workbook_close_preview", workbookId, closeBehavior: closeBehavior ?? "default" },
      after: `closed workbook${closeBehavior ? ` (${closeBehavior})` : ""}`
    });
  }

  private previewWorkbookPendingAction(
    metadata: WorkbookMetadata,
    requestedMode: AgentRunMode,
    input: {
      action: Parameters<AgentOperationStore["create"]>[0]["action"];
      summary: string;
      answer: unknown;
      after: string;
    }
  ): Omit<AgentRunOutput, "telemetry"> {
    const pending = this.createPendingOperation(metadata, {
      action: input.action,
      changes: [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: input.after }],
      summary: input.summary
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: input.answer,
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "not_applicable" },
      changes: pending.changes,
      proof: [],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewFileBackupCreate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const values = input.values as Record<string, unknown> | undefined;
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const targetPath = stringValue(values?.targetPath);
    const mode = fileBackupModeFromInput(values?.mode);
    const request: WorkbookCreateFileBackupRequest = {
      workbookId,
      reason: stringValue(values?.reason) ?? input.request,
      ...(targetPath ? { targetPath } : {}),
      ...(mode ? { mode } : {}),
      ...(typeof values?.pin === "boolean" ? { pin: values.pin } : {})
    };
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "backup.create_file", request },
      changes: [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: "created file backup" }],
      summary: "Prepared full workbook file backup creation."
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "backup_create_file_preview", { workbookId, targetPath: request.targetPath, pin: request.pin ?? false });
  }

  private previewFileBackupRestore(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const backupId = backupIdFromInput(input);
    if (!backupId) {
      return safetyArtifactNeedsInput(metadata, requestedMode, "File backup restore needs target.entity, values.backupId, or operationId.");
    }
    const values = input.values as Record<string, unknown> | undefined;
    const mode = fileRestoreModeFromInput(values?.mode);
    const restoreTargetPath = stringValue(values?.restoreTargetPath);
    const request: WorkbookRestoreFileBackupRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      backupId: backupId as BackupId,
      mode,
      ...(restoreTargetPath ? { restoreTargetPath } : {}),
      ...(typeof values?.force === "boolean" ? { force: values.force } : {})
    };
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "backup.restore_file", request },
      changes: [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: `restore file backup ${backupId} (${mode})` }],
      summary: `Prepared file backup restore for ${backupId}.`
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "backup_restore_file_preview", { backupId, mode });
  }

  private previewBackupPrune(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = backupRetentionRequestFromInput(metadata, input, false);
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "backup.prune", request },
      changes: [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: "pruned persisted backups" }],
      summary: "Prepared backup retention pruning."
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "backup_prune_preview", { criteria: request });
  }

  private previewSheetOperation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    behavior: "create" | "copy" | "rename" | "delete" | "hide" | "unhide" | "protect" | "unprotect" | "clear" | "set_tab_color"
  ): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const values = input.values as Record<string, unknown> | undefined;
    const targetSheetName = explicitSheetName(metadata, input);
    const newSheetName = stringValue(values?.newSheetName ?? values?.sheetName ?? values?.name ?? (behavior === "create" ? input.target?.sheetName : undefined));
    const color = stringValue(values?.color ?? values?.tabColor);
    const password = stringValue(values?.password);
    const protectionOptions = sheetProtectionOptionsFromInput(values);
    const clearApplyTo = sheetClearApplyToFromInput(values?.applyTo);
    let operation: ExcelOperation;
    let summary = "";
    let answer: Record<string, unknown> = { kind: `sheet_${behavior}_preview` };
    let proofSheet = targetSheetName ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "";

    if (behavior === "create") {
      if (!newSheetName) {
        return sheetNeedsInput(metadata, requestedMode, "Sheet creation needs values.sheetName or values.newSheetName.");
      }
      operation = { kind: "sheet.create", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: newSheetName, activate: true };
      summary = `Prepared sheet creation for ${newSheetName}.`;
      answer = { ...answer, sheetName: newSheetName };
      proofSheet = newSheetName;
    } else if (behavior === "copy") {
      if (!targetSheetName || !newSheetName) {
        return sheetNeedsInput(metadata, requestedMode, "Sheet copy needs target.sheetName and values.newSheetName.");
      }
      operation = { kind: "sheet.copy", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sourceSheetName: targetSheetName, newSheetName, position: "after", relativeToSheetName: targetSheetName, activate: true };
      summary = `Prepared sheet copy from ${targetSheetName} to ${newSheetName}.`;
      answer = { ...answer, sourceSheetName: targetSheetName, newSheetName };
    } else if (behavior === "rename") {
      if (!targetSheetName || !newSheetName) {
        return sheetNeedsInput(metadata, requestedMode, "Sheet rename needs target.sheetName and values.newSheetName.");
      }
      operation = { kind: "sheet.rename", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName, newSheetName };
      summary = `Prepared sheet rename from ${targetSheetName} to ${newSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName, newSheetName };
    } else if (behavior === "delete") {
      if (!targetSheetName) {
        return sheetNeedsInput(metadata, requestedMode, "Sheet delete needs target.sheetName.");
      }
      operation = { kind: "sheet.delete", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName };
      summary = `Prepared sheet deletion for ${targetSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName };
    } else if (behavior === "hide" || behavior === "unhide") {
      if (!targetSheetName) {
        return sheetNeedsInput(metadata, requestedMode, `Sheet ${behavior} needs target.sheetName.`);
      }
      operation = { kind: behavior === "hide" ? "sheet.hide" : "sheet.unhide", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName };
      summary = `Prepared sheet ${behavior} for ${targetSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName };
    } else if (behavior === "protect" || behavior === "unprotect") {
      if (!targetSheetName) {
        return sheetNeedsInput(metadata, requestedMode, `Sheet ${behavior} needs target.sheetName.`);
      }
      operation = behavior === "protect"
        ? { kind: "sheet.protect", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName, ...(password ? { password } : {}), ...(protectionOptions ? { options: protectionOptions } : {}) }
        : { kind: "sheet.unprotect", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName, ...(password ? { password } : {}) };
      summary = `Prepared sheet ${behavior} for ${targetSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName, ...(protectionOptions ? { options: protectionOptions } : {}) };
    } else if (behavior === "clear") {
      if (!targetSheetName) {
        return sheetNeedsInput(metadata, requestedMode, "Sheet clear needs target.sheetName.");
      }
      operation = { kind: "sheet.clear", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName, applyTo: clearApplyTo };
      summary = `Prepared sheet clear for ${targetSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName, applyTo: clearApplyTo };
    } else {
      if (!targetSheetName || !color) {
        return sheetNeedsInput(metadata, requestedMode, "Sheet tab-color update needs target.sheetName and values.color.");
      }
      operation = { kind: "sheet.set_tab_color", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "format", reason: input.request, sheetName: targetSheetName, color };
      summary = `Prepared sheet tab color update for ${targetSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName, color };
    }

    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: proofSheet, after: summary.replace(/^Prepared /, "") }], summary, answer);
  }

  private previewFormulaCopyPatterns(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = formulaCopyPatternsRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Formula pattern copy needs values.source and values.destination targets with sheet names.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "formula.copy_patterns", request },
      changes: [{ sheetName: request.targetSheetName, ...(request.targetAddress ? { range: request.targetAddress } : {}), after: "copied formula patterns" }],
      summary: `Prepared formula pattern copy from ${request.sourceSheetName}${request.sourceAddress ? `!${request.sourceAddress}` : ""} to ${request.targetSheetName}${request.targetAddress ? `!${request.targetAddress}` : ""}.`
    });
    return formulaMutationPreviewOutput(metadata, requestedMode, pending, "formula_copy_patterns_preview", request.targetSheetName, request.targetAddress);
  }

  private previewFormulaFill(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, direction: "down" | "right"): Omit<AgentRunOutput, "telemetry"> {
    const request = formulaFillRequestFromInput(metadata, input, direction);
    if (!request) {
      const sheetName = input.target?.sheetName ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name;
      return {
        ...workbookLevelNeedsInput(metadata, requestedMode, `Formula fill ${direction} needs values.source and values.destination ranges on the same sheet.`),
        answer: stripUndefinedRecord({
          kind: `formula_fill_${direction}_needs_input`,
          suggestedIntentAction: direction === "down" ? "fill_formula_down" : "fill_formula_right",
          suggestedValues: {
            source: { ...(sheetName ? { sheetName } : {}), range: "SOURCE_CELL" },
            destination: { ...(sheetName ? { sheetName } : {}), range: direction === "down" ? "TARGET_COLUMN_RANGE" : "TARGET_ROW_RANGE" }
          }
        }),
        agentInstruction: "Retry once with values.source and values.destination in the suggested object shape; do not split formula fill into single-cell writes.",
        maxRecommendedFollowupCalls: 1
      };
    }
    const redirect = this.fragmentationRedirect(metadata, requestedMode, {
      family: direction === "down" ? "formula_fill_down" : "formula_fill_right",
      workbookContextId: metadata.workbookContextId,
      sourceSheetName: request.sheetName,
      sourceAddress: request.sourceAddress,
      targetSheetName: request.sheetName,
      targetAddress: request.targetAddress,
      request: input.request
    });
    if (redirect) return redirect;
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "formula.fill_pattern", request },
      changes: [{ sheetName: request.sheetName, range: request.targetAddress, after: `filled formulas ${direction}` }],
      summary: `Prepared formula fill ${direction} from ${request.sheetName}!${request.sourceAddress} to ${request.sheetName}!${request.targetAddress}.`
    });
    return formulaMutationPreviewOutput(metadata, requestedMode, pending, `formula_fill_${direction}_preview`, request.sheetName, request.targetAddress);
  }

  private previewFormulaRepairPatterns(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = formulaTemplateRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Formula pattern repair needs values.templateId and target.sheetName or values.targetSheetName.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "formula.repair_patterns", request },
      changes: [{ sheetName: request.targetSheetName, after: `repaired formulas from template ${request.templateId}` }],
      summary: `Prepared formula pattern repair on ${request.targetSheetName} from template ${request.templateId}.`
    });
    return formulaMutationPreviewOutput(metadata, requestedMode, pending, "formula_repair_patterns_preview", request.targetSheetName);
  }

  private previewFormulaConvertToValues(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const request: FormulaPatternRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      sheetName: resolved.sheetName,
      address: resolved.range
    };
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "formula.convert_to_values", request },
      changes: [{ sheetName: resolved.sheetName, range: resolved.range, before: "formulas", after: "values" }],
      summary: `Prepared formula-to-values conversion on ${resolved.sheetName}!${resolved.range}.`
    });
    return formulaMutationPreviewOutput(metadata, requestedMode, pending, "formula_convert_to_values_preview", resolved.sheetName, resolved.range);
  }

  private previewNameMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    behavior: "create" | "update" | "delete"
  ): Omit<AgentRunOutput, "telemetry"> {
    const request = nameMutationRequestFromInput(metadata, input, behavior);
    if (!request) {
      return nameRegionNeedsInput(metadata, requestedMode, `Name ${behavior} needs values.name or target.entity, plus values.reference/formula or target range for create/update.`);
    }
    const pending = this.createPendingOperation(metadata, {
      action: behavior === "create"
        ? { kind: "names.create", request: request as NameCreateRequest }
        : behavior === "update"
          ? { kind: "names.update", request: request as NameUpdateRequest }
          : { kind: "names.delete", request: request as NameSelector },
      changes: [{ sheetName: request.sheetName ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", ...(("reference" in request && request.reference) ? { range: stripSheetName(request.reference) } : {}), after: `${behavior} name ${request.name}` }],
      summary: `Prepared name ${behavior} for ${request.name}.`
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, `name_${behavior}_preview`, { name: request.name, ...(request.sheetName ? { sheetName: request.sheetName } : {}) });
  }

  private previewRegionMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    behavior: "register" | "clear_values" | "write_values" | "fill"
  ): Omit<AgentRunOutput, "telemetry"> {
    const request = regionMutationRequestFromInput(metadata, input, behavior);
    if (!request) {
      return nameRegionNeedsInput(metadata, requestedMode, behavior === "register"
        ? "Region registration needs values.name or values.regionName plus target.sheetName/target.range or values.sheetName/values.address."
        : `Region ${behavior.replace("_", " ")} needs values.regionName, values.name, or target.entity${behavior === "clear_values" ? "." : " plus values.values or values.rows."}`);
    }
    const action = behavior === "register"
      ? { kind: "region.register" as const, request: request as RegionRegisterRequest }
      : behavior === "clear_values"
        ? { kind: "region.clear_values" as const, request: request as RegionSelector }
        : behavior === "write_values"
          ? { kind: "region.write_values" as const, request: request as RegionSelector & { values: unknown[][] } }
          : { kind: "region.fill" as const, request: request as RegionSelector & { values: unknown[][]; clearFirst?: boolean } };
    const regionName = "name" in request ? request.name : request.regionName;
    const sheetName = "sheetName" in request ? request.sheetName : metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "";
    const range = "address" in request ? request.address : undefined;
    const changes: NonNullable<AgentRunOutput["changes"]> = [{
      sheetName,
      ...(range ? { range } : {}),
      after: `${behavior.replace("_", " ")} region ${regionName}`
    }];
    const pending = this.createPendingOperation(metadata, {
      action,
      changes,
      summary: `Prepared region ${behavior.replace("_", " ")} for ${regionName}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: `region_${behavior}_preview`,
        regionName,
        ...(range ? { sheetName, range } : {}),
        ...(("values" in request && Array.isArray(request.values)) ? { rowCount: request.values.length, columnCount: request.values.reduce((max, row) => Math.max(max, row.length), 0) } : {})
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: range ? [{ sheetName, range, label: regionName }] : [],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewClearValues(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const redirect = this.fragmentationRedirect(metadata, requestedMode, {
      family: "clear_values",
      workbookContextId: metadata.workbookContextId,
      targetSheetName: resolved.sheetName,
      targetAddress: resolved.range,
      request: input.request
    });
    if (redirect) return redirect;
    const operation: ExcelOperation = {
      kind: "range.clear_values_keep_format",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, before: "values", after: "cleared values; formats preserved" }], `Prepared clear-values preview on ${resolved.sheetName}!${resolved.range}.`, { kind: "clear_values_preview", sheetName: resolved.sheetName, range: resolved.range });
  }

  private previewClearRange(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind: "range.clear",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "structure",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      applyTo: "all"
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, before: "values and formats", after: "cleared" }], `Prepared clear-range preview on ${resolved.sheetName}!${resolved.range}.`, { kind: "clear_range_preview", sheetName: resolved.sheetName, range: resolved.range });
  }

  private previewClearFormats(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind: "range.clear_formats",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, before: "formats", after: "cleared formats" }], `Prepared clear-formats preview on ${resolved.sheetName}!${resolved.range}.`, { kind: "clear_formats_preview", sheetName: resolved.sheetName, range: resolved.range });
  }

  private previewClearStyleDimensions(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const dimensions = styleDimensionsFromAgentInput(input);
    if (dimensions.length === 0) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Style-dimension clearing needs values.dimensions or a request that names borders, fills, fonts, alignment, number formats, row heights, or column widths.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "style-dimension target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
      };
    }
    const operation: ExcelOperation = {
      kind: "range.clear_style_dimensions",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      dimensions
    };
    const label = dimensions.join(", ");
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: resolved.sheetName, range: resolved.range, before: label, after: `cleared ${label}` }],
      `Prepared style-dimension clear preview on ${resolved.sheetName}!${resolved.range}.`,
      { kind: "clear_style_dimensions_preview", sheetName: resolved.sheetName, range: resolved.range, dimensions }
    );
  }

  private previewAutofit(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>, dimension: "columns" | "rows"): Omit<AgentRunOutput, "telemetry"> {
    const redirect = this.fragmentationRedirect(metadata, requestedMode, {
      family: `autofit_${dimension}`,
      workbookContextId: metadata.workbookContextId,
      targetSheetName: resolved.sheetName,
      targetAddress: resolved.range,
      request: input.request
    });
    if (redirect) return redirect;
    const operation: ExcelOperation = {
      kind: dimension === "columns" ? "range.autofit_columns" : "range.autofit_rows",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, after: `autofit ${dimension}` }], `Prepared autofit ${dimension} on ${resolved.sheetName}!${resolved.range}.`, { kind: "autofit_preview", dimension, sheetName: resolved.sheetName, range: resolved.range });
  }

  private previewFreezePanes(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved?: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const sheetName = resolved?.sheetName ?? freezePanesSheetName(metadata, input);
    if (!sheetName) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Freeze panes needs a target sheet.",
        candidates: findAgentCandidates(metadata, input).slice(0, 5),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: ["Provide target.sheetName or use a prepared workbook context with an active sheet."]
      };
    }
    const freeze = freezePanesFromInput(input);
    if (!freeze) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Freeze panes needs a clear request such as unfreeze panes, freeze top row, freeze first column, or values.freezePanes with row/column counts.",
        proof: [{ sheetName, range: resolved?.range ?? usedRangeForSheet(metadata, sheetName), label: "freeze panes target sheet" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
      };
    }
    const operation: ExcelOperation = {
      kind: "sheet.freeze_panes",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      sheetName,
      rows: freeze.rows ?? 0,
      columns: freeze.columns ?? 0
    };
    const after = freeze.rows === 0 && freeze.columns === 0
      ? "unfreeze all panes"
      : `freeze panes at ${freeze.rows ?? 0} row(s) and ${freeze.columns ?? 0} column(s)`;
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName, after }],
      `Prepared freeze panes update on ${sheetName}.`,
      { kind: "freeze_panes_preview", sheetName, freezePanes: { rows: operation.rows, columns: operation.columns } }
    );
  }

  private previewAutoFilterMutation(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    if (isClearFilterRequest(input.request)) {
      return this.previewClearAutoFilter(metadata, input, requestedMode, resolved);
    }
    return this.previewApplyAutoFilter(metadata, input, requestedMode, resolved);
  }

  private previewApplyAutoFilter(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind: "range.apply_autofilter",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, after: "filters enabled" }], `Prepared filter preview on ${resolved.sheetName}!${resolved.range}.`, { kind: "filter_preview", sheetName: resolved.sheetName, range: resolved.range });
  }

  private previewClearAutoFilter(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind: "range.clear_autofilter",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, after: "filters cleared" }], `Prepared clear filter preview on ${resolved.sheetName}!${resolved.range}.`, { kind: "filter_clear_preview", sheetName: resolved.sheetName, range: resolved.range });
  }

  private previewNumberFormatUpdate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const numberFormat = numberFormatMatrixFromInput(input, resolved.range);
    if (!numberFormat) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Number-format updates need values.numberFormats, values.numberFormat, or values.formats. Example: values: { numberFormat: \"dd/mm/yyyy\" }.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "number-format target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: [`Retry with values.numberFormat or a ${resolved.range} shaped values.numberFormats matrix.`]
      };
    }
    const shapeIssue = matrixShapeIssue(resolved.range, numberFormat);
    if (shapeIssue) {
      return {
        status: "VALIDATION_FAILED",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Number-format matrix shape does not match ${resolved.sheetName}!${resolved.range}.`,
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "number-format target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "manual_review",
        warnings: [shapeIssue]
      };
    }
    const operation: ExcelOperation = {
      kind: "range.write_number_formats",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      numberFormat,
      preserveValues: true
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, after: "number formats updated" }], `Prepared number-format preview on ${resolved.sheetName}!${resolved.range}.`, { kind: "number_format_preview", sheetName: resolved.sheetName, range: resolved.range, cellCount: matrixCellCount(numberFormat) });
  }

  private previewRangeCopyMove(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, behavior: "copy" | "move"): Omit<AgentRunOutput, "telemetry"> {
    const endpoints = rangeTransferEndpointsFromInput(metadata, input);
    if (!endpoints) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `Range ${behavior} needs explicit values.source and values.destination targets with sheetName and range.`,
        candidates: findAgentCandidates(metadata, input).slice(0, 5),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
      };
    }
    const operation: ExcelOperation = behavior === "copy"
      ? {
          kind: "range.copy",
          operationId: makeId<OperationId>("op"),
          workbookId: metadata.workbook.workbookId as WorkbookId,
          destructiveLevel: destructiveLevelForRangeCopy(input),
          reason: input.request,
          source: endpoints.source,
          target: endpoints.destination,
          copyType: rangeCopyTypeFromInput(input)
        }
      : {
          kind: "range.move",
          operationId: makeId<OperationId>("op"),
          workbookId: metadata.workbook.workbookId as WorkbookId,
          destructiveLevel: "structure",
          reason: input.request,
          source: endpoints.source,
          target: endpoints.destination
        };
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: endpoints.destination.sheetName, range: endpoints.destination.address, before: undefined, after: `${behavior} from ${endpoints.source.sheetName}!${endpoints.source.address}` }],
      `Prepared range ${behavior} from ${endpoints.source.sheetName}!${endpoints.source.address} to ${endpoints.destination.sheetName}!${endpoints.destination.address}.`,
      { kind: `range_${behavior}_preview`, source: endpoints.source, destination: endpoints.destination }
    );
  }

  private structuralOperationWarning(input: AgentRunInput, resolved: Extract<AgentTargetResolution, { ok: true }>, requestedKind: RangeStructuralOperationKind, actualKind: RangeStructuralOperationKind): string | undefined {
    if (requestedKind === actualKind) {
      return undefined;
    }
    return `Corrected ${String(input.intent?.action ?? "structural operation")} to ${actualKind.replace("range.", "").replace(/_/g, " ")} because ${resolved.range} is a ${structuralAddressShape(resolved.range) ?? "matching"} target.`;
  }

  private previewRangeStructuralOperation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    kind: RangeStructuralOperationKind
  ): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const values = input.values as Record<string, unknown> | undefined;
    const base = {
      kind,
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: kind === "range.clear_values" ? "values" as const : "structure" as const,
      reason: input.request,
      target: { workbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    const operation: ExcelOperation =
      kind === "range.merge"
        ? { ...base, kind, ...(typeof values?.across === "boolean" ? { across: values.across } : {}) }
        : base;
    const actionLabel = kind.replace("range.", "").replace(/_/g, " ");
    const output = this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: resolved.sheetName, range: resolved.range, after: actionLabel }],
      `Prepared ${actionLabel} on ${resolved.sheetName}!${resolved.range}.`,
      { kind: `${kind}_preview`, sheetName: resolved.sheetName, range: resolved.range }
    );
    const corrected = this.structuralOperationWarning(input, resolved, intentStructuralOperationKind(input) ?? kind, kind);
    return corrected ? { ...output, warnings: [corrected, ...output.warnings] } : output;
  }

  private previewWriteStylesMany(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const entries = styleEntriesFromInput(metadata, workbookId, input);
    if (entries.length === 0) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Multi-style writes need values.entries with sheetName, range, and style.");
    }
    if (entries.length === 1) {
      const entry = entries[0]!;
      const redirect = this.fragmentationRedirect(metadata, requestedMode, {
        family: "format_range",
        workbookContextId: metadata.workbookContextId,
        targetSheetName: entry.target.sheetName,
        targetAddress: entry.target.address,
        request: input.request,
        style: entry.style
      });
      if (redirect) return redirect;
    }
    const mergeEntries = shouldMergeRangesFromRequest(input) ? mergeEntriesFromStyleEntries(entries) : [];
    const operations: ExcelOperation[] = [
      ...mergeEntries.map((entry): ExcelOperation => ({
        kind: "range.merge",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "structure",
        reason: input.request,
        target: entry.target,
        across: false
      })),
      {
      kind: "range.write_styles_many",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: input.request,
      entries: entries.map((entry) => ({ target: entry.target, style: entry.style, preserveValues: true as const }))
      }
    ];
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      operations,
      [
        ...mergeEntries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, after: "merged range" })),
        ...entries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, after: "styles updated" }))
      ],
      mergeEntries.length > 0
        ? `Prepared ${mergeEntries.length} merge(s) and style updates for ${entries.length} range(s).`
        : `Prepared style updates for ${entries.length} range(s).`,
      { kind: mergeEntries.length > 0 ? "merge_and_write_styles_many_preview" : "write_styles_many_preview", mergeCount: mergeEntries.length, rangeCount: entries.length }
    );
  }

  private previewMergeRangesWithStyles(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const entries = mergeEntriesFromInput(metadata, workbookId, input);
    if (entries.length === 0) {
      return undefined;
    }
    const mergeEntries = entries.filter((entry) => isMultiCellRange(entry.target.address));
    const styleEntries = entries
      .map((entry) => ({ target: entry.target, style: entry.style ?? defaultStyleForMergeRequest(input) }))
      .filter((entry) => Object.keys(entry.style).length > 0);
    const operations: ExcelOperation[] = [
      ...mergeEntries.map((entry): ExcelOperation => ({
        kind: "range.merge",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "structure",
        reason: input.request,
        target: entry.target,
        across: false
      })),
      ...(styleEntries.length > 0 ? [{
        kind: "range.write_styles_many" as const,
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "format" as const,
        reason: input.request,
        entries: styleEntries.map((entry) => ({ target: entry.target, style: entry.style, preserveValues: true as const }))
      }] : [])
    ];
    if (operations.length === 0) {
      return undefined;
    }
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      operations,
      [
        ...mergeEntries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, after: "merged range" })),
        ...styleEntries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, after: "styles updated" }))
      ],
      `Prepared ${mergeEntries.length} merge(s)${styleEntries.length > 0 ? ` and style updates for ${styleEntries.length} range(s)` : ""}.`,
      { kind: "merge_ranges_preview", mergeCount: mergeEntries.length, rangeCount: entries.length, styledRangeCount: styleEntries.length }
    );
  }

  private previewWriteDataValidation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const validation = dataValidationFromInput(input);
    if (!validation) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Data validation writes need values.validation.source or values.options for the dropdown list.");
    }
    const entries = dataValidationEntriesFromInput(workbookId, input, validation);
    if (entries.length > 0) {
      const operation: ExcelOperation = {
        kind: "range.write_data_validation",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "format",
        reason: input.request,
        target: entries[0]!.target,
        validation: entries[0]!.validation,
        entries
      };
      return this.previewBatchOperation(
        metadata,
        requestedMode,
        [operation],
        entries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, after: "data validation updated" })),
        `Prepared data validation updates for ${entries.length} range(s).`,
        { kind: "write_data_validation_preview", rangeCount: entries.length, entries: entries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, validation: entry.validation })) }
      );
    }
    const operation: ExcelOperation = {
      kind: "range.write_data_validation",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId, sheetName: resolved.sheetName, address: resolved.range },
      validation
    };
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: resolved.sheetName, range: resolved.range, after: "data validation updated" }],
      `Prepared data validation update on ${resolved.sheetName}!${resolved.range}.`,
      { kind: "write_data_validation_preview", sheetName: resolved.sheetName, range: resolved.range, validation }
    );
  }

  private previewWriteConditionalFormatting(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const rule = conditionalFormattingRuleFromInput(input);
    if (!rule) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Conditional formatting writes need values.rule.formula and values.rule.style, or values.formula plus values.style.");
    }
    const operation: ExcelOperation = {
      kind: "range.write_conditional_formatting",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId, sheetName: resolved.sheetName, address: resolved.range },
      rule
    };
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: resolved.sheetName, range: resolved.range, after: { conditionalFormatting: rule } }],
      `Prepared conditional formatting update on ${resolved.sheetName}!${resolved.range}.`,
      { kind: "write_conditional_formatting_preview", sheetName: resolved.sheetName, range: resolved.range, rule }
    );
  }

  private previewRangeColumnReorder(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const columnOrder = columnOrderFromInput(input);
    if (columnOrder.length === 0) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Column reorder needs values.columnOrder with every target column exactly once.");
    }
    const operation: ExcelOperation = {
      kind: "range.reorder_columns",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "structure",
      reason: input.request,
      target: { workbookId, sheetName: resolved.sheetName, address: resolved.range },
      columnOrder
    };
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: resolved.sheetName, range: resolved.range, after: { columnOrder } }],
      `Prepared column reorder on ${resolved.sheetName}!${resolved.range}.`,
      { kind: "range.reorder_columns_preview", sheetName: resolved.sheetName, range: resolved.range, columnOrder }
    );
  }

  private previewBatchOperation(metadata: WorkbookMetadata, requestedMode: AgentRunMode, operations: ExcelOperation[], changes: NonNullable<AgentRunOutput["changes"]>, summary: string, answer: unknown): Omit<AgentRunOutput, "telemetry"> {
    const workflowKind = previewWorkflowKind(answer);
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
      changes,
      summary,
      ...(workflowKind !== undefined ? { workflowKind } : {})
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary,
      answer,
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", safetyFingerprintOnly: true, ...(workflowKind !== undefined ? { workflowKind } : {}) },
      changes,
      proof: changes.flatMap((change) => change.range ? [{ sheetName: change.sheetName, range: change.range, label: "preview target" }] : []).slice(0, 1),
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private async previewFirstStatusMatchUpdate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const sheet = findMentionedSheet(metadata, input) ?? metadata.sheets.find((candidate) => candidate.name === input.target?.sheetName);
    if (!sheet?.usedRange) {
      return undefined;
    }
    const values = await this.readRangeValues(metadata.workbook.workbookId as WorkbookId, sheet.name, sheet.usedRange);
    const headers = values[0]?.map((value) => String(value ?? "").trim().toLowerCase()) ?? [];
    const statusIndex = headers.findIndex((header) => header === "status");
    if (statusIndex < 0) {
      return undefined;
    }
    const rowIndex = values.findIndex((row, index) => index > 0 && String(row[statusIndex] ?? "").trim().toLowerCase() === "open");
    if (rowIndex < 1) {
      return {
        status: "NOT_FOUND",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: `No open item was found on ${sheet.name}.`,
        proof: [{ sheetName: sheet.name, range: sheet.usedRange, label: "searched range" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
      };
    }
    const range = `${columnLetter(statusIndex)}${rowIndex + 1}`;
    return this.previewUpdate(metadata, {
      ...input,
      target: { ...(input.target ?? {}), sheetName: sheet.name, range },
      values: { values: [["Reviewed"]] }
    }, requestedMode);
  }

  private previewTableSort(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
    const table = resolveAgentTable(metadata, input);
    if (!table?.name && !table?.id) {
      return undefined;
    }
    const sortSpec = tableSortSpecFromInput(input, table);
    const amountColumn = table.columns.find((column) => /amount/i.test(column.name));
    const key = sortSpec?.key ?? amountColumn?.index ?? table.columns.findIndex((column) => /amount/i.test(column.name));
    if (key < 0) {
      return undefined;
    }
    const tableName = table.name ?? table.id;
    const sortColumn = table.columns.find((column) => column.index === key) ?? amountColumn;
    const request: TableSortRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName,
      fields: [{
        key,
        ascending: sortSpec?.ascending ?? !/\b(highest|descending|desc|largest|lowest to highest)\b/i.test(input.request),
        ...(sortSpec?.sortOn ? { sortOn: sortSpec.sortOn } : {}),
        ...(sortSpec?.color ? { color: sortSpec.color } : {}),
        ...(sortSpec?.dataOption ? { dataOption: sortSpec.dataOption } : {})
      }]
    };
    const changes: NonNullable<AgentRunOutput["changes"]> = [{ sheetName: table.sheetName, range: table.range, after: `sorted ${tableName} by ${sortColumn?.name ?? "Amount"}` }];
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "table.sort", request },
      changes,
      summary: `Prepared sort preview for ${tableName}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "table_sort_preview", tableName, sheetName: table.sheetName, sortField: sortColumn?.name ?? "Amount", ascending: request.fields[0]?.ascending !== false },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: [{ sheetName: table.sheetName, range: table.range, label: tableName }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewTableUpdateRows(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const rows = tableRowUpdatesFromInput(input);
    if (rows.length === 0) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table row updates need values.rows as { index, values } entries.");
    }
    const request: TableUpdateRowsRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, rows };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.update_rows", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared ${rows.length} table row update(s) for ${tableName}.`,
      answer: { kind: "table_update_rows_preview", tableName, rowCount: rows.length },
      after: "updated table rows"
    });
  }

  private previewTableCreate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const values = input.values as Record<string, unknown> | undefined;
    const sheetName = stringValue(values?.sheetName ?? input.target?.sheetName);
    const address = stringValue(values?.address ?? values?.range ?? input.target?.range);
    if (!sheetName || !address) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Table creation needs target.sheetName and target.range or values.sheetName and values.address.");
    }
    const tableNameInput = stringValue(values?.tableName ?? values?.name);
    const tableValues = explicitCellMatrixFromValues(values);
    const style = stringValue(values?.style);
    const request: TableCreateRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      sheetName,
      address,
      hasHeaders: typeof values?.hasHeaders === "boolean" ? values.hasHeaders : true,
      ...(tableNameInput ? { tableName: tableNameInput } : {}),
      ...(tableValues.length > 0 ? { values: tableValues } : {}),
      ...(style ? { style } : {}),
      ...(typeof values?.showTotals === "boolean" ? { showTotals: values.showTotals } : {})
    };
    const tableName = request.tableName ?? address;
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.create", request },
      tableName,
      sheetName,
      range: address,
      summary: `Prepared table creation for ${tableName}.`,
      answer: { kind: "table_create_preview", tableName: request.tableName, sheetName, range: address, hasHeaders: request.hasHeaders },
      after: `created table ${tableName}`
    });
  }

  private previewTableResize(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const values = input.values as Record<string, unknown> | undefined;
    const address = stringValue(values?.address ?? values?.range ?? values?.newAddress);
    if (!address) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table resize needs values.address, values.range, or values.newAddress.");
    }
    const request: TableResizeRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, address };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.resize", request },
      tableName,
      sheetName: resolved.sheetName,
      range: address,
      summary: `Prepared resize preview for ${tableName}.`,
      answer: { kind: "table_resize_preview", tableName, address },
      after: `resized table to ${address}`
    });
  }

  private previewTableReorderColumns(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const columnOrder = tableColumnOrderFromInput(input);
    if (columnOrder.length === 0) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table column reorder needs values.columnOrder as column names or indexes.");
    }
    const request: TableReorderColumnsRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, columnOrder };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.reorder_columns", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared column reorder preview for ${tableName}.`,
      answer: { kind: "table_reorder_columns_preview", tableName, columnOrder },
      after: `reordered ${columnOrder.length} table columns`
    });
  }

  private previewTableSelectorMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    behavior: "clear_data_keep_formulas" | "clear_filters"
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const request: TableSelector = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName };
    const action = behavior === "clear_data_keep_formulas"
      ? { kind: "table.clear_data_keep_formulas" as const, request }
      : { kind: "table.clear_filters" as const, request };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action,
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: behavior === "clear_data_keep_formulas"
        ? `Prepared clear table data preview for ${tableName}.`
        : `Prepared clear table filters preview for ${tableName}.`,
      answer: { kind: behavior === "clear_data_keep_formulas" ? "table_clear_data_preview" : "table_clear_filters_preview", tableName },
      after: behavior === "clear_data_keep_formulas" ? "cleared table data; formulas preserved" : "cleared table filters"
    });
  }

  private previewTableApplyFilters(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const filters = tableFiltersFromInput(input);
    if (!filters.ok) {
      return tableNeedsInput(metadata, requestedMode, resolved, filters.summary);
    }
    const request: TableApplyFiltersRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, filters: filters.filters };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.apply_filters", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared ${filters.filters.length} table filter(s) for ${tableName}.`,
      answer: { kind: "table_apply_filters_preview", tableName, filterCount: filters.filters.length },
      after: "applied table filters"
    });
  }

  private previewTableApplyView(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const values = input.values as Record<string, unknown> | undefined;
    const filtersResult = Array.isArray(values?.filters) ? tableFiltersFromInput(input) : { ok: true as const, filters: [] };
    if (!filtersResult.ok) {
      return tableNeedsInput(metadata, requestedMode, resolved, filtersResult.summary);
    }
    const sort = tableApplyViewSortFromInput(input, table);
    const clearFilters = typeof values?.clearFilters === "boolean" ? values.clearFilters : false;
    const clearSort = typeof values?.clearSort === "boolean" ? values.clearSort : false;
    if (filtersResult.filters.length === 0 && sort.fields.length === 0 && !clearFilters && !clearSort) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table view previews need values.filters, values.sort.fields, clearFilters, or clearSort.");
    }
    const request: TableApplyViewRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName,
      ...(filtersResult.filters.length > 0 ? { filters: filtersResult.filters } : {}),
      ...(sort.fields.length > 0 ? { sort } : {}),
      ...(clearFilters ? { clearFilters } : {}),
      ...(clearSort ? { clearSort } : {})
    };
    const parts = [
      filtersResult.filters.length > 0 ? `${filtersResult.filters.length} filter(s)` : undefined,
      sort.fields.length > 0 ? `${sort.fields.length} sort field(s)` : undefined,
      clearFilters ? "clear filters first" : undefined,
      clearSort ? "clear sort first" : undefined
    ].filter((part): part is string => Boolean(part));
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.apply_view", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared table view update for ${tableName}: ${parts.join(", ")}.`,
      answer: { kind: "table_apply_view_preview", tableName, filterCount: filtersResult.filters.length, sortFieldCount: sort.fields.length, clearFilters, clearSort },
      after: "applied table view"
    });
  }

  private previewTableTotalRow(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const showTotals = typeof input.values?.showTotals === "boolean" ? input.values.showTotals : !/\b(hide|remove|off|disable)\b/i.test(input.request);
    const request: TableSetTotalRowRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, showTotals };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.set_total_row", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared ${showTotals ? "show" : "hide"} total row preview for ${tableName}.`,
      answer: { kind: "table_total_row_preview", tableName, showTotals },
      after: showTotals ? "shown total row" : "hidden total row"
    });
  }

  private previewTableStyle(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const style = typeof input.values?.style === "string" ? input.values.style : undefined;
    if (!style) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table style previews need values.style.");
    }
    const request: TableSetStyleRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, style };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.set_style", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared table style preview for ${tableName}.`,
      answer: { kind: "table_style_preview", tableName, style },
      after: `set table style ${style}`
    });
  }

  private previewTableCopyStructure(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    const table = tableFromResolution(metadata, resolved);
    const tableName = table?.name ?? resolved.candidate.tableName ?? resolved.candidate.label;
    const values = input.values as Record<string, unknown> | undefined;
    const targetSheetName = stringValue(values?.targetSheetName ?? values?.sheetName);
    const targetAddress = stringValue(values?.targetAddress ?? values?.address ?? values?.range);
    if (!targetSheetName || !targetAddress) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table structure copy needs values.targetSheetName and values.targetAddress.");
    }
    const newTableName = stringValue(values?.newTableName);
    const request: TableCopyStructureRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName,
      targetSheetName,
      targetAddress,
      ...(newTableName ? { newTableName } : {}),
      ...(typeof values?.includeStyle === "boolean" ? { includeStyle: values.includeStyle } : {}),
      ...(typeof values?.includeTotals === "boolean" ? { includeTotals: values.includeTotals } : {}),
      ...(typeof values?.includeFilters === "boolean" ? { includeFilters: values.includeFilters } : {})
    };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.copy_structure", request },
      tableName,
      sheetName: targetSheetName,
      range: targetAddress,
      summary: `Prepared table structure copy from ${tableName}.`,
      answer: { kind: "table_copy_structure_preview", tableName, targetSheetName, targetAddress, newTableName: request.newTableName },
      after: `copied table structure from ${tableName}`
    });
  }

  private previewTablePendingAction(
    metadata: WorkbookMetadata,
    requestedMode: AgentRunMode,
    input: {
      action: Parameters<AgentOperationStore["create"]>[0]["action"];
      tableName: string;
      sheetName: string;
      range: string;
      summary: string;
      answer: unknown;
      after: string;
    }
  ): Omit<AgentRunOutput, "telemetry"> {
    const changes: NonNullable<AgentRunOutput["changes"]> = [{ sheetName: input.sheetName, range: input.range, after: input.after }];
    const pending = this.createPendingOperation(metadata, {
      action: input.action,
      changes,
      summary: input.summary
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: input.answer,
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: [{ sheetName: input.sheetName, range: input.range, label: input.tableName }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private async previewAdvancedMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    kind: AdvancedMutationKind
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    if (kind === "style") {
      const resolved = resolveAgentUpdateTarget(metadata, input);
      if (!resolved.ok) {
        return {
          status: resolved.status,
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: resolved.summary,
          ...(resolved.candidates !== undefined ? { candidates: resolved.candidates } : {}),
          proof: [],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: resolved.nextAction,
          warnings: resolved.warnings
        };
      }
      return this.previewStyleUpdate(metadata, input, requestedMode, resolved);
    }
    if (kind === "template") {
      return this.previewTemplateCleanup(metadata, input, requestedMode);
    }
    return undefined;
  }

  private previewStyleUpdate(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>
  ): Omit<AgentRunOutput, "telemetry"> {
    if (/\b(clear|remove|delete|wipe)\b/i.test(input.request) && styleDimensionsFromAgentInput(input).length > 0) {
      return this.previewClearStyleDimensions(metadata, input, requestedMode, resolved);
    }
    const batchEntries = styleEntriesFromInput(metadata, metadata.workbook.workbookId as WorkbookId, input);
    if (batchEntries.length > 0 && hasStyleBatchInput(input)) {
      return this.previewWriteStylesMany(metadata, input, requestedMode);
    }
    let style = styleFromInput(input);
    if (Object.keys(style).length === 0 && intentAction(input) === "format_range" && hasExactFormatRangeTarget(input)) {
      style = defaultFormatRangeStyle(input);
    }
    if (Object.keys(style).length === 0) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Style update needs at least one supported style property such as fillColor, fontColor, fontBold, alignment, rowHeight, or columnWidth.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "style target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: [
          "No apply-ready style properties were parsed, so no workbook preview was created.",
          "For batched width changes, send intent.action write_styles_many with values.entries containing sheetName, range, and style.columnWidth."
        ]
      };
    }
    const warnings = groupedHeaderStyleWarnings(input, style);
    const redirect = this.fragmentationRedirect(metadata, requestedMode, {
      family: "format_range",
      workbookContextId: metadata.workbookContextId,
      targetSheetName: resolved.sheetName,
      targetAddress: resolved.range,
      request: input.request,
      style
    });
    if (redirect) return redirect;
    const operation: ExcelOperation = {
      kind: "range.write_styles",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      style,
      preserveValues: true
    };
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations: [operation] },
      changes: [{ sheetName: resolved.sheetName, range: resolved.range, after: style }],
      summary: `Prepared style update on ${resolved.sheetName}!${resolved.range}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "style_preview", sheetName: resolved.sheetName, range: resolved.range, style },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "style target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings
    };
  }

  private previewCleanMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    action: AgentCleanMutationAction
  ): Omit<AgentRunOutput, "telemetry"> {
    const requests = cleanPatchRequestsFromInput(metadata, input, resolved);
    if (requests.length > 0) {
      const pending = this.createPendingOperation(metadata, {
        action: { kind: "clean.transform_many", action, requests },
        changes: requests.map((request) => ({ sheetName: request.sheetName, range: cleanOutputAddress(request), after: `cleaned with ${action}` })),
        summary: `Prepared cleaning operation ${action} across ${requests.length} exact range(s).`
      });
      return {
        status: "PREVIEW_READY",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        operationId: pending.operationId,
        confirmationToken: pending.confirmationToken,
        summary: pending.summary,
        answer: { kind: "cleaning_preview", action, requests, grouped: true, rangeCount: requests.length },
        metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched", groupedOperationCount: requests.length },
        changes: pending.changes,
        proof: requests.slice(0, 5).map((request) => ({ sheetName: request.sheetName, range: request.address, label: "cleaning target" })),
        resourceLinks: [operationResource(String(pending.operationId))],
        nextAction: "call_apply_update",
        warnings: []
      };
    }
    const request = cleanRequestFromInput(metadata, input, resolved);
    if (!request) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Cleaning updates need a concrete sheet and range target.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "clean.transform", action, request },
      changes: [{ sheetName: request.sheetName, range: cleanOutputAddress(request), after: `cleaned with ${action}` }],
      summary: `Prepared cleaning operation ${action} on ${request.sheetName}!${request.address}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "cleaning_preview", action, request },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: [{ sheetName: request.sheetName, range: request.address, label: "cleaning target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewStyleCopy(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const requests = styleCopyRequestsFromInput(metadata, input);
    if (requests.length === 0) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Style copy needs values.source and values.destination, or source/target sheet names.");
    }
    const mismatch = requests.map((request) => styleCopyDimensionIssue(request)).find(Boolean);
    if (mismatch) {
      return workbookLevelNeedsInput(metadata, requestedMode, mismatch);
    }
    if (requests.length > 1) {
      const pending = this.createPendingOperation(metadata, {
        action: { kind: "style.copy_dimensions_many", requests },
        changes: requests.map((request) => ({ sheetName: request.targetSheetName, ...(request.targetAddress ? { range: request.targetAddress } : {}), after: `copied style dimensions from ${request.sourceSheetName}` })),
        summary: `Prepared ${requests.length} grouped style copy operation(s).`
      });
      return {
        status: "PREVIEW_READY",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        operationId: pending.operationId,
        confirmationToken: pending.confirmationToken,
        summary: pending.summary,
        answer: { kind: "style_copy_many_preview", requests, copyCount: requests.length },
        metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
        changes: pending.changes,
        proof: requests.slice(0, 8).map((request) => ({ sheetName: request.targetSheetName, range: request.targetAddress ?? usedRangeForSheet(metadata, request.targetSheetName), label: "style copy target" })),
        resourceLinks: [operationResource(String(pending.operationId))],
        nextAction: "call_apply_update",
        warnings: []
      };
    }
    const request = requests[0]!;
    const redirect = this.fragmentationRedirect(metadata, requestedMode, {
      family: "style_copy",
      workbookContextId: metadata.workbookContextId,
      sourceSheetName: request.sourceSheetName,
      sourceAddress: request.sourceAddress,
      targetSheetName: request.targetSheetName,
      targetAddress: request.targetAddress,
      request: input.request,
      dimensions: request.dimensions
    });
    if (redirect) return redirect;
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "style.copy_dimensions", request },
      changes: [{ sheetName: request.targetSheetName, ...(request.targetAddress ? { range: request.targetAddress } : {}), after: `copied style dimensions from ${request.sourceSheetName}` }],
      summary: `Prepared style copy from ${request.sourceSheetName}${request.sourceAddress ? `!${request.sourceAddress}` : ""} to ${request.targetSheetName}${request.targetAddress ? `!${request.targetAddress}` : ""}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "style_copy_preview", request },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: [{ sheetName: request.targetSheetName, range: request.targetAddress ?? usedRangeForSheet(metadata, request.targetSheetName), label: "style copy target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewReplaceStyledTable(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
    const plan = replaceStyledTablePlanFromInput(metadata, input);
    if (!plan) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Replacing a styled table needs target.sheetName plus values.headers and values.row or values.rows.");
    }
    const mismatch = plan.styleCopies.map((request) => styleCopyDimensionIssue(request)).find(Boolean);
    if (mismatch) {
      return workbookLevelNeedsInput(metadata, requestedMode, mismatch);
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "workflow.replace_styled_table", operations: plan.operations, styleCopies: plan.styleCopies },
      changes: plan.changes,
      summary: `Prepared styled table replacement on ${plan.sheetName}!${plan.writeRange}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: `${pending.summary} Apply this preview once; clear, values, style copies, and autofit will run as one ordered workflow.`,
      answer: {
        kind: "replace_range_with_styled_table_preview",
        sheetName: plan.sheetName,
        range: plan.writeRange,
        rowCount: plan.matrix.length,
        columnCount: plan.matrix[0]?.length ?? 0,
        clearCount: plan.clearRanges.length,
        styleCopyCount: plan.styleCopies.length,
        operationCount: plan.operations.length + plan.styleCopies.length
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: [{ sheetName: plan.sheetName, range: plan.writeRange, label: "styled table target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private fragmentationRedirect(metadata: WorkbookMetadata, requestedMode: AgentRunMode, fragment: AgentPreviewFragmentInput): Omit<AgentRunOutput, "telemetry"> | undefined {
    const now = Date.now();
    this.previewFragments.splice(0, this.previewFragments.length, ...this.previewFragments.filter((item) => now - item.createdAt <= AGENT_FRAGMENT_WINDOW_MS));
    const key = previewFragmentKey(fragment);
    const related = this.previewFragments.filter((item) => item.key === key);
    this.previewFragments.push({ ...fragment, key, createdAt: now });
    if (related.length + 1 < AGENT_FRAGMENT_REDIRECT_THRESHOLD) {
      return undefined;
    }
    const fragments = [...related, { ...fragment, key, createdAt: now }];
    const suggested = workflowRedirectSuggestion(fragment, fragments);
    return {
      status: "NEEDS_WORKFLOW_REDIRECT",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: suggested.summary,
      answer: {
        kind: "workflow_redirect",
        reason: "fragmented_operations_detected",
        detectedFamily: fragment.family,
        suggestedIntentAction: suggested.intentAction,
        suggestedRequest: suggested.request,
        suggestedValues: suggested.values
      },
      metrics: {
        operationRisk: fragment.family === "clear_values" ? "destructive" : "safe_format",
        targetFingerprintStatus: "not_applicable",
        workflowKind: suggested.intentAction,
        fragmentationRedirectCount: related.length + 1,
        detectedFamily: fragment.family,
        suggestedWorkflowKind: suggested.intentAction
      },
      proof: fragment.targetSheetName && fragment.targetAddress ? [{ sheetName: fragment.targetSheetName, range: fragment.targetAddress, label: "fragmented operation target" }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "call_preview_update",
      warnings: [suggested.warning]
    };
  }

  private previewStyleRepairConsistency(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = formulaTemplateRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Style repair needs values.templateId and target.sheetName or values.targetSheetName.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "style.repair_consistency", request },
      changes: [{ sheetName: request.targetSheetName, after: `repaired styles from template ${request.templateId}` }],
      summary: `Prepared style consistency repair on ${request.targetSheetName} from template ${request.templateId}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "style_repair_consistency_preview", request },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: [{ sheetName: request.targetSheetName, range: usedRangeForSheet(metadata, request.targetSheetName), label: "style repair target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewTemplateCleanup(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const latestAmbiguous = latestTemplateCandidates(metadata, input);
    if (latestAmbiguous.length > 1) {
      return {
        status: "AMBIGUOUS_TARGET",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Template cleanup needs a concrete source sheet. Multiple monthly sheets could be the latest sheet.",
        candidates: latestAmbiguous,
        proof: latestAmbiguous.flatMap((candidate) => candidate.sheetName && candidate.range ? [{ sheetName: candidate.sheetName, range: candidate.range, label: candidate.label }] : []).slice(0, 5),
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "call_with_target",
        warnings: []
      };
    }
    const sourceSheet = resolveTemplateSourceSheet(metadata, input);
    if (!sourceSheet?.usedRange) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Template cleanup needs a source sheet with a used range.",
        candidates: findAgentCandidates(metadata, input).slice(0, 5),
        proof: [],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "call_with_target",
        warnings: []
      };
    }
    const values = input.values as Record<string, unknown> | undefined;
    const newSheetName = stringValue(values?.newSheetName ?? values?.targetSheetName ?? values?.sheetName) ?? uniqueSheetName(metadata, `${sourceSheet.name} Template`);
    const dataRegions = stringArrayValue(values?.dataRegions).map((region) => normalizeOperationRange(metadata, sourceSheet.name, region));
    const clearRegions = dataRegions.length > 0 ? dataRegions : [sourceSheet.usedRange];
    const operations: ExcelOperation[] = [
      {
        kind: "sheet.copy_clean_data_regions",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "structure",
        reason: input.request,
        sourceSheetName: sourceSheet.name,
        newSheetName,
        dataRegions: clearRegions,
        position: "after",
        relativeToSheetName: sourceSheet.name,
        activate: true
      }
    ];
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
      changes: clearRegions.map((range) => ({ sheetName: newSheetName, range, before: "copied values", after: "cleared data-region values; formatting preserved" })),
      summary: `Prepared template copy ${newSheetName} from ${sourceSheet.name} with ${clearRegions.length} data region(s) cleared.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "template_cleanup_preview", sourceSheetName: sourceSheet.name, newSheetName, dataRegions: clearRegions, operationKind: "sheet.copy_clean_data_regions" },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: [{ sheetName: sourceSheet.name, range: sourceSheet.usedRange, label: "template source" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewTemplateRegister(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = templateCaptureRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Template registration needs values.name and target.sheetName or values.sourceSheetName.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "template.register", request },
      changes: [{ sheetName: request.sourceSheetName, after: `registered template ${request.name}` }],
      summary: `Prepared template registration for ${request.name}.`
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "template_register_preview", { name: request.name, sourceSheetName: request.sourceSheetName, dataRegions: request.dataRegions });
  }

  private previewTemplateUnregister(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const templateId = templateIdFromInput(input);
    if (!templateId) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Template unregister needs values.templateId or target.entity.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "template.unregister", templateId },
      changes: [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: `unregistered template ${templateId}` }],
      summary: `Prepared template unregister for ${templateId}.`
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "template_unregister_preview", { templateId });
  }

  private previewTemplateClearDataRegions(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = templateRepairRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Template data-region clearing needs values.templateId and target.sheetName or values.targetSheetName.");
    }
    const scopedRequest = { ...request, repair: ["dataRegions"] as AddinTemplateRepairRequest["repair"] };
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "template.repair_sheet", request: scopedRequest },
      changes: [{ sheetName: request.targetSheetName, after: `cleared template data regions from ${request.templateId}` }],
      summary: `Prepared template data-region clearing for ${request.targetSheetName}.`
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "template_clear_data_regions_preview", {
      templateId: request.templateId,
      targetSheetName: request.targetSheetName,
      repair: scopedRequest.repair
    });
  }

  private previewTemplateFillRegions(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = templateRegionFillRequestFromInput(metadata, input, this.runtime as Pick<RuntimeService, "getTemplate">);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Template region fill needs values.templateId, target.sheetName or values.targetSheetName, and values.regions or values.regionValues.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations: request.operations },
      changes: request.operations.map((operation) => ({ sheetName: operation.target.sheetName, range: operation.target.address, after: `filled template data region ${operation.target.address}` })),
      summary: `Prepared template region fill for ${request.targetSheetName}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: "template_fill_regions_preview",
        templateId: request.templateId,
        targetSheetName: request.targetSheetName,
        regionCount: request.operations.length
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: request.operations.map((operation) => ({ sheetName: operation.target.sheetName, range: operation.target.address, label: "template data region" })),
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private previewTemplateRepair(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const request = templateRepairRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Template repair needs values.templateId and target.sheetName or values.targetSheetName.");
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "template.repair_sheet", request },
      changes: [{ sheetName: request.targetSheetName, after: `repaired from template ${request.templateId}` }],
      summary: `Prepared template repair for ${request.targetSheetName}.`
    });
    return backupLifecyclePreviewOutput(metadata, requestedMode, pending, "template_repair_preview", { templateId: request.templateId, targetSheetName: request.targetSheetName, repair: request.repair });
  }

  private previewFormulaUpdate(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    matrix: CellMatrix
  ): Omit<AgentRunOutput, "telemetry"> {
    const normalized = formulaMatrixFromInput(input, resolved.range, matrix);
    if (!normalized.ok) {
      return formulaNeedsInputOutput(metadata, requestedMode, resolved, normalized.summary, normalized.suggestedValues, normalized.warnings);
    }
    const formulas = normalized.formulas;
    const formulaCount = formulas.flat().filter((value) => value !== null && value !== undefined && value !== "").length;
    if (formulaCount === 0) {
      return formulaNeedsInputOutput(metadata, requestedMode, resolved, "Formula updates need explicit formula values before OpenWorkbook can preview the update.");
    }
    const operations: ExcelOperation[] = [];
    const formulaOperation: ExcelOperation = {
      kind: "range.write_formulas",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      formulas: formulas as Array<Array<string | null>>,
      preserveFormats: true
    };
    operations.push(formulaOperation);
    const style = styleFromInput(input);
    const hasStyle = Object.keys(style).length > 0;
    if (hasStyle) {
      operations.push({
        kind: "range.write_styles",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "format",
        reason: input.request,
        target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
        style,
        preserveValues: true
      });
    }
    const changes: NonNullable<AgentRunOutput["changes"]> = formulas.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
      sheetName: resolved.sheetName,
      cell: cellAddressFor(resolved.range, rowIndex, columnIndex),
      range: resolved.range,
      after: value
    }))).filter((change) => typeof change.after === "string" && change.after.length > 0);
    if (hasStyle) {
      changes.unshift({ sheetName: resolved.sheetName, range: resolved.range, after: { style } });
    }
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
      changes,
      summary: `Prepared ${formulaCount} formula update(s) on ${resolved.sheetName}!${resolved.range}${hasStyle ? " with formatting" : ""}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: "formula_update_preview",
        sheetName: resolved.sheetName,
        range: resolved.range,
        target: { sheetName: resolved.sheetName, range: resolved.range },
        formulaCount,
        formulaPattern: normalized.formulaPattern,
        scannedRows: normalized.dimensions.rows,
        changedCells: formulaCount,
        examples: formulaPreviewExamples(resolved.range, formulas),
        operationCount: operations.length,
        ...(hasStyle ? { style } : {})
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      agentInstruction: "Apply this formula preview with apply_update, or answer with the preview details. Do not fetch full rows, chunk-read, or split formula writes.",
      maxRecommendedFollowupCalls: 1,
      warnings: normalized.warnings
    };
  }

  private previewTableAppend(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    matrix: CellMatrix
  ): Omit<AgentRunOutput, "telemetry"> {
    if (!hasCellValues(matrix)) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Table append needs row values before OpenWorkbook can preview the update.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: resolved.candidate.label }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
      };
    }
    if (containsFormulaLikeValue(matrix)) {
      return {
        status: "VALIDATION_FAILED",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Preview blocked formula-like values in a table append.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: resolved.candidate.label }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "manual_review",
        warnings: ["Use a formula-aware workflow for values that start with '='."]
      };
    }
    const tableName = resolved.candidate.tableName ?? resolved.candidate.label;
    const table = tableFromResolution(metadata, resolved);
    const expectedColumnCount = table?.columns.length ?? 0;
    if (expectedColumnCount > 0) {
      const invalidRowIndex = matrix.findIndex((row) => row.length !== expectedColumnCount);
      if (invalidRowIndex >= 0) {
        return {
          status: "VALIDATION_FAILED",
          mode: requestedMode,
          workbookContextId: metadata.workbookContextId,
          summary: `Table append row ${invalidRowIndex + 1} has ${matrix[invalidRowIndex]?.length ?? 0} value(s), but ${tableName} has ${expectedColumnCount} column(s).`,
          proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: tableName }],
          resourceLinks: [contextResource(metadata.workbookContextId)],
          nextAction: "ask_user",
          warnings: ["Provide one value per table column or explicitly confirm which missing cells should be blank."]
        };
      }
    }
    const request: TableAppendRowsRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName,
      values: matrix
    };
    const changes = matrix.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
      sheetName: resolved.sheetName,
      range: resolved.range,
      columnName: String(columnIndex + 1),
      before: undefined,
      after: value,
      cell: `table:${tableName}:newRow${rowIndex + 1}:col${columnIndex + 1}`
    })));
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "table.append_rows", request },
      changes,
      summary: `Prepared ${matrix.length} row append(s) to table ${tableName}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: {
        kind: "table_append_preview",
        source: "preview",
        tableName,
        sheetName: resolved.sheetName,
        rowCount: matrix.length,
        columnCount: matrix.reduce((max, row) => Math.max(max, row.length), 0)
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: tableName }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
    };
  }

  private operationStatus(input: AgentRunInput): Omit<AgentRunOutput, "telemetry"> {
    const operationId = input.operationId ? String(input.operationId) : "";
    const pending = this.operations.get(operationId);
    if (!pending) {
      return { status: "NOT_FOUND", mode: "operation_status", summary: "No pending or terminal operation was found for the supplied operationId.", proof: [], resourceLinks: [], nextAction: "ask_user", warnings: [] };
    }
    const mismatchWarning = operationWorkflowMismatchWarning(input, pending.workflowKind);
    return {
      status: pending.applyStatus === "applying" ? "IN_PROGRESS" : "SUCCESS",
      mode: "operation_status",
      workbookContextId: pending.workbookContextId,
      operationId: pending.operationId,
      summary: mismatchWarning ?? `Operation ${pending.operationId} is ${pending.applyStatus ?? "previewed"}.`,
      answer: this.getOperationResource(String(pending.operationId)),
      changes: pending.changes,
      proof: [],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: pending.applyStatus === "previewed" ? "call_apply_update" : "answer_now",
      warnings: mismatchWarning ? [mismatchWarning] : []
    };
  }

  private cancelOperation(input: AgentRunInput): Omit<AgentRunOutput, "telemetry"> {
    const operationId = input.operationId ? String(input.operationId) : "";
    const pending = this.operations.get(operationId);
    if (!pending) {
      return { status: "NOT_FOUND", mode: "cancel_operation", summary: "No previewed operation was found for the supplied operationId.", proof: [], resourceLinks: [], nextAction: "ask_user", warnings: [] };
    }
    if (pending.applyStatus !== "previewed") {
      return {
        status: pending.applyStatus === "applying" ? "IN_PROGRESS" : "VALIDATION_FAILED",
        mode: "cancel_operation",
        workbookContextId: pending.workbookContextId,
        operationId: pending.operationId,
        summary: `Operation ${pending.operationId} cannot be cancelled because it is ${pending.applyStatus ?? "not previewed"}.`,
        answer: this.getOperationResource(String(pending.operationId)),
        changes: pending.changes,
        proof: [],
        resourceLinks: [operationResource(String(pending.operationId))],
        nextAction: pending.applyStatus === "applying" ? "fetch_resource" : "answer_now",
        warnings: ["Only previewed operations can be cancelled."]
      };
    }
    const cancelled = this.operations.cancel(operationId);
    return {
      status: "SUCCESS",
      mode: "cancel_operation",
      workbookContextId: pending.workbookContextId,
      operationId: pending.operationId,
      summary: `Cancelled previewed operation ${pending.operationId}.`,
      answer: {
        kind: "agent_operation_cancelled",
        operationId: pending.operationId,
        workbookContextId: pending.workbookContextId,
        cancelled: Boolean(cancelled)
      },
      changes: pending.changes,
      proof: [],
      resourceLinks: [],
      nextAction: "answer_now",
      warnings: []
    };
  }

  private async applyUpdate(input: AgentRunInput): Promise<Omit<AgentRunOutput, "telemetry">> {
    const operationId = input.operationId ? String(input.operationId) : "";
    const pending = this.operations.get(operationId);
    if (!pending) {
      return { status: "NOT_FOUND", mode: "apply_update", summary: "No pending previewed operation was found for the supplied operationId.", proof: [], resourceLinks: [], nextAction: "ask_user", warnings: [] };
    }
    if (pending.terminalOutput) {
      return pending.terminalOutput;
    }
    if (pending.applyStatus === "applying") {
      return {
        status: "IN_PROGRESS",
        mode: "apply_update",
        workbookContextId: pending.workbookContextId,
        operationId: pending.operationId,
        summary: "Workbook update is still applying. Retry apply_update with the same operationId and confirmationToken to fetch the result.",
        changes: pending.changes,
        proof: [],
        resourceLinks: [operationResource(String(pending.operationId))],
        nextAction: "fetch_resource",
        warnings: []
      };
    }
    if (input.confirmationToken !== pending.confirmationToken) {
      return {
        status: "NEEDS_INPUT",
        mode: "apply_update",
        workbookContextId: pending.workbookContextId,
        operationId: pending.operationId,
        summary: "Apply requires the confirmationToken returned by preview_update.",
        changes: pending.changes,
        proof: [],
        resourceLinks: [operationResource(String(pending.operationId))],
        nextAction: "call_apply_update",
        warnings: []
      };
    }
    const current = await this.metadataBuilder.getOrBuild({ workbookContextId: pending.workbookContextId });
    if (pending.sourceFingerprintHash && current.metadata.fingerprint.structureHash !== pending.sourceFingerprintHash) {
      return {
        status: "STALE_CONTEXT",
        mode: "apply_update",
        workbookContextId: pending.workbookContextId,
        operationId: pending.operationId,
        summary: "Workbook structure changed after preview. Refresh the preview before applying.",
        metrics: { operationRisk: pending.risk, targetFingerprintStatus: "changed" },
        changes: pending.changes,
        proof: [],
        resourceLinks: [contextResource(current.metadata.workbookContextId), operationResource(String(pending.operationId))],
        nextAction: "retry_after_refresh",
        warnings: ["Target fingerprints changed after preview."]
      };
    }
    if (pending.sourceTargetFingerprintHash && current.metadata.fingerprint.structureHash === pending.sourceFingerprintHash) {
      const currentTargetFingerprint = targetFingerprintHash(current.metadata, pending.changes);
      if (currentTargetFingerprint !== pending.sourceTargetFingerprintHash) {
        return {
          status: "STALE_CONTEXT",
          mode: "apply_update",
          workbookContextId: pending.workbookContextId,
          operationId: pending.operationId,
          summary: "Workbook target metadata changed after preview. Refresh the preview before applying.",
          metrics: { operationRisk: pending.risk, targetFingerprintStatus: "changed" },
          changes: pending.changes,
          proof: [],
          resourceLinks: [contextResource(current.metadata.workbookContextId), operationResource(String(pending.operationId))],
          nextAction: "retry_after_refresh",
          warnings: ["Target fingerprints changed after preview."]
        };
      }
    }
    this.operations.markApplying(operationId);
    let result: unknown;
    try {
      result = await this.applyPendingAction(pending, operationId);
    } catch (error) {
      const output: Omit<AgentRunOutput, "telemetry"> = {
        status: "VALIDATION_FAILED",
        mode: "apply_update",
        workbookContextId: pending.workbookContextId,
        operationId: pending.operationId,
        summary: "Workbook update failed.",
        answer: { kind: "apply_update_result", ok: false, validationOk: false, validationIssueCount: 1, changeCount: pending.changes.length, operationRisk: pending.risk, updateRisk: pending.updateRisk },
        metrics: { operationRisk: pending.risk, updateRisk: pending.updateRisk, targetFingerprintStatus: "matched" },
        changes: pending.changes,
        proof: [],
        resourceLinks: [operationResource(String(pending.operationId))],
        nextAction: "manual_review",
        warnings: [error instanceof Error ? error.message : String(error)]
      };
      this.operations.markCompleted(operationId, output);
      return output;
    }
    const applyFailed = (result as { ok?: boolean }).ok === false;
    const validation = !applyFailed ? await this.runtime.validateWorkbook({ workbookId: pending.workbookId }) : undefined;
    const issueCount = validation?.issues?.length ?? 0;
    const validationFailed = validation?.ok === false;
    const resultRecord = result as { transactionId?: string; backups?: string[]; rollbackAvailable?: boolean; telemetry?: unknown; warnings?: unknown[]; results?: unknown[]; error?: unknown; formulaPreservation?: unknown };
    const resultWarnings = Array.isArray(resultRecord.warnings) ? resultRecord.warnings.map(operationWarningMessage) : [];
    const errorWarning = applyErrorMessage(resultRecord.error);
    const cacheImpact = !applyFailed ? this.recordApplyCacheImpact(pending) : undefined;
    const invalidated = applyFailed ? { invalidatedContextIds: [] as string[], invalidatedResourceUris: [] as string[] } : this.applyContextInvalidation(pending);
    const operationJournalRef = cacheImpact ? {
      workbookContextId: cacheImpact.journalEntry.workbookContextId,
      operationId: cacheImpact.journalEntry.operationId,
      contextVersion: cacheImpact.journalEntry.contextVersion,
      appliedAt: cacheImpact.journalEntry.appliedAt
    } : undefined;
    const permissionFollowup = this.structuralPermissionFollowup(resultWarnings, pending);
    const output: Omit<AgentRunOutput, "telemetry"> = {
      status: applyFailed || validationFailed ? "VALIDATION_FAILED" : "SUCCESS",
      mode: "apply_update",
      workbookContextId: pending.workbookContextId,
      operationId: pending.operationId,
      summary: applyFailed
        ? "Workbook update failed."
        : validationFailed
          ? "Workbook update applied, but workbook validation reported issues."
          : pending.summary.replace("Prepared", "Applied"),
      answer: {
        kind: "apply_update_result",
        ok: !applyFailed,
        validationOk: !validationFailed,
        validationIssueCount: issueCount,
        changeCount: pending.changes.length,
        transactionId: resultRecord.transactionId,
        backupIds: resultRecord.backups ?? [],
        rollbackAvailable: resultRecord.rollbackAvailable ?? false,
        partialFailure: applyFailed && Array.isArray(resultRecord.results) && resultRecord.results.some((step) => Boolean(step) && typeof step === "object" && (step as { ok?: unknown }).ok !== false),
        operationRisk: pending.risk,
        updateRisk: pending.updateRisk,
        cacheImpact,
        telemetry: resultRecord.telemetry,
        ...(resultRecord.formulaPreservation !== undefined ? { formulaPreservation: resultRecord.formulaPreservation } : {}),
        ...(applyFailed && resultRecord.results !== undefined ? { stepResults: resultRecord.results } : {})
      },
      metrics: { operationRisk: pending.risk, updateRisk: pending.updateRisk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      ...(operationJournalRef ? { operationJournalRef } : {}),
      proof: pending.changes.flatMap((change) => change.range ? [{ sheetName: change.sheetName, range: change.range, label: "applied target" }] : []).slice(0, 1),
      resourceLinks: resultRecord.transactionId ? [{ uri: `excel://transactions/${resultRecord.transactionId}`, name: "transaction", description: "Applied workbook transaction.", mimeType: "application/json" }] : [],
      invalidatedContextIds: invalidated.invalidatedContextIds,
      invalidatedResourceUris: invalidated.invalidatedResourceUris,
      nextAction: permissionFollowup ? "ask_user" : applyFailed || validationFailed ? "manual_review" : "answer_now",
      ...(permissionFollowup ? {
        requiredFollowup: permissionFollowup,
        taskOutcome: "needs_user_input" as const,
        agentInstruction: "Ask the user for approval if needed, then call excel.agent.run with intent.action set_permissions and the provided permission values. After permission succeeds, create a fresh preview for the original workflow; do not retry this stale failed operationId.",
        finalAnswer: "This structural update needs workbook structure permission before it can be applied."
      } : {}),
      warnings: [
        ...resultWarnings,
        ...(errorWarning && !resultWarnings.includes(errorWarning) ? [errorWarning] : []),
        ...(permissionFollowup ? ["Enable structure permission with intent.action set_permissions, then create a fresh preview before applying."] : []),
        ...(validation?.issues?.slice(0, 5).map((issue) => issue.message) ?? [])
      ]
    };
    this.operations.markCompleted(operationId, output);
    return output;
  }

  private applyPendingAction(pending: NonNullable<ReturnType<AgentOperationStore["get"]>>, operationId: string) {
    return this.runtime.runWithAgentExecutionContext(
      pending.agentId ? { agentId: pending.agentId, ...(pending.agentName !== undefined ? { agentName: pending.agentName } : {}), clientType: "mcp" } : undefined,
      () => this.applyPendingActionInContext(pending, operationId)
    );
  }

  private structuralPermissionFollowup(
    warnings: string[],
    pending: NonNullable<ReturnType<AgentOperationStore["get"]>>
  ): AgentRequiredFollowup | undefined {
    const warningText = warnings.join(" ");
    const blockedByPermissions = /\b(DESTRUCTIVE_ACTION_BLOCKED|PERMISSION_DENIED)\b/.test(warningText)
      || warningText.includes("Structure and workbook actions are disabled");
    if (!blockedByPermissions || (pending.risk !== "structure_change" && pending.risk !== "destructive")) {
      return undefined;
    }
    return {
      mode: "answer",
      nextAction: "answer_now",
      instruction: "Call excel.agent.run with intent.action set_permissions and values.permissions {\"allowWrites\":true,\"allowDestructiveActions\":true,\"scopeToWorkbook\":true,\"requireConfirmationFor\":[]}. After it succeeds, create a fresh preview for the original structural workflow and apply that fresh preview."
    };
  }

  private applyPendingActionInContext(pending: NonNullable<ReturnType<AgentOperationStore["get"]>>, operationId: string) {
    switch (pending.action.kind) {
      case "batch":
        return this.runtime.applyBatch({ workbookId: pending.workbookId, operations: pending.action.operations, mode: "apply", idempotencyKey: `agent:${operationId}` });
      case "table.append_rows":
        return this.runtime.appendTableRows(pending.action.request);
      case "table.update_rows":
        return this.runtime.updateTableRows(pending.action.request);
      case "table.create":
        return this.runtime.createTable(pending.action.request);
      case "table.resize":
        return this.runtime.resizeTable(pending.action.request);
      case "table.reorder_columns":
        return this.runtime.reorderTableColumns(pending.action.request);
      case "table.clear_data_keep_formulas":
        return this.runtime.clearTableDataKeepFormulas(pending.action.request);
      case "table.clear_filters":
        return this.runtime.clearTableFilters(pending.action.request);
      case "table.apply_filters":
        return this.runtime.applyTableFilters(pending.action.request);
      case "table.sort":
        return this.runtime.sortTable(pending.action.request);
      case "table.apply_view":
        return this.runtime.applyTableView(pending.action.request);
      case "table.set_total_row":
        return this.runtime.setTableTotalRow(pending.action.request);
      case "table.set_style":
        return this.runtime.setTableStyle(pending.action.request);
      case "table.copy_structure":
        return this.runtime.copyTableStructure(pending.action.request);
      case "template.register":
        return this.runtime.registerTemplate(pending.action.request);
      case "template.unregister":
        return this.runtime.unregisterTemplate(pending.action.templateId);
      case "template.repair_sheet":
        return this.runtime.repairSheetFromTemplate(pending.action.request);
      case "style.copy_dimensions":
        return this.runtime.copyStyleDimensions(pending.action.request, { idempotencyKey: `agent:${operationId}:style:0` });
      case "style.copy_dimensions_many":
        return this.applyStyleCopyRequests(pending.action.requests, operationId);
      case "workflow.replace_styled_table":
        return this.applyReplaceStyledTableWorkflow(pending.workbookId, pending.action.operations, pending.action.styleCopies, operationId);
      case "style.repair_consistency":
        return this.runtime.repairStyleFromTemplate(pending.action.request);
      case "visual_readability.apply":
        return this.applyVisualReadabilityPlan(pending, operationId);
      case "clean.transform":
        return this.applyCleanMutation(pending.action.action, pending.action.request);
      case "clean.transform_many":
        return this.applyCleanMutations(pending.action.action, pending.action.requests);
      case "workbook.snapshot":
        return this.runtime.createWorkbookSnapshot(pending.action.request);
      case "workbook.create_backup":
        return this.runtime.createWorkbookBackup(pending.action.request);
      case "snapshot.refresh":
        return this.runtime.refreshSnapshot({ snapshotId: pending.action.snapshotId, reason: pending.summary });
      case "snapshot.invalidate":
        return this.runtime.invalidateSnapshot(pending.action.snapshotId);
      case "snapshot.delete":
        return this.runtime.deleteSnapshot(pending.action.snapshotId);
      case "backup.create_file":
        return this.runtime.createFileBackup(pending.action.request);
      case "backup.restore_file":
        return this.runtime.restoreFileBackup({
          ...pending.action.request,
          confirmationToken: pending.confirmationToken
        });
      case "backup.prune":
        return this.runtime.pruneFileBackups(pending.action.request);
      case "backup.pin":
        return this.runtime.pinFileBackup(pending.action.backupId, true);
      case "backup.unpin":
        return this.runtime.pinFileBackup(pending.action.backupId, false);
      case "backup.delete":
        return this.runtime.deleteFileBackup(pending.action.backupId);
      case "workbook.restore_backup":
        return this.runtime.restoreWorkbookBackup(pending.action.backupId, pending.confirmationToken);
      case "workbook.import_local_config":
        return this.runtime.importWorkbookLocalConfig(pending.action.request);
      case "workbook.embed_local_config":
        return this.runtime.embedWorkbookLocalConfig(
          pending.action.workbookId,
          pending.action.includePermissions !== undefined ? { includePermissions: pending.action.includePermissions } : {}
        );
      case "workbook.import_embedded_local_config":
        return this.runtime.importWorkbookEmbeddedLocalConfig(pending.action.request);
      case "workbook.close":
        return this.runtime.closeWorkbook(pending.action.workbookId, pending.action.closeBehavior);
      case "formula.copy_patterns":
        return this.runtime.copyFormulaPatterns(pending.action.request);
      case "formula.fill_pattern":
        return this.runtime.fillFormulaPattern(pending.action.request);
      case "formula.repair_patterns":
        return this.runtime.repairFormulasFromTemplate(pending.action.request);
      case "formula.convert_to_values":
        return this.runtime.convertFormulasToValues(pending.action.request);
      case "names.create":
        return this.runtime.createName(pending.action.request);
      case "names.update":
        return this.runtime.updateName(pending.action.request);
      case "names.delete":
        return this.runtime.deleteName(pending.action.request);
      case "region.register":
        return this.runtime.registerRegion(pending.action.request);
      case "region.clear_values":
        return this.runtime.clearRegionValues(pending.action.request);
      case "region.write_values":
        return this.runtime.writeRegionValues(pending.action.request);
      case "region.fill":
        return this.runtime.fillRegion(pending.action.request);
    }
  }

  private async applyVisualReadabilityPlan(pending: NonNullable<ReturnType<AgentOperationStore["get"]>>, operationId: string) {
    if (pending.action.kind !== "visual_readability.apply") {
      return { ok: false, warnings: ["Internal visual readability apply received the wrong operation kind."] };
    }
    const { request, operations } = pending.action;
    const before = await this.readVisualFormulaSnapshot(request.workbookId, request.sheetName, request.formulaRanges, `${operationId}:before`);
    if (before.ok === false) {
      return before.result;
    }
    const applied = await this.runtime.applyBatch({ workbookId: request.workbookId, operations, mode: "apply", idempotencyKey: `agent:${operationId}:visual_readability` });
    const appliedRecord = applied && typeof applied === "object" ? applied as unknown as Record<string, unknown> : {};
    if (appliedRecord.ok === false) {
      return applied;
    }
    const after = await this.readVisualFormulaSnapshot(request.workbookId, request.sheetName, request.formulaRanges, `${operationId}:after`);
    if (after.ok === false) {
      return after.result;
    }
    const formulaDiff = compareVisualFormulaSnapshots(before.snapshot, after.snapshot);
    const appliedWarnings = Array.isArray(appliedRecord.warnings) ? appliedRecord.warnings.filter((warning): warning is string => typeof warning === "string") : [];
    const formulaWarning = formulaDiff.changedCount > 0
      ? `Formula preservation failed: ${formulaDiff.changedCount} formula cell(s) changed during visual readability apply.`
      : undefined;
    return {
      ...appliedRecord,
      ok: appliedRecord.ok !== false && formulaDiff.changedCount === 0,
      warnings: formulaWarning ? [...appliedWarnings, formulaWarning] : appliedWarnings,
      formulaPreservation: {
        checkedRanges: request.formulaRanges,
        formulasChecked: formulaDiff.checkedCount,
        formulasChanged: formulaDiff.changedCount,
        unchanged: formulaDiff.changedCount === 0
      },
      telemetry: {
        ...(appliedRecord.telemetry && typeof appliedRecord.telemetry === "object" ? appliedRecord.telemetry as Record<string, unknown> : {}),
        visualReadabilityApply: true,
        visualReadabilityRuleCount: request.ruleCount,
        visualReadabilitySkippedRuleCount: request.skippedRuleCount,
        formulasChecked: formulaDiff.checkedCount,
        formulasChanged: formulaDiff.changedCount
      }
    };
  }

  private async readVisualFormulaSnapshot(workbookId: WorkbookId, sheetName: string, formulaRanges: string[], idempotencyKey: string): Promise<
    | { ok: true; snapshot: Map<string, string> }
    | { ok: false; result: { ok: false; warnings: string[]; telemetry: Record<string, unknown> } }
  > {
    if (formulaRanges.length === 0) {
      return { ok: true, snapshot: new Map() };
    }
    const operations: ExcelOperation[] = formulaRanges.map((address) => ({
      kind: "range.read_full",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "none",
      reason: "Verify visual readability formula preservation.",
      target: { workbookId, sheetName, address },
      facets: ["formulas"],
      includeFormulas: true
    }));
    const result = await this.runtime.applyBatch({ workbookId, operations, mode: "validate", idempotencyKey });
    const record = result && typeof result === "object" ? result as unknown as Record<string, unknown> : {};
    if (record.ok === false) {
      return {
        ok: false,
        result: {
          ok: false,
          warnings: ["Could not read formulas for visual readability preservation check."],
          telemetry: { visualReadabilityApply: true, formulaPreservationReadFailed: true }
        }
      };
    }
    return { ok: true, snapshot: visualFormulaSnapshotFromBatchResult(formulaRanges, result) };
  }

  private async applyStyleCopyRequests(requests: StyleCopyRequest[], operationId: string) {
    if (requests.length === 0) {
      return { ok: true, warnings: [], telemetry: { styleCopyCount: 0 } };
    }
    const runtime = this.runtime as RuntimeService & {
      copyStyleDimensionsMany?: (input: { workbookId: WorkbookId; requests: StyleCopyRequest[] }, options?: { idempotencyKey?: string }) => Promise<unknown>;
    };
    if (typeof runtime.copyStyleDimensionsMany === "function") {
      return runtime.copyStyleDimensionsMany({ workbookId: requests[0]!.workbookId, requests }, { idempotencyKey: `agent:${operationId}:style:many` });
    }
    const results = [];
    for (const [index, request] of requests.entries()) {
      results.push(await this.runtime.copyStyleDimensions(request, { idempotencyKey: `agent:${operationId}:style:${index}` }));
    }
    return combineApplyResults(results);
  }

  private async applyReplaceStyledTableWorkflow(workbookId: WorkbookId, operations: ExcelOperation[], styleCopies: StyleCopyRequest[], operationId: string) {
    const results = [];
    if (operations.length > 0) {
      results.push(await this.runtime.applyBatch({ workbookId, operations, mode: "apply", idempotencyKey: `agent:${operationId}:values` }));
    }
    if (styleCopies.length > 0) {
      results.push(await this.applyStyleCopyRequests(styleCopies, operationId));
    }
    return combineApplyResults(results);
  }

  private applyCleanMutation(action: AgentCleanMutationAction, request: AgentCleanRequest) {
    switch (action) {
      case "normalize_headers":
        return this.runtime.cleanNormalizeHeaders(request);
      case "trim_whitespace":
        return this.runtime.cleanTrimWhitespace(request);
      case "remove_duplicates":
        return this.runtime.cleanRemoveDuplicates(request);
      case "parse_dates":
        return this.runtime.cleanParseDates(request);
      case "parse_numbers":
        return this.runtime.cleanParseNumbers(request);
      case "standardize_currency":
        return this.runtime.cleanStandardizeCurrency(request);
      case "fill_missing_values":
        return this.runtime.cleanFillMissingValues(request);
      case "split_column":
        return this.runtime.cleanSplitColumn(request as AgentCleanRequest & { columnIndex: number; targetAddress: string });
      case "merge_columns":
        return this.runtime.cleanMergeColumns(request as AgentCleanRequest & { columnIndexes: number[]; targetAddress: string });
    }
  }

  private async applyCleanMutations(action: AgentCleanMutationAction, requests: AgentCleanRequest[]) {
    const results = [];
    for (const request of requests) {
      results.push(await this.applyCleanMutation(action, request));
    }
    return combineApplyResults(results);
  }

  private async rollback(input: AgentRunInput): Promise<Omit<AgentRunOutput, "telemetry">> {
    const transactionId = input.target?.entity?.startsWith("tx_") ? input.target.entity : input.transactionId ?? input.operationId;
    const backupId = input.target?.entity?.startsWith("backup_") ? input.target.entity : undefined;
    const result = transactionId
      ? await this.runtime.rollbackTransaction(transactionId as TransactionId, input.confirmationToken)
      : backupId
        ? await this.runtime.restoreBackup(backupId as BackupId, input.confirmationToken)
        : undefined;
    if (!result) {
      return { status: "NEEDS_INPUT", mode: "rollback", summary: "Rollback requires an operationId/transactionId or a backup identifier.", proof: [], resourceLinks: [], nextAction: "ask_user", warnings: [] };
    }
    if (result.ok) {
      this.invalidateWorkbook(result.diffSummary?.changedRanges?.[0]?.workbookId ?? "");
    }
    const workbookId = result.diffSummary?.changedRanges?.[0]?.workbookId;
    const validation = workbookId ? await this.runtime.validateWorkbook({ workbookId }) : undefined;
    return {
      status: result.ok === false || validation?.ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: "rollback",
      summary: result.ok === false ? "Rollback failed." : "Rollback completed and validation was attempted.",
      answer: { result, validation },
      proof: result.diffSummary?.changedRanges?.map((range) => ({ sheetName: range.sheetName, range: range.address, label: "rollback scope" })) ?? [],
      resourceLinks: [],
      nextAction: result.ok === false || validation?.ok === false ? "manual_review" : "answer_now",
      warnings: validation?.issues?.slice(0, 5).map((issue) => issue.message) ?? []
    };
  }

  private async readAndProfileRange(workbookId: WorkbookId, sheetName: string, address: string, runMetrics: AgentRunMetrics) {
    const values = await this.readRangeValues(workbookId, sheetName, address, runMetrics);
    return profileValues(values, address);
  }

  private async readRangeSnapshot(
    workbookId: WorkbookId,
    sheetName: string,
    address: string,
    facets: NonNullable<Extract<ExcelOperation, { kind: "range.read_full" }>["facets"]>,
    runMetrics?: AgentRunMetrics
  ): Promise<RangeSnapshot | undefined> {
    const operation: ExcelOperation = {
      kind: "range.read_full",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "none",
      reason: "Agent exact range diagnostics read",
      target: { workbookId, sheetName, address },
      facets
    };
    const result = await this.runtime.applyBatch({ workbookId, mode: "apply", operations: [operation] });
    const snapshot = operationReadSnapshots(result)[0]?.snapshot;
    if (runMetrics) {
      runMetrics.internalReadCount += 1;
      runMetrics.fullReadCellCount += matrixCellCount(snapshot?.values ?? snapshot?.text ?? []);
      runMetrics.fullReadUsed = true;
    }
    return snapshot as RangeSnapshot | undefined;
  }

  private async readAndProfileRanges(workbookId: WorkbookId, targets: Array<{ sheetName: string; range: string }>, runMetrics: AgentRunMetrics) {
    const operations: ExcelOperation[] = targets.map((target) => ({
      kind: "range.read_full",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "none",
      reason: "Agent comparison read",
      target: { workbookId, sheetName: target.sheetName, address: target.range },
      facets: ["values", "text"]
    }));
    const result = await this.runtime.applyBatch({ workbookId, mode: "apply", operations });
    const readData = operationReadSnapshots(result);
    return targets.map((target, index) => {
      const values = readData[index]?.snapshot?.values ?? [];
      runMetrics.internalReadCount += 1;
      runMetrics.fullReadCellCount += matrixCellCount(values);
      return { ...target, profile: profileValues(values, target.range) };
    });
  }

  private async readRangeValues(workbookId: WorkbookId, sheetName: string, address: string, runMetrics?: AgentRunMetrics): Promise<CellMatrix> {
    const requestedCells = cellCountFromAddress(address);
    if (requestedCells !== undefined && requestedCells > 10000) {
      if (runMetrics) {
        runMetrics.internalReadCount += 1;
      }
      return [];
    }
    const snapshot = await this.runtime.snapshotRanges(workbookId, [{ workbookId, sheetName, address }]);
    const values = (snapshot as { rangeSnapshots?: Array<{ values?: CellMatrix }> }).rangeSnapshots?.[0]?.values ?? [];
    if (runMetrics) {
      runMetrics.internalReadCount += 1;
      runMetrics.fullReadCellCount += matrixCellCount(values);
    }
    return values;
  }

  private async readColumnSnapshot(workbookId: WorkbookId, sheetName: string, address: string): Promise<{ values: CellMatrix; formulas: CellMatrix }> {
    const snapshot = await this.runtime.snapshotRanges(workbookId, [{ workbookId, sheetName, address }]);
    const first = (snapshot as { rangeSnapshots?: Array<{ values?: CellMatrix; formulas?: CellMatrix }> }).rangeSnapshots?.[0];
    return {
      values: first?.values ?? [],
      formulas: first?.formulas ?? []
    };
  }

  private createPendingOperation(
    metadata: WorkbookMetadata,
    input: {
      action: Parameters<AgentOperationStore["create"]>[0]["action"];
      changes: NonNullable<AgentRunOutput["changes"]>;
      summary: string;
      workflowKind?: string;
    }
  ) {
    const risk = classifyAgentActionRisk(input.action);
    const updateRisk = assessAgentUpdateRisk(input.action);
    const agentContext = this.runtime.currentAgentExecutionContext();
    return this.operations.create({
      workbookContextId: metadata.workbookContextId,
      workbookId: metadata.workbook.workbookId as WorkbookId,
      action: input.action,
      ...(input.workflowKind !== undefined ? { workflowKind: input.workflowKind } : {}),
      changes: input.changes,
      summary: input.summary,
      risk,
      updateRisk,
      ...(agentContext?.agentId !== undefined ? { agentId: agentContext.agentId } : {}),
      ...(agentContext?.agentName !== undefined ? { agentName: agentContext.agentName } : {}),
      sourceFingerprintHash: metadata.fingerprint.structureHash,
      sourceTargetFingerprintHash: targetFingerprintHash(metadata, input.changes)
    });
  }
}

function formulaMutationPreviewOutput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  pending: ReturnType<AgentOrchestrator["createPendingOperation"]>,
  kind: string,
  sheetName: string,
  range?: string
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "PREVIEW_READY",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    operationId: pending.operationId,
    confirmationToken: pending.confirmationToken,
    summary: pending.summary,
    answer: { kind, sheetName, ...(range ? { range } : {}) },
    metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
    changes: pending.changes,
    proof: range ? [{ sheetName, range, label: "formula mutation target" }] : [],
    resourceLinks: [operationResource(String(pending.operationId))],
    nextAction: "call_apply_update",
    warnings: []
  };
}

function previewWorkflowKind(answer: unknown): string | undefined {
  if (!answer || typeof answer !== "object" || Array.isArray(answer)) {
    return undefined;
  }
  return stringValue((answer as Record<string, unknown>).kind);
}

function invalidPreviewOperationReuseOutput(input: AgentRunInput): Omit<AgentRunOutput, "telemetry"> {
  const operationId = String(input.operationId ?? "");
  return {
    status: "VALIDATION_FAILED",
    mode: "preview_update",
    ...(operationId ? { operationId } : {}),
    summary: "preview_update cannot reuse an existing operationId. Use operation_status for an existing preview, apply_update to apply it, or call preview_update without operationId to create a fresh preview.",
    answer: {
      kind: "invalid_preview_operation_reuse",
      operationId,
      requestedMode: "preview_update",
      validModes: ["operation_status", "apply_update"]
    },
    proof: [],
    resourceLinks: operationId ? [operationResource(operationId)] : [],
    nextAction: "ask_user",
    warnings: ["Do not continue a new preview workflow with an operationId from a different preview."]
  };
}

function operationWorkflowMismatchWarning(input: AgentRunInput, storedWorkflowKind: string | undefined): string | undefined {
  const requestedWorkflowKind = requestedWorkflowKindFromInput(input);
  if (!requestedWorkflowKind || !storedWorkflowKind || requestedWorkflowKind === storedWorkflowKind) {
    return undefined;
  }
  return `The supplied operationId belongs to ${storedWorkflowKind}, but this request appears to be for ${requestedWorkflowKind}. Create a fresh preview_update for the requested workflow instead of reusing this operationId.`;
}

function requestedWorkflowKindFromInput(input: AgentRunInput): string | undefined {
  const action = intentAction(input);
  if (action === "grouped_header" || shouldPreviewGroupedHeader(input)) {
    return "grouped_header_preview";
  }
  if (action === "improve_visual_readability") {
    return "visual_readability_preview";
  }
  return undefined;
}

function backupLifecyclePreviewOutput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  pending: ReturnType<AgentOrchestrator["createPendingOperation"]>,
  kind: string,
  details: Record<string, unknown>
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "PREVIEW_READY",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    operationId: pending.operationId,
    confirmationToken: pending.confirmationToken,
    summary: pending.summary,
    answer: { kind, ...details },
    metrics: { operationRisk: pending.risk, targetFingerprintStatus: "not_applicable" },
    changes: pending.changes,
    proof: [],
    resourceLinks: [operationResource(String(pending.operationId))],
    nextAction: "call_apply_update",
    warnings: []
  };
}

function templateMetadataOutput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  kind: string,
  summary: string,
  answer: Record<string, unknown>
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    answer: { kind, ...answer },
    metrics: { source: `runtime_${kind}` },
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: []
  };
}

interface AgentRunMetrics {
  internalReadCount: number;
  fullReadCellCount: number;
  fullReadUsed?: boolean;
  safetyFingerprintOnly?: boolean;
  semanticEntryCount?: number;
  route: IntentRoute;
  intent: NormalizedAgentIntent;
  autoApplied?: boolean;
  safetyDecision?: string;
  previewOperationId?: string;
  operationRisk?: AgentOperationRisk;
  actionHandlerId?: AgentActionHandlerId | string;
  autoApplyBlockedReason?: string;
  targetFingerprintStatus?: "matched" | "changed" | "not_applicable";
  metadataFreshnessReason?: string;
  metadataDetailLevel?: "structure" | "sampled";
  validationStatus: "passed" | "failed" | "not_run";
}

interface AgentValuePatch {
  target: AgentRunTarget;
  values: CellMatrix;
  reason?: string;
}

interface AgentSemanticValuePatch {
  sectionId?: string;
  sectionLabel?: string;
  sheetName?: string;
  rowMatch: {
    column?: string;
    value: unknown;
    contains?: boolean;
  };
  columnMatch: string;
  value: unknown;
  reason?: string;
}

type AgentPreviewFragmentFamily =
  | "style_copy"
  | "clear_values"
  | "autofit_columns"
  | "autofit_rows"
  | "write_values"
  | "format_range"
  | "formula_fill_down"
  | "formula_fill_right";

interface AgentPreviewFragmentInput {
  family: AgentPreviewFragmentFamily;
  workbookContextId: string;
  sourceSheetName?: string | undefined;
  sourceAddress?: string | undefined;
  targetSheetName?: string | undefined;
  targetAddress?: string | undefined;
  request: string;
  dimensions?: StyleDimension[] | undefined;
  style?: unknown;
  values?: unknown;
}

interface AgentPreviewFragment extends AgentPreviewFragmentInput {
  key: string;
  createdAt: number;
}

function previewFragmentKey(fragment: AgentPreviewFragmentInput): string {
  return [
    fragment.workbookContextId,
    fragment.family,
    fragment.sourceSheetName ?? "",
    fragment.sourceAddress ?? "",
    fragment.targetSheetName ?? ""
  ].join("|");
}

function workflowRedirectSuggestion(fragment: AgentPreviewFragmentInput, fragments: AgentPreviewFragmentInput[] = [fragment]): {
  intentAction: AgentIntentAction;
  request: string;
  values: Record<string, unknown>;
  summary: string;
  warning: string;
} {
  const groupedTargetRange = groupedFragmentTargetRange(fragments) ?? fragment.targetAddress;
  if (fragment.family === "style_copy") {
    return {
      intentAction: "copy_style_from_template",
      request: "Preview one grouped style copy operation for all target ranges instead of splitting by column chunks.",
      values: {
        styleCopies: fragments.map((entry) => ({
          source: { sheetName: fragment.sourceSheetName, range: fragment.sourceAddress },
          destination: { sheetName: entry.targetSheetName, range: entry.targetAddress },
          ...(entry.dimensions?.length ? { dimensions: entry.dimensions } : fragment.dimensions?.length ? { dimensions: fragment.dimensions } : {})
        }))
      },
      summary: "Repeated style-copy previews look like one broad formatting task. Use one grouped style-copy preview/apply workflow instead of chunked calls.",
      warning: "Fragmented style-copy previews were redirected to a grouped workflow."
    };
  }
  if (fragment.family === "clear_values") {
    return {
      intentAction: "clear_range",
      request: "Preview one clear-range operation for the stale layout, or use replace_range_with_styled_table when rewriting a table.",
      values: { clearRange: groupedTargetRange },
      summary: "Repeated value clears can leave stale borders and fills. Use clear_range or replace_range_with_styled_table for old layout cleanup.",
      warning: "Fragmented value-only cleanup was redirected because it can leave stale formatting."
    };
  }
  if (fragment.family === "write_values") {
    return {
      intentAction: "write_values",
      request: "Preview one grouped values.patches operation for adjacent value writes instead of splitting ranges.",
      values: {
        patches: fragments.map((entry) => ({
          target: { sheetName: entry.targetSheetName, range: entry.targetAddress },
          values: entry.values ?? []
        }))
      },
      summary: "Repeated value-write previews look like one related write task. Use one grouped values.patches preview/apply workflow.",
      warning: "Fragmented value-write previews were redirected to a grouped patch workflow."
    };
  }
  if (fragment.family === "format_range") {
    return {
      intentAction: "write_styles_many",
      request: "Preview one grouped style write operation for adjacent format ranges instead of splitting format calls.",
      values: {
        entries: fragments.map((entry) => ({
          target: { sheetName: entry.targetSheetName, range: entry.targetAddress },
          style: entry.style ?? "STYLE_FOR_THIS_RANGE"
        }))
      },
      summary: "Repeated format previews look like one related styling task. Use one write_styles_many preview/apply workflow.",
      warning: "Fragmented format previews were redirected to a grouped style workflow."
    };
  }
  if (fragment.family === "formula_fill_down" || fragment.family === "formula_fill_right") {
    const direction = fragment.family === "formula_fill_down" ? "down" : "right";
    return {
      intentAction: direction === "down" ? "fill_formula_down" : "fill_formula_right",
      request: `Preview one formula fill ${direction} operation for the full affected range instead of splitting adjacent fills.`,
      values: {
        source: { sheetName: fragment.sourceSheetName, range: fragment.sourceAddress },
        destination: { sheetName: fragment.targetSheetName, range: groupedTargetRange }
      },
      summary: `Repeated formula-fill-${direction} previews look like one formula fill task. Use one full-range formula fill preview/apply workflow.`,
      warning: `Fragmented formula-fill-${direction} previews were redirected to a grouped formula workflow.`
    };
  }
  return {
    intentAction: "autofit",
    request: "Preview one autofit operation for the full affected range instead of splitting adjacent ranges.",
    values: { targetRange: groupedTargetRange },
    summary: "Repeated autofit previews look like one broad layout task. Use one full-range autofit preview/apply workflow.",
    warning: "Fragmented autofit previews were redirected to a grouped workflow."
  };
}

function groupedFragmentTargetRange(fragments: AgentPreviewFragmentInput[]): string | undefined {
  const parsed = fragments
    .map((fragment) => fragment.targetAddress ? rangeShape(fragment.targetAddress) : undefined)
    .filter((shape): shape is NonNullable<ReturnType<typeof rangeShape>> => Boolean(shape));
  if (parsed.length === 0) {
    return undefined;
  }
  const startRow = Math.min(...parsed.map((shape) => shape.startRow));
  const startColumn = Math.min(...parsed.map((shape) => shape.startColumn));
  const endRow = Math.max(...parsed.map((shape) => shape.endRow));
  const endColumn = Math.max(...parsed.map((shape) => shape.endColumn));
  return addressFromBounds(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
}

function shouldTrackFragmentedValueWrite(request: string): boolean {
  return /\b(extract(?:ed)?|ocr|field|field\/value|booking|invoice|shipment|manifest|chunk|part|first|second|next|adjacent|split)\b/i.test(request);
}

function workbookOverviewIntent(input: AgentRunInput): ReturnType<typeof detectWorkbookOverviewIntent> {
  if (input.target?.candidateId || input.target?.sheetName || input.target?.tableName || input.target?.range) {
    return emptyWorkbookOverviewIntent();
  }
  return detectWorkbookOverviewIntent(input.request);
}

function operationReadSnapshots(result: unknown): Array<{ snapshot?: { values?: CellMatrix; text?: string[][] } }> {
  const typed = result as {
    data?: Array<{ snapshot?: { values?: CellMatrix; text?: string[][] } }>;
    readData?: Array<{ snapshot?: { values?: CellMatrix; text?: string[][] } }>;
  };
  return typed.readData ?? typed.data ?? [];
}

function isTargetFingerprintStatus(value: unknown): value is "matched" | "changed" | "not_applicable" {
  return value === "matched" || value === "changed" || value === "not_applicable";
}

function targetFingerprintHash(metadata: WorkbookMetadata, changes: NonNullable<AgentRunOutput["changes"]>): string {
  const targets = changes
    .filter((change) => change.sheetName || change.range)
    .map((change) => ({ sheetName: change.sheetName, range: change.range }))
    .sort((left, right) => `${left.sheetName}!${left.range ?? ""}`.localeCompare(`${right.sheetName}!${right.range ?? ""}`));
  const targetSheets = new Set(targets.map((target) => target.sheetName).filter(Boolean));
  const targetTables = metadata.tables
    .filter((table) => targetSheets.has(table.sheetName) || targets.some((target) => target.range && table.range === target.range))
    .map((table) => ({
      id: table.id,
      name: table.name,
      sheetName: table.sheetName,
      range: table.range,
      columns: table.columns.map((column) => ({ name: column.name, index: column.index, inferredType: column.inferredType }))
    }));
  const targetNamedRanges = metadata.namedRanges
    .filter((name) => !name.sheetName || targetSheets.has(name.sheetName) || targets.some((target) => target.range && name.range === target.range))
    .map((name) => ({ name: name.name, sheetName: name.sheetName, range: name.range }));
  const targetFormulaRegions = metadata.formulaRegions
    .filter((region) => targetSheets.has(region.sheetName) || targets.some((target) => target.range && rangesOverlapAddresses(target.range, region.range)))
    .map((region) => ({ id: region.id, sheetName: region.sheetName, range: region.range }));
  const targetSheetMetadata = metadata.sheets
    .filter((sheet) => targetSheets.has(sheet.name))
    .map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      usedRange: sheet.usedRange,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount
    }));
  return JSON.stringify({
    targets,
    sheets: targetSheetMetadata,
    tables: targetTables,
    namedRanges: targetNamedRanges,
    formulaRegions: targetFormulaRegions,
    selection: metadata.selection ? {
      sheetName: metadata.selection.sheetName,
      address: metadata.selection.address
    } : undefined
  });
}

function detectWorkbookOverviewIntent(requestText: string) {
  const request = requestText.toLowerCase();
  const workbookIntent = /\b(workbook|file|xlsx|excel file)\b/.test(request);
  const tableIntent = /\b(which|list|show|what)\b.*\btables?\b.*\b(workbook|file)\b|\btables?\b.*\b(workbook|file)\b/.test(request);
  const namedRangeIntent = /\bnamed ranges?\b.*\b(workbook|file)\b|\b(workbook|file)\b.*\bnamed ranges?\b/.test(request);
  const blankIntent = /\bblank|empty\b/.test(request);
  const sheetCountIntent = /\bhow many sheets?|sheet count\b/.test(request);
  const sheetListIntent = /\blist\b.*\bsheets?\b|\beach sheet\b/.test(request);
  const fileReviewIntent = /\b(look into|review|inspect|summari[sz]e|analy[sz]e|check)\b/.test(request) && workbookIntent;
  const aboutIntent = (/\bwhat(?:'s| is)? it about\b|\babout\b/.test(request) && workbookIntent) || fileReviewIntent;
  return { aboutIntent, tableIntent, namedRangeIntent, blankIntent, sheetCountIntent, sheetListIntent };
}

function emptyWorkbookOverviewIntent() {
  return { aboutIntent: false, tableIntent: false, namedRangeIntent: false, blankIntent: false, sheetCountIntent: false, sheetListIntent: false };
}

function hasWorkbookOverviewIntent(intent: ReturnType<typeof detectWorkbookOverviewIntent>): boolean {
  return intent.aboutIntent || intent.tableIntent || intent.namedRangeIntent || intent.blankIntent || intent.sheetCountIntent || intent.sheetListIntent;
}

function isWorkbookDumpRequest(requestText: string): boolean {
  const request = requestText.toLowerCase();
  return isEveryCellDumpRequest(requestText)
    || /\b(show|print|dump|list|read|return)\b.*\b(entire|whole|full|all)\b.*\b(workbook|excel file|file)\b/.test(request)
    || /\b(entire|whole|full|all)\b.*\b(workbook|excel file|file)\b.*\b(show|print|dump|list|read|return)\b/.test(request);
}

function isEveryCellDumpRequest(requestText: string): boolean {
  const request = requestText.toLowerCase();
  return /\b(print|dump|show|return|list|read)\b.*\bevery\s+cell\b/.test(request)
    || /\bevery\s+cell\b.*\b(print|dump|show|return|list|read)\b/.test(request)
    || /\ball\s+cells?\b.*\b(all|every)\s+sheets?\b/.test(request);
}

function shouldBuildSampledMetadata(input: AgentRunInput, effectiveMode: AgentRunMode, overviewIntent: ReturnType<typeof detectWorkbookOverviewIntent>, route: IntentRoute): boolean {
  if (input.detailLevel === "workbook_summary" || input.detailLevel === "semantic_index" || input.detailLevel === "sheet_summary") {
    return false;
  }
  if (input.detailLevel === "full_table" && !isExplicitFullDataRequest(input.request)) {
    return false;
  }
  if (input.detailLevel === "table_sample" || input.detailLevel === "full_table") {
    return true;
  }
  if (/\b(sections?|blocks?|areas?)\b/i.test(input.request)) {
    return true;
  }
  if (route.metadataPolicy === "structure_only") {
    return false;
  }
  if (route.metadataPolicy === "sampled_required") {
    return true;
  }
  if (isSheetOverviewRequest(input.request)) {
    return false;
  }
  if (isWorkbookDumpRequest(input.request) || isLargeTargetRangeRequest(input)) {
    return false;
  }
  if (effectiveMode === "prepare") {
    return false;
  }
  if ((effectiveMode === "answer" || effectiveMode === "find") && hasWorkbookOverviewIntent(overviewIntent)) {
    return false;
  }
  return true;
}

function isLargeTargetRangeRequest(input: AgentRunInput): boolean {
  const range = input.target?.range;
  if (!range) {
    return false;
  }
  const requestedCells = cellCountFromAddress(range);
  return requestedCells !== undefined && requestedCells > AGENT_LARGE_RANGE_CELL_LIMIT;
}

async function freezePaneStatusAnswerOutput(
    runtime: RuntimeService,
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    runMetrics: AgentRunMetrics
  ): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
  if (!isFreezePaneStatusQuestion(input)) {
    return undefined;
  }
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  const sheetName = input.target?.sheetName ?? metadata.selection?.sheetName ?? metadata.sheets.find((sheet) => sheet.usedRange)?.name ?? metadata.sheets[0]?.name;
  const sheet = sheetName ? metadata.sheets.find((candidate) => candidate.name === sheetName) : undefined;
  const range = sheet?.usedRange ?? input.target?.range;
  if (!sheetName) {
    return {
      status: "NEEDS_INPUT",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: "Freeze pane status needs a target sheet.",
      proof: [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "ask_user",
      warnings: ["Provide target.sheetName or prepare a workbook context with an active sheet."]
    };
  }
  runMetrics.internalReadCount += 1;
  const result = await runtime.getStyleFingerprint({
    workbookId,
    sheetName,
    ...(range !== undefined ? { address: range } : {}),
    maxCellSamples: 0
  });
  if ((result as { ok?: boolean }).ok === false) {
    return formulaRuntimeErrorOutput(metadata, requestedMode, `Freeze pane status is unavailable for ${sheetName}.`, result);
  }
  const fingerprint = (result as { fingerprint?: unknown }).fingerprint ?? result;
  const freezePanes = freezePanesFromFingerprint(fingerprint);
  const readable = freezePanes.readable !== false;
  const frozen = freezePanes.frozen === true;
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: readable
      ? frozen
        ? freezePaneSummary(sheetName, freezePanes)
        : `No frozen panes are active on ${sheetName}.`
      : "Freeze pane status cannot be read from the current Excel host API path.",
    answer: {
      kind: "freeze_panes_status",
      source: "runtime_style_fingerprint",
      sheetName,
      ...freezePanes,
      canApplyFreezePanes: true
    },
    metrics: { source: "runtime_style_fingerprint" },
    proof: sheetName ? [{ sheetName, range: range ?? "A1", label: "freeze panes status target sheet" }] : [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: readable ? [] : ["Current freeze pane split was not readable through the live Office.js capture path; ask to set or unfreeze panes if you want a deterministic change."]
  };
}

function isFreezePaneStatusQuestion(input: AgentRunInput): boolean {
  const request = freezePaneQuestionText(input).toLowerCase();
  return /\b(which|what|where|show|tell|check|read|inspect|current|currently|is|are|has|have|status)\b/.test(request)
    && /\b(freeze|frozen)\b/.test(request)
    && /\b(panes?|rows?|columns?|cols?|header|top|first)\b/.test(request);
}

function freezePaneQuestionText(input: AgentRunInput): string {
  const intent: Record<string, unknown> = isRecord(input.intent) ? input.intent : {};
  return [
    input.request,
    typeof intent.reason === "string" ? intent.reason : undefined,
    ...(Array.isArray(intent.targetHints) ? intent.targetHints.filter((hint: unknown): hint is string => typeof hint === "string") : [])
  ].filter((part): part is string => typeof part === "string" && part.trim().length > 0).join(" ");
}

function freezePanesFromFingerprint(fingerprint: unknown): Record<string, unknown> {
  const record = isRecord(fingerprint) ? fingerprint : {};
  const dimensions = isRecord(record.dimensions) ? record.dimensions : {};
  const freezePanes = isRecord(dimensions.freezePanes) ? dimensions.freezePanes : {};
  return freezePanesFromDimension(freezePanes);
}

function freezePaneSummary(sheetName: string, freezePanes: Record<string, unknown>): string {
  const columns = typeof freezePanes.columns === "number" ? freezePanes.columns : undefined;
  const rows = typeof freezePanes.rows === "number" ? freezePanes.rows : undefined;
  const lastFrozenColumn = typeof freezePanes.lastFrozenColumn === "string" ? freezePanes.lastFrozenColumn : undefined;
  const firstUnfrozenColumn = typeof freezePanes.firstUnfrozenColumn === "string" ? freezePanes.firstUnfrozenColumn : undefined;
  const parts: string[] = [];
  if (columns !== undefined && columns > 0) {
    parts.push(lastFrozenColumn && firstUnfrozenColumn
      ? `columns A:${lastFrozenColumn} are frozen; first unfrozen column is ${firstUnfrozenColumn}`
      : `${columns} column(s) are frozen`);
  }
  if (rows !== undefined && rows > 0) {
    parts.push(`rows 1:${rows} are frozen; first unfrozen row is ${rows + 1}`);
  }
  return parts.length > 0
    ? `Freeze panes on ${sheetName}: ${parts.join("; ")}.`
    : `Freeze panes are active on ${sheetName}.`;
}

function workbookOverviewAnswer(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  const intent = workbookOverviewIntent(input);
  if (!hasWorkbookOverviewIntent(intent)) {
    return undefined;
  }
  if (shouldDeferWorkbookOverviewToSelection(metadata, input)) {
    return undefined;
  }
  const blankSheets = metadata.sheets.filter((sheet) => !sheet.usedRange || (sheet.rowCount ?? 0) <= 1 || (sheet.columnCount ?? 0) <= 1);
  const answer = {
    kind: "workbook_overview",
    source: "cached_metadata",
    workbook: metadata.workbook,
    ...(metadata.selection ? { selection: metadata.selection } : {}),
    sheetCount: metadata.sheets.length,
    tableCount: metadata.tables.length,
    namedRangeCount: metadata.namedRanges.length,
    sectionCount: metadata.sections.length,
    semanticIndex: buildSemanticWorkbookIndex(metadata, { maxEntries: 12 }),
    sheets: metadata.sheets.map((sheet) => ({
      name: sheet.name,
      kind: sheet.kind,
      usedRange: sheet.usedRange,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      tableCount: sheet.tableIds.length,
      sectionCount: sheet.sectionIds.length
    })),
    tables: metadata.tables.map((table) => ({ name: table.name, sheetName: table.sheetName, range: table.range, columnCount: table.columns.length })),
    namedRanges: metadata.namedRanges,
    blankSheets: blankSheets.map((sheet) => ({ name: sheet.name, usedRange: sheet.usedRange }))
  };
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: workbookOverviewSummary(metadata, intent, blankSheets.length),
    answer,
    metrics: { source: "cached_metadata", sheetCount: metadata.sheets.length, tableCount: metadata.tables.length, namedRangeCount: metadata.namedRanges.length, sectionCount: metadata.sections.length, semanticEntryCount: answer.semanticIndex.entryCount },
    proof: metadata.sheets.slice(0, 5).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "used range" }] : []),
    resourceLinks: [contextResource(metadata.workbookContextId), semanticIndexResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    taskOutcome: "final_answer",
    finalAnswer: workbookOverviewSummary(metadata, intent, blankSheets.length),
    agentInstruction: "This workbook overview is complete for a summary request. Answer now from cached metadata; do not fetch fullResultUri, chunk-read sheets, or call low-level MCP resources unless the user asks for all raw rows or exact cell values.",
    maxRecommendedFollowupCalls: 0,
    warnings: []
  };
}

function shouldDeferWorkbookOverviewToSelection(metadata: WorkbookMetadata, input: AgentRunInput): boolean {
  if (!metadata.selection?.sheetName || !metadata.selection.address) {
    return false;
  }
  if (input.detailLevel === "workbook_summary" || input.detailLevel === "semantic_index" || input.detailLevel === "sheet_summary") {
    return false;
  }
  if (input.target?.candidateId || input.target?.tableName || input.target?.sheetName || input.target?.range) {
    return false;
  }
  if (/\b(all|every|entire|whole|full)\s+(?:table|sheet|rows?|data|values?|workbook)\b/i.test(input.request)) {
    return false;
  }
  return requestExplicitlyAsksSelectionContext(input.request);
}

function isSheetOverviewRequest(requestText: string): boolean {
  if (isExplicitFullDataRequest(requestText)) {
    return false;
  }
  if (/\b(actual\s+values?|raw\s+values?|rows?|records?|cells?|sample\s+data|first\s+\d+|last\s+\d+)\b/i.test(requestText)) {
    return false;
  }
  if (/\b(style|styling|format|formats|font|fonts|color|colors|border|borders|alignment|fills?|number formats?|formula|formulas?)\b/i.test(requestText)) {
    return false;
  }
  return /\b(?:look(?:\s+at|\s+into)?|inspect|review|summari[sz]e|analy[sz]e|check|show|describe)\b.{0,60}\b(?:worksheet|sheet|active sheet|current sheet)\b/i.test(requestText)
    || /\b(?:worksheet|sheet|active sheet|current sheet)\b.{0,60}\b(?:look|inspect|review|summary|summari[sz]e|overview|structure|about)\b/i.test(requestText);
}

function requestExplicitlyAsksSelectionContext(requestText: string): boolean {
  if (/\bthis\s+(?:workbook|worksheet|sheet|file|xlsx|excel file)\b/i.test(requestText)) {
    return false;
  }
  return /\b(selection|selected|highlighted|active cell|current cell|this cell|this row|this range|this column|here)\b/i.test(requestText)
    || /\bwhat\b.{0,30}\b(?:this|here)\b/i.test(requestText)
    || /\b(?:check|inspect|look(?:\s+at)?|analy[sz]e|review)\b.{0,30}\b(?:this|here)\b/i.test(requestText);
}

function detailLevelAnswerOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  if (input.detailLevel === "workbook_summary") {
    return workbookSummaryDetailOutput(metadata, requestedMode);
  }
  if (input.detailLevel === "semantic_index") {
    return semanticIndexDetailOutput(metadata, input, requestedMode);
  }
  if (input.detailLevel === "sheet_summary") {
    return sheetSummaryDetailOutput(metadata, input, requestedMode);
  }
  if (input.detailLevel === "style_overview") {
    return undefined;
  }
  if (input.detailLevel === "workbook_design_overview") {
    return undefined;
  }
  if (input.detailLevel === "full_table" && !isExplicitFullDataRequest(input.request)) {
    const output = sheetSummaryDetailOutput(metadata, input, requestedMode);
    return {
      ...output,
      summary: `${output.summary} Full-table detail was not fetched because the request was an overview-style inspection.`,
      warnings: [...output.warnings, "Use full_table only when the user explicitly asks for all rows, every value, or full table contents."]
    };
  }
  return undefined;
}

function sheetOverviewAnswerOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  if (!isSheetOverviewRequest(input.request)) {
    return undefined;
  }
  return sheetSummaryDetailOutput(metadata, input, requestedMode);
}

function workbookDesignOverviewAnswerOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  const action = intentAction(input);
  if (action && action !== "workbook_design_overview") {
    return undefined;
  }
  if (action !== "workbook_design_overview" && input.detailLevel !== "workbook_design_overview" && !isWorkbookDesignOverviewRequest(input.request)) {
    return undefined;
  }
  const target = resolveWorkbookDesignTarget(metadata, input);
  if (!target) {
    return {
      status: "NEEDS_INPUT",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: "Workbook design overview needs a target sheet or table.",
      answer: {
        kind: "workbook_design_overview_needs_target",
        source: "cached_metadata",
        candidateSheets: metadata.sheets.filter((sheet) => sheet.usedRange || sheet.headers.length > 0 || sheet.tableIds.length > 0).slice(0, 8).map((sheet) => ({
          sheetName: sheet.name,
          kind: sheet.kind,
          usedRange: sheet.usedRange
        }))
      },
      proof: [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "ask_user",
      warnings: ["Select a sheet/table or provide target.sheetName before requesting a workbook design overview."]
    };
  }
  const { sheet, table, range, columns } = target;
  const detected = detectVisualReadabilityStructure(metadata, sheet, range);
  const visualColumns = inferVisualReadabilityColumns(metadata, sheet, detected);
  const relatedSheets = workbookDesignRelatedSheets(metadata, sheet, columns);
  const columnRecommendations = columns.map((column) => workbookDesignColumnRecommendation(metadata, sheet, column, relatedSheets));
  const dropdownCount = columnRecommendations.filter((column) => column.recommendedBehavior === "dropdown_list").length;
  const lookupCount = columnRecommendations.filter((column) => column.lookupRecommendation !== undefined).length;
  const formatCount = columnRecommendations.filter((column) => column.formatRecommendation !== undefined).length;
  const answer = stripUndefinedRecord({
    kind: "workbook_design_overview",
    source: "cached_metadata_semantic_design",
    workbook: { name: metadata.workbook.name, sheetCount: metadata.workbook.sheetCount },
    sheet: { name: sheet.name, kind: sheet.kind, usedRange: sheet.usedRange, rowCount: sheet.rowCount, columnCount: sheet.columnCount },
    table: table ? { name: table.name, range: table.range, headerRange: table.headerRange, dataRange: table.dataRange, columnCount: table.columns.length } : undefined,
    target: { sheetName: sheet.name, range, tableName: table?.name },
    dataState: workbookDesignDataState(sheet, table),
    inspectionPolicy: {
      valuesRead: false,
      fullReadCellCount: 0,
      guidance: "Use this overview for template/design recommendations. Do not broad-read empty data rows just to infer column roles; use targeted validation/reference workflows only after the user chooses a recommendation."
    },
    relatedSheets,
    columnRecommendations,
    groupSuggestions: suggestedColumnGroups(visualColumns, detected),
    summary: {
      columnCount: columns.length,
      formatRecommendations: formatCount,
      dropdownCandidates: dropdownCount,
      lookupCandidates: lookupCount
    },
    nextWorkflows: workbookDesignNextWorkflows(dropdownCount, lookupCount)
  });
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `Returned workbook design overview for ${sheet.name}!${range}: ${columns.length} column recommendation(s), ${dropdownCount} dropdown candidate(s), ${lookupCount} lookup/reference candidate(s).`,
    answer,
    metrics: { source: "cached_metadata_semantic_design", fullReadCellCount: 0, internalReadCount: 0, columnCount: columns.length, relatedSheetCount: relatedSheets.length },
    proof: [{ sheetName: sheet.name, range, label: "workbook design overview" }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: []
  };
}

function isWorkbookDesignOverviewRequest(requestText: string): boolean {
  if (isExplicitFullDataRequest(requestText)) {
    return false;
  }
  if (isExplicitCellValueReadRequest(requestText.toLowerCase())) {
    return false;
  }
  if (isDataValidationReadRequest(requestText.toLowerCase())) {
    return false;
  }
  if (/\b(style|styling|visual|readability|font|fonts|color|colors|border|borders|fills?|theme|header)\b/i.test(requestText)) {
    return false;
  }
  const asksDesign = /\b(workbook|sheet|table|column[-\s]?by[-\s]?column|columns?)\b/i.test(requestText)
    && /\b(design|data\s+entry|template|dropdown|lookup|reference|validation|column[-\s]?by[-\s]?column)\b/i.test(requestText)
    && /\b(review|recommend|decide|should\s+be|for\s+each|each\s+column|every\s+column|column[-\s]?by[-\s]?column)\b/i.test(requestText);
  const asksColumnTypes = /\b(free\s+text|date|money|number|id|code|dropdown|lookup|reference)\b/i.test(requestText)
    && /\b(each|per|every|columns?|sheet|table)\b/i.test(requestText);
  return asksDesign || asksColumnTypes;
}

function resolveWorkbookDesignTarget(metadata: WorkbookMetadata, input: AgentRunInput): { sheet: WorkbookMetadata["sheets"][number]; table?: TableMetadata; range: string; columns: ColumnMetadata[] } | undefined {
  const targetSheetName = stringValue(input.target?.sheetName)
    ?? stringValue(input.target?.entity)
    ?? findMentionedSheet(metadata, input)?.name
    ?? (requestMentionsActiveSheet(input.request) ? metadata.workbook.activeSheet : undefined)
    ?? metadata.workbook.activeSheet;
  const sheet = targetSheetName
    ? metadata.sheets.find((candidate) => sameText(candidate.name, targetSheetName))
    : undefined;
  if (!sheet) {
    return undefined;
  }
  const tableName = stringValue(input.target?.tableName);
  const table = (tableName ? metadata.tables.find((candidate) => sameText(candidate.name, tableName) && sameText(candidate.sheetName, sheet.name)) : undefined)
    ?? metadata.tables.find((candidate) => sameText(candidate.sheetName, sheet.name));
  const bestHeader = sheet.headers.slice().sort((left, right) => right.confidence - left.confidence)[0];
  const columns = table?.columns && table.columns.length > 0 ? table.columns : bestHeader?.columns ?? [];
  const range = stripSheetName(input.target?.range ?? table?.range ?? bestHeader?.range ?? sheet.usedRange ?? "A1:A1");
  return { sheet, ...(table ? { table } : {}), range, columns };
}

function workbookDesignDataState(sheet: WorkbookMetadata["sheets"][number], table?: TableMetadata) {
  const rowCount = sheet.rowCount ?? rowCountFromAddress(sheet.usedRange) ?? rowCountFromAddress(table?.range);
  const tableDataRows = rowCountFromAddress(table?.dataRange);
  const looksTemplateLike = Boolean(table && (tableDataRows ?? 0) >= 20);
  return stripUndefinedRecord({
    kind: looksTemplateLike ? "template_or_structured_table" : sheet.kind === "template" ? "template" : "metadata_only",
    rowCount,
    tableDataRows,
    recommendation: looksTemplateLike
      ? "Treat the sheet as a template/structured table for design review. Infer from headers and related sheets; avoid repeated data-row reads unless the user asks for actual values."
      : "Metadata is enough for initial design recommendations; sample values are optional proof, not required for column role decisions."
  });
}

function rowCountFromAddress(address: string | undefined): number | undefined {
  if (!address) return undefined;
  const parsed = tryParseA1Address(stripSheetName(address));
  return parsed ? parsed.endRow - parsed.startRow + 1 : undefined;
}

function workbookDesignRelatedSheets(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number], columns: ColumnMetadata[]): Array<{ sheetName: string; kind: string; usedRange?: string; confidence: number; reasons: string[] }> {
  const targetHeaders = new Set(columns.map((column) => normalizeHeaderName(column.name)).filter(Boolean));
  const targetTokens = new Set(columns.flatMap((column) => workbookDesignHeaderTokens(column.name)));
  return metadata.sheets
    .filter((candidate) => candidate.name !== sheet.name)
    .map((candidate) => {
      const candidateColumns = candidate.headers.flatMap((header) => header.columns);
      const exactOverlap = candidateColumns.filter((column) => targetHeaders.has(normalizeHeaderName(column.name))).length;
      const tokenOverlap = candidateColumns.flatMap((column) => workbookDesignHeaderTokens(column.name)).filter((token) => targetTokens.has(token)).length;
      const nameScore = workbookDesignSheetNameScore(candidate.name, columns);
      const kindScore = candidate.kind === "lookup" ? 2 : candidate.kind === sheet.kind ? 1 : 0;
      const score = exactOverlap * 3 + Math.min(tokenOverlap, 4) + nameScore + kindScore;
      const reasons = [
        exactOverlap > 0 ? `${exactOverlap} matching header(s)` : undefined,
        tokenOverlap > 0 ? `${Math.min(tokenOverlap, 4)} related header token(s)` : undefined,
        nameScore > 0 ? "sheet name matches a column domain" : undefined,
        candidate.kind === "lookup" ? "lookup sheet" : undefined
      ].filter((reason): reason is string => Boolean(reason));
      return { candidate, score, reasons };
    })
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score || left.candidate.name.localeCompare(right.candidate.name))
    .slice(0, 8)
    .map((entry) => stripUndefinedRecord({
      sheetName: entry.candidate.name,
      kind: entry.candidate.kind,
      usedRange: entry.candidate.usedRange,
      confidence: Math.max(0.35, Math.min(0.95, entry.score / 10)),
      reasons: entry.reasons
    }) as { sheetName: string; kind: string; usedRange?: string; confidence: number; reasons: string[] });
}

function workbookDesignSheetNameScore(sheetName: string, columns: ColumnMetadata[]): number {
  const normalizedSheet = normalizeComparableText(sheetName);
  let score = 0;
  for (const column of columns) {
    const normalized = normalizeComparableText(column.name);
    if ((/customer|ลูกค้า/.test(normalized) && /customer|ลูกค้า/.test(normalizedSheet))
      || (/driver|truck|คนขับ|ทะเบียน/.test(normalized) && /driver|truck|รถ|คนขับ/.test(normalizedSheet))
      || (/booking|บุ๊ค|จอง/.test(normalized) && /booking|บุ๊ค|จอง/.test(normalizedSheet))) {
      score += 4;
    }
  }
  return score;
}

function workbookDesignHeaderTokens(value: string): string[] {
  return normalizeComparableText(value)
    .split(/[^a-z0-9ก-๙]+/i)
    .filter((token) => token.length >= 3)
    .slice(0, 8);
}

function workbookDesignColumnRecommendation(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number], column: ColumnMetadata, relatedSheets: ReturnType<typeof workbookDesignRelatedSheets>) {
  const header = column.name;
  const normalized = normalizeComparableText(header);
  const role = visualColumnRole(column);
  const lookup = workbookDesignLookupRecommendation(metadata, column, relatedSheets);
  const dropdown = workbookDesignDropdownRecommendation(column, lookup);
  const behavior = lookup ? "lookup_reference" : dropdown ? "dropdown_list" : workbookDesignBehavior(column, role, normalized);
  return stripUndefinedRecord({
    column: column.letter,
    header,
    currentInferredType: column.inferredType,
    role,
    recommendedBehavior: behavior,
    formatRecommendation: workbookDesignFormatRecommendation(column, role, normalized),
    dropdownRecommendation: dropdown,
    lookupRecommendation: lookup,
    applySafety: lookup ? "separate_preview_required" : dropdown ? "validation_bucket_or_separate_preview" : "safe_visual_formatting",
    rationale: workbookDesignColumnRationale(column, role, behavior)
  });
}

function workbookDesignBehavior(column: ColumnMetadata, role: string, normalized: string): string {
  if (role === "date") return "date";
  if (role === "money" || role === "number") return "number_money";
  if (role === "id" || /เลข|number|no|code|id|ทะเบียน|phone|โทร|tax|ภาษี/.test(normalized)) return "id_text_code";
  if (role === "status" || role === "category") return "dropdown_list";
  return "free_text";
}

function workbookDesignFormatRecommendation(column: ColumnMetadata, role: string, normalized: string) {
  if (role === "date" || /วันที่|date/.test(normalized)) {
    return { type: "date", numberFormat: "dd/mm/yyyy", reason: "Date-like header should sort/filter as dates." };
  }
  if (role === "money" || role === "number" || /ราคา|ยอด|ค่า|ภาษี|amount|price|fee|tax|total|net|gross/.test(normalized)) {
    return { type: "money", numberFormat: "#,##0.00", reason: "Money-like columns should align right and use a consistent numeric format." };
  }
  if (role === "id" || /เลข|booking|บุ๊ค|no|number|code|id|ทะเบียน|phone|โทร|tax id|เลขประจำตัว/.test(normalized)) {
    return { type: "text_code", numberFormat: "@", reason: "Identifiers should stay text so leading zeros, hyphens, and registration codes are preserved." };
  }
  return undefined;
}

function workbookDesignDropdownRecommendation(column: ColumnMetadata, lookup: unknown) {
  if (lookup) {
    return undefined;
  }
  const normalized = normalizeComparableText(column.name);
  if (/สถานะ|status|state|stage/.test(normalized)) {
    return {
      source: "suggested_static_options",
      options: /จ่าย|payment|paid/.test(normalized) ? ["ยังไม่จ่าย", "จ่ายแล้ว", "รอตรวจสอบ"] : ["รอดำเนินการ", "กำลังดำเนินการ", "เสร็จแล้ว", "ยกเลิก"],
      nextWorkflow: { intentAction: "write_data_validation", mode: "preview_update" }
    };
  }
  if (/container size|ขนาดตู้/.test(normalized)) {
    return {
      source: "suggested_static_options",
      options: ["20GP", "40GP", "40HQ"],
      nextWorkflow: { intentAction: "write_data_validation", mode: "preview_update" }
    };
  }
  if (/type|category|ประเภท/.test(normalized)) {
    return {
      source: "needs_source_or_existing_values",
      options: [],
      nextWorkflow: { intentAction: "read_data_validation", mode: "answer" }
    };
  }
  return undefined;
}

function workbookDesignLookupRecommendation(metadata: WorkbookMetadata, column: ColumnMetadata, relatedSheets: ReturnType<typeof workbookDesignRelatedSheets>) {
  const normalized = normalizeComparableText(column.name);
  const target = relatedSheets.find((sheet) => {
      const sheetName = normalizeComparableText(sheet.sheetName);
      return (/ลูกค้า|customer/.test(normalized) && /customer|ลูกค้า/.test(sheetName))
        || (/booking|บุ๊ค|จอง/.test(normalized) && /booking|บุ๊ค|จอง/.test(sheetName))
        || (/driver|truck|คนขับ|ทะเบียน/.test(normalized) && /driver|truck|รถ|คนขับ/.test(sheetName))
      || (sheet.kind === "lookup" && workbookDesignRelatedSheetHasHeader(metadata, sheet.sheetName, column.name));
  });
  if (!target) {
    return undefined;
  }
  const targetSheet = metadata.sheets.find((sheet) => sheet.name === target.sheetName);
  const keyColumn = workbookDesignLookupKeyColumn(targetSheet, column);
  return stripUndefinedRecord({
    sourceSheetName: target.sheetName,
    sourceRange: target.usedRange,
    keyColumn,
    confidence: target.confidence,
    nextWorkflow: { intentAction: "write_data_validation", mode: "preview_update", note: "Preview dropdown/source-list or lookup formula separately before applying." },
    reason: `${column.name} appears related to ${target.sheetName}; use that sheet as the source of truth instead of free text when possible.`
  });
}

function workbookDesignRelatedSheetHasHeader(metadata: WorkbookMetadata, sheetName: string, header: string): boolean {
  const normalized = normalizeHeaderName(header);
  if (!normalized) return false;
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  return Boolean(sheet?.headers.some((entry) => entry.columns.some((column) => normalizeHeaderName(column.name) === normalized)));
}

function workbookDesignLookupKeyColumn(sheet: WorkbookMetadata["sheets"][number] | undefined, sourceColumn: ColumnMetadata): string | undefined {
  const normalizedSource = normalizeComparableText(sourceColumn.name);
  const columns = sheet?.headers.flatMap((header) => header.columns) ?? [];
  const preferred = columns.find((column) => {
    const normalized = normalizeComparableText(column.name);
    return (/customer|ลูกค้า/.test(normalizedSource) && /customer|ลูกค้า|ชื่อ/.test(normalized))
      || (/booking|บุ๊ค/.test(normalizedSource) && /booking|บุ๊ค|เลข/.test(normalized))
      || (/driver|truck|คนขับ|ทะเบียน/.test(normalizedSource) && /driver|truck|คนขับ|ทะเบียน|ชื่อ/.test(normalized));
  }) ?? columns[0];
  return preferred ? `${preferred.letter}:${preferred.name}` : undefined;
}

function workbookDesignColumnRationale(column: ColumnMetadata, role: string, behavior: string): string {
  if (behavior === "lookup_reference") return "Header matches a related lookup/source sheet.";
  if (behavior === "dropdown_list") return "Header is status/category-like and should use controlled values.";
  if (behavior === "date") return "Header and metadata indicate a date column.";
  if (behavior === "number_money") return "Header and metadata indicate money/number values.";
  if (behavior === "id_text_code") return "Identifier-like values should be preserved as text.";
  return column.importance !== undefined && column.importance >= 0.85 ? `High-importance ${role} column; keep easy to scan.` : "No strong controlled-value or lookup signal; keep manual text entry.";
}

function workbookDesignNextWorkflows(dropdownCount: number, lookupCount: number) {
  const workflows: Array<Record<string, unknown>> = [
    { intentAction: "improve_visual_readability", mode: "preview_update", purpose: "Apply safe visual formatting, widths, alignment, filters, and number formats." }
  ];
  if (dropdownCount > 0) {
    workflows.push({ intentAction: "write_data_validation", mode: "preview_update", purpose: "Preview dropdown/data-validation rules for selected columns." });
  }
  if (lookupCount > 0) {
    workflows.push({ intentAction: "derive_values", mode: "preview_update", purpose: "Preview lookup/reference formulas or source-list behavior separately." });
  }
  return workflows;
}

function semanticIndexDetailOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
  const semanticIndex = buildSemanticWorkbookIndex(metadata, { maxEntries: input.budget?.maxExamples ?? 25 });
  const candidates = semanticIndex.entries.map((entry) => ({
    id: entry.id,
    kind: entry.sourceKind === "selection" ? "range" as const : entry.sourceKind,
    label: entry.label,
    ...(entry.sheetName !== undefined ? { sheetName: entry.sheetName } : {}),
    ...(entry.tableName !== undefined ? { tableName: entry.tableName } : {}),
    ...(entry.range !== undefined ? { range: entry.range } : {}),
    semanticRole: entry.role,
    aliases: entry.aliases.slice(0, 8),
    confidence: entry.confidence,
    reason: `${entry.role} semantic index entry from ${entry.sourceKind}.`,
    nextRequestHint: `Retry with target.candidateId "${entry.id}".`
  }));
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `Returned semantic workbook index for ${metadata.workbook.name} from cached metadata.`,
    answer: semanticIndex,
    metrics: { source: "cached_metadata", detailLevel: "semantic_index", fullReadCellCount: 0, semanticEntryCount: semanticIndex.entryCount },
    candidates,
    proof: semanticIndex.entries.flatMap((entry) => entry.sheetName && entry.range ? [{ sheetName: entry.sheetName, range: entry.range, label: entry.label }] : []).slice(0, 5),
    resourceLinks: [contextResource(metadata.workbookContextId), semanticIndexResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: metadata.detailLevel === "sampled" ? [] : ["Semantic roles are structure-only; confidence improves when sampled metadata is available."]
  };
}

function isExplicitFullDataRequest(requestText: string): boolean {
  const request = requestText.toLowerCase();
  return /\b(full|entire|complete|all|every)\b.*\b(table|rows?|values?|cells?|data|contents?)\b/.test(request)
    || /\b(read|show|list|return|dump)\b.*\b(all|every)\b/.test(request)
    || /\ba1:[a-z]+\d+\b/i.test(requestText);
}

function workbookSummaryDetailOutput(metadata: WorkbookMetadata, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
  const summary = workbookCompactSummary(metadata);
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `Returned workbook summary for ${metadata.workbook.name} from cached metadata.`,
    answer: {
      kind: "workbook_summary",
      source: "cached_metadata",
      workbook: metadata.workbook,
      sheetCount: metadata.sheets.length,
      tableCount: metadata.tables.length,
      namedRangeCount: metadata.namedRanges.length,
      semanticIndex: buildSemanticWorkbookIndex(metadata, { maxEntries: 12 }),
      sheets: summary.sheets.map((sheet) => ({
        ...sheet,
        contextHints: compactWorkbookContextHints(metadata, String(sheet.name ?? ""))
      }))
    },
    metrics: { source: "cached_metadata", detailLevel: "workbook_summary", fullReadCellCount: 0 },
    proof: metadata.sheets.slice(0, 5).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "used range" }] : []),
    resourceLinks: [contextResource(metadata.workbookContextId), semanticIndexResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    taskOutcome: "final_answer",
    agentInstruction: "This workbook summary is complete for an overview request. Answer now from cached metadata; do not fetch fullResultUri, chunk-read sheets, or call low-level MCP resources unless the user asks for all raw rows or exact cell values.",
    maxRecommendedFollowupCalls: 0,
    warnings: []
  };
}

function sheetSummaryDetailOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
  const requestedSheet = stringValue(input.target?.sheetName)
    ?? stringValue((input.values as Record<string, unknown> | undefined)?.sheetName)
    ?? sheetNameFromRequest(metadata, input.request)
    ?? metadata.workbook.activeSheet
    ?? metadata.sheets[0]?.name;
  const sheet = requestedSheet ? metadata.sheets.find((candidate) => normalizeComparableText(candidate.name) === normalizeComparableText(requestedSheet)) : undefined;
  if (!sheet) {
    return {
      status: "AMBIGUOUS_TARGET",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: "Sheet summary needs a sheet target.",
      candidates: metadata.sheets.slice(0, 10).map((candidate, index) => ({
        id: candidate.id,
        kind: "sheet" as const,
        label: candidate.name,
        sheetName: candidate.name,
        ...(candidate.usedRange !== undefined ? { range: candidate.usedRange } : {}),
        confidence: Math.max(0.1, 0.9 - index * 0.05),
        reason: "Available sheet"
      })),
      proof: [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "call_with_target",
      warnings: []
    };
  }
  const tableIds = new Set(sheet.tableIds);
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `Returned sheet summary for ${sheet.name} from cached metadata.`,
    answer: {
      kind: "sheet_summary",
      source: "cached_metadata",
      sheet,
      contextHints: compactWorkbookContextHints(metadata, sheet.name),
      tables: metadata.tables.filter((table) => tableIds.has(table.id)),
      namedRanges: metadata.namedRanges.filter((name) => name.sheetName === sheet.name),
      sections: metadata.sections.filter((section) => section.sheetName === sheet.name).map(compactSectionMapForAgent),
      summaryBlocks: metadata.summaryBlocks.filter((block) => block.sheetName === sheet.name),
      formulaRegions: metadata.formulaRegions.filter((region) => region.sheetName === sheet.name)
    },
    metrics: { source: "cached_metadata", detailLevel: "sheet_summary", fullReadCellCount: 0 },
    proof: sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "used range" }] : [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    taskOutcome: "final_answer",
    agentInstruction: "This sheet summary is complete for an overview request. Answer now from cached metadata; do not fetch fullResultUri, chunk-read the sheet, or call low-level MCP resources unless the user asks for all raw rows or exact cell values.",
    maxRecommendedFollowupCalls: 0,
    warnings: metadata.detailLevel === "sampled" ? [] : ["Sheet sections and formula regions are richer after sampled metadata is available."]
  };
}

function compactSectionMapForAgent(section: WorkbookMetadata["sections"][number]): Record<string, unknown> {
  const dataRange = dataRangeForSection(section);
  const keyColumns = section.columns
    .filter((column) => ["description", "dimension", "identifier", "vendor", "account", "category", "status"].includes(column.role ?? ""))
    .slice(0, 6);
  const editableColumns = section.columns
    .filter((column) => ["measure", "amount", "status", "category", "note", "unknown"].includes(column.role ?? "") || /price|rate|available|propose|amount|cost|note|status/i.test(column.name))
    .slice(0, 8);
  return stripUndefinedRecord({
    id: section.id,
    fingerprint: sectionMapFingerprint(section, dataRange),
    label: section.label,
    kind: section.kind,
    sheetName: section.sheetName,
    range: section.range,
    headerRange: section.headerRange,
    headerRow: section.headerRow,
    dataRange,
    rowCount: section.rowCount,
    columnCount: section.columnCount,
    nonEmptyCellCount: section.nonEmptyCellCount,
    confidence: section.confidence,
    labels: section.labels.slice(0, 8),
    columns: section.columns.slice(0, 16).map(compactColumnAnchor),
    keyColumns: keyColumns.map(compactColumnAnchor),
    editableColumns: editableColumns.map(compactColumnAnchor),
    nextRequestHints: [
      section.headerRange ? `Resolve edits by section ${section.id}, row label, and column header before using raw coordinates.` : undefined,
      dataRange ? `Read exact examples from ${section.sheetName}!${dataRange} only when row values are needed.` : undefined
    ].filter(Boolean)
  });
}

function sectionMapFingerprint(section: WorkbookMetadata["sections"][number], dataRange: string | undefined): string {
  return hashStable({
    id: section.id,
    sheetName: section.sheetName,
    range: section.range,
    headerRange: section.headerRange,
    headerRow: section.headerRow,
    dataRange,
    columns: section.columns.map((column) => ({
      index: column.index,
      letter: column.letter,
      normalizedName: column.normalizedName,
      role: column.role,
      inferredType: column.inferredType
    })),
    labels: section.labels.slice(0, 12)
  });
}

function compactColumnAnchor(column: ColumnMetadata): Record<string, unknown> {
  return stripUndefinedRecord({
    name: column.name,
    normalizedName: column.normalizedName,
    index: column.index,
    letter: column.letter,
    role: column.role,
    inferredType: column.inferredType,
    importance: column.importance
  });
}

function dataRangeForSection(section: WorkbookMetadata["sections"][number]): string | undefined {
  const parsed = tryParseA1Address(stripSheetName(section.range));
  if (!parsed) {
    return undefined;
  }
  const startRow = Math.max(parsed.startRow, (section.headerRow ?? parsed.startRow - 1) + 1);
  if (startRow > parsed.endRow) {
    return undefined;
  }
  return `${numberToColumn(parsed.startColumn)}${startRow}:${numberToColumn(parsed.endColumn)}${parsed.endRow}`;
}

function sheetNameFromRequest(metadata: WorkbookMetadata, request: string): string | undefined {
  const normalizedRequest = normalizeComparableText(request);
  const exact = metadata.sheets.find((sheet) => normalizedRequest.includes(normalizeComparableText(sheet.name)))?.name;
  if (exact) {
    return exact;
  }
  return findAgentCandidates(metadata, { request, mode: "answer" })
    .find((candidate) => candidate.kind === "sheet" && candidate.sheetName)?.sheetName;
}

function emptyLiveReadDiagnosticOutput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  requestedMode: AgentRunMode,
  sheetName: string,
  range: string,
  label: string,
  profile: ReturnType<typeof profileValues>
): Omit<AgentRunOutput, "telemetry"> | undefined {
  const requestedCells = cellCountFromAddress(range);
  if ((requestedCells !== undefined && requestedCells <= 1) || profile.metrics.nonEmptyCount !== 0 || !metadataSuggestsRangeHasData(metadata, sheetName, range)) {
    return undefined;
  }
  if (emptyResultKindForProfile(metadata, sheetName, range, profile) === "fresh_sheet") {
    return undefined;
  }
  return {
    status: "ERROR",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `OpenWorkbook tried to read ${sheetName}!${range}, but the live read returned no values even though workbook metadata indicates this range should contain data.`,
    answer: {
      kind: "live_read_diagnostic",
      sheetName,
      range,
      label,
      expectedDataFromMetadata: true,
      recommendation: "Reload the OpenWorkbook Local taskpane, then retry the same excel.agent.run request. Do not fall back to offline .xlsx parsing unless the user explicitly asks for saved-file analysis."
    },
    metrics: profile.metrics,
    proof: [{ sheetName, range, label }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "manual_review",
    taskOutcome: "cannot_complete",
    agentInstruction: "Report this Open Workbook live-read failure to the user. Do not use Python/openpyxl/offline parsing while the live Excel add-in is connected unless the user explicitly requests offline file analysis.",
    maxRecommendedFollowupCalls: 0,
    finalAnswer: `Open Workbook tried to read live values from ${sheetName}!${range}, but the returned payload was empty while workbook metadata says the range should contain data. Treat this as an Open Workbook live-read failure, not an empty sheet.`,
    warnings: [
      "Open Workbook live Excel read returned empty data for a metadata-backed non-empty range.",
      "Do not use Python/openpyxl/offline parsing as a fallback for live Excel state unless the user explicitly asks."
    ]
  };
}

function shouldAnalyzeReferenceSheet(input: AgentRunInput): boolean {
  const request = input.request.toLowerCase();
  if (intentAction(input) === "find_similar_rows" || intentAction(input) === "find_style_references") {
    return false;
  }
  if (isDataValidationReadRequest(request)) {
    return false;
  }
  if (!/\b(reference|same as|similar to|prior|previous|last month|before|template|convention|pattern|skim|learn from|look into)\b/.test(request)) {
    return false;
  }
  return /\b(analy[sz]e|skim|look|inspect|learn|understand|pattern|relation|relationship|how many|count|counts?|types?|style|styling|header|formula|condition|format|note|columns?)\b/.test(request);
}

function resolveReferenceAnalysisTarget(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { sheetName: string; range: string; tableName?: string; label: string } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const requestedSheet = stringValue(values?.referenceSheetName ?? values?.referenceSheet ?? values?.sourceSheetName ?? values?.sourceSheet)
    ?? stringValue(input.target?.sheetName)
    ?? sheetNameFromRequest(metadata, input.request);
  const requestedRange = stringValue(values?.referenceRange ?? values?.sourceRange ?? values?.range ?? input.target?.range);
  const requestedTable = stringValue(values?.referenceTableName ?? values?.sourceTableName ?? input.target?.tableName);
  if (requestedTable) {
    const table = metadata.tables.find((candidate) => typeof candidate.name === "string" && normalizeComparableText(candidate.name) === normalizeComparableText(requestedTable));
    if (table) {
      const tableName = table.name ?? requestedTable;
      return {
        sheetName: table.sheetName,
        range: clampReferenceAnalysisRange(table.dataRange ?? table.range),
        tableName,
        label: tableName
      };
    }
  }
  if (!requestedSheet) {
    return undefined;
  }
  const sheet = metadata.sheets.find((candidate) => normalizeComparableText(candidate.name) === normalizeComparableText(requestedSheet));
  if (!sheet) {
    return undefined;
  }
  return {
    sheetName: sheet.name,
    range: clampReferenceAnalysisRange(requestedRange ?? sheet.usedRange ?? "A1:A1"),
    label: sheet.name
  };
}

function clampReferenceAnalysisRange(range: string): string {
  const parsed = tryParseA1Address(stripSheetName(range));
  if (!parsed) {
    return stripSheetName(range);
  }
  const maxColumns = 36;
  const maxCells = 9000;
  const columnCount = Math.min(maxColumns, parsed.endColumn - parsed.startColumn + 1);
  const maxRows = Math.max(1, Math.floor(maxCells / columnCount));
  const endRow = Math.min(parsed.endRow, parsed.startRow + maxRows - 1);
  const endColumn = Math.min(parsed.endColumn, parsed.startColumn + columnCount - 1);
  return `${numberToColumn(parsed.startColumn)}${parsed.startRow}:${numberToColumn(endColumn)}${endRow}`;
}

function referenceStyleTargets(metadata: WorkbookMetadata, sheetName: string, range: string): Array<{ label: string; range: string }> {
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const header = sheet?.headers[0]?.range;
  const parsed = tryParseA1Address(stripSheetName(range));
  const dataRange = parsed && parsed.endRow > parsed.startRow
    ? `${numberToColumn(parsed.startColumn)}${parsed.startRow + 1}:${numberToColumn(parsed.endColumn)}${Math.min(parsed.endRow, parsed.startRow + 12)}`
    : undefined;
  return [
    ...(header ? [{ label: "header_style", range: header }] : []),
    ...(dataRange ? [{ label: "data_row_style", range: dataRange }] : [{ label: "range_style", range }])
  ];
}

function referenceAnalysisFromSnapshot(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  reference: { sheetName: string; range: string; tableName?: string; label: string },
  snapshot: RangeSnapshot | undefined
) {
  const values = snapshot?.values ?? [];
  const formulas = normalizeFormulaOnlyMatrix(snapshot?.formulas ?? []);
  const numberFormats = snapshot?.numberFormat ?? [];
  const parsed = tryParseA1Address(stripSheetName(reference.range));
  const headers = referenceHeaders(metadata, reference.sheetName, values);
  const dataRows = values.slice(headers.sourceRowIndex + 1).filter((row) => row.some(hasCellValue));
  const requestedColumns = requestedReferenceColumns(input, headers.headers);
  const columnProfiles = headers.headers
    .map((header, index) => referenceColumnProfile(header, index, dataRows, numberFormats, requestedColumns.has(index)))
    .filter((profile): profile is Record<string, unknown> => profile !== undefined)
    .sort((left, right) => Number(Boolean(right.requested)) - Number(Boolean(left.requested)) || Number(right.nonEmptyCount ?? 0) - Number(left.nonEmptyCount ?? 0))
    .slice(0, 16);
  const relationships = referenceColumnRelationships(headers.headers, dataRows, requestedColumns);
  const formulaPatterns = referenceFormulaPatterns(headers.headers, formulas, parsed?.startRow ?? 1);
  const rowConditionPatterns = referenceRowConditionPatterns(headers.headers, dataRows);
  const evidenceRows = referenceEvidenceRows(headers.headers, dataRows, requestedColumns, 6);
  return {
    objectives: referenceObjectives(input.request),
    structure: {
      sheetName: reference.sheetName,
      tableName: reference.tableName,
      range: reference.range,
      rowCount: values.length,
      dataRowCount: dataRows.length,
      columnCount: headers.headers.length,
      headerRow: parsed ? parsed.startRow + headers.sourceRowIndex : undefined,
      headers: headers.headers.map((name, index) => ({ name, index, letter: numberToColumn(index + (parsed?.startColumn ?? 1)) }))
    },
    columnProfiles,
    relationships,
    formulaPatterns,
    rowConditionPatterns,
    evidenceRows,
    warnings: values.length === 0
      ? ["No values were returned from the reference range; ask the user to reload the taskpane or choose another reference range."]
      : reference.range !== stripSheetName(reference.range)
        ? []
        : []
  };
}

function referenceHeaders(metadata: WorkbookMetadata, sheetName: string, values: CellMatrix): { headers: string[]; sourceRowIndex: number } {
  const sheetHeaders = metadata.sheets.find((sheet) => sheet.name === sheetName)?.headers[0]?.columns.map((column) => column.name);
  const firstNonEmptyIndex = values.findIndex((row) => row.some(hasCellValue));
  const rowIndex = firstNonEmptyIndex >= 0 ? firstNonEmptyIndex : 0;
  const valueHeaders = (values[rowIndex] ?? []).map((value, index) => String(value ?? sheetHeaders?.[index] ?? `Column ${index + 1}`).trim() || `Column ${index + 1}`);
  const headers = valueHeaders.length > 0 ? valueHeaders : sheetHeaders ?? [];
  return { headers, sourceRowIndex: rowIndex };
}

function requestedReferenceColumns(input: AgentRunInput, headers: string[]): Set<number> {
  const values = input.values as Record<string, unknown> | undefined;
  const requested = new Set<number>();
  const add = (name: string) => {
    const normalized = normalizeComparableText(name);
    const index = headers.findIndex((header) => normalizeComparableText(header) === normalized || normalizeComparableText(header).includes(normalized));
    if (index >= 0) requested.add(index);
  };
  const columns = values?.columns ?? values?.referenceColumns ?? values?.sourceColumns;
  if (Array.isArray(columns)) {
    columns.forEach((column) => typeof column === "string" ? add(column) : undefined);
  }
  const request = normalizeComparableText(input.request);
  headers.forEach((header, index) => {
    const normalized = normalizeComparableText(header);
    if (normalized && request.includes(normalized)) requested.add(index);
    if (/\b(type|category|status|description|desc|note|amount|actual|cash|variance|formula|date|style|header)\b/.test(normalized) && request.match(/\b(type|category|status|description|desc|note|amount|actual|cash|variance|formula|date|style|header)\b/)) {
      requested.add(index);
    }
  });
  return requested;
}

function referenceColumnProfile(header: string, index: number, rows: CellMatrix, numberFormats: string[][], requested: boolean): Record<string, unknown> | undefined {
  const values = rows.map((row) => row[index]).filter(hasCellValue);
  if (values.length === 0 && !requested) {
    return undefined;
  }
  const counts = countBy(values.map(referenceCellKey), (value) => value);
  const topValues = Object.entries(counts)
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
    .slice(0, 8)
    .map(([value, count]) => ({ value, count }));
  const numeric = values.map(numericCellValue).filter((value): value is number => value !== undefined);
  const formats = new Set(numberFormats.slice(1).map((row) => row[index]).filter((value): value is string => typeof value === "string" && value.length > 0));
  return stripUndefinedRecord({
    name: header,
    index,
    letter: numberToColumn(index + 1),
    requested,
    nonEmptyCount: values.length,
    uniqueCount: Object.keys(counts).length,
    topValues,
    numericSummary: numeric.length > 0 ? {
      count: numeric.length,
      min: Math.min(...numeric),
      max: Math.max(...numeric),
      sum: numeric.reduce((total, value) => total + value, 0)
    } : undefined,
    numberFormats: [...formats].slice(0, 4)
  });
}

function referenceColumnRelationships(headers: string[], rows: CellMatrix, requestedColumns: Set<number>): Array<Record<string, unknown>> {
  const candidateIndexes = new Set<number>(requestedColumns);
  headers.forEach((header, index) => {
    if (/\b(description|desc|type|category|status|note|amount|cash|actual|variance|date)\b/i.test(header)) {
      candidateIndexes.add(index);
    }
  });
  const indexes = [...candidateIndexes].slice(0, 8);
  const relationships: Array<Record<string, unknown>> = [];
  for (let left = 0; left < indexes.length; left += 1) {
    for (let right = left + 1; right < indexes.length; right += 1) {
      const leftIndex = indexes[left]!;
      const rightIndex = indexes[right]!;
      const examples = rows
        .filter((row) => hasCellValue(row[leftIndex]) && hasCellValue(row[rightIndex]))
        .slice(0, 5)
        .map((row) => ({
          [headers[leftIndex] ?? `Column ${leftIndex + 1}`]: row[leftIndex],
          [headers[rightIndex] ?? `Column ${rightIndex + 1}`]: row[rightIndex]
        }));
      if (examples.length > 0) {
        relationships.push({
          columns: [headers[leftIndex] ?? `Column ${leftIndex + 1}`, headers[rightIndex] ?? `Column ${rightIndex + 1}`],
          examples
        });
      }
    }
  }
  return relationships.slice(0, 12);
}

function referenceFormulaPatterns(headers: string[], formulas: CellMatrix, startRow: number): Array<Record<string, unknown>> {
  const byColumn = new Map<number, Map<string, number>>();
  formulas.forEach((row, rowIndex) => {
    row.forEach((formula, columnIndex) => {
      if (typeof formula !== "string" || !formula.startsWith("=")) return;
      const normalized = formula.replace(/\d+/g, "#");
      const counts = byColumn.get(columnIndex) ?? new Map<string, number>();
      counts.set(normalized, (counts.get(normalized) ?? 0) + 1);
      byColumn.set(columnIndex, counts);
    });
  });
  return [...byColumn.entries()].map(([columnIndex, counts]) => ({
    column: headers[columnIndex] ?? `Column ${columnIndex + 1}`,
    letter: numberToColumn(columnIndex + 1),
    formulaCount: [...counts.values()].reduce((total, count) => total + count, 0),
    patterns: [...counts.entries()].sort((left, right) => right[1] - left[1]).slice(0, 4).map(([formulaR1C1Like, count]) => ({ formulaR1C1Like, count })),
    firstFormulaRow: startRow
  })).slice(0, 8);
}

function referenceRowConditionPatterns(headers: string[], rows: CellMatrix): Array<Record<string, unknown>> {
  const amountIndexes = headers
    .map((header, index) => ({ header, index }))
    .filter((entry) => /\b(amount|cash|actual|variance|total|price|cost|fee)\b/i.test(entry.header))
    .map((entry) => entry.index)
    .slice(0, 4);
  return amountIndexes.flatMap((index) => {
    const numeric = rows.map((row) => numericCellValue(row[index])).filter((value): value is number => value !== undefined);
    if (numeric.length === 0) return [];
    return [stripUndefinedRecord({
      column: headers[index] ?? `Column ${index + 1}`,
      positiveCount: numeric.filter((value) => value > 0).length,
      zeroCount: numeric.filter((value) => value === 0).length,
      negativeCount: numeric.filter((value) => value < 0).length,
      suggestedUse: /variance/i.test(headers[index] ?? "") ? "Row note/style conditions often follow positive, zero, and negative variance states." : undefined
    })];
  });
}

function referenceEvidenceRows(headers: string[], rows: CellMatrix, requestedColumns: Set<number>, limit: number): Array<Record<string, unknown>> {
  const indexes = requestedColumns.size > 0
    ? [...requestedColumns].slice(0, 10)
    : headers.map((_header, index) => index).filter((index) => rows.some((row) => hasCellValue(row[index]))).slice(0, 10);
  return rows.slice(0, limit).map((row, rowIndex) => stripUndefinedRecord({
    rowOffset: rowIndex + 1,
    values: Object.fromEntries(indexes.map((index) => [headers[index] ?? `Column ${index + 1}`, row[index] ?? ""]))
  }));
}

function referenceObjectives(request: string): string[] {
  const objectives = [
    /\bhow many|counts?|unique|types?\b/i.test(request) ? "counts/groupings" : undefined,
    /\brelation|relationship|map|mapping|between\b/i.test(request) ? "column relationships" : undefined,
    /\bformula|formulas?\b/i.test(request) ? "formula patterns" : undefined,
    /\bstyle|styling|format|header|color|border|font\b/i.test(request) ? "style conventions" : undefined,
    /\bcondition|conditional|row\b/i.test(request) ? "row conditions" : undefined,
    /\bnote|notes?|description|desc|text\b/i.test(request) ? "text/note conventions" : undefined
  ].filter((objective): objective is string => objective !== undefined);
  return objectives.length > 0 ? objectives : ["reference patterns"];
}

function hasCellValue(value: unknown): boolean {
  return value !== null && value !== undefined && value !== "";
}

function referenceCellKey(value: unknown): string {
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value ?? "").trim();
}

type MatchUpdatePredicate = { operator: "contains" | "equals"; value: string };

type MatchUpdateColumn = {
  name: string;
  index: number;
  letter: string;
};

type MatchUpdateSpec = {
  sheetName: string;
  searchedRange: string;
  headerRow: number;
  headers: HeaderMetadata["columns"];
  matchColumn: MatchUpdateColumn;
  predicate: MatchUpdatePredicate;
  updateColumn: MatchUpdateColumn;
  updateValue: unknown;
};

function resolveMatchUpdateSpec(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { ok: true } & MatchUpdateSpec | { ok: false; summary: string; candidates?: AgentCandidate[]; warnings: string[] } | undefined {
  const parsedRequest = resolveMatchUpdateRequest(input.request, input.values);
  if (!parsedRequest) {
    return undefined;
  }
  const sheet = resolveMatchUpdateSheet(metadata, input);
  if (!sheet.ok) {
    return {
      ok: false,
      summary: sheet.summary,
      ...(sheet.candidates !== undefined ? { candidates: sheet.candidates } : {}),
      warnings: sheet.warnings
    };
  }
  const header = bestHeaderForSheet(metadata, sheet.sheetName, [...parsedRequest.matchHints, ...parsedRequest.updateHints]);
  if (!header) {
    return {
      ok: false,
      summary: `Could not find a header row on ${sheet.sheetName} for the match/update task.`,
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      warnings: ["Ask the user for an exact table/range or header row before updating."]
    };
  }
  const matchColumn = findMatchUpdateColumn(header.columns, parsedRequest.matchHints, "match");
  if (!matchColumn) {
    return {
      ok: false,
      summary: `Could not resolve the match column for "${parsedRequest.matchHints.join(" / ")}".`,
      candidates: header.columns.slice(0, 8).map((column) => columnCandidate(sheet.sheetName, header, column)),
      warnings: ["Ask the user which column to search before reading chunks."]
    };
  }
  const updateColumn = findMatchUpdateColumn(header.columns, parsedRequest.updateHints, "update");
  if (!updateColumn) {
    return {
      ok: false,
      summary: `Could not resolve the update column for "${parsedRequest.updateHints.join(" / ")}".`,
      candidates: header.columns.slice(0, 8).map((column) => columnCandidate(sheet.sheetName, header, column)),
      warnings: ["Ask the user which column to update before trying cell writes."]
    };
  }
  const searchedRange = matchUpdateSearchRange(sheet.usedRange, header);
  if (!searchedRange) {
    return {
      ok: false,
      summary: `Could not resolve a bounded search range on ${sheet.sheetName}.`,
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      warnings: ["Ask the user for a specific table or range."]
    };
  }
  return {
    ok: true,
    sheetName: sheet.sheetName,
    searchedRange,
    headerRow: header.row,
    headers: header.columns,
    matchColumn,
    predicate: parsedRequest.predicate,
    updateColumn,
    updateValue: parsedRequest.updateValue
  };
}

function resolveMatchUpdateRequest(request: string, values?: AgentRunInput["values"]): {
  matchHints: string[];
  predicate: MatchUpdatePredicate;
  updateHints: string[];
  updateValue: unknown;
} | undefined {
  if (values && hasStructuredMutationPayload({ request, values })) {
    return undefined;
  }
  const text = request.trim();
  const truckMatch = text.match(/\b(?:all\s+)?(?:transfer|transfers|transfer\s+from\/to|payer|payee|vendor|customer|account)\s+(?:to|from|contains?|for)?\s*["']?([A-Z0-9][\p{L}\p{M}\p{N}\s._&+-]{1,60}?)["']?\s+(?:is|are|=|as|to)\s+(?:truck(?:\s+id)?|vehicle)\s+["']?([\w-]{2,20})["']?/iu);
  if (truckMatch?.[1] && truckMatch[2]) {
    return {
      matchHints: ["Transfer From/To", "transfer", "payer", "payee", "account"],
      predicate: { operator: "contains", value: cleanupMatchText(truckMatch[1]) },
      updateHints: ["Truck ID", "truck", "vehicle"],
      updateValue: truckMatch[2].trim()
    };
  }
  const setWhereMatch = text.match(/\b(?:where|rows?\s+where|all\s+rows?\s+where)\s+([\p{L}\p{M}\p{N}\s/_-]{2,40})\s+(contains?|equals?|=|is)\s+["']?([^"',]+?)["']?\s*,?\s*(?:set|update|write)\s+([\p{L}\p{M}\p{N}\s/_-]{2,40})\s+(?:to|=|as)\s+["']?([^"']+?)["']?$/iu);
  if (setWhereMatch?.[1] && setWhereMatch[3] && setWhereMatch[4] && setWhereMatch[5]) {
    return {
      matchHints: [setWhereMatch[1].trim()],
      predicate: { operator: /equal|=|is/i.test(setWhereMatch[2] ?? "") ? "equals" : "contains", value: cleanupMatchText(setWhereMatch[3]) },
      updateHints: [setWhereMatch[4].trim()],
      updateValue: cleanupMatchText(setWhereMatch[5])
    };
  }
  return undefined;
}

function resolveExactRowSearchRequest(input: AgentRunInput): { predicate: MatchUpdatePredicate; columnOffset?: number } | undefined {
  if (intentAction(input) === "search_range") {
    return undefined;
  }
  if (resolveMatchUpdateRequest(input.request, input.values)) {
    return undefined;
  }
  const request = input.request;
  if (!/\b(search|find|where|contains?|matching|matches)\b/i.test(request)) {
    return undefined;
  }
  const valueMatch = request.match(/\bcontain(?:s|ing)?\s+["']?([^"',]+?)["']?(?:\s|$)/i)
    ?? request.match(/\b(?:for|matching|match(?:es)?)\s+["']?([A-Z0-9][\p{L}\p{M}\p{N}\s._&+-]{1,60}?)["']?(?:\s|$)/iu);
  const value = valueMatch?.[1] ? cleanupMatchText(valueMatch[1]) : undefined;
  if (!value || COMMON_EXACT_SEARCH_VALUES.has(value.toLowerCase())) {
    return undefined;
  }
  const columnMatch = request.match(/\bcolumn\s+([A-Z]{1,3})\b/i);
  let columnOffset: number | undefined;
  if (columnMatch?.[1] && input.target?.range) {
    const parsed = tryParseA1Address(stripSheetName(input.target.range));
    const column = columnToNumber(columnMatch[1].toUpperCase());
    if (parsed && column >= parsed.startColumn && column <= parsed.endColumn) {
      columnOffset = column - parsed.startColumn;
    }
  }
  return stripUndefinedRecord({
    predicate: { operator: "contains" as const, value },
    columnOffset
  }) as { predicate: MatchUpdatePredicate; columnOffset?: number };
}

const COMMON_EXACT_SEARCH_VALUES = new Set(["row", "rows", "column", "data", "value", "values", "all"]);

function cleanupMatchText(value: string): string {
  return value.replace(/\s+(?:is|are|as|to|for)$/i, "").trim();
}

function resolveMatchUpdateSheet(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { ok: true; sheetName: string; usedRange: string } | { ok: false; summary: string; candidates?: AgentCandidate[]; warnings: string[] } {
  if (input.target?.sheetName) {
    const exact = metadata.sheets.find((sheet) => normalizeComparableText(sheet.name) === normalizeComparableText(input.target!.sheetName!));
    if (exact?.usedRange) {
      return { ok: true, sheetName: exact.name, usedRange: exact.usedRange };
    }
  }
  const request = normalizeComparableText(input.request);
  const mentioned = metadata.sheets.filter((sheet) => sheet.usedRange && request.includes(normalizeComparableText(sheet.name)));
  if (mentioned.length === 1) {
    return { ok: true, sheetName: mentioned[0]!.name, usedRange: mentioned[0]!.usedRange! };
  }
  const transactionSheets = metadata.sheets.filter((sheet) => sheet.usedRange && sheet.kind === "transaction");
  if (transactionSheets.length === 1) {
    return { ok: true, sheetName: transactionSheets[0]!.name, usedRange: transactionSheets[0]!.usedRange! };
  }
  const candidates: AgentCandidate[] = (mentioned.length > 1 ? mentioned : transactionSheets)
    .slice(0, 8)
    .flatMap((sheet) => typeof sheet.usedRange === "string" ? [{
      id: sheet.id,
      kind: "sheet" as const,
      label: sheet.name,
      sheetName: sheet.name,
      range: sheet.usedRange,
      confidence: mentioned.includes(sheet) ? 0.95 : 0.7
    }] : []);
  return {
    ok: false,
    summary: "Match/update needs one sheet or table scope before searching rows.",
    ...(candidates.length > 0 ? { candidates } : {}),
    warnings: ["Ask the user which sheet to update; do not broad-read workbook chunks."]
  };
}

function bestHeaderForSheet(metadata: WorkbookMetadata, sheetName: string, hints: string[] = []): HeaderMetadata | undefined {
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  return sheet?.headers
    .slice()
    .sort((left, right) => headerMatchUpdateScore(right, hints) - headerMatchUpdateScore(left, hints) || right.confidence - left.confidence || left.row - right.row)[0];
}

function headerMatchUpdateScore(header: HeaderMetadata, hints: string[]): number {
  return header.columns.reduce((total, column) => total + matchUpdateColumnScore(column, hints, "match"), 0);
}

function findMatchUpdateColumn(columns: HeaderMetadata["columns"], hints: string[], role: "match" | "update"): MatchUpdateColumn | undefined {
  const scored = columns
    .map((column) => ({ column, score: matchUpdateColumnScore(column, hints, role) }))
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score || left.column.index - right.column.index);
  const winner = scored[0];
  if (!winner) {
    return undefined;
  }
  return { name: winner.column.name || winner.column.letter, index: winner.column.index, letter: winner.column.letter };
}

function matchUpdateColumnScore(column: HeaderMetadata["columns"][number], hints: string[], role: "match" | "update"): number {
  const haystack = normalizeComparableText([column.name, column.normalizedName, column.role, column.letter].filter(Boolean).join(" "));
  let score = 0;
  for (const hint of hints) {
    const normalizedHint = normalizeComparableText(hint);
    if (!normalizedHint) continue;
    if (haystack === normalizedHint) score += 20;
    if (haystack.includes(normalizedHint) || normalizedHint.includes(haystack)) score += 10;
    for (const token of normalizedHint.split(/\s+/).filter((entry) => entry.length > 2)) {
      if (haystack.includes(token)) score += 3;
    }
  }
  if (role === "match" && /\btransfer|from|to|payer|payee|account|vendor|customer|description|note\b/.test(haystack)) score += 6;
  if (role === "update" && /\btruck|vehicle|id|status|type|description|note\b/.test(haystack)) score += 6;
  return score;
}

function columnCandidate(sheetName: string, header: HeaderMetadata, column: HeaderMetadata["columns"][number]): AgentCandidate {
  return stripUndefinedRecord({
    id: `column:${sheetName}:${column.letter}`,
    kind: "range",
    label: column.name || column.letter,
    sheetName,
    range: `${column.letter}${header.row}:${column.letter}${header.row}`,
    semanticRole: column.role,
    confidence: column.importance ?? header.confidence
  }) as AgentCandidate;
}

function matchUpdateSearchRange(usedRange: string, header: HeaderMetadata): string | undefined {
  const parsed = tryParseA1Address(stripSheetName(usedRange));
  if (!parsed) {
    return undefined;
  }
  const startColumn = Math.min(...header.columns.map((column) => column.index + 1), parsed.startColumn);
  const endColumn = Math.max(...header.columns.map((column) => column.index + 1), parsed.endColumn);
  return `${numberToColumn(startColumn)}${header.row}:${numberToColumn(endColumn)}${parsed.endRow}`;
}

function cellMatchesPredicate(value: unknown, predicate: MatchUpdatePredicate): boolean {
  const actual = normalizeComparableText(String(value ?? ""));
  const expected = normalizeComparableText(predicate.value);
  if (!actual || !expected) {
    return false;
  }
  return predicate.operator === "equals" ? actual === expected : actual.includes(expected);
}

function valueByHeader(headers: HeaderMetadata["columns"], row: CellMatrix[number], pattern: RegExp): unknown {
  const column = headers.find((candidate) => pattern.test(candidate.name) || pattern.test(candidate.normalizedName));
  return column ? row[column.index] : undefined;
}

function metadataSuggestsRangeHasData(metadata: WorkbookMetadata, sheetName: string, range: string): boolean {
  const normalizedSheet = normalizeComparableText(sheetName);
  const normalizedRange = normalizeAddressForCompare(range);
  const parsed = tryParseA1Address(normalizedRange);
  if (!parsed) {
    return false;
  }
  const overlaps = (candidateRange: string | undefined) => {
    if (!candidateRange) return false;
    return rangesOverlapAddresses(normalizedRange, normalizeAddressForCompare(candidateRange));
  };
  if (metadata.tables.some((table) => normalizeComparableText(table.sheetName) === normalizedSheet && overlaps(table.range))) {
    return true;
  }
  const sheet = metadata.sheets.find((candidate) => normalizeComparableText(candidate.name) === normalizedSheet);
  if (!sheet) {
    return false;
  }
  if (sheet.headers.some((header) => overlaps(header.range))) {
    return true;
  }
  if (sheet.usedRange && overlaps(sheet.usedRange) && parsed.startRow <= 2) {
    return true;
  }
  return false;
}

function workbookDumpGuardOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  if (!isWorkbookDumpRequest(input.request)) {
    return undefined;
  }
  const summary = workbookCompactSummary(metadata);
  const askToNarrow = isEveryCellDumpRequest(input.request);
  return {
    status: askToNarrow ? "NEEDS_INPUT" : "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: askToNarrow
      ? "Full workbook cell dumps are not returned inline. Choose a sheet, table, or range, or ask for a compact workbook summary."
      : `Returned a compact workbook summary for ${metadata.workbook.name} instead of dumping every cell.`,
    answer: {
      kind: "workbook_dump_guard",
      source: "cached_metadata",
      workbook: metadata.workbook,
      sheetCount: metadata.sheets.length,
      tableCount: metadata.tables.length,
      namedRangeCount: metadata.namedRanges.length,
      sectionCount: metadata.sections.length,
      sheets: summary.sheets,
      alternatives: ["summary", "specific sheet", "specific range"],
      refusedFullCellDump: true
    },
    metrics: { source: "cached_metadata", sheetCount: metadata.sheets.length, tableCount: metadata.tables.length, fullReadCellCount: 0 },
    proof: metadata.sheets.slice(0, 5).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "used range" }] : []),
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: askToNarrow ? "ask_user" : "answer_now",
    warnings: ["Full workbook cell dumps are intentionally blocked to avoid excessive context; use workbookContextId or a scoped target for follow-up reads."]
  };
}

function emptyResultKindForProfile(
  metadata: WorkbookMetadata,
  sheetName: string,
  range: string,
  profile: ReturnType<typeof profileValues>
): "fresh_sheet" | "empty_range" | "trailing_empty_rows" | undefined {
  if (profile.metrics.nonEmptyCount !== 0) {
    const trailingRows = typeof profile.emptySummary?.trailingRows === "number" ? profile.emptySummary.trailingRows : 0;
    return trailingRows > 0 ? "trailing_empty_rows" : undefined;
  }
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const normalizedRange = stripSheetName(range);
  if (sheet?.usedRange
    && stripSheetName(sheet.usedRange) === normalizedRange
    && sheet.tableIds.length === 0
    && sheet.headers.length === 0
    && sheet.sectionIds.length === 0) {
    return "fresh_sheet";
  }
  return "empty_range";
}

function largeRangeGuardOutput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  requestedMode: AgentRunMode,
  sheetName: string,
  range: string,
  label: string
): Omit<AgentRunOutput, "telemetry"> | undefined {
  const requestedCells = cellCountFromAddress(range);
  if (requestedCells === undefined || requestedCells <= AGENT_LARGE_RANGE_CELL_LIMIT) {
    return undefined;
  }
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const searchLike = /\b(where|contains?|equals?|matching|find|search|rows?\s+where|filter)\b/i.test(input.request);
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: searchLike
      ? `The requested search over ${sheetName}!${range} is too large for an inline read; use a targeted search intent instead of broad-reading the range.`
      : `Returned a compact summary for ${sheetName}!${range}; the requested ${requestedCells.toLocaleString()} cells exceed the inline read limit.`,
    answer: {
      kind: "large_range_guard",
      source: "cached_metadata",
      sheetName,
      range,
      requestedCells,
      inlineCellLimit: AGENT_LARGE_RANGE_CELL_LIMIT,
      usedRange: sheet?.usedRange,
      rowCount: sheet?.rowCount,
      columnCount: sheet?.columnCount,
      recommendation: searchLike
        ? "Retry once with intent.action search_range or find_similar_rows and the specific column/text predicates; do not fetch fullResultUri or broad-read the sheet."
        : isExplicitFullDataRequest(input.request)
          ? "Ask for a smaller range, specific columns, filters, or an export-style workflow."
          : "Ask for a smaller range, a table, specific columns, or a compact schema/profile.",
      suggestedIntentAction: searchLike ? "search_range" : undefined,
      suggestedNextRequest: searchLike ? `Search ${sheetName} with exact column/text predicates and return bounded exact matching rows.` : undefined
    },
    metrics: { source: "cached_metadata", requestedCells, inlineCellLimit: AGENT_LARGE_RANGE_CELL_LIMIT, fullReadCellCount: 0 },
    candidates: findAgentCandidates(metadata, input).slice(0, 5),
    proof: [{ sheetName, range: sheet?.usedRange ?? range, label }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: (searchLike || isExplicitFullDataRequest(input.request)) ? "ask_user" : "answer_now",
    warnings: [searchLike
      ? "Large search-like read was blocked; ask the user for exact predicates or call one targeted search intent. Do not broad-read follow-up chunks."
      : isExplicitFullDataRequest(input.request)
        ? "Large full-data read was blocked; ask the user for a smaller range, specific columns, filters, or an export-style workflow."
        : "Large range read was summarized from cached metadata without reading cell values."]
  };
}

function selectionAwareProof(
  metadata: WorkbookMetadata,
  resolved: AgentTargetResolution & { ok: true },
  range: string,
  label: string
): AgentProofReference[] {
  const proof: AgentProofReference[] = [{ sheetName: resolved.sheetName, range, label }];
  const selection = metadata.selection;
  if (
    resolved.candidate.id === "selection:implicit" &&
    selection?.sheetName === resolved.sheetName &&
    selection.address &&
    selection.address !== range
  ) {
    proof.push({ sheetName: selection.sheetName, range: selection.address, label: "current Excel selection" });
  }
  return proof;
}

function workbookCompactSummary(metadata: WorkbookMetadata): { sheets: Array<Record<string, unknown>> } {
  return {
    sheets: metadata.sheets.map((sheet) => stripUndefinedRecord({
      name: sheet.name,
      kind: sheet.kind,
      usedRange: sheet.usedRange,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      tableCount: sheet.tableIds.length,
      sectionCount: sheet.sectionIds.length
    }))
  };
}

function workbookOverviewSummary(
  metadata: WorkbookMetadata,
  intent: { aboutIntent: boolean; tableIntent: boolean; namedRangeIntent: boolean; blankIntent: boolean; sheetCountIntent: boolean; sheetListIntent: boolean },
  blankSheetCount: number
): string {
  if (intent.tableIntent) {
    return `This workbook has ${metadata.tables.length} table(s): ${metadata.tables.map((table) => table.name ?? table.id).join(", ") || "none"}.`;
  }
  if (intent.namedRangeIntent) {
    return `This workbook has ${metadata.namedRanges.length} named range(s).`;
  }
  if (intent.blankIntent) {
    return `Found ${blankSheetCount} blank or mostly empty sheet(s).`;
  }
  if (intent.sheetCountIntent) {
    return `${metadata.workbook.name} has ${metadata.sheets.length} sheet(s).`;
  }
  if (intent.sheetListIntent) {
    return `Listed ${metadata.sheets.length} sheet(s) from cached workbook metadata.`;
  }
  return `${metadata.workbook.name} appears to be an Excel workbook with ${metadata.sheets.length} sheet(s), ${metadata.tables.length} table(s), and ${metadata.namedRanges.length} named range(s).`;
}

function sectionAnswerOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  if (!isSectionInventoryRequest(input.request)) {
    return undefined;
  }
  const sheet = resolveSectionSheet(metadata, input);
  if (!sheet) {
    const candidates = findAgentCandidates(metadata, input).filter((candidate) => candidate.kind === "sheet").slice(0, 5);
    return {
      status: candidates.length > 0 ? "AMBIGUOUS_TARGET" : "NEEDS_INPUT",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: "Section inventory needs one sheet target.",
      ...(candidates.length > 0 ? { candidates } : {}),
      proof: [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: candidates.length > 0 ? "call_with_target" : "ask_user",
      warnings: []
    };
  }
  const detectedSections = metadata.sections.filter((section) => section.sheetName === sheet.name);
  const namedSections = metadata.namedRanges
    .filter((name) => name.sheetName === sheet.name && /\b(section|block|area)\b/i.test(name.name.replace(/([a-z])([A-Z])/g, "$1 $2")))
    .map((name, index) => ({
      id: `name:${name.name}`,
      sheetName: name.sheetName ?? sheet.name,
      label: name.name.replace(/([a-z])([A-Z])/g, "$1 $2"),
      kind: "region" as const,
      range: name.range,
      headerRange: undefined,
      columns: [],
      rowCount: undefined,
      confidence: 0.7 - index * 0.01
    }));
  const sections = [
    ...detectedSections.map((section) => ({
      id: section.id,
      sheetName: section.sheetName,
      label: section.label,
      kind: section.kind,
      range: section.range,
      headerRange: section.headerRange,
      columns: section.columns,
      rowCount: section.rowCount,
      confidence: section.confidence
    })),
    ...namedSections
  ];
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `Found ${sections.length} section(s) on ${sheet.name} from bounded sheet metadata.`,
    answer: {
      kind: "sheet_sections",
      source: "cached_metadata",
      sheetName: sheet.name,
      sectionCount: sections.length,
      sections: sections.map((section) => ({
        id: section.id,
        label: section.label,
        kind: section.kind,
        range: section.range,
        headerRange: section.headerRange,
        columnCount: section.columns.length,
        rowCount: section.rowCount,
        confidence: section.confidence
      }))
    },
    metrics: { source: "cached_metadata", sectionCount: sections.length },
    candidates: sections.map((section) => ({
      id: section.id,
      kind: "region" as const,
      label: section.label,
      sheetName: section.sheetName,
      range: section.range,
      confidence: section.confidence
    })).slice(0, input.budget?.maxExamples ?? 10),
    proof: sections.slice(0, 5).map((section) => ({ sheetName: section.sheetName, range: section.range, label: section.label })),
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: metadata.detailLevel === "sampled" ? [] : ["Section inventory is best after sampled metadata is available."]
  };
}

function isSectionInventoryRequest(request: string): boolean {
  if (/\b(rows?|records?|data|sample|examples?|values?|contents?)\b/i.test(request)) {
    return false;
  }
  return /\b(how many|what|which|list|show|inventory|count)\b.*\b(sections?|blocks?|areas?)\b/i.test(request)
    || /\b(sections?|blocks?|areas?)\b.*\b(on|in)\b.*\bsheet\b/i.test(request);
}

function resolveSectionSheet(metadata: WorkbookMetadata, input: AgentRunInput): WorkbookMetadata["sheets"][number] | undefined {
  if (input.target?.sheetName) {
    const normalized = normalizeComparableText(input.target.sheetName);
    return metadata.sheets.find((sheet) => normalizeComparableText(sheet.name) === normalized);
  }
  if (requestMentionsActiveSheetForSections(input.request) && metadata.workbook.activeSheet) {
    return metadata.sheets.find((sheet) => sheet.name === metadata.workbook.activeSheet);
  }
  const request = normalizeComparableText(input.request);
  const mentioned = metadata.sheets.filter((sheet) => request.includes(normalizeComparableText(sheet.name)));
  if (mentioned.length === 1) {
    return mentioned[0];
  }
  return metadata.sheets.length === 1 ? metadata.sheets[0] : undefined;
}

function requestMentionsActiveSheetForSections(request: string): boolean {
  return /\b(active|current|this)\s+sheet\b/i.test(request) || /\bthis\s+sheet\b/i.test(request);
}

function resolveComparisonTargets(metadata: WorkbookMetadata, request: string): Array<{ sheetName: string; range: string }> {
  if (!/\b(compare|versus|vs\.?|higher|lower|more|less|between|both)\b/i.test(request)) {
    return [];
  }
  const normalizedRequest = normalizeComparableText(request);
  const exactMatches = metadata.sheets
    .filter((sheet) => sheet.usedRange && normalizedRequest.includes(normalizeComparableText(sheet.name)))
    .sort((left, right) => normalizeComparableText(right.name).length - normalizeComparableText(left.name).length)
    .filter((sheet, index, sheets) => {
      const normalizedSheet = normalizeComparableText(sheet.name);
      return !sheets.slice(0, index).some((other) => normalizeComparableText(other.name).includes(normalizedSheet));
    })
    .map((sheet) => ({ sheetName: sheet.name, range: preferredComparisonRange(metadata, sheet, request) ?? sheet.usedRange! }));
  if (exactMatches.length >= 2) {
    return dedupeBy(exactMatches, (target) => target.sheetName).slice(0, 4);
  }
  const matches = metadata.sheets
    .filter((sheet) => sheet.usedRange && requestMentionsComparisonSheet(normalizedRequest, sheet.name))
    .map((sheet) => ({ sheetName: sheet.name, range: preferredComparisonRange(metadata, sheet, request) ?? sheet.usedRange! }));
  return dedupeBy(matches, (target) => target.sheetName).slice(0, 4);
}

function requestMentionsComparisonSheet(normalizedRequest: string, sheetName: string): boolean {
  const normalizedSheetName = normalizeComparableText(sheetName);
  if (normalizedRequest.includes(normalizedSheetName)) {
    return true;
  }
  const month = /^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*(?:\s+\d{4})?$/i.exec(sheetName.trim())?.[1]?.toLowerCase();
  if (!month) {
    return false;
  }
  const aliases = month === "sept" ? ["sep", "sept", "september"] : monthAliases(month);
  return aliases.some((alias) => new RegExp(`\\b${alias}\\b`, "i").test(normalizedRequest));
}

function monthAliases(month: string): string[] {
  const aliases: Record<string, string[]> = {
    jan: ["jan", "january"],
    feb: ["feb", "february"],
    mar: ["mar", "march"],
    apr: ["apr", "april"],
    may: ["may"],
    jun: ["jun", "june"],
    jul: ["jul", "july"],
    aug: ["aug", "august"],
    sep: ["sep", "sept", "september"],
    oct: ["oct", "october"],
    nov: ["nov", "november"],
    dec: ["dec", "december"]
  };
  return aliases[month] ?? [month];
}

function preferredComparisonRange(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number], request: string): string | undefined {
  if (!/\b(performance|perform|business|summary|summarize|kpi|metric|cash|profit|loss|revenue|expense|spend|spent|received|billed)\b/i.test(request)) {
    return undefined;
  }
  const monthlySummaryRange = monthlySummaryColumnRange(metadata, sheet);
  if (monthlySummaryRange) {
    return monthlySummaryRange;
  }
  const sheetSections = metadata.sections.filter((section) => section.sheetName === sheet.name && section.range);
  const ranked = sheetSections
    .map((section) => ({ section, score: comparisonSectionScore(section) }))
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score || left.section.range.localeCompare(right.section.range));
  const best = ranked[0]?.section;
  if (!best) {
    return undefined;
  }
  return mergedSectionColumnRange(sheetSections, best.range) ?? best.range;
}

function monthlySummaryColumnRange(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number]): string | undefined {
  if (!/\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{4}\b/i.test(sheet.name) || !sheet.usedRange) {
    return undefined;
  }
  const used = tryParseA1Address(sheet.usedRange);
  if (!used || used.endColumn < columnToNumber("AJ")) {
    return undefined;
  }
  const summaryColumnSections = metadata.sections
    .filter((section) => section.sheetName === sheet.name)
    .filter((section) => {
      const parsed = tryParseA1Address(section.range);
      return parsed && parsed.startColumn >= columnToNumber("AG") && parsed.endColumn <= columnToNumber("AJ");
    });
  const merged = mergedSectionColumnRange(summaryColumnSections, summaryColumnSections[0]?.range ?? "");
  if (merged) {
    return merged;
  }
  return `AG1:AJ${Math.min(used.endRow, 20)}`;
}

function comparisonSectionScore(section: WorkbookMetadata["sections"][number]): number {
  const text = normalizeComparableText([
    section.label,
    section.kind,
    ...section.labels,
    ...section.columns.map((column) => column.name)
  ].join(" "));
  let score = 0;
  const hasPerformanceTerms = /kpi|summary|metric|profit|loss|revenue|expense|cash|spend/.test(text);
  if (section.kind === "summary" && hasPerformanceTerms) score += 5;
  if (hasPerformanceTerms) score += 4;
  if (/invoice|transaction|truck|container|proof|booking/.test(text)) score -= 6;
  if (score > 0) {
    score += Math.min(2, section.confidence ?? 0);
  }
  return score;
}

function mergedSectionColumnRange(sections: WorkbookMetadata["sections"], bestRange: string): string | undefined {
  const best = tryParseA1Address(bestRange);
  if (!best) {
    return undefined;
  }
  const sameColumnSections = sections
    .map((section) => ({ section, parsed: tryParseA1Address(section.range) }))
    .filter((entry): entry is { section: WorkbookMetadata["sections"][number]; parsed: NonNullable<ReturnType<typeof tryParseA1Address>> } =>
      Boolean(entry.parsed && entry.parsed.startColumn === best.startColumn && entry.parsed.endColumn === best.endColumn)
    );
  if (sameColumnSections.length <= 1) {
    return undefined;
  }
  const startRow = Math.min(...sameColumnSections.map((entry) => entry.parsed.startRow));
  const endRow = Math.max(...sameColumnSections.map((entry) => entry.parsed.endRow));
  return `${numberToColumn(best.startColumn)}${startRow}:${numberToColumn(best.endColumn)}${endRow}`;
}

function comparisonAnswerOutput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  requestedMode: AgentRunMode,
  profiles: Array<{ sheetName: string; range: string; profile: ReturnType<typeof profileValues> }>
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "SUCCESS",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `Compared ${profiles.map((profile) => profile.sheetName).join(" and ")} using targeted compact reads.`,
    answer: {
      kind: "comparison_profile",
      source: "live_read",
      request: input.request,
      sheets: profiles.map((item) => ({
        sheetName: item.sheetName,
        range: item.range,
        shape: item.profile.shape,
        metrics: item.profile.metrics,
        sample: item.profile.sample,
        ...(item.profile.emptySummary ? { emptySummary: item.profile.emptySummary } : {}),
        ...(item.profile.sparseRows ? { sparseRows: item.profile.sparseRows } : {}),
        ...(item.profile.rows ? { rows: item.profile.rows } : {})
      })),
      numericComparison: compareNumericMetrics(profiles),
      alignedRows: alignComparisonRows(profiles)
    },
    metrics: { comparedSheetCount: profiles.length, source: "live_read" },
    proof: profiles.map((profile) => ({ sheetName: profile.sheetName, range: profile.range, label: "comparison source" })),
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: []
  };
}

function adjustReadRangeForSemanticColumn(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  sheetName: string,
  range: string
): { range: string; column?: ColumnMetadata; warnings: string[] } {
  const match = resolveRequestedColumn(metadata, input, sheetName);
  if (!match) {
    return { range, warnings: [] };
  }
  const current = columnRangeBounds(range, metadata, sheetName);
  if (!current) {
    return { range, column: match.column, warnings: [] };
  }
  const targetColumnNumber = match.column.index + 1;
  const targetRange = `${match.column.letter}${current.startRow}:${match.column.letter}${current.endRow}`;
  const sameSingleColumn = current.startColumn === targetColumnNumber && current.endColumn === targetColumnNumber;
  if (sameSingleColumn) {
    return { range, column: match.column, warnings: [] };
  }
  if ((current.startColumn === current.endColumn || input.target?.column) && match.confidence >= 0.86) {
    return {
      range: targetRange,
      column: match.column,
      warnings: [`Adjusted target range from ${range} to ${targetRange} because "${match.column.name}" matched column ${match.column.letter}.`]
    };
  }
  return { range, column: match.column, warnings: [] };
}

function resolveRequestedColumn(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  sheetName: string
): { column: ColumnMetadata; confidence: number } | undefined {
  const requestedNames = requestedColumnNames(input);
  if (requestedNames.length === 0) {
    return undefined;
  }
  const columns = metadataColumnsForSheet(metadata, sheetName);
  if (columns.length === 0) {
    return undefined;
  }
  const scored = columns
    .map((column) => ({ column, confidence: requestedNames.reduce((best, name) => Math.max(best, columnNameMatchScore(name, column)), 0) }))
    .filter((entry) => entry.confidence >= 0.86)
    .sort((left, right) => right.confidence - left.confidence || left.column.index - right.column.index);
  const best = scored[0];
  const second = scored[1];
  if (!best || (second && best.confidence === second.confidence && best.column.normalizedName === second.column.normalizedName)) {
    return undefined;
  }
  return best;
}

function requestedColumnNames(input: AgentRunInput): string[] {
  const values = [
    input.target?.column,
    ...(((input.intent as { targetHints?: string[] } | undefined)?.targetHints ?? [])),
    ...quotedColumnNames(input.request),
    ...namedColumnPhrases(input.request)
  ].filter((value): value is string => typeof value === "string" && value.trim().length > 0);
  return [...new Set(values.map((value) => value.trim()))];
}

function quotedColumnNames(request: string): string[] {
  return [...request.matchAll(/["“']([^"“”']{2,80})["”']/g)].map((match) => match[1]!).filter((value) => /\s|\/|_/.test(value));
}

function namedColumnPhrases(request: string): string[] {
  const phrases: string[] = [];
  const patterns = [
    /\b(?:column|field|header)\s+(?:named|called)\s+([^.,;:\n]{2,80})/gi,
    /\b([^.,;:\n]{2,80})\s+(?:column|field|header)\b/gi
  ];
  for (const pattern of patterns) {
    for (const match of request.matchAll(pattern)) {
      const value = match[1]?.trim();
      if (value && !/^[A-Z]{1,3}$/i.test(value)) {
        const cleaned = value.replace(/\b(?:the|a|an|specifically|from|to|for|all|unique|non-empty|nonempty|values?|count|counts?|how|many)\b/gi, " ").replace(/\s+/g, " ").trim();
        phrases.push(cleaned);
        const tail = /\b(?:in|on|by|from)\s+(.+)$/i.exec(cleaned)?.[1]?.trim();
        if (tail) {
          phrases.push(tail);
        }
      }
    }
  }
  return phrases.filter((value) => value.length >= 2);
}

function metadataColumnsForSheet(metadata: WorkbookMetadata, sheetName: string): ColumnMetadata[] {
  const byKey = new Map<string, ColumnMetadata>();
  const add = (column: ColumnMetadata) => {
    const key = `${column.index}:${normalizeHeaderName(column.name)}`;
    if (!byKey.has(key)) {
      byKey.set(key, column);
    }
  };
  metadata.tables.filter((table) => table.sheetName === sheetName).forEach((table) => table.columns.forEach(add));
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  sheet?.headers.forEach((header) => header.columns.forEach(add));
  metadata.sections.filter((section) => section.sheetName === sheetName).forEach((section) => section.columns.forEach(add));
  return [...byKey.values()];
}

function columnNameMatchScore(requestedName: string, column: ColumnMetadata): number {
  const requested = normalizeHeaderName(requestedName);
  const columnName = normalizeHeaderName(column.name);
  if (!requested || !columnName || /^[a-z]{1,3}$/i.test(requestedName.trim())) {
    return 0;
  }
  if (requested === columnName || requested === column.normalizedName || requested.replace(/_/g, "") === columnName.replace(/_/g, "")) {
    return 1;
  }
  const requestedTokens = new Set(headerTokens(requestedName));
  const columnTokens = headerTokens(column.name);
  if (requestedTokens.size === 0 || columnTokens.length === 0) {
    return 0;
  }
  const overlap = columnTokens.filter((token) => requestedTokens.has(token)).length;
  if (overlap === 0) {
    return 0;
  }
  const precision = overlap / columnTokens.length;
  const recall = overlap / requestedTokens.size;
  return 0.2 + precision * 0.45 + recall * 0.35;
}

function headerTokens(value: string): string[] {
  return normalizeHeaderName(value.replace(/[\/_-]+/g, " ")).split("_").filter(Boolean);
}

function columnRangeBounds(range: string, metadata: WorkbookMetadata, sheetName: string): { startColumn: number; endColumn: number; startRow: number; endRow: number } | undefined {
  const normalized = stripSheetName(range).replace(/\$/g, "").toUpperCase();
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const defaultEndRow = Math.max(1, sheet?.rowCount ?? 1);
  const cellRange = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(normalized);
  if (cellRange?.[1] && cellRange[2]) {
    const endColumn = cellRange[3] ?? cellRange[1];
    const endRow = cellRange[4] ?? cellRange[2];
    return {
      startColumn: columnToNumber(cellRange[1]),
      endColumn: columnToNumber(endColumn),
      startRow: Number(cellRange[2]),
      endRow: Number(endRow)
    };
  }
  const columnOnly = /^([A-Z]{1,3})(?::([A-Z]{1,3}))?$/.exec(normalized);
  if (columnOnly?.[1]) {
    return {
      startColumn: columnToNumber(columnOnly[1]),
      endColumn: columnToNumber(columnOnly[2] ?? columnOnly[1]),
      startRow: 1,
      endRow: defaultEndRow
    };
  }
  return undefined;
}

function aggregateProfileForRequest(
  input: AgentRunInput,
  profile: ReturnType<typeof profileValues>,
  column?: ColumnMetadata
): Record<string, unknown> | undefined {
  if (!isAggregateValueRequest(input.request)) {
    return undefined;
  }
  const values = profileValuesForAggregation(profile, column);
  if (values.length === 0) {
    return undefined;
  }
  const counts = new Map<string, { value: unknown; count: number }>();
  for (const value of values) {
    const key = String(value);
    const current = counts.get(key);
    counts.set(key, { value, count: (current?.count ?? 0) + 1 });
  }
  const valueCounts = [...counts.values()]
    .sort((left, right) => right.count - left.count || String(left.value).localeCompare(String(right.value)));
  return stripUndefinedRecord({
    kind: "range_value_counts",
    source: "live_read",
    column: column ? { name: column.name, letter: column.letter, index: column.index } : undefined,
    uniqueCount: counts.size,
    valueCount: values.length,
    valueCounts
  });
}

function isAggregateValueRequest(request: string): boolean {
  return /\b(unique|distinct|dedupe|deduplicate|count|counts?|how many|frequency|frequencies|group by)\b/i.test(request);
}

function profileValuesForAggregation(profile: ReturnType<typeof profileValues>, column?: ColumnMetadata): unknown[] {
  const rows = Array.isArray(profile.rows)
    ? profile.rows as unknown[][]
    : Array.isArray(profile.sparseRows)
      ? (profile.sparseRows as Array<{ cells?: Array<{ value?: unknown }> }>).flatMap((row) => row.cells?.map((cell) => cell.value) ?? [])
      : [];
  const values = rows.flat().filter((value) => value !== null && value !== undefined && value !== "");
  if (column && values.length > 0 && normalizeHeaderName(String(values[0])) === normalizeHeaderName(column.name)) {
    return values.slice(1);
  }
  return values;
}

function connectionReadinessSummary(connectionState: string): string {
  if (connectionState === "stale") {
    return "Open Workbook has a stale Excel add-in session, so agents cannot reach the active workbook.";
  }
  if (connectionState === "connected_no_workbook") {
    return "Open Workbook is connected to Excel, but there is no active workbook.";
  }
  return "Open Workbook backend is running, but no Excel add-in session is connected.";
}

function connectionReadinessWarnings(connectionState: string): string[] {
  if (connectionState === "stale") {
    return ["Reload or reopen the OpenWorkbook Local taskpane in Excel, then retry. Restart Excel only if the taskpane cannot reconnect."];
  }
  if (connectionState === "connected_no_workbook") {
    return ["Open or activate a workbook in Excel, then retry the request."];
  }
  return ["Open Excel and load the OpenWorkbook Local taskpane, then retry the request."];
}

function activeWorkbookIdFromStatus(status: unknown): WorkbookId | undefined {
  const activeWorkbook = (status as { activeWorkbook?: { workbookId?: unknown } } | undefined)?.activeWorkbook;
  return typeof activeWorkbook?.workbookId === "string" ? activeWorkbook.workbookId as WorkbookId : undefined;
}

function compactAgentCollaborationSummary(collaboration: unknown): Record<string, unknown> {
  const typed = collaboration as {
    agents?: Array<Record<string, unknown>>;
    tasks?: Array<Record<string, unknown>>;
    locks?: Array<Record<string, unknown>>;
    transactions?: Array<Record<string, unknown>>;
    conflicts?: Array<Record<string, unknown>>;
    events?: Array<Record<string, unknown>>;
  };
  const agents = Array.isArray(typed.agents) ? typed.agents : [];
  const tasks = Array.isArray(typed.tasks) ? typed.tasks : [];
  const locks = Array.isArray(typed.locks) ? typed.locks : [];
  const transactions = Array.isArray(typed.transactions) ? typed.transactions : [];
  const conflicts = Array.isArray(typed.conflicts) ? typed.conflicts : [];
  const events = Array.isArray(typed.events) ? typed.events : [];
  const activeLocks = locks.filter((lock) => lock.status === "active");
  const activeTransactions = transactions.filter((transaction) => ["queued", "applying", "blocked"].includes(String(transaction.status ?? transaction.transactionStatus ?? "")));
  const openTasks = tasks.filter((task) => !["completed", "failed", "cancelled"].includes(String(task.status ?? "")));
  return {
    agentCount: agents.length,
    taskCount: tasks.length,
    openTaskCount: openTasks.length,
    activeTaskCount: openTasks.length,
    activeLockCount: activeLocks.length,
    activeTransactionCount: activeTransactions.length,
    conflictCount: conflicts.length,
    openConflictCount: conflicts.length,
    agents: agents.slice(0, 8).map((agent) => ({
      agentId: agent.agentId,
      agentName: agent.agentName,
      status: agent.status,
      clientType: agent.clientType,
      lastSeenAt: agent.lastSeenAt
    })),
    tasks: openTasks.slice(0, 8).map((task) => ({
      taskId: task.taskId,
      goal: task.goal,
      status: task.status,
      agentId: task.agentId ?? task.assignedAgentId,
      assignedAgentId: task.assignedAgentId,
      progress: task.progress,
      currentStep: task.currentStep
    })),
    locks: activeLocks.slice(0, 8).map((lock) => ({
      lockId: lock.lockId,
      ownerAgentId: lock.ownerAgentId,
      taskId: lock.taskId,
      transactionId: lock.transactionId,
      mode: lock.mode,
      scope: lock.scope,
      expiresAt: lock.expiresAt
    })),
    transactions: activeTransactions.slice(0, 8).map((transaction) => ({
      transactionId: transaction.transactionId,
      agentId: transaction.agentId,
      taskId: transaction.taskId,
      status: transaction.status ?? transaction.transactionStatus,
      queuePosition: transaction.queuePosition,
      progressMessage: transaction.progressMessage
    })),
    conflicts: conflicts.slice(0, 8).map((conflict) => ({
      conflictId: conflict.conflictId,
      code: conflict.code,
      message: conflict.message,
      ownerAgentId: conflict.ownerAgentId,
      taskId: conflict.taskId,
      transactionId: conflict.transactionId,
      lockId: conflict.lockId,
      lockExpiresAt: conflict.lockExpiresAt
    })),
    recentEvents: events.slice(0, 8).map((event) => ({
      eventId: event.eventId,
      type: event.type,
      agentId: event.agentId,
      taskId: event.taskId,
      transactionId: event.transactionId,
      message: event.message,
      createdAt: event.createdAt
    }))
  };
}

function profileValues(values: CellMatrix, address?: string) {
  const flattened = values.flat().filter((value) => value !== null && value !== undefined && value !== "");
  const numeric = flattened.map((value) => typeof value === "number" ? value : typeof value === "string" ? Number(value.replace(/[$,]/g, "")) : NaN).filter(Number.isFinite);
  const sum = numeric.reduce((total, value) => total + value, 0);
  const nonEmptyRows = values
    .map((row, rowIndex) => ({ row, rowIndex }))
    .filter((entry) => entry.row.some((value) => value !== null && value !== undefined && value !== ""));
  const includeRows = matrixCellCount(values) <= 500 || nonEmptyRows.length <= 75;
  const shape = { rows: values.length, columns: maxMatrixColumns(values) };
  const emptySummary = emptyMatrixSummary(values);
  const sparseRows = shouldReturnSparseRows(shape.rows * shape.columns, flattened.length, nonEmptyRows.length)
    ? sparseValueRows(values, address)
    : undefined;
  const sample = firstNonEmptyRows(values, 5).map((row) => compactSampleRow(row, 8));
  const rowMetadata = rangeRowMetadata(address, values, nonEmptyRows.map((entry) => entry.rowIndex));
  return {
    kind: "range_profile" as const,
    source: "live_read" as const,
    shape,
    metrics: {
      nonEmptyCount: flattened.length,
      numericCount: numeric.length,
      measureColumns: measureColumnMetrics(values),
      ...(numeric.length > 0 ? { sum, min: Math.min(...numeric), max: Math.max(...numeric), average: sum / numeric.length } : {})
    },
    sample,
    ...(emptySummary.emptyCells > 0 ? { emptySummary } : {}),
    ...(rowMetadata.length > 0 ? { rowMetadata } : {}),
    ...(sparseRows ? { sparseRows } : {}),
    ...(includeRows && !sparseRows && nonEmptyRows.length > 0 ? { rows: nonEmptyRows.map((entry) => trimTrailingEmptyCells(entry.row)) } : {}),
    warning: flattened.length === 0 ? "No non-empty cells were found in the targeted range." : undefined
  };
}

function measureColumnMetrics(values: CellMatrix): Array<Record<string, unknown>> | undefined {
  const headers = values[0] ?? [];
  if (headers.length === 0 || values.length <= 1) {
    return undefined;
  }
  const metrics = headers
    .map((header, index) => ({ header: String(header ?? "").trim(), index }))
    .filter((entry) => isMeasureColumnName(entry.header))
    .map((entry): Record<string, unknown> | undefined => {
      const numeric = values.slice(1)
        .map((row) => numericCellValue(row[entry.index]))
        .filter((value): value is number => value !== undefined);
      if (numeric.length === 0) {
        return undefined;
      }
      const sum = numeric.reduce((total, value) => total + value, 0);
      return stripUndefinedRecord({
        name: entry.header,
        letter: numberToColumn(entry.index + 1),
        count: numeric.length,
        sum,
        min: Math.min(...numeric),
        max: Math.max(...numeric),
        average: sum / numeric.length
      });
    })
    .filter((entry): entry is Record<string, unknown> => entry !== undefined);
  return metrics.length > 0 ? metrics : undefined;
}

function numericCellValue(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return undefined;
  }
  const normalized = Number(value.replace(/[$,]/g, ""));
  return Number.isFinite(normalized) ? normalized : undefined;
}

function isMeasureColumnName(name: string): boolean {
  const normalized = normalizeComparableText(name);
  if (!normalized || /\b(id|no|number|date|booking|container|truck|account|ref|filename|size|status|type|direction)\b/.test(normalized)) {
    return false;
  }
  return /\b(amount|price|cost|fee|total|gross|net|collect|revenue|expense|spend|cash|tax|balance|profit|loss|thb)\b/.test(normalized);
}

function compactTablePayloadFromResult(result: unknown): {
  headers?: unknown[];
  values?: CellMatrix;
  formulas?: CellMatrix;
  text?: CellMatrix;
  numberFormat?: string[][];
} {
  if (!result || typeof result !== "object") {
    return {};
  }
  const raw = result as {
    table?: unknown;
    headers?: unknown;
    values?: unknown;
    formulas?: unknown;
    text?: unknown;
    numberFormat?: unknown;
  };
  const table = raw.table && typeof raw.table === "object" ? raw.table as Record<string, unknown> : {};
  const headers = Array.isArray(raw.headers) ? raw.headers : Array.isArray(table.headers) ? table.headers : undefined;
  const values = matrixFromUnknown(raw.values) ?? matrixFromUnknown(table.values);
  const formulas = matrixFromUnknown(raw.formulas) ?? matrixFromUnknown(table.formulas);
  const text = matrixFromUnknown(raw.text) ?? matrixFromUnknown(table.text);
  const numberFormat = stringMatrixFromUnknown(raw.numberFormat) ?? stringMatrixFromUnknown(table.numberFormat);
  return {
    ...(headers ? { headers } : {}),
    ...(values ? { values } : {}),
    ...(formulas ? { formulas } : {}),
    ...(text ? { text } : {}),
    ...(numberFormat ? { numberFormat } : {})
  };
}

function matrixFromUnknown(value: unknown): CellMatrix | undefined {
  return Array.isArray(value) && value.every((row) => Array.isArray(row)) ? value as CellMatrix : undefined;
}

function stringMatrixFromUnknown(value: unknown): string[][] | undefined {
  return Array.isArray(value) && value.every((row) => Array.isArray(row))
    ? value.map((row) => row.map((cell) => String(cell ?? "")))
    : undefined;
}

function shouldReturnSparseRows(cellCount: number, nonEmptyCount: number, nonEmptyRowCount: number): boolean {
  if (nonEmptyCount === 0) return false;
  if (cellCount >= 50 && nonEmptyCount / cellCount <= 0.4) return true;
  return cellCount >= 500 && nonEmptyRowCount <= 75;
}

function emptyMatrixSummary(values: CellMatrix) {
  const rows = values.length;
  const columns = maxMatrixColumns(values);
  const cellCount = rows * columns;
  const nonEmptyCells = values.reduce((total, row) => total + row.filter((value) => value !== null && value !== undefined && value !== "").length, 0);
  return {
    sourceRows: rows,
    sourceColumns: columns,
    sourceCells: cellCount,
    nonEmptyCells,
    emptyCells: Math.max(0, cellCount - nonEmptyCells),
    trailingRows: trailingEmptyRows(values),
    trailingColumns: trailingEmptyColumns(values)
  };
}

function sparseValueRows(values: CellMatrix, address?: string) {
  const origin = address ? tryParseA1Address(address) : undefined;
  const startRow = origin?.startRow ?? 1;
  const startColumn = origin?.startColumn ?? 1;
  return values.flatMap((row, rowIndex) => {
    const cells = row.flatMap((value, columnIndex) => {
      if (value === null || value === undefined || value === "") {
        return [];
      }
      const columnNumber = startColumn + columnIndex;
      const rowNumber = startRow + rowIndex;
      return [{
        column: numberToColumn(columnNumber),
        address: `${numberToColumn(columnNumber)}${rowNumber}`,
        value
      }];
    });
    return cells.length > 0 ? [{ row: startRow + rowIndex, cells }] : [];
  });
}

function maxMatrixColumns(values: CellMatrix): number {
  return values.reduce((max, row) => Math.max(max, row.length), 0);
}

function trailingEmptyRows(values: CellMatrix): number {
  let count = 0;
  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (values[index]?.some((value) => value !== null && value !== undefined && value !== "")) {
      break;
    }
    count += 1;
  }
  return count;
}

function trailingEmptyColumns(values: CellMatrix): number {
  const columns = maxMatrixColumns(values);
  let count = 0;
  for (let columnIndex = columns - 1; columnIndex >= 0; columnIndex -= 1) {
    if (values.some((row) => row[columnIndex] !== null && row[columnIndex] !== undefined && row[columnIndex] !== "")) {
      break;
    }
    count += 1;
  }
  return count;
}

function trimTrailingEmptyCells(row: CellMatrix[number]): CellMatrix[number] {
  let end = row.length;
  while (end > 0 && (row[end - 1] === null || row[end - 1] === undefined || row[end - 1] === "")) {
    end -= 1;
  }
  return row.slice(0, end);
}

function compactSampleRow(row: CellMatrix[number], maxCells: number): CellMatrix[number] {
  const trimmed = trimTrailingEmptyCells(row);
  if (trimmed.length <= maxCells) {
    return trimmed;
  }
  if (trimmed.slice(0, maxCells).some((value) => value !== null && value !== undefined && value !== "")) {
    return trimmed.slice(0, maxCells);
  }
  const firstValueIndex = trimmed.findIndex((value) => value !== null && value !== undefined && value !== "");
  if (firstValueIndex === -1) {
    return [];
  }
  const start = Math.max(0, Math.min(firstValueIndex, trimmed.length - maxCells));
  return trimmed.slice(start, start + maxCells);
}

function alignComparisonRows(profiles: Array<{ sheetName: string; profile: ReturnType<typeof profileValues> }>) {
  if (profiles.length < 2 || profiles.some((profile) => !profile.profile.rows)) {
    return [];
  }
  const labels = new Set<string>();
  const bySheet = profiles.map((profile) => {
    const rows = new Map<string, CellMatrix[number]>();
    for (const row of profile.profile.rows ?? []) {
      const label = typeof row[0] === "string" ? row[0].trim() : "";
      if (!label || /^metric$/i.test(label)) continue;
      const key = normalizeComparableText(label);
      labels.add(key);
      rows.set(key, row);
    }
    return { sheetName: profile.sheetName, rows };
  });
  return [...labels].slice(0, 75).map((key) => {
    const values = bySheet.map((sheet) => ({
      sheetName: sheet.sheetName,
      label: String(sheet.rows.get(key)?.[0] ?? ""),
      value: sheet.rows.get(key)?.[1] ?? null
    }));
    const numericValues = values.map((value) => typeof value.value === "number" ? value.value : typeof value.value === "string" ? Number(value.value.replace(/[$,]/g, "")) : NaN);
    const delta = numericValues.length >= 2 && numericValues.slice(0, 2).every(Number.isFinite) ? numericValues[1]! - numericValues[0]! : undefined;
    return {
      key,
      values,
      ...(delta !== undefined ? { delta } : {})
    };
  });
}

function compareNumericMetrics(profiles: Array<{ sheetName: string; profile: ReturnType<typeof profileValues> }>) {
  const rows = profiles.map((item) => ({ sheetName: item.sheetName, sum: item.profile.metrics.sum ?? 0, numericCount: item.profile.metrics.numericCount }));
  const highest = [...rows].sort((left, right) => right.sum - left.sum)[0];
  return { rows, highestSumSheet: highest?.sheetName };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function styleSummaryFromFingerprint(fingerprint: unknown) {
  const record = isRecord(fingerprint) ? fingerprint : {};
  const dimensions = isRecord(record.dimensions) ? record.dimensions : {};
  return {
    sheetName: record.sheetName,
    range: record.address,
    rowCount: record.rowCount,
    columnCount: record.columnCount,
    truncated: record.truncated === true,
    fills: dimensions.fills,
    fonts: dimensions.fonts,
    alignment: dimensions.alignment,
    numberFormats: dimensions.numberFormats,
    borders: dimensions.borders,
    rowHeights: dimensions.rowHeights,
    columnWidths: dimensions.columnWidths,
    freezePanes: dimensions.freezePanes,
    conditionalFormatting: dimensions.conditionalFormatting,
    dataValidation: dimensions.dataValidation
  };
}

function isStyleOverviewRequest(request: string): boolean {
  if (/\bstyle\s+summar(y|ies)\b/i.test(request) || /\bread\b.{0,20}\bstyles?\b/i.test(request)) {
    return false;
  }
  return /\b(style|styling|formatting|visual|readability|design)\b/i.test(request)
    && /\b(overview|summary|review|inspect|look|suggest|recommend|improve|better|standardi[sz]e|best practice)\b/i.test(request);
}

function styleOverviewFromMetadata(metadata: WorkbookMetadata, sheetName: string, range: string, styleSummary: ReturnType<typeof styleSummaryFromFingerprint>) {
  const sheet = metadata.sheets.find((candidate) => sameText(candidate.name, sheetName));
  const tables = metadata.tables.filter((table) => sameText(table.sheetName, sheetName));
  const primaryTable = tables.find((table) => rangeOverlapsLoose(table.range, range)) ?? tables[0];
  const detected = detectVisualReadabilityStructure(metadata, sheet ?? {
    id: `sheet:${sheetName}`,
    index: 0,
    name: sheetName,
    kind: "unknown",
    usedRange: range,
    tableIds: [],
    sectionIds: [],
    summaryBlockIds: [],
    formulaRegionIds: [],
    headers: [],
    isHidden: false
  } as WorkbookMetadata["sheets"][number], range);
  const columns = sheet ? inferVisualReadabilityColumns(metadata, sheet, detected) : columnsFromTable(primaryTable);
  const columnGroups = suggestedColumnGroups(columns, detected);
  const groupedHeaderSuggestion = groupedHeaderSuggestionFromColumns(detected, columnGroups);
  const freezePanes = freezePaneOverview(styleSummary.freezePanes);
  return stripUndefinedRecord({
    kind: "style_overview",
    source: "cached_metadata_and_style_fingerprint",
    sheetName,
    range,
    freezePanes,
    table: primaryTable ? stripUndefinedRecord({
      name: primaryTable.name,
      range: primaryTable.range,
      headerRange: primaryTable.headerRange,
      dataRange: primaryTable.dataRange,
      columnCount: primaryTable.columns.length
    }) : undefined,
    detected: stripUndefinedRecord({
      headerRow: detected.headerRow,
      headerRange: detected.headerRange,
      dataRange: detected.dataRange,
      tableRanges: detected.tableRanges,
      hasFilter: detected.hasFilter,
      confidence: detected.confidence
    }),
    currentStyle: stripUndefinedRecord({
      header: styleOverviewHeaderStyle(styleSummary),
      fills: compactStyleDimension(styleSummary.fills),
      fonts: compactStyleDimension(styleSummary.fonts),
      borders: compactStyleDimension(styleSummary.borders),
      alignment: compactStyleDimension(styleSummary.alignment),
      rowHeights: compactStyleDimension(styleSummary.rowHeights),
      columnWidths: compactStyleDimension(styleSummary.columnWidths),
      freezePanes,
      numberFormats: compactStyleDimension(styleSummary.numberFormats),
      conditionalFormatting: compactStyleDimension(styleSummary.conditionalFormatting),
      dataValidation: compactStyleDimension(styleSummary.dataValidation)
    }),
    columnRoles: columns.slice(0, 32).map((column) => stripUndefinedRecord({
      column: column.column,
      header: column.header,
      role: column.role,
      inferredType: column.inferredType,
      confidence: column.confidence,
      freezePane: freezePaneColumnAnnotation(column.column, freezePanes)
    })),
    columnGroupSuggestions: columnGroups,
    groupedHeaderSuggestion,
    recommendations: styleOverviewRecommendations(groupedHeaderSuggestion, columns, styleSummary),
    recommendedWorkflow: groupedHeaderSuggestion
      ? { intentAction: "grouped_header", mode: "preview_update", requiresConfirmation: true }
      : { intentAction: "improve_visual_readability", mode: "preview_update", requiresConfirmation: true }
  });
}

function freezePaneOverview(rawFreezePanes: unknown): Record<string, unknown> | undefined {
  if (!isRecord(rawFreezePanes)) {
    return undefined;
  }
  const freezePanes = freezePanesFromDimension(rawFreezePanes);
  const readable = freezePanes.readable !== false;
  const frozen = freezePanes.frozen === true;
  return stripUndefinedRecord({
    readable,
    frozen,
    rows: numericRecordValue(freezePanes, "rows"),
    columns: numericRecordValue(freezePanes, "columns"),
    lastFrozenRow: numericRecordValue(freezePanes, "lastFrozenRow"),
    firstUnfrozenRow: numericRecordValue(freezePanes, "firstUnfrozenRow"),
    lastFrozenColumn: stringRecordValue(freezePanes, "lastFrozenColumn"),
    firstUnfrozenColumn: stringRecordValue(freezePanes, "firstUnfrozenColumn"),
    summary: readable
      ? frozen
        ? freezePaneSummary("this sheet", freezePanes).replace(/^Freeze panes on this sheet: /, "")
        : "No frozen panes are active."
      : stringRecordValue(freezePanes, "message")
  });
}

function freezePaneColumnAnnotation(column: string, freezePanes: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  if (!freezePanes || freezePanes.readable === false) {
    return undefined;
  }
  const columnIndex = columnToNumber(column);
  if (!Number.isFinite(columnIndex) || columnIndex <= 0) {
    return undefined;
  }
  const frozenColumnCount = numericRecordValue(freezePanes, "columns") ?? 0;
  const lastFrozenColumn = stringRecordValue(freezePanes, "lastFrozenColumn");
  const firstUnfrozenColumn = stringRecordValue(freezePanes, "firstUnfrozenColumn");
  return stripUndefinedRecord({
    isFrozen: frozenColumnCount > 0 && columnIndex <= frozenColumnCount,
    isLastFrozenColumn: lastFrozenColumn !== undefined && sameText(column, lastFrozenColumn),
    isFirstUnfrozenColumn: firstUnfrozenColumn !== undefined && sameText(column, firstUnfrozenColumn)
  });
}

function freezePanesFromDimension(freezePanes: Record<string, unknown>): Record<string, unknown> {
  if (freezePanes.readable !== true && freezePanes.frozen !== true && freezePanes.frozen !== false) {
    return {
      ...freezePanes,
      readable: false,
      message: "Freeze pane location was not captured by the loaded Excel taskpane."
    };
  }
  return freezePanes;
}

function numericRecordValue(record: Record<string, unknown>, key: string): number | undefined {
  const value = record[key];
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function stringRecordValue(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key];
  return typeof value === "string" && value.trim() ? value : undefined;
}

function columnsFromTable(table: TableMetadata | undefined): VisualReadabilityColumnRole[] {
  return table?.columns.map((column) => ({
    column: column.letter,
    header: column.name,
    role: visualColumnRole(column),
    inferredType: column.inferredType,
    confidence: Math.max(0.45, Math.min(0.98, column.importance ?? 0.7)),
    signals: visualColumnSignals(column)
  })) ?? [];
}

function rangeOverlapsLoose(left: string | undefined, right: string | undefined): boolean {
  if (!left || !right) return false;
  return rangesOverlapAddresses(stripSheetName(left), stripSheetName(right));
}

function styleOverviewHeaderStyle(styleSummary: ReturnType<typeof styleSummaryFromFingerprint>) {
  return stripUndefinedRecord({
    fills: firstStyleCells(styleSummary.fills, "fillColor"),
    fonts: firstStyleCells(styleSummary.fonts, "fontBold"),
    alignment: firstStyleCells(styleSummary.alignment, "horizontalAlignment"),
    borders: compactStyleDimension(styleSummary.borders)
  });
}

function firstStyleCells(dimension: unknown, key: string) {
  if (!isRecord(dimension)) return undefined;
  const cells = isRecord(dimension.cells) || Array.isArray(dimension.cells) ? dimension.cells : undefined;
  if (!Array.isArray(cells)) return undefined;
  return cells.filter((cell) => isRecord(cell) && cell.rowIndex === 0 && cell[key] !== undefined).slice(0, 8);
}

function suggestedColumnGroups(columns: VisualReadabilityColumnRole[], detected: VisualReadabilityDetectedStructure): Array<{ label: string; startColumn: string; endColumn: string; columns: string[]; role: string; fillColor: string; headerFillColor: string }> {
  const palette = [
    { fillColor: "#1A3C6E", headerFillColor: "#D9EAF7" },
    { fillColor: "#0F6B78", headerFillColor: "#D7EEF2" },
    { fillColor: "#548235", headerFillColor: "#E2EFDA" },
    { fillColor: "#C65911", headerFillColor: "#FCE4D6" },
    { fillColor: "#8064A2", headerFillColor: "#EDE7F6" },
    { fillColor: "#666666", headerFillColor: "#E7E6E6" }
  ];
  return visualColumnGroups(columns, detected).map((group, index) => {
    const role = normalizeComparableText(group.label).replace(/\s+/g, "_") || "group";
    const color = palette[index % palette.length]!;
    return {
      label: group.label,
      startColumn: group.startColumn,
      endColumn: group.endColumn,
      columns: columnsInSpan(columns, group.startColumn, group.endColumn),
      role,
      fillColor: color.fillColor,
      headerFillColor: color.headerFillColor
    };
  });
}

function groupedHeaderGroupsFromInput(input: AgentRunInput, columns: VisualReadabilityColumnRole[], detected: VisualReadabilityDetectedStructure): Array<{ label: string; startColumn: string; endColumn: string; fillColor: string; headerFillColor: string }> {
  const values = input.values as Record<string, unknown> | undefined;
  const groupedHeader = isRecord(values?.groupedHeader) ? values.groupedHeader as Record<string, unknown> : values;
  const rawGroups = Array.isArray(groupedHeader?.groups) ? groupedHeader.groups : undefined;
  const inferred = suggestedColumnGroups(columns, detected);
  if (!rawGroups) {
    return inferred.map(({ label, startColumn, endColumn, fillColor, headerFillColor }) => ({ label, startColumn, endColumn, fillColor, headerFillColor }));
  }
  return rawGroups.flatMap((raw, index) => {
    if (!isRecord(raw)) return [];
    const columnSpan = groupedHeaderColumnSpan(raw);
    const startColumn = stringValue(raw.startColumn ?? raw.start ?? raw.from ?? raw.column ?? columnSpan?.startColumn);
    const endColumn = stringValue(raw.endColumn ?? raw.end ?? raw.to ?? raw.column ?? columnSpan?.endColumn ?? startColumn);
    const label = stringValue(raw.label ?? raw.name ?? raw.title) ?? inferred[index]?.label;
    if (!startColumn || !endColumn || !label) return [];
    return [{
      label,
      startColumn: startColumn.toUpperCase(),
      endColumn: endColumn.toUpperCase(),
      fillColor: colorString(raw.fillColor ?? raw.color) ?? inferred[index]?.fillColor ?? "#1A3C6E",
      headerFillColor: colorString(raw.headerFillColor ?? raw.bodyFillColor ?? raw.lightFillColor) ?? inferred[index]?.headerFillColor ?? "#D9EAF7"
    }];
  });
}

function groupedHeaderColumnSpan(raw: Record<string, unknown>): { startColumn: string; endColumn: string } | undefined {
  const columns = Array.isArray(raw.columns)
    ? raw.columns.map((column) => stringValue(column)).filter((column): column is string => Boolean(column && /^[A-Z]+$/i.test(column.trim())))
    : [];
  if (columns.length > 0) {
    const indexes = columns.map((column) => columnToNumber(column.trim().toUpperCase())).filter((index) => Number.isFinite(index));
    if (indexes.length > 0) {
      return { startColumn: columnLetter(Math.min(...indexes) - 1), endColumn: columnLetter(Math.max(...indexes) - 1) };
    }
  }
  const range = stringValue(raw.range ?? raw.address);
  if (!range) {
    return undefined;
  }
  const normalized = stripSheetName(range).replace(/\$/g, "").trim();
  const parsed = tryParseA1Address(normalized);
  if (parsed) {
    return { startColumn: columnLetter(parsed.startColumn - 1), endColumn: columnLetter(parsed.endColumn - 1) };
  }
  const columnRange = /^([A-Z]+)\s*(?::|-|\bto\b)\s*([A-Z]+)$/i.exec(normalized);
  const rangeStart = columnRange?.[1];
  const rangeEnd = columnRange?.[2];
  if (rangeStart && rangeEnd) {
    return { startColumn: rangeStart.toUpperCase(), endColumn: rangeEnd.toUpperCase() };
  }
  return undefined;
}

function groupedHeaderRowHeight(input: AgentRunInput, defaultHeight: number, key = "groupRowHeight"): number {
  const values = input.values as Record<string, unknown> | undefined;
  const groupedHeader = isRecord(values?.groupedHeader) ? values.groupedHeader as Record<string, unknown> : values;
  return numberValue(groupedHeader?.[key]) ?? defaultHeight;
}

function columnsInSpan(columns: VisualReadabilityColumnRole[], startColumn: string, endColumn: string): string[] {
  const start = columnToNumber(startColumn);
  const end = columnToNumber(endColumn);
  return columns
    .filter((column) => {
      const current = columnToNumber(column.column);
      return current >= start && current <= end;
    })
    .map((column) => column.column);
}

function groupedHeaderSuggestionFromColumns(detected: VisualReadabilityDetectedStructure, groups: ReturnType<typeof suggestedColumnGroups>) {
  const multiColumnGroups = groups.filter((group) => columnToNumber(group.endColumn) > columnToNumber(group.startColumn));
  if (!detected.headerRange || groups.length < 2 || groups.reduce((total, group) => total + group.columns.length, 0) < 6 || multiColumnGroups.length === 0) {
    return undefined;
  }
  return {
    kind: "grouped_header_suggestion",
    targetHeaderRange: detected.headerRange,
    levels: 1,
    insertAboveHeader: true,
    styleExistingHeader: true,
    requiresStructuralPreview: true,
    defaultApplyBehavior: "suggest_only",
    groups: groups.map((group) => ({
      label: group.label,
      startColumn: group.startColumn,
      endColumn: group.endColumn,
      fillColor: group.fillColor,
      headerFillColor: group.headerFillColor,
      merge: columnToNumber(group.endColumn) > columnToNumber(group.startColumn)
    })),
    operationsNeeded: ["insert_rows", "write_values_many", "merge_range", "write_styles_many"]
  };
}

function styleOverviewRecommendations(groupedHeaderSuggestion: unknown, columns: VisualReadabilityColumnRole[], styleSummary: ReturnType<typeof styleSummaryFromFingerprint>) {
  const recommendations: Array<Record<string, unknown>> = [];
  if (groupedHeaderSuggestion) {
    recommendations.push({
      id: "grouped_header",
      category: "structural_style",
      title: "Add a grouped visual header above the table header.",
      applySafety: "preview_required"
    });
  }
  if (columns.length >= 8) {
    recommendations.push({
      id: "freeze_header",
      category: "layout_format",
      title: "Freeze the header area for wide-sheet scanning.",
      applySafety: "opt_in"
    });
  }
  if (styleSummary.truncated) {
    recommendations.push({
      id: "style_sample_limited",
      category: "inspection",
      title: "Style sampling was compact; inspect a smaller header/body range before high-fidelity template repair.",
      applySafety: "read_only"
    });
  }
  return recommendations;
}

function styleWarnings(result: unknown): string[] {
  const fingerprint = isRecord(result) && isRecord(result.fingerprint) ? result.fingerprint : result;
  const warnings = isRecord(fingerprint) && Array.isArray(fingerprint.warnings) ? fingerprint.warnings : [];
  return warnings.map((warning) => isRecord(warning) && typeof warning.message === "string" ? warning.message : String(warning)).slice(0, 5);
}

function formulaReadAnswerFromSnapshot(snapshot: RangeSnapshot | undefined, patterns: unknown, sheetName: string, range: string) {
  const values = snapshot?.values ?? [];
  const text = snapshot?.text ?? [];
  const rawFormulas = snapshot?.formulas ?? [];
  const snapshotFormulas = normalizeFormulaOnlyMatrix(rawFormulas);
  const patternFormulas = formulaMatrixFromPatternCells(patterns, "formula");
  const patternFormulasR1C1 = formulaMatrixFromPatternCells(patterns, "formulaR1C1");
  const formulas = mergeFormulaMatrices(snapshotFormulas, matrixFromUnknown(isRecord(patterns) ? patterns.formulas : undefined) ?? [], patternFormulas);
  const formulasR1C1 = mergeFormulaMatrices(matrixFromUnknown(isRecord(patterns) ? patterns.formulasR1C1 : undefined) ?? [], patternFormulasR1C1);
  const patternMatrix = mergePatternMatrices(matrixFromUnknown(isRecord(patterns) ? patterns.patternMatrix : undefined) ?? [], patterns);
  const parsed = tryParseA1Address(stripSheetName(range));
  const rowCount = Math.max(values.length, text.length, formulas.length, formulasR1C1.length, patternMatrix.length);
  const columnCount = Math.max(maxMatrixColumns(values), maxMatrixColumns(text), maxMatrixColumns(formulas), maxMatrixColumns(formulasR1C1), maxMatrixColumns(patternMatrix));
  const formulaColumns = new Set<number>();
  for (const row of [formulas, formulasR1C1].flat()) {
    row.forEach((formula, index) => {
      if (formulaLike(formula)) formulaColumns.add(index);
    });
  }
  const cells: Array<Record<string, unknown>> = [];
  let formulaCount = 0;
  let hardcodedCount = 0;
  let blankCount = 0;
  let formulaResultBlankCount = 0;
  let hardcodedInFormulaColumnCount = 0;
  const missingFormulaGaps: Array<Record<string, unknown>> = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const value = values[rowIndex]?.[columnIndex];
      const displayedText = text[rowIndex]?.[columnIndex];
      const formula = formulas[rowIndex]?.[columnIndex];
      const formulaR1C1 = formulasR1C1[rowIndex]?.[columnIndex];
      const patternHash = patternMatrix[rowIndex]?.[columnIndex];
      const hasFormulaProof = formulaLike(formula) || formulaLike(formulaR1C1) || typeof patternHash === "string";
      const status = hasFormulaProof
        ? "formula"
        : isBlankishAutoApplyValue(value) && isBlankishAutoApplyValue(displayedText)
          ? "blank"
          : "hardcoded";
      if (status === "formula") formulaCount += 1;
      if (status === "hardcoded") hardcodedCount += 1;
      if (status === "blank") blankCount += 1;
      if (status === "formula" && isBlankishAutoApplyValue(value) && isBlankishAutoApplyValue(displayedText)) formulaResultBlankCount += 1;
      const cell = parsed ? `${numberToColumn(parsed.startColumn + columnIndex)}${parsed.startRow + rowIndex}` : undefined;
      if (status !== "formula" && formulaColumns.has(columnIndex)) {
        hardcodedInFormulaColumnCount += 1;
        if (missingFormulaGaps.length < 10) {
          missingFormulaGaps.push(stripUndefinedRecord({ cell, formulaStatus: status, value, displayedText }));
        }
      }
      if (cells.length < 25) {
        cells.push(stripUndefinedRecord({
          ...(cell ? { cell } : {}),
          value,
          displayedText,
          formula: formulaLike(formula) ? formula : undefined,
          formulaR1C1: formulaLike(formulaR1C1) ? formulaR1C1 : undefined,
          patternHash: typeof patternHash === "string" ? patternHash : undefined,
          formulaStatus: status
        }));
      }
    }
  }
  return stripUndefinedRecord({
    kind: "formula_read",
    source: "live_read",
    sheetName,
    range,
    shape: { rows: rowCount, columns: columnCount },
    formulaCount,
    hardcodedCount,
    blankCount,
    formulaResultBlankCount,
    hardcodedInFormulaColumnCount,
    missingFormulaGaps,
    cells,
    values,
    text,
    formulas,
    formulasR1C1,
    patternMatrix
  }) as {
    kind: "formula_read";
    formulaCount: number;
    hardcodedCount: number;
    blankCount: number;
    hardcodedInFormulaColumnCount: number;
    [key: string]: unknown;
  };
}

function formulaReadSummary(answer: { formulaCount: number; hardcodedCount: number; blankCount: number }, sheetName: string, range: string): string {
  if (answer.formulaCount > 0) {
    return `Read formula proof for ${sheetName}!${range}: ${answer.formulaCount} formula cell(s), ${answer.hardcodedCount} hardcoded cell(s), ${answer.blankCount} blank cell(s).`;
  }
  return `Read formula proof for ${sheetName}!${range}: no formula cells found; ${answer.hardcodedCount} hardcoded cell(s), ${answer.blankCount} blank cell(s).`;
}

function formulaMatrixFromPatternCells(patterns: unknown, field: "formula" | "formulaR1C1"): CellMatrix {
  if (!isRecord(patterns) || !Array.isArray(patterns.cells)) {
    return [];
  }
  const rows = typeof patterns.rowCount === "number" ? patterns.rowCount : 0;
  const columns = typeof patterns.columnCount === "number" ? patterns.columnCount : 0;
  const matrix: CellMatrix = Array.from({ length: rows }, () => Array.from({ length: columns }, () => null));
  for (const rawCell of patterns.cells) {
    if (!isRecord(rawCell)) continue;
    const rowIndex = typeof rawCell.rowIndex === "number" ? rawCell.rowIndex : undefined;
    const columnIndex = typeof rawCell.columnIndex === "number" ? rawCell.columnIndex : undefined;
    const formula = typeof rawCell[field] === "string" ? rawCell[field] : undefined;
    if (rowIndex === undefined || columnIndex === undefined || !formula) continue;
    matrix[rowIndex] ??= [];
    matrix[rowIndex]![columnIndex] = formula;
  }
  return matrix;
}

function mergeFormulaMatrices(...matrices: CellMatrix[]): CellMatrix {
  const rowCount = Math.max(...matrices.map((matrix) => matrix.length), 0);
  const columnCount = Math.max(...matrices.map(maxMatrixColumns), 0);
  return Array.from({ length: rowCount }, (_row, rowIndex) =>
    Array.from({ length: columnCount }, (_column, columnIndex) => {
      for (const matrix of matrices) {
        const value = matrix[rowIndex]?.[columnIndex];
        if (formulaLike(value)) return String(value);
      }
      return null;
    })
  );
}

function mergePatternMatrices(matrix: CellMatrix, patterns: unknown): CellMatrix {
  const cellMatrix = matrix.map((row) => [...row]);
  if (!isRecord(patterns) || !Array.isArray(patterns.cells)) {
    return cellMatrix;
  }
  for (const rawCell of patterns.cells) {
    if (!isRecord(rawCell)) continue;
    const rowIndex = typeof rawCell.rowIndex === "number" ? rawCell.rowIndex : undefined;
    const columnIndex = typeof rawCell.columnIndex === "number" ? rawCell.columnIndex : undefined;
    const hash = typeof rawCell.patternHash === "string" ? rawCell.patternHash : undefined;
    if (rowIndex === undefined || columnIndex === undefined || !hash) continue;
    cellMatrix[rowIndex] ??= [];
    cellMatrix[rowIndex]![columnIndex] = hash;
  }
  return cellMatrix;
}

function normalizeFormulaOnlyMatrix(matrix: CellMatrix): CellMatrix {
  return matrix.map((row) => row.map((value) => formulaLike(value) ? value : null));
}

function formatDiagnosticsFromSnapshot(snapshot: RangeSnapshot | undefined, sheetName: string, range: string) {
  const values = snapshot?.values ?? [];
  const text = snapshot?.text ?? [];
  const formulas = snapshot?.formulas ?? [];
  const numberFormat = snapshot?.numberFormat ?? [];
  const cells: Array<Record<string, unknown>> = [];
  const issues: Array<{ code: string; severity: "info" | "warning" | "error"; message: string; cell?: string; suggestedAction?: string; suggestedValues?: Record<string, unknown> }> = [];
  const parsed = tryParseA1Address(stripSheetName(range));
  const rowCount = Math.max(values.length, text.length, formulas.length, numberFormat.length);
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const columnCount = Math.max(values[rowIndex]?.length ?? 0, text[rowIndex]?.length ?? 0, formulas[rowIndex]?.length ?? 0, numberFormat[rowIndex]?.length ?? 0);
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const cell = parsed ? `${numberToColumn(parsed.startColumn + columnIndex)}${parsed.startRow + rowIndex}` : undefined;
      const rawValue = values[rowIndex]?.[columnIndex];
      const displayedText = text[rowIndex]?.[columnIndex];
      const formula = formulas[rowIndex]?.[columnIndex];
      const format = numberFormat[rowIndex]?.[columnIndex];
      const detectedType = detectCellType(rawValue, displayedText, formula);
      cells.push({
        ...(cell ? { cell } : {}),
        rawValue,
        displayedText,
        formula,
        numberFormat: format,
        detectedType
      });
      if (typeof rawValue === "string" && isDateLikeText(rawValue)) {
        issues.push({
          code: /\b\d{1,2}[/-]\d{1,2}[/-]\d{2}\b/.test(rawValue) ? "DATE_TEXT_TWO_DIGIT_YEAR" : "DATE_TEXT",
          severity: "warning",
          message: `${cell ?? "Cell"} contains date-like text${/\b\d{1,2}[/-]\d{1,2}[/-]\d{2}\b/.test(rawValue) ? " with a two-digit year" : ""}.`,
          ...(cell ? { cell } : {}),
          suggestedAction: "parse_dates",
          suggestedValues: { numberFormat: "dd/mm/yyyy" }
        });
      }
      if (typeof rawValue === "number" && (format === undefined || format === "General") && isDateLikeText(displayedText)) {
        issues.push({
          code: "DATE_SERIAL_GENERAL_FORMAT",
          severity: "warning",
          message: `${cell ?? "Cell"} is numeric but appears to display as a date without an explicit date number format.`,
          ...(cell ? { cell } : {}),
          suggestedAction: "write_number_formats",
          suggestedValues: { numberFormat: "dd/mm/yyyy" }
        });
      }
      if (typeof rawValue === "string" && /^=/.test(rawValue)) {
        issues.push({
          code: "FORMULA_STORED_AS_TEXT",
          severity: "warning",
          message: `${cell ?? "Cell"} starts with '=' but is stored as text.`,
          ...(cell ? { cell } : {}),
          suggestedAction: "write_formulas"
        });
      }
    }
  }
  const uniqueFormats = [...new Set(numberFormat.flat().filter((value): value is string => typeof value === "string"))];
  if (uniqueFormats.length > 1) {
    issues.push({
      code: "INCONSISTENT_NUMBER_FORMATS",
      severity: "info",
      message: `Range ${sheetName}!${range} contains ${uniqueFormats.length} number formats.`,
      suggestedAction: "write_number_formats",
      suggestedValues: { numberFormats: uniqueFormats }
    });
  }
  return {
    kind: "format_diagnostics",
    sheetName,
    range,
    cells,
    style: snapshot?.style,
    issues,
    suggestedNextAction: issues[0]?.suggestedAction ?? "answer_now"
  };
}

function detectCellType(rawValue: unknown, displayedText: unknown, formula: unknown): string {
  if (typeof formula === "string" && formula.startsWith("=")) return "formula";
  if (typeof rawValue === "number") return "number";
  if (typeof rawValue === "boolean") return "boolean";
  if (typeof rawValue === "string" && isDateLikeText(rawValue)) return "date_text";
  if (typeof displayedText === "string" && isDateLikeText(displayedText)) return "date_display";
  if (typeof rawValue === "string") return "text";
  if (rawValue === null || rawValue === undefined || rawValue === "") return "blank";
  return typeof rawValue;
}

function isDateLikeText(value: unknown): value is string {
  return typeof value === "string" && /^\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s*$/.test(value);
}

function styleFromInput(input: AgentRunInput): NonNullable<RangeSnapshot["style"]> {
  const values = input.values as Record<string, unknown> | undefined;
  const structured = normalizeStyleRecord(values?.style);
  const flattened = normalizeStyleRecord(values);
  const requested = styleFromRequest(input.request);
  const style = { ...requested, ...flattened, ...structured };
  if (isGroupedHeaderStyleRequest(input) && !hasExplicitFillColor(input, flattened, structured) && style.fillColor === "#D9EAF7") {
    return {
      ...style,
      fillColor: "#1A3C6E",
      fontColor: "#FFFFFF",
      fontBold: true,
      horizontalAlignment: "center"
    };
  }
  return style;
}

function defaultFormatRangeStyle(input: AgentRunInput): NonNullable<RangeSnapshot["style"]> {
  const base: NonNullable<RangeSnapshot["style"]> = {
    fillColor: "#D9EAF7",
    fontColor: "#1F2937",
    fontBold: true,
    horizontalAlignment: "center"
  };
  if (isGroupedHeaderStyleRequest(input)) {
    return {
      ...base,
      fillColor: "#1A3C6E",
      fontColor: "#FFFFFF"
    };
  }
  return base;
}

function hasExactFormatRangeTarget(input: AgentRunInput): boolean {
  return typeof input.target?.range === "string"
    || typeof input.target?.address === "string"
    || (typeof input.target?.row === "number" && Number.isInteger(input.target.row) && input.target.row > 0);
}

function normalizeStyleRecord(value: unknown): NonNullable<RangeSnapshot["style"]> {
  if (!value || typeof value !== "object") {
    return {};
  }
  const record = value as Record<string, unknown>;
  const fillColor = colorString(record.fillColor ?? record.fill ?? record.backgroundColor ?? record.background);
  const fontColor = colorString(record.fontColor ?? record.textColor);
  const style: NonNullable<RangeSnapshot["style"]> = {};
  if (fillColor !== undefined) style.fillColor = fillColor;
  if (fontColor !== undefined) style.fontColor = fontColor;
  const fontBold = booleanValue(record.fontBold ?? record.bold);
  if (fontBold !== undefined) style.fontBold = fontBold;
  const fontItalic = booleanValue(record.fontItalic ?? record.italic);
  if (fontItalic !== undefined) style.fontItalic = fontItalic;
  const fontName = stringValue(record.fontName);
  if (fontName !== undefined) style.fontName = fontName;
  const fontSize = numberValue(record.fontSize);
  if (fontSize !== undefined) style.fontSize = fontSize;
  const horizontalAlignment = stringValue(record.horizontalAlignment ?? record.align);
  if (horizontalAlignment !== undefined) style.horizontalAlignment = horizontalAlignment;
  const verticalAlignment = stringValue(record.verticalAlignment);
  if (verticalAlignment !== undefined) style.verticalAlignment = verticalAlignment;
  const rowHeight = numberValue(record.rowHeight);
  if (rowHeight !== undefined) style.rowHeight = rowHeight;
  const columnWidth = numberValue(record.columnWidth);
  if (columnWidth !== undefined) style.columnWidth = agentColumnWidthToOfficeWidth(columnWidth);
  const borders = record.borders && typeof record.borders === "object"
    ? record.borders as NonNullable<RangeSnapshot["style"]>["borders"]
    : undefined;
  if (borders !== undefined) {
    style.borders = borders;
  }
  return style;
}

function agentColumnWidthToOfficeWidth(width: number): number {
  if (!Number.isFinite(width) || width <= 0) {
    return width;
  }
  if (width > 60) {
    return width;
  }

  const maxDigitPixelWidth = 7;
  const padding = Math.trunc(128 / maxDigitPixelWidth);
  const pixels = Math.trunc(((256 * width + padding) / 256) * maxDigitPixelWidth);
  return Math.round(pixels * 0.75 * 100) / 100;
}

function styleFromRequest(request: string): NonNullable<RangeSnapshot["style"]> {
  const lower = request.toLowerCase();
  const fillColor =
    /\b(fill|background|highlight|turn|make|set|color|colour)\b/.test(lower)
      ? colorNearContext(lower, ["fill", "background", "highlight"]) ?? colorFromText(lower, ["fill", "background", "highlight", "turn", "make", "set", "color", "colour"])
      : undefined;
  const fontColor =
    /\b(font|text)\b/.test(lower)
      ? colorNearContext(lower, ["font", "text"]) ?? colorFromText(lower, ["font", "text"])
      : undefined;
  const horizontalAlignment = /\b(?:center|centered|centre|centred)\s+(?:align|aligned|alignment)|\b(?:align|aligned)\s+(?:center|centre)\b/.test(lower)
    ? "center"
    : undefined;
  return {
    ...(lower.includes("header") ? {
      fontBold: true,
      ...(fillColor ? {} : { fillColor: "#D9EAF7" }),
      ...(fontColor || fillColor ? {} : { fontColor: "#1F2937" }),
      horizontalAlignment: "center"
    } : {}),
    ...(fillColor ? { fillColor } : {}),
    ...(fontColor ? { fontColor } : fillColor === "#000000" ? { fontColor: "#FFFFFF" } : {}),
    ...(/\bbold\b/.test(lower) ? { fontBold: true } : {}),
    ...(/\bitalic\b/.test(lower) ? { fontItalic: true } : {}),
    ...(horizontalAlignment ? { horizontalAlignment } : {}),
    ...(/\bborders?\b/.test(lower) && /\b(add|apply|draw|thin|outline|all sides?)\b/.test(lower)
      ? { borders: { style: "continuous" as const, weight: "thin" as const } }
      : {})
  };
}

function isGroupedHeaderStyleRequest(input: AgentRunInput): boolean {
  const request = input.request.toLowerCase();
  const targetRange = stripSheetName(input.target?.range ?? input.target?.address ?? "");
  return /\b(grouped|grouping|category|top)\s+(?:header|headers|row)|\bheader\s+(?:group|grouping|category)|\brow\s*1\b/i.test(request)
    || /^A1(?::[A-Z]+1)?$/i.test(targetRange);
}

function hasExplicitFillColor(
  input: AgentRunInput,
  flattened: NonNullable<RangeSnapshot["style"]>,
  structured: NonNullable<RangeSnapshot["style"]>
): boolean {
  return flattened.fillColor !== undefined
    || structured.fillColor !== undefined
    || (/\b(fill|background|highlight|turn|make|set|color|colour|dark|darker|light|lighter|blue|black|white|red|green|yellow|orange|purple|gray|grey|#[0-9a-f]{6})\b/i.test(input.request)
      && colorFromText(input.request.toLowerCase(), ["fill", "background", "highlight", "turn", "make", "set", "color", "colour"]) !== undefined);
}

function groupedHeaderStyleWarnings(input: AgentRunInput, style: NonNullable<RangeSnapshot["style"]>): string[] {
  if (!isGroupedHeaderStyleRequest(input)) {
    return [];
  }
  if (style.fillColor === "#1A3C6E" && /\b(match|same|row\s*2|actual\s+header|column\s+header)\b/i.test(input.request)) {
    return ["Grouped header row styling was kept darker than the actual column header row so row 1 remains visually distinct from row 2."];
  }
  if (style.fillColor === "#D9EAF7") {
    return ["Grouped header row is using the same light fill as the actual column header row; use a darker fill such as #1A3C6E to keep the hierarchy distinct."];
  }
  return [];
}

function colorString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (/^#[0-9a-fA-F]{6}$/.test(trimmed)) return trimmed.toUpperCase();
  return colorFromText(trimmed.toLowerCase());
}

function colorFromText(text: string, contextWords?: string[]): string | undefined {
  const hex = text.match(/#[0-9a-fA-F]{6}/)?.[0];
  if (hex) return hex.toUpperCase();
  const colorMap: Record<string, string> = {
    black: "#000000",
    white: "#FFFFFF",
    yellow: "#FFFF00",
    red: "#FF0000",
    green: "#00B050",
    blue: "#0070C0",
    gray: "#808080",
    grey: "#808080",
    orange: "#FFC000"
  };
  for (const [name, color] of Object.entries(colorMap)) {
    if (!new RegExp(`\\b${name}\\b`).test(text)) continue;
    if (!contextWords || contextWords.some((word) => new RegExp(`\\b${word}\\b`).test(text))) {
      return color;
    }
  }
  return undefined;
}

function colorNearContext(text: string, contextWords: string[]): string | undefined {
  const colorNames = ["black", "white", "yellow", "red", "green", "blue", "gray", "grey", "orange"];
  const candidates = [
    ...[...text.matchAll(/#[0-9a-fA-F]{6}/g)].map((match) => ({ value: match[0].toUpperCase(), index: match.index ?? 0, explicit: true })),
    ...colorNames.flatMap((color) => [...text.matchAll(new RegExp(`\\b${color}\\b`, "g"))].map((match) => ({ value: color, index: match.index ?? 0, explicit: false })))
  ];
  let best: { value: string; distance: number; explicit: boolean } | undefined;
  for (const context of contextWords) {
    for (const contextMatch of text.matchAll(new RegExp(`\\b${context}\\b`, "g"))) {
      for (const candidate of candidates) {
        const distance = Math.abs(candidate.index - (contextMatch.index ?? 0));
        if (
          distance <= 40 &&
          (!best || (candidate.explicit && !best.explicit) || (candidate.explicit === best.explicit && distance < best.distance))
        ) {
          best = { value: candidate.value, distance, explicit: candidate.explicit };
        }
      }
    }
  }
  return best ? colorString(best.value) : undefined;
}

function booleanValue(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

function dataValidationEntriesFromInput(
  workbookId: WorkbookId,
  input: AgentRunInput,
  defaultValidation: Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"]
): Array<{
  target: A1Range;
  validation: Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"];
}> {
  const values = input.values as Record<string, unknown> | undefined;
  const rawInput = input as unknown as Record<string, unknown>;
  const rawEntries = Array.isArray(values?.entries)
    ? values.entries
    : Array.isArray(values?.patches)
      ? values.patches
    : Array.isArray(rawInput.entries)
      ? rawInput.entries as unknown[]
      : [];
  const entries: Array<{
    target: A1Range;
    validation: Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"];
  }> = [];
  for (const rawEntry of rawEntries) {
    if (!rawEntry || typeof rawEntry !== "object") {
      continue;
    }
    const entry = rawEntry as Record<string, unknown>;
    const target = entry.target && typeof entry.target === "object" ? entry.target as Record<string, unknown> : {};
    const sheetName = stringValue(entry.sheetName ?? target.sheetName ?? input.target?.sheetName ?? values?.sheetName);
    const address = stringValue(entry.address ?? entry.range ?? target.address ?? target.range);
    if (!sheetName || !address) {
      continue;
    }
    entries.push({
      target: { workbookId, sheetName, address: unqualifiedAddress(address) },
      validation: dataValidationFromRecord(entry, defaultValidation)
    });
  }
  return entries;
}

function dataValidationFromInput(input: AgentRunInput): Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"] | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const validation = values?.validation && typeof values.validation === "object" ? values.validation as Record<string, unknown> : undefined;
  const source = validation?.source ?? validation?.formula1 ?? values?.source ?? values?.options ?? values?.allowedValues;
  const optionsOrFormula = dataValidationSourceFromValue(source);
  const options = optionsOrFormula !== undefined
    ? optionsOrFormula
    : optionsFromRequest(input.request);
  if (Array.isArray(options) && options.length === 0) {
    return undefined;
  }
  if (typeof options === "string" && options.trim().length === 0) {
    return undefined;
  }
  return {
    type: "list",
    source: options,
    inCellDropDown: booleanValue(validation?.inCellDropDown ?? values?.inCellDropDown) ?? true,
    ignoreBlanks: booleanValue(validation?.ignoreBlanks ?? values?.ignoreBlanks) ?? true
  };
}

function dataValidationSourceFromValue(source: unknown): string | string[] | undefined {
  if (Array.isArray(source)) {
    const values = source.filter((item): item is string => typeof item === "string" && item.trim().length > 0).map((item) => item.trim());
    if (values.length === 1) {
      const onlyValue = values[0];
      const formula = onlyValue ? normalizeDataValidationRangeReference(onlyValue) : undefined;
      return formula ?? values;
    }
    return values;
  }
  if (typeof source !== "string") {
    return undefined;
  }
  const trimmed = source.trim();
  if (!trimmed) {
    return undefined;
  }
  if (trimmed.startsWith("=")) {
    return normalizeDataValidationFormulaSource(trimmed);
  }
  const formula = normalizeDataValidationRangeReference(trimmed);
  if (formula) {
    return formula;
  }
  return trimmed.split(",").map((item) => item.trim()).filter(Boolean);
}

function normalizeDataValidationFormulaSource(source: string): string {
  const formula = normalizeDataValidationRangeReference(source);
  if (formula) {
    return formula;
  }
  return source;
}

function normalizeDataValidationRangeReference(source: string): string | undefined {
  const trimmed = source.trim();
  const body = trimmed.startsWith("=") ? trimmed.slice(1).trim() : trimmed;
  const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)$/iu.exec(body);
  if (!match) {
    return undefined;
  }
  const rawSheet = match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2] ?? "";
  const address = match[3] ?? "";
  const sheet = rawSheet.trim();
  if (!sheet) {
    return undefined;
  }
  if (!/[\s()[\]{}.,]/u.test(sheet)) {
    return `=${sheet}!${address}`;
  }
  return `='${sheet.replace(/'/g, "''")}'!${address}`;
}

function dataValidationSourceFromValueWithFallback(source: unknown, fallback: string | string[]): string | string[] {
  return dataValidationSourceFromValue(source) ?? fallback;
}

function dataValidationFromRecord(
  record: Record<string, unknown>,
  fallback: Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"]
): Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"] {
  const validation = record.validation && typeof record.validation === "object" ? record.validation as Record<string, unknown> : undefined;
  const source = validation?.source ?? validation?.formula1 ?? record.source ?? record.options ?? record.allowedValues ?? fallback.source;
  const parsedSource = dataValidationSourceFromValueWithFallback(source, fallback.source);
  const next: Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"] = {
    type: "list",
    source: parsedSource
  };
  const inCellDropDown = booleanValue(validation?.inCellDropDown ?? record.inCellDropDown) ?? fallback.inCellDropDown;
  if (inCellDropDown !== undefined) next.inCellDropDown = inCellDropDown;
  const ignoreBlanks = booleanValue(validation?.ignoreBlanks ?? record.ignoreBlanks) ?? fallback.ignoreBlanks;
  if (ignoreBlanks !== undefined) next.ignoreBlanks = ignoreBlanks;
  const prompt = validation?.prompt && typeof validation.prompt === "object"
    ? validation.prompt as Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"]["prompt"]
    : fallback.prompt;
  if (prompt !== undefined) next.prompt = prompt;
  const errorAlert = validation?.errorAlert && typeof validation.errorAlert === "object"
    ? validation.errorAlert as Extract<ExcelOperation, { kind: "range.write_data_validation" }>["validation"]["errorAlert"]
    : fallback.errorAlert;
  if (errorAlert !== undefined) next.errorAlert = errorAlert;
  return next;
}

function unqualifiedAddress(address: string): string {
  const bangIndex = address.lastIndexOf("!");
  return bangIndex >= 0 ? address.slice(bangIndex + 1) : address;
}

function optionsFromRequest(request: string): string[] {
  const including = request.match(/\b(?:including|include|values?|options?|allowed values?)[:\s]+([^.;]+)/i)?.[1];
  if (!including) {
    return [];
  }
  return including.split(/,|\bor\b/i).map((item) => item.trim()).filter((item) => /^[A-Za-z0-9' -]+$/.test(item));
}

function conditionalFormattingRuleFromInput(input: AgentRunInput): Extract<ExcelOperation, { kind: "range.write_conditional_formatting" }>["rule"] | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const rawRule = values?.rule && typeof values.rule === "object" ? values.rule as Record<string, unknown> : undefined;
  const formula = stringValue(rawRule?.formula ?? values?.formula) ?? formulaFromRequest(input.request);
  const style = normalizeStyleRecord(rawRule?.style ?? values?.style ?? values);
  if (!formula || Object.keys(style).length === 0) {
    return undefined;
  }
  return { type: "custom", formula, style };
}

function formulaFromRequest(request: string): string | undefined {
  const explicit = request.match(/=\s*[^.。\n]+/)?.[0]?.trim();
  if (explicit) {
    return explicit;
  }
  const fortyHq = request.match(/\b40HQ\b/i);
  const column = request.match(/\b(?:column|col)\s+([A-Z])\b/i)?.[1] ?? request.match(/\$([A-Z])\d+/)?.[1];
  if (fortyHq && column) {
    return `=$${column.toUpperCase()}2="40HQ"`;
  }
  return undefined;
}

function columnOrderFromInput(input: AgentRunInput): Array<string | number> {
  const values = input.values as Record<string, unknown> | undefined;
  const raw = values?.columnOrder ?? values?.columns ?? values?.order;
  return Array.isArray(raw)
    ? raw.filter((item): item is string | number => typeof item === "string" || typeof item === "number")
    : [];
}

function styleDimensionsFromAgentInput(input: AgentRunInput): StyleDimension[] {
  const values = input.values as Record<string, unknown> | undefined;
  const rawDimensions = values?.dimensions;
  const dimensions = new Set<StyleDimension>();
  if (Array.isArray(rawDimensions)) {
    for (const rawDimension of rawDimensions) {
      const dimension = normalizeStyleDimension(rawDimension);
      if (dimension) dimensions.add(dimension);
    }
  }
  const request = input.request.toLowerCase();
  const requestDimensions: Array<[RegExp, StyleDimension]> = [
    [/\bborders?\b/, "borders"],
    [/\bfills?\b|\bbackgrounds?\b/, "fills"],
    [/\bfonts?\b|\bbold\b|\bitalic\b/, "fonts"],
    [/\balignment\b|\balign\b/, "alignment"],
    [/\bnumber\s*formats?\b|\bdate formats?\b|\bcurrency formats?\b/, "numberFormats"],
    [/\brow heights?\b/, "rowHeights"],
    [/\bcolumn widths?\b/, "columnWidths"]
  ];
  for (const [pattern, dimension] of requestDimensions) {
    if (pattern.test(request)) dimensions.add(dimension);
  }
  return [...dimensions];
}

function normalizeStyleDimension(value: unknown): StyleDimension | undefined {
  if (typeof value !== "string") return undefined;
  const normalized = value.toLowerCase().replace(/[\s_-]+/g, "");
  switch (normalized) {
    case "border":
    case "borders":
      return "borders";
    case "fill":
    case "fills":
    case "background":
    case "backgrounds":
      return "fills";
    case "font":
    case "fonts":
      return "fonts";
    case "alignment":
    case "align":
      return "alignment";
    case "numberformat":
    case "numberformats":
    case "dateformat":
    case "dateformats":
      return "numberFormats";
    case "rowheight":
    case "rowheights":
      return "rowHeights";
    case "columnwidth":
    case "columnwidths":
      return "columnWidths";
    default:
      return undefined;
  }
}

function styleEntriesFromInput(
  metadata: WorkbookMetadata,
  workbookId: WorkbookId,
  input: AgentRunInput
): Array<{ target: A1Range; style: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }> {
  const values = input.values as Record<string, unknown> | undefined;
  const rawInput = input as unknown as Record<string, unknown>;
  const rawEntries = firstArrayValue(values?.entries, values?.patches, rawInput.entries, rawInput.patches);
  const entries: Array<{ target: A1Range; style: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }> = [];
  if (Array.isArray(rawEntries)) {
    for (const rawEntry of rawEntries) {
      if (!rawEntry || typeof rawEntry !== "object") {
        continue;
      }
      const entry = rawEntry as Record<string, unknown>;
      const target = entry.target && typeof entry.target === "object" ? entry.target as Record<string, unknown> : {};
      const sheetName = stringValue(entry.sheetName ?? target.sheetName ?? input.target?.sheetName ?? values?.sheetName);
      const rawAddress = stringValue(entry.address ?? entry.range ?? target.address ?? target.range);
      const style = entry.style && typeof entry.style === "object"
        ? normalizeStyleRecord(entry.style)
        : normalizeStyleRecord(styleLikeValuePatchCell(entry) ?? entry);
      const address = sheetName && rawAddress ? normalizeStyleTargetAddress(metadata, sheetName, rawAddress) : undefined;
      if (sheetName && address && Object.keys(style).length > 0) {
        entries.push({ target: { workbookId, sheetName, address }, style });
      }
    }
  }
  const sheetName = stringValue(input.target?.sheetName ?? values?.sheetName);
  const rawAddress = stringValue(input.target?.range ?? input.target?.address ?? values?.address ?? values?.range);
  const address = sheetName && rawAddress ? normalizeStyleTargetAddress(metadata, sheetName, rawAddress) : undefined;
  const style = normalizeStyleRecord(values?.style ?? values);
  if (entries.length === 0 && sheetName && address && Object.keys(style).length > 0) {
    entries.push({ target: { workbookId, sheetName, address }, style });
  }
  return entries;
}

function mergeEntriesFromInput(
  metadata: WorkbookMetadata,
  workbookId: WorkbookId,
  input: AgentRunInput
): Array<{ target: A1Range; style?: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }> {
  const values = input.values as Record<string, unknown> | undefined;
  const rawInput = input as unknown as Record<string, unknown>;
  const rawEntries = firstArrayValue(values?.mergeRanges, values?.merges, values?.entries, rawInput.mergeRanges, rawInput.merges, rawInput.entries);
  const entries: Array<{ target: A1Range; style?: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }> = [];
  if (Array.isArray(rawEntries)) {
    for (const rawEntry of rawEntries) {
      if (!rawEntry || typeof rawEntry !== "object") {
        continue;
      }
      const entry = rawEntry as Record<string, unknown>;
      const target = entry.target && typeof entry.target === "object" ? entry.target as Record<string, unknown> : {};
      const sheetName = stringValue(entry.sheetName ?? target.sheetName ?? input.target?.sheetName ?? values?.sheetName);
      const rawAddress = stringValue(entry.address ?? entry.range ?? target.address ?? target.range);
      const style = entry.style && typeof entry.style === "object"
        ? normalizeStyleRecord(entry.style)
        : normalizeStyleRecord(styleLikeValuePatchCell(entry) ?? entry);
      const address = sheetName && rawAddress ? stripSheetName(rawAddress).trim() : undefined;
      if (sheetName && address) {
        entries.push({
          target: { workbookId, sheetName, address },
          ...(Object.keys(style).length > 0 ? { style } : {})
        });
      }
    }
  }
  const sheetName = stringValue(input.target?.sheetName ?? values?.sheetName);
  const rawAddress = stringValue(input.target?.range ?? input.target?.address ?? values?.address ?? values?.range);
  if (entries.length === 0 && sheetName && rawAddress) {
    const style = normalizeStyleRecord(values?.style ?? values);
    entries.push({
      target: { workbookId, sheetName, address: stripSheetName(rawAddress).trim() },
      ...(Object.keys(style).length > 0 ? { style } : {})
    });
  }
  return entries;
}

function mergeEntriesFromStyleEntries(
  entries: Array<{ target: A1Range; style: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }>
): Array<{ target: A1Range }> {
  return entries
    .filter((entry) => isMultiCellRange(entry.target.address))
    .map((entry) => ({ target: entry.target }));
}

function isMultiCellRange(address: string): boolean {
  const parsed = tryParseA1Address(stripSheetName(address));
  return Boolean(parsed && (parsed.endRow > parsed.startRow || parsed.endColumn > parsed.startColumn));
}

function shouldMergeRangesFromRequest(input: AgentRunInput): boolean {
  return /\bmerge(?:d|s)?\b/i.test(input.request);
}

function hasMergeBatchInput(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  const rawInput = input as unknown as Record<string, unknown>;
  return shouldMergeRangesFromRequest(input)
    && (Array.isArray(values?.mergeRanges)
      || Array.isArray(values?.merges)
      || Array.isArray(rawInput.mergeRanges)
      || Array.isArray(rawInput.merges));
}

function defaultStyleForMergeRequest(input: AgentRunInput): Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] {
  if (!/\b(center|centered|centre|centred|middle|align|alignment)\b/i.test(input.request)) {
    return {};
  }
  return {
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true
  };
}

function styleLikeValuePatchCell(entry: Record<string, unknown>): unknown {
  const values = entry.values;
  if (!Array.isArray(values) || values.length !== 1 || !Array.isArray(values[0]) || values[0].length !== 1) {
    return undefined;
  }
  const cell = values[0][0];
  if (!cell || typeof cell !== "object" || Array.isArray(cell)) {
    return undefined;
  }
  const style = normalizeStyleRecord(cell);
  return Object.keys(style).length > 0 ? cell : undefined;
}

function hasStyleBatchInput(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  const rawInput = input as unknown as Record<string, unknown>;
  return Array.isArray(values?.entries)
    || Array.isArray(values?.patches)
    || Array.isArray(rawInput.entries)
    || Array.isArray(rawInput.patches);
}

function firstArrayValue(...values: unknown[]): unknown[] | undefined {
  return values.find((value): value is unknown[] => Array.isArray(value));
}

function normalizeStyleTargetAddress(metadata: WorkbookMetadata, sheetName: string, address: string): string {
  const normalized = stripSheetName(address).trim();
  const wholeColumn = /^([A-Z]+):\1$/i.exec(normalized);
  if (!wholeColumn) {
    return normalized;
  }
  const used = tryParseA1Address(stripSheetName(usedRangeForSheet(metadata, sheetName)));
  const endRow = Math.max(1, used?.endRow ?? 1);
  const column = wholeColumn[1]!.toUpperCase();
  return `${column}1:${column}${endRow}`;
}

function rangeMetadataMethodForAction(action: AgentIntentAction | undefined): string | undefined {
  switch (action) {
    case "read_hyperlinks":
      return "range.read_hyperlinks";
    case "read_comments":
      return "range.read_comments";
    case "read_notes":
      return "range.read_notes";
    case "read_merged_cells":
      return "range.read_merged_cells";
    case "read_data_validation":
      return "range.read_data_validation";
    case "read_conditional_formatting":
      return "range.read_conditional_formatting";
    case "search_range":
      return "range.search";
    case "find_blank_cells":
      return "range.find_blank_cells";
    case "find_range_errors":
      return "range.find_errors";
    default:
      return undefined;
  }
}

function dataValidationSummaryAnswer(result: unknown, metadata: WorkbookMetadata, sheetName: string, range: string): Record<string, unknown> {
  const typed = result && typeof result === "object" ? result as Record<string, unknown> : {};
  const data = typed.data && typeof typed.data === "object" ? typed.data as Record<string, unknown> : {};
  const rules = dataValidationRulesFromMetadata(data);
  const summaries = rules.map((rule, index) => {
    const source = dataValidationSource(rule);
    const parsed = parseDataValidationOptions(source);
    const type = stringValue(rule.type ?? data.type);
    const inCellDropDown = booleanValue(rule.inCellDropDown ?? data.inCellDropDown);
    return stripUndefinedRecord({
      index,
      address: stringValue(rule.address ?? data.address ?? range),
      type,
      inCellDropDown,
      source,
      sourceKind: dataValidationSourceKind(source),
      options: parsed.options,
      optionCount: parsed.optionCount,
      sourceComplete: parsed.sourceComplete,
      sourceRange: parsed.sourceRange
    });
  });
  const first = summaries[0] as Record<string, unknown> | undefined;
  const type = stringValue(first?.type ?? data.type);
  const mixedRange = typeof type === "string" && /inconsistent/i.test(type);
  const fieldContext = validationFieldContext(metadata, sheetName, range, summaries);
  return stripUndefinedRecord({
    kind: "data_validation_summary",
    source: "runtime_range_metadata",
    method: "range.read_data_validation",
    sheetName,
    range,
    ruleCount: summaries.length,
    rules: summaries,
    type,
    inCellDropDown: first?.inCellDropDown,
    sourceFormula: typeof first?.source === "string" ? first.source : undefined,
    options: first?.options,
    optionCount: first?.optionCount,
    sourceComplete: summaries.length > 0 ? summaries.every((rule) => rule.sourceComplete === true) : false,
    sourceRange: first?.sourceRange,
    fieldContext,
    validationRangeStatus: mixedRange ? "mixed_or_inconsistent_range" : undefined,
    guidance: mixedRange
      ? "This multi-cell range has mixed/inconsistent validation. Do not conclude the dropdown option is missing or the rule is broken from this range summary. Read one representative data cell with the dropdown, such as the selected cell or a known row in the Transaction Type column, to inspect its exact validation source."
      : "Use this inline validation summary to answer dropdown option questions. Add missing dropdown values by updating the source-list cells when sourceRange is present; otherwise preview a write_data_validation rule rewrite."
  });
}

function validationFieldContext(
  metadata: WorkbookMetadata,
  sheetName: string,
  range: string,
  summaries: Array<Record<string, unknown>>
): Array<Record<string, unknown>> | undefined {
  if (summaries.length === 0) {
    return undefined;
  }
  return summaries.map((summary) => {
    const address = stringValue(summary.address) ?? range;
    const bounds = columnRangeBounds(address, metadata, sheetName) ?? columnRangeBounds(range, metadata, sheetName);
    const column = bounds ? metadataColumnsForSheet(metadata, sheetName).find((candidate) => candidate.index === bounds.startColumn - 1) : undefined;
    const validation = fieldValidationContext(summary);
    const allowedValues = Array.isArray(summary.options) ? summary.options.filter((value): value is string => typeof value === "string") : undefined;
    return stripUndefinedRecord({
      field: column?.name ?? (bounds ? columnLetter(bounds.startColumn - 1) : undefined),
      range: address,
      headerRange: column ? headerRangeForColumn(metadata, sheetName, column) : undefined,
      semanticType: column?.role,
      dataType: column?.inferredType,
      hasValidation: true,
      allowedValues,
      allowedValueCount: typeof summary.optionCount === "number" ? summary.optionCount : allowedValues?.length,
      allowedValuesTruncated: false,
      validation,
      examples: allowedValues?.slice(0, 3)
    });
  });
}

function tableFieldContext(
  table: TableMetadata,
  columns: TableFieldContextColumn[],
  data: { values?: CellMatrix | undefined; text?: CellMatrix | undefined; formulas?: CellMatrix | undefined; numberFormat?: CellMatrix | undefined }
): Array<Record<string, unknown>> | undefined {
  if (columns.length === 0) {
    return undefined;
  }
  const matrix = data.values ?? data.text ?? [];
  const contexts = columns.slice(0, 24).map((column, projectedIndex) => {
    const values = matrix.map((row) => row[projectedIndex]);
    const presentValues = values.filter(hasCellValue);
    const distinct = compactDistinctValues(presentValues, 5);
    const examples = compactDistinctValues(presentValues, 2);
    const formats = data.numberFormat ? compactDistinctValues(data.numberFormat.map((row) => row[projectedIndex]).filter(hasCellValue), 2) : undefined;
    const formulas = data.formulas ? data.formulas.map((row) => row[projectedIndex]).filter(hasCellValue) : [];
    return stripUndefinedRecord({
      field: column.name,
      range: tableColumnRange(table.dataRange ?? table.range, column.letter),
      headerRange: tableColumnHeaderRange(table.range, column.letter),
      semanticType: column.role,
      dataType: column.inferredType,
      numberFormats: formats && formats.length > 0 ? formats : undefined,
      hasFormulas: formulas.length > 0 ? true : undefined,
      currentDistinctValues: distinct.length > 0 ? distinct : undefined,
      currentDistinctValueCount: distinct.length,
      blankCount: matrix.length > 0 ? values.length - presentValues.length : undefined,
      examples: examples.length > 0 ? examples : undefined
    });
  });
  return contexts.length > 0 ? contexts : undefined;
}

type TableFieldContextColumn = {
  index: number;
  letter: string;
  name: string;
  inferredType: ColumnMetadata["inferredType"];
  role?: ColumnMetadata["role"] | undefined;
  importance?: number | undefined;
};

function compactDistinctValues(values: unknown[], limit: number): unknown[] {
  const seen = new Set<string>();
  const distinct: unknown[] = [];
  for (const value of values) {
    const compact = compactFieldContextValue(value);
    const key = typeof compact === "string" ? compact.trim() : JSON.stringify(compact);
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    distinct.push(compact);
    if (distinct.length >= limit) {
      break;
    }
  }
  return distinct;
}

function compactFieldContextValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }
  const normalized = value.replace(/\s+/g, " ").trim();
  return normalized.length <= 48 ? normalized : `${normalized.slice(0, 47)}…`;
}

function tableColumnRange(range: string | undefined, columnLetterValue: string | undefined): string | undefined {
  if (!range || !columnLetterValue) {
    return undefined;
  }
  const parsed = rangeShape(stripSheetName(range));
  return parsed ? `${columnLetterValue}${parsed.startRow}:${columnLetterValue}${parsed.endRow}` : undefined;
}

function tableColumnHeaderRange(range: string | undefined, columnLetterValue: string | undefined): string | undefined {
  if (!range || !columnLetterValue) {
    return undefined;
  }
  const parsed = rangeShape(stripSheetName(range));
  return parsed ? `${columnLetterValue}${parsed.startRow}` : undefined;
}

function fieldValidationContext(summary: Record<string, unknown>): Record<string, unknown> | undefined {
  const type = stringValue(summary.type);
  const sourceKind = stringValue(summary.sourceKind);
  const options = Array.isArray(summary.options) ? summary.options.filter((value): value is string => typeof value === "string") : undefined;
  const sourceRange = stringValue(summary.sourceRange);
  const source = stringValue(summary.source);
  if (sourceKind === "inline_list") {
    return stripUndefinedRecord({
      type,
      sourceType: "inline",
      options,
      optionCount: typeof summary.optionCount === "number" ? summary.optionCount : options?.length,
      truncated: false
    });
  }
  if (sourceKind === "range_formula" && sourceRange) {
    const dynamic = /\b(INDIRECT|OFFSET|CHOOSE|INDEX|FILTER)\b/i.test(sourceRange);
    return stripUndefinedRecord({
      type,
      sourceType: dynamic ? "formula" : "range",
      sourceRange: dynamic ? undefined : sourceRange,
      formula: dynamic ? source : undefined,
      optionsResolved: dynamic ? false : options !== undefined,
      reason: dynamic ? "Dynamic or dependent dropdown; options depend on formula evaluation." : undefined,
      options,
      optionCount: typeof summary.optionCount === "number" ? summary.optionCount : options?.length,
      truncated: false
    });
  }
  return stripUndefinedRecord({
    type,
    sourceType: sourceKind,
    optionsResolved: options !== undefined,
    options,
    optionCount: typeof summary.optionCount === "number" ? summary.optionCount : options?.length,
    truncated: false
  });
}

function fieldHasEnforceableDropdownContext(field: Record<string, unknown>): boolean {
  const validation = field.validation && typeof field.validation === "object" ? field.validation as Record<string, unknown> : {};
  const sourceType = stringValue(validation.sourceType);
  const semanticType = stringValue(field.semanticType);
  const dataType = stringValue(field.dataType);
  const fieldName = stringValue(field.field) ?? "";
  const allowedValues = Array.isArray(field.allowedValues) ? field.allowedValues : [];
  return allowedValues.length > 0
    && sourceType !== "formula"
    && (semanticType === "status" || semanticType === "category" || dataType === "status" || /\b(status|state|stage|category|type)\b/i.test(fieldName));
}

function uniqueProposedValidationValues(values: CellMatrix): string[] {
  const seen = new Set<string>();
  const proposed: string[] = [];
  for (const row of values) {
    for (const cell of row) {
      if (cell === null || cell === undefined || cell === "") {
        continue;
      }
      const value = String(cell).trim();
      const normalized = normalizeValidationOption(value);
      if (!normalized || seen.has(normalized)) {
        continue;
      }
      seen.add(normalized);
      proposed.push(value);
    }
  }
  return proposed;
}

function normalizeValidationOption(value: string): string {
  return value.trim().toLowerCase().normalize("NFKC");
}

function headerRangeForColumn(metadata: WorkbookMetadata, sheetName: string, column: ColumnMetadata): string | undefined {
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const header = sheet?.headers.find((candidate) => candidate.columns.some((item) => item.index === column.index));
  return header ? `${column.letter}${header.row}` : undefined;
}

function dataValidationRulesFromMetadata(data: Record<string, unknown>): Record<string, unknown>[] {
  if (Array.isArray(data.rules)) {
    return data.rules.filter((rule): rule is Record<string, unknown> => Boolean(rule && typeof rule === "object"));
  }
  const rule = data.rule && typeof data.rule === "object" ? data.rule as Record<string, unknown> : undefined;
  if (rule) {
    const list = rule.list && typeof rule.list === "object" ? rule.list as Record<string, unknown> : undefined;
    return [stripUndefinedRecord({
      address: data.address,
      type: stringValue(data.type) ?? (list ? "list" : undefined),
      source: list?.source ?? rule.source,
      inCellDropDown: list?.inCellDropDown,
      ignoreBlanks: data.ignoreBlanks,
      valid: data.valid
    })];
  }
  if (data.type || data.source) {
    return [data];
  }
  return [];
}

function dataValidationSource(rule: Record<string, unknown>): unknown {
  const source = rule.source ?? rule.formula1;
  if (source !== undefined) {
    return source;
  }
  const list = rule.list && typeof rule.list === "object" ? rule.list as Record<string, unknown> : undefined;
  return list?.source;
}

function dataValidationSourceKind(source: unknown): string | undefined {
  if (Array.isArray(source)) {
    return "inline_list";
  }
  if (typeof source !== "string") {
    return undefined;
  }
  const trimmed = source.trim();
  if (trimmed.startsWith("=")) {
    return "range_formula";
  }
  return "inline_list";
}

function parseDataValidationOptions(source: unknown): { options?: string[]; optionCount: number; sourceComplete: boolean; sourceRange?: string } {
  if (Array.isArray(source)) {
    const options = source.filter((item): item is string => typeof item === "string" && item.trim().length > 0).map((item) => item.trim());
    return { options, optionCount: options.length, sourceComplete: true };
  }
  if (typeof source !== "string") {
    return { optionCount: 0, sourceComplete: false };
  }
  const trimmed = source.trim();
  if (trimmed.startsWith("=")) {
    return { optionCount: 0, sourceComplete: false, sourceRange: trimmed.slice(1) };
  }
  const options = trimmed.split(",").map((item) => item.trim()).filter(Boolean);
  return { options, optionCount: options.length, sourceComplete: true };
}

function hasRangeMetadataReadIntent(input: AgentRunInput): boolean {
  return rangeMetadataMethodForAction(intentAction(input) ?? inferredRangeMetadataReadAction(input)) !== undefined;
}

function inferredRangeMetadataReadAction(input: AgentRunInput): AgentIntentAction | undefined {
  const request = input.request.toLowerCase();
  const readVerb = /\b(read|show|check|inspect|review|tell|what|which|whether|does|do|has|have|list|find)\b/.test(request);
  if (readVerb && isDataValidationReadRequest(request)) {
    return "read_data_validation";
  }
  if (readVerb && /\b(conditional\s+format|conditional\s+formatting)\b/.test(request)) {
    return "read_conditional_formatting";
  }
  if (readVerb && /\b(merged cells?|merge ranges?|merged ranges?)\b/.test(request)) {
    return "read_merged_cells";
  }
  if (readVerb && /\b(hyperlinks?|links?)\b/.test(request)) {
    return "read_hyperlinks";
  }
  if (readVerb && /\b(comments?)\b/.test(request)) {
    return "read_comments";
  }
  if (readVerb && /\b(notes?)\b/.test(request)) {
    return "read_notes";
  }
  return undefined;
}

function isDataValidationReadRequest(request: string): boolean {
  if (isExplicitCellValueReadRequest(request)) {
    return false;
  }
  return /\b(data\s+validation|validation|dropdown|drop\s*down|select\s+list|selection\s+list|allowed values?)\b/.test(request);
}

function isExplicitCellValueReadRequest(request: string): boolean {
  return /\b(read|show|list|fetch|give|check|tell)\b/.test(request)
    && /\b(actual\s+cell\s+values?|cell\s+values?|raw\s+values?|literal\s+text|text\s+strings?|stored\s+in\s+cells?|cells?\s+contain|contains?\s+the\s+text|values?\s+in\s+(?:column|range)|column\s+[a-z]+\s*,?\s+rows?\s+\d+)\b/i.test(request);
}

function shouldSummarizeGroupedHeader(input: AgentRunInput): boolean {
  const request = input.request.toLowerCase();
  const readVerb = /\b(read|show|check|inspect|review|tell|what|which|summari[sz]e|examine|look(?:\s+at)?)\b/.test(request);
  return readVerb
    && /\bgroup(?:ed)?\s+headers?\b|\bheader\s+groups?\b|\bmerged\s+headers?\b/.test(request)
    && !/\b(apply|update|change|set|make|merge\s+these|preview|fix|create|insert|delete|remove)\b/.test(request);
}

function groupedHeaderSummaryRange(metadata: WorkbookMetadata, sheetName: string, normalizedRange: string, input: AgentRunInput): string {
  const explicitRange = stringValue(input.target?.range ?? input.target?.address);
  if (explicitRange && isSingleRowRange(explicitRange)) {
    return stripSheetName(explicitRange);
  }
  const requestedRow = numberValue(/\brow\s*(\d+)\b/i.exec(input.request)?.[1]) ?? 1;
  const parsed = tryParseA1Address(stripSheetName(normalizedRange))
    ?? tryParseA1Address(stripSheetName(usedRangeForSheet(metadata, sheetName)));
  if (!parsed) {
    return `A${requestedRow}:XFD${requestedRow}`;
  }
  return `${numberToColumn(parsed.startColumn)}${requestedRow}:${numberToColumn(parsed.endColumn)}${requestedRow}`;
}

function isSingleRowRange(address: string): boolean {
  const parsed = tryParseA1Address(stripSheetName(address));
  return Boolean(parsed && parsed.startRow === parsed.endRow);
}

function mergedRangesFromMetadataResult(result: unknown): string[] {
  const data = result && typeof result === "object" ? (result as Record<string, unknown>).data : undefined;
  if (!data || typeof data !== "object") {
    return [];
  }
  const record = data as Record<string, unknown>;
  if (record.isNullObject === true) {
    return [];
  }
  const address = stringValue(record.address);
  if (!address) {
    return [];
  }
  return address
    .split(/\s*,\s*/)
    .map((part) => stripSheetName(part).replace(/\$/g, "").trim())
    .filter((part) => part.length > 0);
}

function groupedHeaderSpansFromSnapshot(headerRange: string, snapshot: RangeSnapshot | undefined, mergedRanges: string[]) {
  return mergedRanges.map((range) => ({
    range,
    label: groupedHeaderLabelForRange(headerRange, snapshot, range),
    merged: true
  }));
}

function groupedHeaderUnmergedLabels(headerRange: string, snapshot: RangeSnapshot | undefined, mergedRanges: string[]) {
  const parsedHeader = tryParseA1Address(stripSheetName(headerRange));
  if (!parsedHeader) {
    return [];
  }
  const covered = new Set<number>();
  for (const range of mergedRanges) {
    const parsed = tryParseA1Address(stripSheetName(range));
    if (!parsed) continue;
    for (let column = parsed.startColumn; column <= parsed.endColumn; column += 1) {
      covered.add(column);
    }
  }
  const textRow = (snapshot?.text ?? snapshot?.values ?? [])[0] ?? [];
  const labels: Array<{ cell: string; label: unknown; merged: false }> = [];
  for (let column = parsedHeader.startColumn; column <= parsedHeader.endColumn; column += 1) {
    if (covered.has(column)) continue;
    const offset = column - parsedHeader.startColumn;
    const value = textRow[offset];
    if (value !== undefined && value !== null && String(value).trim().length > 0) {
      labels.push({ cell: `${numberToColumn(column)}${parsedHeader.startRow}`, label: value, merged: false });
    }
  }
  return labels;
}

function groupedHeaderLabelForRange(headerRange: string, snapshot: RangeSnapshot | undefined, range: string): unknown {
  const parsedHeader = tryParseA1Address(stripSheetName(headerRange));
  const parsedRange = tryParseA1Address(stripSheetName(range));
  if (!parsedHeader || !parsedRange) {
    return undefined;
  }
  const offset = parsedRange.startColumn - parsedHeader.startColumn;
  return (snapshot?.text?.[0]?.[offset] ?? snapshot?.values?.[0]?.[offset]) as unknown;
}

function searchResultHasNoMatches(result: unknown): boolean {
  if ((result as { ok?: boolean }).ok === false) {
    return false;
  }
  const record = result && typeof result === "object" ? result as Record<string, unknown> : {};
  const data = record.data && typeof record.data === "object" ? record.data as Record<string, unknown> : record;
  if (typeof data.count === "number") {
    return data.count === 0;
  }
  if (Array.isArray(data.matches)) {
    return data.matches.length === 0;
  }
  if (Array.isArray(data.areas)) {
    return data.areas.length === 0;
  }
  return false;
}

function workflowPlanForAction(action: AgentIntentAction | undefined) {
  const plans: Record<string, {
    workflow: string;
    title: string;
    mutatesWorkbook: boolean;
    steps: string[];
    requiredCapabilities: string[];
    continuation: string;
    warnings: string[];
  }> = {
    prepare_session: {
      workflow: "excel.workflow.prepare_session",
      title: "Prepare workbook session",
      mutatesWorkbook: false,
      steps: ["Read runtime status", "Read active workbook context", "Summarize workbook map", "Summarize collaboration state"],
      requiredCapabilities: ["excel.runtime.get_status", "excel.runtime.get_active_context", "excel.runtime.get_capabilities", "excel.workbook.get_workbook_map", "excel.collab.get_status"],
      continuation: "Use the returned workbookContextId for follow-up answer, preview, validation, or rollback calls.",
      warnings: []
    },
    create_formula_sheet: {
      workflow: "excel.workflow.create_formula_sheet",
      title: "Create formula sheet",
      mutatesWorkbook: true,
      steps: ["Create or copy the target sheet", "Write constants", "Write formulas", "Apply number formats", "Validate formulas"],
      requiredCapabilities: ["excel.sheet.create", "excel.range.write_values", "excel.range.write_formulas", "excel.range.write_number_formats", "excel.validate.formulas"],
      continuation: "Use preview_update/apply_update with explicit sheet, values, formulas, and formats; validate formulas after apply.",
      warnings: ["This planning route does not apply workbook mutations directly."]
    },
    create_template_report: {
      workflow: "excel.workflow.create_template_report",
      title: "Create template report",
      mutatesWorkbook: true,
      steps: ["Resolve the registered template", "Create a report sheet from the template", "Clear declared data regions", "Fill declared regions", "Compare and repair styles", "Validate against the template"],
      requiredCapabilities: ["excel.template.get", "excel.template.create_sheet_from_template", "excel.template.clear_data_regions", "excel.template.fill_regions", "excel.style.compare_fingerprint", "excel.style.repair_consistency", "excel.template.validate_sheet_against_template"],
      continuation: "Use copy_template_sheet, clear_template_data_regions, fill_template_regions, repair_style_consistency, and validate_sheet_against_template through preview/apply where required.",
      warnings: ["Template data-region filling must use explicit region values; the backend will not infer new data."]
    },
    create_pivot_chart_summary: {
      workflow: "excel.workflow.create_pivot_chart_summary",
      title: "Create PivotTable and chart summary",
      mutatesWorkbook: true,
      steps: ["Check PivotTable capability status", "Create or refresh the PivotTable", "Create or refresh the chart", "Validate the PivotTable source"],
      requiredCapabilities: ["excel.pivot.get_capability_matrix", "excel.pivot.create", "excel.pivot.refresh", "excel.chart.create", "excel.chart.refresh", "excel.pivot.validate_source"],
      continuation: "Report host capability limits honestly; PivotTable and chart execution remains host-limited until the active Excel runtime exposes deterministic support.",
      warnings: ["PivotTable and chart operations are host-limited in the current orchestration plan."]
    },
    preview_risky_edit: {
      workflow: "excel.workflow.preview_risky_edit",
      title: "Preview risky edit with proof",
      mutatesWorkbook: true,
      steps: ["Capture a before snapshot", "Create a scoped plan", "Preview or apply the scoped plan", "Capture an after snapshot", "Compare snapshots", "Prepare rollback preview"],
      requiredCapabilities: ["excel.snapshot.create", "excel.plan.create", "excel.plan.preview", "excel.plan.apply", "excel.snapshot.compare_compact", "excel.transaction.preview_rollback"],
      continuation: "Use preview_update/apply_update for the scoped edit, then compare snapshots or request rollback proof if needed.",
      warnings: ["Sparse/null-padded broad writes remain blocked unless explicitly confirmed as intentional."]
    },
    inspect_analyze: {
      workflow: "excel.workflow.inspect_analyze",
      title: "Inspect and analyze tabular data",
      mutatesWorkbook: false,
      steps: ["Read a compact target range or table", "Profile shape and missing values", "Detect duplicate rows", "Summarize numeric columns", "Store detailed proof behind the workbook context"],
      requiredCapabilities: ["excel.range.read_compact", "excel.table.read_compact"],
      continuation: "Answer from the compact analysis summary or ask for a narrower target when the request is ambiguous.",
      warnings: []
    },
    rollback_validate: {
      workflow: "excel.workflow.rollback_validate",
      title: "Rollback and validate workbook",
      mutatesWorkbook: true,
      steps: ["Resolve transaction or backup target", "Run rollback or backup restore", "Recalculate workbook", "Run compact workbook validation", "Return validation proof"],
      requiredCapabilities: ["excel.transaction.rollback", "excel.workbook.restore_backup", "excel.workbook.calculate", "excel.validate.workbook"],
      continuation: "Use rollback mode or the specific backup/transaction action with confirmation, then validate the workbook.",
      warnings: ["Rollback and restore require explicit confirmation tokens."]
    }
  };
  return plans[action ?? ""];
}

function workflowTargetFromInput(metadata: WorkbookMetadata, input: AgentRunInput): { sheetName: string; range: string } | undefined {
  const target = input.target;
  if (target?.sheetName) {
    return { sheetName: target.sheetName, range: target.range ?? usedRangeForSheet(metadata, target.sheetName) };
  }
  if (target?.tableName) {
    const table = metadata.tables.find((candidate) => candidate.name === target.tableName);
    if (table) {
      return { sheetName: table.sheetName, range: table.dataRange ?? table.range };
    }
  }
  const firstSheet = metadata.sheets.find((sheet) => sheet.usedRange);
  return firstSheet?.usedRange ? { sheetName: firstSheet.name, range: firstSheet.usedRange } : undefined;
}

function resolveTemplateSourceSheet(metadata: WorkbookMetadata, input: AgentRunInput): WorkbookMetadata["sheets"][number] | undefined {
  if (input.target?.sheetName) {
    const normalized = normalizeComparableText(input.target.sheetName);
    return metadata.sheets.find((sheet) => normalizeComparableText(sheet.name) === normalized);
  }
  const mentioned = metadata.sheets.find((sheet) => normalizeComparableText(input.request).includes(normalizeComparableText(sheet.name)));
  if (mentioned) {
    return mentioned;
  }
  if (/\blatest\b/i.test(input.request)) {
    const monthly = monthSheetCandidates(metadata);
    if (monthly.length === 1) {
      return monthly[0];
    }
    return undefined;
  }
  return undefined;
}

function latestTemplateCandidates(metadata: WorkbookMetadata, input: AgentRunInput): AgentCandidate[] {
  if (input.target?.sheetName || input.target?.candidateId || !/\blatest\b/i.test(input.request)) {
    return [];
  }
  const request = normalizeComparableText(input.request);
  if (metadata.sheets.some((sheet) => request.includes(normalizeComparableText(sheet.name)))) {
    return [];
  }
  return monthSheetCandidates(metadata).map((sheet, index) => ({
    id: `sheet:${sheet.name}`,
    kind: "sheet" as const,
    label: sheet.name,
    sheetName: sheet.name,
    ...(sheet.usedRange !== undefined ? { range: sheet.usedRange } : {}),
    confidence: 0.88 - index * 0.02
  }));
}

function monthSheetCandidates(metadata: WorkbookMetadata): WorkbookMetadata["sheets"] {
  return metadata.sheets.filter((sheet) => isMonthLikeSheetName(sheet.name));
}

function isMonthLikeSheetName(name: string): boolean {
  return /^(jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)(\s+\d{4})?$/i.test(name.trim());
}

function uniqueSheetName(metadata: WorkbookMetadata, baseName: string): string {
  const existing = new Set(metadata.sheets.map((sheet) => sheet.name));
  if (!existing.has(baseName)) {
    return baseName;
  }
  for (let index = 2; index < 100; index += 1) {
    const candidate = `${baseName} ${index}`;
    if (!existing.has(candidate)) {
      return candidate;
    }
  }
  return `${baseName} ${Date.now()}`;
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function explicitSheetName(metadata: WorkbookMetadata, input: AgentRunInput): string | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const rawName = stringValue(input.target?.sheetName)
    ?? stringValue(values?.sourceSheetName)
    ?? stringValue(values?.sheetName);
  if (rawName) {
    const normalized = normalizeComparableText(rawName);
    return metadata.sheets.find((sheet) => normalizeComparableText(sheet.name) === normalized)?.name ?? rawName;
  }
  if (requestMentionsActiveSheet(input.request)) {
    return metadata.workbook.activeSheet;
  }
  return findMentionedSheet(metadata, input)?.name;
}

function sheetNeedsInput(metadata: WorkbookMetadata, requestedMode: AgentRunMode, summary: string): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: []
  };
}

function formulaRuntimeErrorOutput(metadata: WorkbookMetadata, requestedMode: AgentRunMode, summary: string, result: unknown): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "ERROR",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    answer: result,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "manual_review",
    warnings: []
  };
}

function formulaWarnings(warnings: unknown): string[] {
  return Array.isArray(warnings)
    ? warnings.map((warning) => typeof warning === "string" ? warning : (warning as { message?: unknown })?.message).filter((message): message is string => typeof message === "string")
    : [];
}

function nameSelectorFromInput(metadata: WorkbookMetadata, input: AgentRunInput): { name: string; sheetName?: string } | undefined {
  const raw = nameOrRegionNameFromInput(input);
  if (!raw) {
    return undefined;
  }
  const match = metadata.namedRanges.find((name) => normalizeComparableText(name.name) === normalizeComparableText(raw));
  return {
    name: match?.name ?? raw,
    ...(match?.sheetName !== undefined ? { sheetName: match.sheetName } : {})
  };
}

function nameMutationRequestFromInput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  behavior: "create" | "update" | "delete"
): NameCreateRequest | NameUpdateRequest | NameSelector | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const name = stringValue(values?.name ?? input.target?.entity);
  if (!name) {
    return undefined;
  }
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  const sheetName = stringValue(values?.sheetName ?? input.target?.sheetName);
  const selector: NameSelector = { workbookId, name, ...(sheetName ? { sheetName } : {}) };
  if (behavior === "delete") {
    return selector;
  }
  const targetReference = sheetName && input.target?.range ? `${sheetName}!${input.target.range}` : undefined;
  const reference = stringValue(values?.reference ?? values?.range ?? targetReference);
  const formula = stringValue(values?.formula);
  if (!reference && !formula) {
    return undefined;
  }
  const comment = stringValue(values?.comment);
  return {
    ...selector,
    ...(reference ? { reference } : {}),
    ...(formula ? { formula } : {}),
    ...(comment ? { comment } : {}),
    ...(typeof values?.visible === "boolean" ? { visible: values.visible } : {})
  };
}

function regionMutationRequestFromInput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  behavior: "register" | "clear_values" | "write_values" | "fill"
): RegionRegisterRequest | RegionSelector | (RegionSelector & { values: unknown[][]; clearFirst?: boolean }) | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  const regionName = stringValue(values?.regionName ?? values?.name ?? input.target?.entity);
  if (!regionName) {
    return undefined;
  }
  if (behavior === "register") {
    const sheetName = stringValue(values?.sheetName ?? input.target?.sheetName);
    const address = stringValue(values?.address ?? values?.range ?? input.target?.range);
    if (!sheetName || !address) {
      return undefined;
    }
    const kind = regionKindFromInput(values?.kind);
    const description = stringValue(values?.description);
    const templateId = stringValue(values?.templateId);
    return {
      workbookId,
      name: regionName,
      sheetName,
      address,
      ...(kind ? { kind } : {}),
      ...(description ? { description } : {}),
      ...(templateId ? { templateId: templateId as TemplateId } : {}),
      ...(typeof values?.createNamedRange === "boolean" ? { createNamedRange: values.createNamedRange } : {})
    };
  }
  const selector: RegionSelector = { workbookId, regionName };
  if (behavior === "clear_values") {
    return selector;
  }
  const matrix = explicitCellMatrixFromValues(values);
  if (!hasCellValues(matrix)) {
    return undefined;
  }
  return {
    ...selector,
    values: matrix,
    ...(behavior === "fill" && typeof values?.clearFirst === "boolean" ? { clearFirst: values.clearFirst } : {})
  };
}

function regionKindFromInput(value: unknown): RegionRegisterRequest["kind"] | undefined {
  return typeof value === "string" && ["data", "header", "formula", "output", "template", "table", "named-range", "other"].includes(value)
    ? value as RegionRegisterRequest["kind"]
    : undefined;
}

function explicitCellMatrixFromValues(values: Record<string, unknown> | undefined): CellMatrix {
  if (Array.isArray(values?.rows)) return values.rows as CellMatrix;
  if (Array.isArray(values?.values)) return values.values as CellMatrix;
  return [];
}

function regionNameFromInput(input: AgentRunInput): string | undefined {
  return nameOrRegionNameFromInput(input);
}

function nameOrRegionNameFromInput(input: AgentRunInput): string | undefined {
  if (input.target?.candidateId?.startsWith("name:")) {
    return input.target.candidateId.slice("name:".length);
  }
  const values = input.values as Record<string, unknown> | undefined;
  return stringValue(input.target?.entity)
    ?? stringValue(values?.name)
    ?? stringValue(values?.regionName)
    ?? stringValue(values?.namedItem);
}

function nameRegionNeedsInput(metadata: WorkbookMetadata, requestedMode: AgentRunMode, summary: string): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: []
  };
}

function backupRangesFromInput(workbookId: WorkbookId, input: AgentRunInput): A1Range[] {
  const ranges = (input.values as { ranges?: unknown } | undefined)?.ranges;
  if (Array.isArray(ranges)) {
    return ranges.flatMap((range) => {
      if (!range || typeof range !== "object") {
        return [];
      }
      const item = range as { sheetName?: unknown; address?: unknown; range?: unknown };
      const sheetName = stringValue(item.sheetName);
      const address = stringValue(item.address ?? item.range);
      return sheetName && address ? [{ workbookId, sheetName, address }] : [];
    });
  }
  const sheetName = stringValue(input.target?.sheetName);
  const address = stringValue(input.target?.range);
  return sheetName && address ? [{ workbookId, sheetName, address }] : [];
}

function findMentionedSheet(metadata: WorkbookMetadata, input: AgentRunInput): WorkbookMetadata["sheets"][number] | undefined {
  if (input.target?.sheetName) {
    const normalized = normalizeComparableText(input.target.sheetName);
    return metadata.sheets.find((sheet) => normalizeComparableText(sheet.name) === normalized);
  }
  const request = normalizeComparableText(input.request);
  return metadata.sheets.find((sheet) => request.includes(normalizeComparableText(sheet.name)));
}

function requestMentionsActiveSheet(request: string): boolean {
  return /\b(active|current|this)\s+sheet\b/i.test(request) || /\bthis\s+sheet\b/i.test(request);
}

function normalizeOperationRange(metadata: WorkbookMetadata, sheetName: string, range: string): string {
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const rowCount = Math.max(1, sheet?.rowCount ?? 1);
  const columnOnly = /^([A-Z]+):([A-Z]+)$/i.exec(range.replace(/\$/g, ""));
  if (columnOnly?.[1] && columnOnly[2]) {
    return `${columnOnly[1].toUpperCase()}1:${columnOnly[2].toUpperCase()}${rowCount}`;
  }
  return range;
}

function structuralRangeOperationKind(
  input: AgentRunInput,
  resolved: Extract<AgentTargetResolution, { ok: true }>,
  requestedKind: RangeStructuralOperationKind
): RangeStructuralOperationKind {
  if (!isRowColumnStructuralKind(requestedKind)) {
    return requestedKind;
  }
  const requestedRow = requestMentionsStructuralRows(input.request) || input.target?.row !== undefined;
  const requestedColumn = requestMentionsStructuralColumns(input.request) || input.target?.column !== undefined;
  const shape = structuralAddressShape(input.target?.range ?? input.target?.address ?? resolved.range);
  if (requestedRow && !requestedColumn) {
    return requestedKind.includes("insert") ? "range.insert_rows" : "range.delete_rows";
  }
  if (requestedColumn && !requestedRow) {
    return requestedKind.includes("insert") ? "range.insert_columns" : "range.delete_columns";
  }
  if (shape === "row" && requestedKind.endsWith("_columns")) {
    return requestedKind.includes("insert") ? "range.insert_rows" : "range.delete_rows";
  }
  if (shape === "column" && requestedKind.endsWith("_rows")) {
    return requestedKind.includes("insert") ? "range.insert_columns" : "range.delete_columns";
  }
  return requestedKind;
}

function intentStructuralOperationKind(input: AgentRunInput): RangeStructuralOperationKind | undefined {
  switch (input.intent?.action) {
    case "insert_rows":
      return "range.insert_rows";
    case "delete_rows":
      return "range.delete_rows";
    case "insert_columns":
      return "range.insert_columns";
    case "delete_columns":
      return "range.delete_columns";
    default:
      return undefined;
  }
}

function isRowColumnStructuralKind(kind: RangeStructuralOperationKind): boolean {
  return kind === "range.insert_rows" || kind === "range.delete_rows" || kind === "range.insert_columns" || kind === "range.delete_columns";
}

function requestMentionsStructuralRows(request: string): boolean {
  return /\b(?:insert|delete|remove|drop)\b.*\brows?\b|\brows?\b.*\b(?:insert|delete|remove|drop)\b|\b(?:this|selected|active|current)\s+row\b/i.test(request);
}

function requestMentionsStructuralColumns(request: string): boolean {
  return /\b(?:insert|delete|remove|drop)\b.*\b(?:cols?|columns?)\b|\b(?:cols?|columns?)\b.*\b(?:insert|delete|remove|drop)\b|\b(?:this|selected|active|current)\s+col(?:umn)?\b/i.test(request);
}

function structuralAddressShape(address: string): "row" | "column" | undefined {
  const normalized = stripSheetName(address).replace(/\$/g, "").replace(/\s+/g, "").toUpperCase();
  if (/^\d+:\d+$/.test(normalized)) {
    return "row";
  }
  if (/^[A-Z]+:[A-Z]+$/.test(normalized)) {
    return "column";
  }
  const parsed = tryParseA1Address(normalized);
  if (!parsed) {
    return undefined;
  }
  if (parsed.startRow === parsed.endRow && parsed.startColumn !== parsed.endColumn) {
    return "row";
  }
  if (parsed.startColumn === parsed.endColumn && parsed.startRow !== parsed.endRow) {
    return "column";
  }
  return undefined;
}

interface SimilarRowRange {
  sheetName: string;
  range: string;
  reason: string;
}

interface SimilarRowSignals {
  tokens: string[];
  numbers: number[];
  predicates: SimilarRowPredicate[];
}

interface SimilarRowMatch {
  sheetName: string;
  range: string;
  sheetRowNumber: number;
  rowIndex: number;
  values: unknown[];
  columns: Array<{ letter: string; name: string; role?: string; value: unknown }>;
  score: number;
  matchedSignals: string[];
  matchedColumns: Array<{ letter: string; name: string; value: unknown; signals: string[] }>;
  whyMatched: string;
}

interface SimilarRowPredicate {
  label: string;
  value: string | number;
  match: "contains" | "equals" | "number_equals";
  headerPattern?: RegExp;
  role?: string;
}

interface StyleReferenceCandidate {
  sheetName: string;
  range: string;
  label: string;
  sourceKind: "section" | "table" | "sheet" | "header";
  confidence: number;
  reason: string;
  nextAction: string;
}

function compactWorkbookContextHints(metadata: WorkbookMetadata, sheetName: string, table?: TableMetadata): string[] {
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const columns = table?.columns ?? sheet?.headers.flatMap((header) => header.columns) ?? [];
  const labelColumns = columns
    .filter((column) => column.inferredType === "status" || /status|label|category|type|class|tag|allowed|approval|state/i.test(column.name))
    .map((column) => `${column.letter}:${column.name}`)
    .slice(0, 6);
  const hints = [
    labelColumns.length > 0 ? `label/status columns: ${labelColumns.join(", ")}` : undefined,
    labelColumns.length > 0 ? "dropdown rules: call read_data_validation on the exact label/status column before guessing allowed values" : undefined,
    sheet?.usedRange ? `style context: call read_style_summary/read_style_fingerprint on ${sheet.name}!${table?.range ?? sheet.usedRange}` : undefined,
    likelyHistoricalSheets(metadata, sheetName).length > 0 ? `related sheets: ${likelyHistoricalSheets(metadata, sheetName).slice(0, 4).join(", ")}` : undefined,
    sheet?.kind === "transaction" || table ? "historical labels: call find_similar_rows from the current row/range for prior examples" : undefined
  ];
  return hints.filter((hint): hint is string => typeof hint === "string").slice(0, 6);
}

function likelyHistoricalSheets(metadata: WorkbookMetadata, sheetName: string): string[] {
  const source = metadata.sheets.find((sheet) => sheet.name === sheetName);
  if (!source) {
    return [];
  }
  const sourceHeaders = new Set(source.headers.flatMap((header) => header.columns.map((column) => normalizeHeaderName(column.name))));
  const sourcePeriod = periodScore(source.name);
  return metadata.sheets
    .filter((sheet) => sheet.name !== sheetName)
    .map((sheet) => {
      const headerOverlap = sheet.headers
        .flatMap((header) => header.columns)
        .filter((column) => sourceHeaders.has(normalizeHeaderName(column.name))).length;
      const kindScore = sheet.kind === source.kind ? 3 : sheet.kind === "transaction" || source.kind === "transaction" ? 1 : 0;
      const period = periodScore(sheet.name);
      const periodDistance = sourcePeriod !== undefined && period !== undefined ? Math.abs(sourcePeriod - period) : 12;
      const periodBonus = periodDistance > 0 && periodDistance <= 2 ? 4 - periodDistance : 0;
      return { sheet, score: headerOverlap + kindScore + periodBonus };
    })
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score || left.sheet.name.localeCompare(right.sheet.name))
    .map((entry) => entry.sheet.name);
}

function similarRowSignals(input: AgentRunInput, values: CellMatrix): SimilarRowSignals {
  const rawValues = values.flat().filter((value) => value !== null && value !== undefined && String(value).trim().length > 0);
  const rawText = [input.request, ...rawValues.map((value) => String(value))].join(" ");
  const tokens = Array.from(new Set(tokenizeReferenceText(rawText)
    .filter((token) => !/^\d{4}$/.test(token) && !COMMON_SIMILAR_ROW_TOKENS.has(token))))
    .slice(0, 24);
  const numbers = rawValues
    .map((value) => typeof value === "number" ? value : Number(String(value).replace(/,/g, "")))
    .filter((value) => Number.isFinite(value))
    .slice(0, 12);
  const requestNumbers = Array.from(input.request.matchAll(/(?<![\w-])-?\d[\d,]*(?:\.\d+)?(?![\w-])/g))
    .map((match) => Number(match[0].replace(/,/g, "")))
    .filter((value) => Number.isFinite(value));
  const predicates = referencePredicatesFromRequest(input.request, [...numbers, ...requestNumbers]);
  return { tokens, numbers: Array.from(new Set([...numbers, ...requestNumbers])).slice(0, 16), predicates };
}

const COMMON_SIMILAR_ROW_TOKENS = new Set(["find", "similar", "rows", "row", "current", "this", "that", "with", "from", "last", "month", "sheet", "transaction", "transactions", "show", "look", "more", "into", "how", "were", "was", "label", "labeled"]);

function tokenizeReferenceText(value: string): string[] {
  return (value.match(/[\p{L}\p{M}\p{N}][\p{L}\p{M}\p{N}_-]{1,}/gu) ?? [])
    .map((token) => token.toLowerCase())
    .filter((token) => token.length > 1);
}

function referencePredicatesFromRequest(request: string, numbers: number[]): SimilarRowPredicate[] {
  const lower = request.toLowerCase();
  const predicates: SimilarRowPredicate[] = [];
  if (/\binflow\b|\bcash in\b|เงินเข้า|เติมเงิน|เพิ่มเงิน|add(?:ing)? fund|fund(?:ing)?|capital/i.test(request)) {
    predicates.push({ label: "Direction is Inflow", value: "inflow", match: "equals", headerPattern: /direction|flow|cash\s*in\s*out/i, role: "status" });
  }
  if (/\boutflow\b|\bcash out\b|จ่าย|ค่า|ซ่อม/i.test(request)) {
    predicates.push({ label: "Direction is Outflow", value: "outflow", match: "equals", headerPattern: /direction|flow|cash\s*in\s*out/i, role: "status" });
  }
  for (const match of request.matchAll(/\bX\d{3,}\b/gi)) {
    predicates.push({ label: `Transfer contains ${match[0].toUpperCase()}`, value: match[0].toLowerCase(), match: "contains", headerPattern: /transfer|from|to|payer|payee|vendor|customer|account/i });
  }
  if (/\bprach\b|yothapra/i.test(request)) {
    predicates.push({ label: "Transfer contains PRACH", value: "prach", match: "contains", headerPattern: /transfer|from|to|payer|payee|vendor|customer|account/i });
  }
  for (const match of request.matchAll(/\b\d{2,3}-\d{4}\b/g)) {
    predicates.push({ label: `Identifier contains ${match[0]}`, value: match[0].toLowerCase(), match: "contains", headerPattern: /truck|vehicle|id|job|container/i, role: "identifier" });
  }
  for (const number of numbers.filter((value) => Math.abs(value) >= 1)) {
    predicates.push({ label: `Amount equals ${number}`, value: number, match: "number_equals", headerPattern: /amount|price|variance|total|net|cash|actual|fee|tax|lifting|collect/i, role: "amount" });
  }
  return dedupeBy(predicates, (predicate) => `${predicate.match}:${String(predicate.value).toLowerCase()}:${predicate.headerPattern?.source ?? predicate.role ?? ""}`).slice(0, 10);
}

function shouldRunReferenceRowSearch(input: AgentRunInput): boolean {
  if (intentAction(input) === "find_style_references") {
    return false;
  }
  if (resolveMatchUpdateRequest(input.request)) {
    return false;
  }
  const request = input.request.toLowerCase();
  if (isDataValidationReadRequest(request)) {
    return false;
  }
  return /\b(reference|similar|look back|last month|prior|previous|before|how did we|how we|label(?:ed)?|classif(?:y|ied|ication))\b/.test(request)
    && !shouldRunStyleReferenceSearch(input);
}

function shouldRunStyleReferenceSearch(input: AgentRunInput): boolean {
  const request = input.request.toLowerCase();
  return /\b(style|styling|format|formatting|template|look like|same as|copy.*from)\b/.test(request)
    && /\b(reference|example|before|previous|last month|prior|source|template|same as|look like)\b/.test(request);
}

function shouldSearchResolvedRange(input: AgentRunInput, metadata: WorkbookMetadata, sheetName: string, sourceRange: string): boolean {
  const parsed = tryParseA1Address(stripSheetName(sourceRange));
  const rowCount = parsed ? parsed.endRow - parsed.startRow + 1 : 1;
  if (rowCount <= 3) {
    return false;
  }
  const request = normalizeComparableText(input.request);
  const sheetMentioned = request.includes(normalizeComparableText(sheetName));
  const targetSheet = input.target?.sheetName && normalizeComparableText(input.target.sheetName) === normalizeComparableText(sheetName);
  const hasRelated = likelyHistoricalSheets(metadata, sheetName).length > 0;
  return Boolean(sheetMentioned || targetSheet || !hasRelated);
}

function similarRowCandidateRanges(metadata: WorkbookMetadata, sourceSheetName: string, sourceRange: string, input: AgentRunInput): SimilarRowRange[] {
  const ranges: SimilarRowRange[] = [];
  const sourceParsed = tryParseA1Address(stripSheetName(sourceRange));
  const relatedSheets = new Set(likelyHistoricalSheets(metadata, sourceSheetName));
  const targetHints = similarRowTargetHints(input);
  const explicitSheets = explicitSimilarRowReferenceSheets(metadata, sourceSheetName, input);
  for (const table of metadata.tables) {
    if (table.sheetName === sourceSheetName && table.range === sourceRange) {
      continue;
    }
    if (explicitSheets.size > 0 && !explicitSheets.has(table.sheetName)) {
      continue;
    }
    if (table.dataRange) {
      ranges.push({
        sheetName: table.sheetName,
        range: clampRangeForSimilarRows(table.dataRange),
        reason: explicitSheets.has(table.sheetName) ? "requested reference table" : relatedSheets.has(table.sheetName) ? "related table" : "table"
      });
    }
  }
  for (const section of metadata.sections) {
    if (section.sheetName === sourceSheetName && section.range === sourceRange) {
      continue;
    }
    if (explicitSheets.size > 0 && !explicitSheets.has(section.sheetName)) {
      continue;
    }
    if (section.kind === "table-like" || section.columns.length > 0 || relatedSheets.has(section.sheetName)) {
      ranges.push({
        sheetName: section.sheetName,
        range: clampRangeForSimilarRows(section.range),
        reason: explicitSheets.has(section.sheetName) ? "requested reference section" : relatedSheets.has(section.sheetName) ? "related section" : "table-like section"
      });
    }
  }
  for (const sheet of metadata.sheets) {
    if (!sheet.usedRange || sheet.name === sourceSheetName) {
      continue;
    }
    const mentioned = targetHints.some((hint) => hint.includes(normalizeComparableText(sheet.name)));
    if (explicitSheets.size > 0 && !explicitSheets.has(sheet.name)) {
      continue;
    }
    if (relatedSheets.has(sheet.name) || mentioned || sheet.kind === "transaction") {
      const sameShape = sourceParsed ? clampRangeForSimilarRows(sheet.usedRange, sourceParsed.endColumn - sourceParsed.startColumn + 1) : clampRangeForSimilarRows(sheet.usedRange);
      ranges.push({ sheetName: sheet.name, range: sameShape, reason: mentioned ? "requested reference sheet" : relatedSheets.has(sheet.name) ? "related sheet" : "candidate sheet" });
    }
  }
  return dedupeSimilarRanges(ranges);
}

function similarRowTargetHints(input: AgentRunInput): string[] {
  return [
    input.request,
    ...(Array.isArray(input.intent?.targetHints) ? input.intent.targetHints : [])
  ].map(normalizeComparableText);
}

function explicitSimilarRowReferenceSheets(metadata: WorkbookMetadata, sourceSheetName: string, input: AgentRunInput): Set<string> {
  const hints = similarRowTargetHints(input);
  const explicitSheets = new Set<string>();
  for (const sheet of metadata.sheets) {
    const sheetName = normalizeComparableText(sheet.name);
    if (sheetName.length > 0 && hints.some((hint) => hint.includes(sheetName))) {
      explicitSheets.add(sheet.name);
    }
  }
  if (explicitSheets.size > 1 && explicitSheets.has(sourceSheetName)) {
    explicitSheets.delete(sourceSheetName);
  }
  return explicitSheets;
}

function dedupeSimilarRanges(ranges: SimilarRowRange[]): SimilarRowRange[] {
  const seen = new Set<string>();
  return ranges.filter((range) => {
    const key = `${range.sheetName}!${range.range}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function clampRangeForSimilarRows(range: string, maxColumns = 20): string {
  const parsed = tryParseA1Address(stripSheetName(range));
  if (!parsed) {
    return range;
  }
  const endRow = Math.min(parsed.endRow, parsed.startRow + 199);
  const endColumn = Math.min(parsed.endColumn, parsed.startColumn + maxColumns - 1);
  return `${numberToColumn(parsed.startColumn)}${parsed.startRow}:${numberToColumn(endColumn)}${endRow}`;
}

function rankSimilarRows(metadata: WorkbookMetadata, candidate: SimilarRowRange, values: CellMatrix, signals: SimilarRowSignals): SimilarRowMatch[] {
  const parsed = tryParseA1Address(stripSheetName(candidate.range));
  const startRow = parsed?.startRow ?? 1;
  const columns = columnsForCandidateRange(metadata, candidate.sheetName, candidate.range);
  return values
    .map((row, rowIndex) => {
      const rowText = normalizeReferenceCellText(row.map((value) => String(value ?? "")).join(" "));
      const tokenMatches = signals.tokens.filter((token) => rowText.includes(token));
      const numberMatches = signals.numbers.filter((number) => row.some((value) => Number(String(value).replace(/,/g, "")) === number));
      const predicateMatches = matchReferencePredicates(row, columns, signals.predicates);
      const matchedColumns = matchedReferenceColumns(row, columns, [...tokenMatches, ...numberMatches.map((number) => String(number))], predicateMatches);
      const score = tokenMatches.length * 3
        + numberMatches.length * 2
        + predicateMatches.length * 5
        + (candidate.reason.includes("related") ? 2 : 0)
        + importantReferenceColumnBonus(columns, row);
      const sheetRowNumber = startRow + rowIndex;
      const projectedColumns = projectReferenceColumns(row, columns, matchedColumns);
      return {
        sheetName: candidate.sheetName,
        range: parsed ? `${numberToColumn(parsed.startColumn)}${sheetRowNumber}:${numberToColumn(parsed.endColumn)}${sheetRowNumber}` : candidate.range,
        sheetRowNumber,
        rowIndex,
        values: projectedColumns.map((column) => column.value),
        columns: projectedColumns,
        score,
        matchedSignals: [...predicateMatches.map((match) => match.label), ...tokenMatches, ...numberMatches.map((number) => String(number))].slice(0, 12),
        matchedColumns,
        whyMatched: referenceMatchReason(candidate, predicateMatches, tokenMatches, numberMatches)
      };
    })
    .filter((match) => match.score > 0 && match.values.some((value) => value !== null && value !== undefined && String(value).trim().length > 0));
}

function normalizeReferenceCellText(value: string): string {
  return value.toLowerCase().normalize("NFKC");
}

function columnsForCandidateRange(metadata: WorkbookMetadata, sheetName: string, range: string): ColumnMetadata[] {
  const parsed = tryParseA1Address(stripSheetName(range));
  const startColumn = parsed?.startColumn ?? 1;
  const endColumn = parsed?.endColumn ?? startColumn + 25;
  const table = metadata.tables.find((candidate) => candidate.sheetName === sheetName && rangesOverlap(candidate.range, range));
  const section = metadata.sections.find((candidate) => candidate.sheetName === sheetName && rangesOverlap(candidate.range, range));
  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const sourceColumns = table?.columns.length ? table.columns : section?.columns.length ? section.columns : sheet?.headers[0]?.columns ?? [];
  const filtered = sourceColumns.filter((column) => column.index + 1 >= startColumn && column.index + 1 <= endColumn);
  if (filtered.length > 0) {
    return filtered.map((column) => ({ ...column, index: column.index - (startColumn - 1) }));
  }
  return Array.from({ length: Math.max(0, endColumn - startColumn + 1) }, (_unused, index) => ({
    index,
    letter: numberToColumn(startColumn + index),
    name: numberToColumn(startColumn + index),
    normalizedName: normalizeHeaderName(numberToColumn(startColumn + index)),
    inferredType: "unknown" as const,
    role: "unknown" as const,
    importance: 0.2
  }));
}

function rangesOverlap(left: string, right: string): boolean {
  try {
    return rangesOverlapAddresses(stripSheetName(left), stripSheetName(right));
  } catch {
    return false;
  }
}

function matchReferencePredicates(row: unknown[], columns: ColumnMetadata[], predicates: SimilarRowPredicate[]): SimilarRowPredicate[] {
  return predicates.filter((predicate) => columns.some((column) => {
    const value = row[column.index];
    if (value === undefined || value === null || value === "") {
      return false;
    }
    const headerMatches = predicate.headerPattern?.test(column.name) || predicate.headerPattern?.test(column.normalizedName) || (predicate.role !== undefined && column.role === predicate.role);
    if (!headerMatches) {
      return false;
    }
    if (predicate.match === "number_equals") {
      return Number(String(value).replace(/,/g, "")) === predicate.value;
    }
    const comparable = normalizeReferenceCellText(String(value));
    const expected = normalizeReferenceCellText(String(predicate.value));
    return predicate.match === "equals" ? comparable === expected : comparable.includes(expected);
  }));
}

function matchedReferenceColumns(
  row: unknown[],
  columns: ColumnMetadata[],
  signals: string[],
  predicateMatches: SimilarRowPredicate[]
): SimilarRowMatch["matchedColumns"] {
  const matched = columns.flatMap((column) => {
    const value = row[column.index];
    if (value === undefined || value === null || value === "") {
      return [];
    }
    const text = normalizeReferenceCellText(String(value));
    const valueSignals = signals.filter((signal) => text.includes(normalizeReferenceCellText(signal)));
    const predicateSignals = predicateMatches
      .filter((predicate) => predicate.headerPattern?.test(column.name) || predicate.headerPattern?.test(column.normalizedName) || (predicate.role !== undefined && column.role === predicate.role))
      .map((predicate) => predicate.label);
    const allSignals = [...valueSignals, ...predicateSignals];
    return allSignals.length > 0 ? [{ letter: column.letter, name: column.name, value, signals: allSignals.slice(0, 5) }] : [];
  });
  return dedupeBy(matched, (entry) => `${entry.letter}:${entry.name}`).slice(0, 8);
}

function projectReferenceColumns(
  row: unknown[],
  columns: ColumnMetadata[],
  matchedColumns: SimilarRowMatch["matchedColumns"]
): SimilarRowMatch["columns"] {
  const matchedLetters = new Set(matchedColumns.map((column) => column.letter));
  const preferred = columns
    .filter((column) => matchedLetters.has(column.letter) || isImportantReferenceColumn(column))
    .sort((left, right) => Number(matchedLetters.has(right.letter)) - Number(matchedLetters.has(left.letter)) || (right.importance ?? roleImportance(right.role ?? "unknown")) - (left.importance ?? roleImportance(left.role ?? "unknown")) || left.index - right.index)
    .slice(0, 12)
    .sort((left, right) => left.index - right.index);
  const selected = preferred.length > 0 ? preferred : columns.slice(0, 12);
  return selected.map((column) => stripUndefinedRecord({
    letter: column.letter,
    name: column.name,
    ...(column.role !== undefined ? { role: column.role } : {}),
    value: row[column.index]
  }) as SimilarRowMatch["columns"][number]);
}

function isImportantReferenceColumn(column: ColumnMetadata): boolean {
  const name = `${column.name} ${column.normalizedName}`;
  return ["date", "description", "vendor", "account", "amount", "status", "category", "identifier"].includes(column.role ?? "")
    || /date|description|type|direction|amount|actual|variance|transfer|from|to|truck|job|vendor|customer|account|note/i.test(name);
}

function importantReferenceColumnBonus(columns: ColumnMetadata[], row: unknown[]): number {
  return columns.some((column) => isImportantReferenceColumn(column) && row[column.index] !== undefined && row[column.index] !== null && row[column.index] !== "") ? 1 : 0;
}

function referenceMatchReason(candidate: SimilarRowRange, predicates: SimilarRowPredicate[], tokens: string[], numbers: number[]): string {
  const reasons = [
    candidate.reason,
    ...predicates.map((predicate) => predicate.label),
    ...(tokens.length > 0 ? [`text matched: ${tokens.slice(0, 4).join(", ")}`] : []),
    ...(numbers.length > 0 ? [`number matched: ${numbers.slice(0, 4).join(", ")}`] : [])
  ];
  return reasons.slice(0, 4).join("; ");
}

function styleReferenceCandidates(metadata: WorkbookMetadata, input: AgentRunInput): StyleReferenceCandidate[] {
  const request = normalizeComparableText(input.request);
  const targetSheet = input.target?.sheetName;
  const related = targetSheet ? new Set(likelyHistoricalSheets(metadata, targetSheet)) : new Set<string>();
  const requestedSheets = new Set(metadata.sheets
    .filter((sheet) => request.includes(normalizeComparableText(sheet.name)))
    .map((sheet) => sheet.name));
  const candidates: StyleReferenceCandidate[] = [
    ...metadata.sections.flatMap((section) => {
      if (!section.range || section.kind === "notes") {
        return [];
      }
      const score = (related.has(section.sheetName) ? 0.3 : 0)
        + (requestedSheets.has(section.sheetName) ? 0.3 : 0)
        + (section.kind === "table-like" ? 0.2 : 0)
        + Math.min(0.2, section.confidence / 5);
      return [{
        sheetName: section.sheetName,
        range: section.headerRange ?? section.range,
        label: section.label,
        sourceKind: "section" as const,
        confidence: Math.min(0.98, 0.45 + score),
        reason: `${section.kind} section${related.has(section.sheetName) ? " on related sheet" : ""}`,
        nextAction: "Use copy_style_from_template or read_style_summary on this source range before applying style."
      }];
    }),
    ...metadata.tables.map((table) => ({
      sheetName: table.sheetName,
      range: table.headerRange ?? table.range,
      label: table.name ?? table.id,
      sourceKind: "table" as const,
      confidence: Math.min(0.96, 0.62 + (related.has(table.sheetName) ? 0.2 : 0) + (requestedSheets.has(table.sheetName) ? 0.12 : 0)),
      reason: related.has(table.sheetName) ? "table on related sheet" : "workbook table",
      nextAction: "Use copy_style_from_template for matching table/header/body dimensions."
    })),
    ...metadata.sheets.flatMap<StyleReferenceCandidate>((sheet) => {
      if (!sheet.usedRange) return [];
      const headers = sheet.headers.flatMap((header) => ({
        sheetName: sheet.name,
        range: header.range,
        label: `${sheet.name} header row`,
        sourceKind: "header" as const,
        confidence: Math.min(0.93, 0.58 + (related.has(sheet.name) ? 0.18 : 0) + (requestedSheets.has(sheet.name) ? 0.12 : 0)),
        reason: "header style candidate",
        nextAction: "Use this as a source for header formatting only."
      }));
      return headers.length > 0 ? headers : [{
        sheetName: sheet.name,
        range: sheet.usedRange,
        label: `${sheet.name} used range`,
        sourceKind: "sheet" as const,
        confidence: Math.min(0.82, 0.38 + (related.has(sheet.name) ? 0.2 : 0) + (requestedSheets.has(sheet.name) ? 0.16 : 0)),
        reason: related.has(sheet.name) ? "related sheet style candidate" : "sheet style candidate",
        nextAction: "Read style summary first; avoid copying whole-sheet styles unless user asks broadly."
      }];
    })
  ];
  return dedupeBy(candidates, (candidate) => `${candidate.sheetName}!${candidate.range}`)
    .filter((candidate) => candidate.confidence >= 0.5 || requestedSheets.has(candidate.sheetName))
    .sort((left, right) => right.confidence - left.confidence || left.sheetName.localeCompare(right.sheetName));
}

function compactStyleReferenceFingerprint(result: unknown): unknown {
  const fingerprint = result && typeof result === "object" ? (result as { fingerprint?: unknown }).fingerprint : undefined;
  const typed = fingerprint && typeof fingerprint === "object" ? fingerprint as Record<string, unknown> : undefined;
  const dimensions = typed?.dimensions && typeof typed.dimensions === "object" ? typed.dimensions as Record<string, unknown> : undefined;
  return stripUndefinedRecord({
    address: typed?.address,
    fills: compactStyleDimension(dimensions?.fills),
    fonts: compactStyleDimension(dimensions?.fonts),
    borders: compactStyleDimension(dimensions?.borders),
    numberFormats: compactStyleDimension(dimensions?.numberFormats),
    dataValidation: compactStyleDimension(dimensions?.dataValidation)
  });
}

function periodScore(value: string): number | undefined {
  const monthMatch = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{4})\b/i.exec(value);
  if (!monthMatch?.[1] || !monthMatch[2]) {
    return undefined;
  }
  const month = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(monthMatch[1].toLowerCase().slice(0, 3));
  const year = Number(monthMatch[2]);
  return Number.isFinite(year) && month >= 0 ? year * 12 + month : undefined;
}

function resolveAgentTable(metadata: WorkbookMetadata, input: AgentRunInput): TableMetadata | undefined {
  if (input.target?.tableName) {
    const normalized = normalizeAgentLookup(input.target.tableName);
    const exact = metadata.tables.find((table) => normalizeAgentLookup(table.name ?? table.id) === normalized);
    if (exact) {
      return exact;
    }
  }
  if (input.target?.sheetName) {
    const normalizedSheet = normalizeComparableText(input.target.sheetName);
    const exactRange = input.target.range;
    return metadata.tables.find((table) => normalizeComparableText(table.sheetName) === normalizedSheet && (!exactRange || table.range === exactRange))
      ?? metadata.tables.find((table) => normalizeComparableText(table.sheetName) === normalizedSheet);
  }
  const request = normalizeComparableText(input.request);
  return metadata.tables.find((table) => request.includes(normalizeComparableText(table.name ?? "")) || request.includes(normalizeComparableText(table.sheetName)));
}

function tableFromResolution(metadata: WorkbookMetadata, resolved: Extract<AgentTargetResolution, { ok: true }>): TableMetadata | undefined {
  if (resolved.candidate.tableName) {
    const normalized = normalizeAgentLookup(resolved.candidate.tableName);
    const exact = metadata.tables.find((table) => normalizeAgentLookup(table.name ?? table.id) === normalized);
    if (exact) {
      return exact;
    }
  }
  return metadata.tables.find((table) => table.id === resolved.candidate.id)
    ?? metadata.tables.find((table) => table.sheetName === resolved.sheetName && table.range === resolved.range);
}

function explicitlyRequestedTable(metadata: WorkbookMetadata, input: AgentRunInput): TableMetadata | undefined {
  if (!/\btables?\b/i.test(input.request) && !input.target?.tableName) {
    return undefined;
  }
  const requestedName = input.target?.tableName ? normalizeComparableText(input.target.tableName) : undefined;
  const request = normalizeComparableText(input.request);
  return metadata.tables.find((table) => {
    const name = normalizeComparableText(table.name ?? table.id);
    return name !== "" && (requestedName === name || request.includes(name));
  });
}

function requestExplicitlyNamesTable(input: AgentRunInput, table: TableMetadata): boolean {
  if (input.target?.tableName) {
    return normalizeComparableText(input.target.tableName) === normalizeComparableText(table.name ?? table.id);
  }
  if (!/\btables?\b/i.test(input.request)) {
    return false;
  }
  const name = normalizeComparableText(table.name ?? table.id);
  return name !== "" && normalizeComparableText(input.request).includes(name);
}

function tableRowUpdatesFromInput(input: AgentRunInput): TableUpdateRowsRequest["rows"] {
  const raw = input.values?.rows;
  if (!Array.isArray(raw)) {
    return [];
  }
  return raw.flatMap((row, index) => {
    if (row && typeof row === "object" && !Array.isArray(row)) {
      const candidate = row as { index?: unknown; values?: unknown };
      if (typeof candidate.index === "number" && Array.isArray(candidate.values)) {
        return [{ index: candidate.index, values: candidate.values as CellMatrix[number] }];
      }
      return [];
    }
    return Array.isArray(row) ? [{ index, values: row as CellMatrix[number] }] : [];
  });
}

function tableColumnOrderFromInput(input: AgentRunInput): TableReorderColumnsRequest["columnOrder"] {
  const raw = input.values?.columnOrder ?? input.values?.columns;
  return Array.isArray(raw)
    ? raw.filter((column): column is string | number => typeof column === "string" || typeof column === "number")
    : [];
}

function tableFiltersFromInput(input: AgentRunInput): { ok: true; filters: TableApplyFiltersRequest["filters"] } | { ok: false; summary: string } {
  const raw = input.values?.filters;
  if (!Array.isArray(raw) || raw.length === 0) {
    return { ok: false, summary: "Table filter previews need values.filters, for example [{ column: \"Status\", value: \"Open\" }]." };
  }
  const filters = raw.flatMap((filter) => {
    const normalized = normalizeTableFilterSpec(filter);
    return normalized ? [normalized] : [];
  });
  if (filters.length !== raw.length) {
    return { ok: false, summary: "Table filter previews need each filter to include column plus criteria, criterion, or value." };
  }
  return { ok: true, filters };
}

function isClearFilterRequest(request: string): boolean {
  return /\b(clear|remove|reset)\b/i.test(request) && /\b(filters?|autofilter|auto\s*filter)\b/i.test(request);
}

function normalizeTableFilterSpec(filter: unknown): TableApplyFiltersRequest["filters"][number] | undefined {
  if (!filter || typeof filter !== "object" || Array.isArray(filter)) {
    return undefined;
  }
  const candidate = filter as Record<string, unknown>;
  const column = typeof candidate.column === "string" || typeof candidate.column === "number" ? candidate.column : undefined;
  if (column === undefined) {
    return undefined;
  }
  if (candidate.criteria && typeof candidate.criteria === "object") {
    return { column, criteria: candidate.criteria };
  }
  const raw = Object.prototype.hasOwnProperty.call(candidate, "criterion")
    ? candidate.criterion
    : Object.prototype.hasOwnProperty.call(candidate, "value")
      ? candidate.value
      : undefined;
  if (raw === undefined || raw === null) {
    return undefined;
  }
  const values = Array.isArray(raw) ? raw : [raw];
  if (values.length === 0) {
    return undefined;
  }
  return { column, criteria: { filterOn: "Values", values } };
}

function tableSortSpecFromInput(input: AgentRunInput, table: TableMetadata): TableSortRequest["fields"][number] | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const rawColumn = values?.sortBy ?? values?.column ?? values?.key;
  const key = tableSortKey(table, rawColumn);
  if (key === undefined) {
    return undefined;
  }
  return {
    key,
    ascending: sortAscendingFromInput(values, input.request),
    ...(isTableSortOn(values?.sortOn) ? { sortOn: values.sortOn } : {}),
    ...(typeof values?.color === "string" ? { color: values.color } : {}),
    ...(isTableSortDataOption(values?.dataOption) ? { dataOption: values.dataOption } : {})
  };
}

function tableApplyViewSortFromInput(input: AgentRunInput, table: TableMetadata | undefined): NonNullable<TableApplyViewRequest["sort"]> {
  const values = input.values as Record<string, unknown> | undefined;
  const rawSort = values?.sort && typeof values.sort === "object" && !Array.isArray(values.sort) ? values.sort as Record<string, unknown> : undefined;
  const rawFields = Array.isArray(rawSort?.fields)
    ? rawSort.fields
    : Array.isArray(values?.fields)
      ? values.fields
      : undefined;
  const explicitFields = rawFields?.flatMap((field) => normalizeTableSortFieldSpec(field, table)) ?? [];
  const method = tableSortMethodFromInput(rawSort?.method ?? values?.method);
  if (explicitFields.length > 0) {
    return {
      fields: explicitFields,
      ...(typeof rawSort?.matchCase === "boolean" ? { matchCase: rawSort.matchCase } : typeof values?.matchCase === "boolean" ? { matchCase: values.matchCase } : {}),
      ...(method ? { method } : {})
    };
  }
  const inferred = table ? tableSortSpecFromInput(input, table) : undefined;
  return inferred ? { fields: [inferred] } : { fields: [] };
}

function normalizeTableSortFieldSpec(field: unknown, table: TableMetadata | undefined): TableSortField[] {
  if (!field || typeof field !== "object" || Array.isArray(field)) {
    return [];
  }
  const candidate = field as Record<string, unknown>;
  const rawKey = candidate.key ?? candidate.column ?? candidate.name;
  const key = typeof rawKey === "number" && Number.isInteger(rawKey) && rawKey >= 0
    ? rawKey
    : table ? tableSortKey(table, rawKey) : undefined;
  if (key === undefined) {
    return [];
  }
  return [{
    key,
    ...(typeof candidate.ascending === "boolean" ? { ascending: candidate.ascending } : {}),
    ...(isTableSortOn(candidate.sortOn) ? { sortOn: candidate.sortOn } : {}),
    ...(typeof candidate.color === "string" ? { color: candidate.color } : {}),
    ...(isTableSortDataOption(candidate.dataOption) ? { dataOption: candidate.dataOption } : {})
  }];
}

function tableSortMethodFromInput(value: unknown): NonNullable<TableApplyViewRequest["sort"]>["method"] | undefined {
  return value === "PinYin" || value === "StrokeCount" ? value : undefined;
}

function tableSortKey(table: TableMetadata, rawColumn: unknown): number | undefined {
  if (typeof rawColumn === "number" && Number.isInteger(rawColumn) && rawColumn >= 0) {
    return rawColumn;
  }
  const column = stringValue(rawColumn);
  if (!column) {
    return undefined;
  }
  const normalized = normalizeAgentLookup(column);
  return table.columns.find((candidate) => normalizeAgentLookup(candidate.name) === normalized)?.index
    ?? table.columns.find((candidate) => normalizeComparableText(candidate.name).includes(normalizeComparableText(column)))?.index;
}

function sortAscendingFromInput(values: Record<string, unknown> | undefined, request: string): boolean {
  if (typeof values?.ascending === "boolean") {
    return values.ascending;
  }
  const direction = stringValue(values?.direction ?? values?.order);
  if (direction && /\b(desc|descending|high|highest|largest|z-a)\b/i.test(direction)) {
    return false;
  }
  if (direction && /\b(asc|ascending|low|lowest|smallest|a-z)\b/i.test(direction)) {
    return true;
  }
  return !/\b(highest|descending|desc|largest|lowest to highest)\b/i.test(request);
}

function isTableSortOn(value: unknown): value is NonNullable<TableSortRequest["fields"][number]["sortOn"]> {
  return value === "Value" || value === "CellColor" || value === "FontColor" || value === "Icon";
}

function isTableSortDataOption(value: unknown): value is NonNullable<TableSortRequest["fields"][number]["dataOption"]> {
  return value === "Normal" || value === "TextAsNumber";
}

function tableReadColumnsFromInput(input: AgentRunInput): Array<string | number> {
  const raw = input.values?.columns ?? input.values?.projectedColumns;
  return Array.isArray(raw)
    ? raw.filter((column): column is string | number => typeof column === "string" || typeof column === "number")
    : [];
}

function tableReadRowOffset(input: AgentRunInput): number {
  const values = input.values as Record<string, unknown> | undefined;
  const explicit = nonNegativeInteger(values?.rowOffset);
  if (explicit !== undefined) {
    return explicit;
  }
  const rowStart = positiveIntegerValue(values?.rowStart);
  if (rowStart !== undefined && rowStart > 0) {
    return rowStart - 1;
  }
  const requested = tableRowWindowFromRequest(input.request);
  return requested?.rowStart !== undefined && requested.rowStart > 0 ? requested.rowStart - 1 : 0;
}

function compactTableRowLimit(input: AgentRunInput, columnCount: number, rowOffset = 0): number {
  const values = input.values as Record<string, unknown> | undefined;
  const rowEnd = positiveIntegerValue(values?.rowEnd);
  if (rowEnd !== undefined && rowEnd > rowOffset) {
    return rowEnd - rowOffset;
  }
  const requestedWindow = tableRowWindowFromRequest(input.request);
  if (requestedWindow?.rowEnd !== undefined && requestedWindow.rowEnd > rowOffset) {
    return requestedWindow.rowEnd - rowOffset;
  }
  const requested = nonNegativeInteger(values?.rowLimit ?? values?.maxRows ?? input.budget?.maxExamples);
  if (requested !== undefined && requested > 0) {
    return requested;
  }
  const firstRows = /\bfirst\s+(\d{1,5})\s+(?:table\s+)?rows?\b/i.exec(input.request);
  if (firstRows?.[1]) {
    return Math.min(10_000, Math.max(1, Number(firstRows[1])));
  }
  const maxCells = input.budget?.maxPayloadBytes ? Math.max(25, Math.floor(input.budget.maxPayloadBytes / 80)) : undefined;
  if (maxCells !== undefined && columnCount > 0) {
    return Math.max(1, Math.min(50, Math.floor(maxCells / columnCount)));
  }
  if (input.detailLevel === "table_sample") {
    return 20;
  }
  if (input.detailLevel === "full_table") {
    return 10_000;
  }
  return 50;
}

function tableRowWindowFromRequest(request: string): { rowStart: number; rowEnd: number } | undefined {
  const match = /\brows?\s+(\d{1,5})\s*(?:-|to|through|thru)\s*(\d{1,5})\b/i.exec(request);
  if (!match?.[1] || !match[2]) {
    return undefined;
  }
  const rowStart = Math.max(1, Number(match[1]));
  const rowEnd = Math.max(rowStart, Number(match[2]));
  return { rowStart, rowEnd };
}

function isReadOnlyInspectionRequest(request: string): boolean {
  return /\b(?:read|show|inspect|look(?:\s+at|\s+into)?|review|summari[sz]e|analy[sz]e|list|print|describe)\b/i.test(request)
    && !/\b(?:add|write|update|change|set|replace|append|insert|delete|clear|format|style|apply|create|modify|fix|remove|sort|filter)\b/i.test(request);
}

function hasStructuredMutationPayload(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  if (!values) {
    return false;
  }
  return ["values", "rows", "patches", "formulas", "style", "validation", "rule", "entries", "options", "totalRow", "showTotals"].some((key) => values[key] !== undefined);
}

function shouldRouteStructuredWritePayloadToPreview(input: AgentRunInput): boolean {
  if (!hasStructuredMutationPayload(input)) {
    return false;
  }
  if (input.target?.range || input.target?.tableName) {
    return true;
  }
  const patches = input.values?.patches;
  return Array.isArray(patches) && patches.some((patch) => Boolean(patch?.target?.range || patch?.target?.tableName));
}

function tableNeedsInput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  resolved: Extract<AgentTargetResolution, { ok: true }>,
  summary: string
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: resolved.candidate.label }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: []
  };
}

function normalizeComparableText(value: string): string {
  return value.trim().toLowerCase().normalize("NFKC").replace(/\s+/g, " ").replace(/[^\p{L}\p{M}\p{N}_ ]/gu, "");
}

function dedupeBy<T>(values: T[], keyFor: (value: T) => string): T[] {
  const seen = new Set<string>();
  return values.filter((value) => {
    const key = keyFor(value);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function applyOutputBudget(output: Omit<AgentRunOutput, "telemetry">, input: AgentRunInput, results: AgentResultStore, metadataCache: WorkbookMetadataCache): Omit<AgentRunOutput, "telemetry"> {
  const responseMode = responseModeFromInput(input);
  const defaults = defaultResponseBudget(responseMode);
  const maxExamples = Math.max(0, input.budget?.maxExamples ?? defaults.maxExamples);
  const maxPayloadBytes = input.budget?.maxPayloadBytes ?? defaults.maxPayloadBytes;
  const maxEstimatedTokens = input.budget?.maxEstimatedTokens ?? defaults.maxEstimatedTokens;
  const stored = responseMode === "verbose" || !answerNeedsResultResource(output.answer)
    ? undefined
    : results.create({
      workbookContextId: output.workbookContextId === undefined ? undefined : String(output.workbookContextId),
      freshness: resultFreshnessForOutput(output, metadataCache),
      summary: output.summary,
      answer: output.answer
    });
  const resourceLinks = stored ? appendUniqueResource(output.resourceLinks, resultResource(stored.resultId)) : output.resourceLinks;
  const answer = responseMode === "verbose" ? output.answer : compactAnswerForResponseMode(output.answer, responseMode, input, stored?.resourceUri, stored?.fullResourceUri, output.proof);
  const continuation = continuationForOutput(output, responseMode, stored);
  const budgeted = stripUndefinedOptionals({
    ...output,
    ...(answer !== undefined ? { answer } : { answer: undefined }),
    resourceLinks,
    ...(continuation !== undefined ? { continuation } : {}),
    proof: output.proof.slice(0, Math.min(output.proof.length, Math.max(1, maxExamples))),
    warnings: output.warnings.slice(0, Math.max(5, maxExamples)),
    ...(output.candidates ? { candidates: output.candidates.slice(0, maxExamples).map((candidate) => responseMode === "verbose" ? candidate : compactCandidate(candidate)) } : {}),
    ...(output.changes ? { changes: output.changes.slice(0, maxExamples) } : {})
  });
  if (budgeted.answer && typeof budgeted.answer === "object" && "sample" in budgeted.answer && Array.isArray((budgeted.answer as { sample?: unknown[] }).sample)) {
    budgeted.answer = {
      ...budgeted.answer,
      sample: (budgeted.answer as { sample: unknown[] }).sample.slice(0, Math.max(1, Math.min(3, maxExamples)))
    };
  }
  let compact = stripUndefinedOptionals(budgeted);
  const byteBudget = maxPayloadBytes ?? (maxEstimatedTokens ? maxEstimatedTokens * 4 : undefined);
  if (byteBudget && Buffer.byteLength(JSON.stringify(compact)) > byteBudget) {
    const exactReadCompact = compactExactReadForBudget(compact, byteBudget);
    if (exactReadCompact) {
      return exactReadCompact;
    }
    const compactContinuation = compactContinuationForBudget(compact.continuation);
    const { continuation: _continuation, ...compactWithoutContinuation } = compact;
    compact = stripUndefinedOptionals({
      ...compactWithoutContinuation,
      ...(compact.answer ? { answer: minimalAnswerForBudget(compact.answer, compact.continuation, compact.workbookContextId) } : {}),
      ...(compact.candidates ? { candidates: compact.candidates.slice(0, Math.min(compact.candidates.length, 3)).map(compactCandidate) } : {}),
      proof: compact.proof.slice(0, Math.min(compact.proof.length, 3)),
      ...(compact.changes ? { changes: compact.changes.slice(0, Math.min(compact.changes.length, 3)) } : {}),
      ...(compactContinuation !== undefined ? { continuation: compactContinuation } : {}),
      warnings: [...compact.warnings, "Agent response was compacted to satisfy the requested payload/token budget."]
    });
  }
  if (byteBudget && Buffer.byteLength(JSON.stringify(compact)) > byteBudget && compact.answer !== undefined) {
    const compactContinuation = compactContinuationForBudget(compact.continuation, true);
    const { continuation: _continuation, ...compactWithoutContinuation } = compact;
    compact = stripUndefinedOptionals({
      ...compactWithoutContinuation,
      answer: minimalAnswerForBudget(compact.answer, compact.continuation, compact.workbookContextId),
      proof: compact.proof.slice(0, 1),
      ...(compact.candidates ? { candidates: compact.candidates.slice(0, 1).map(compactCandidate) } : {}),
      ...(compact.finalAnswer ? { finalAnswer: truncateForBudget(compact.finalAnswer, 360) } : {}),
      ...(compactContinuation !== undefined ? { continuation: compactContinuation } : {}),
      warnings: [...compact.warnings, "Answer details were reduced to a minimal inline summary; use resourceLinks as MCP/Open Workbook handles, not HTTP URLs."]
    });
  }
  return stripUndefinedOptionals(compact);
}

function enforceFullDataFollowupAfterBudget(output: Omit<AgentRunOutput, "telemetry">, input: AgentRunInput): Omit<AgentRunOutput, "telemetry"> {
  if (!isExplicitFullDataRequest(input.request)) {
    return output;
  }
  const answer = output.answer && typeof output.answer === "object" ? output.answer as Record<string, unknown> : undefined;
  if (!answer || (answer.kind !== "table_compact_read" && answer.kind !== "range_profile")) {
    return output;
  }
  const actualInlineRows = countMatrixRows(answer.rows) ?? countMatrixRows(answer.values) ?? countMatrixRows(answer.valuesPreview) ?? countMatrixRows(answer.sparseRows) ?? 0;
  const totalRowCount = typeof answer.totalRowCount === "number" ? answer.totalRowCount : undefined;
  const actuallyComplete = totalRowCount !== undefined
    ? actualInlineRows >= totalRowCount && answer.previewTruncated !== true
    : answer.inlineIsComplete === true;
  if (actuallyComplete) {
    return output;
  }
  const fullResultUri = stringValue(answer.fullResultUri ?? output.continuation?.fullResultUri);
  const resultUri = stringValue(answer.resultUri ?? output.continuation?.resultUri);
  const inlineRowCount = actualInlineRows;
  const target = [answer.sheetName, answer.tableName ?? answer.range].filter(Boolean).join(" ");
  const summary = fullResultUri
    ? `Full data was requested, but only ${inlineRowCount ?? 0}${totalRowCount !== undefined ? ` of ${totalRowCount}` : ""} row(s) are inline for ${target || "the requested result"}. Fetch the stored full result once.`
    : `Full data was requested, but the result is not complete inline. Ask the user for a narrower range, columns, filters, or export-style workflow.`;
  const { taskOutcome: _taskOutcome, ...outputWithoutTaskOutcome } = output;
  return stripUndefinedOptionals({
    ...outputWithoutTaskOutcome,
    summary,
    nextAction: fullResultUri ? "fetch_resource" : "ask_user",
    finalAnswer: summary,
    agentInstruction: fullResultUri
      ? "Call excel.agent.run once with continuation.fullResultUri to fetch the stored full result; do not answer as if the inline preview is complete and do not broad-read another chunk."
      : "Ask the user to narrow the full-data request by row range, columns, filters, or export workflow; do not broad-read more chunks.",
    maxRecommendedFollowupCalls: fullResultUri ? 1 : 0,
    ...(fullResultUri ? {
      requiredFollowup: {
        mode: "answer" as const,
        nextAction: "fetch_resource" as const,
        instruction: "Call excel.agent.run once with continuation.fullResultUri or the returned excel:// handle. Never use Webfetch/browser for excel:// handles."
      },
      continuation: {
        ...(output.continuation ?? {}),
        ...(resultUri ? { resultUri } : {}),
        fullResultUri
      }
    } : {}),
    warnings: [
      ...output.warnings,
      fullResultUri
        ? "Inline rows are a preview only; the user asked for all data, so fetch the stored full result before presenting all rows."
        : "Inline rows are incomplete for a full-data request; ask the user to narrow the request."
    ]
  });
}

function compactExactReadForBudget(output: Omit<AgentRunOutput, "telemetry">, byteBudget: number): Omit<AgentRunOutput, "telemetry"> | undefined {
  const answer = output.answer;
  if (!answer || typeof answer !== "object" || Array.isArray(answer)) {
    return undefined;
  }
  const typed = answer as Record<string, unknown>;
  const encodedValues = Array.isArray(typed.encodedValues) ? typed.encodedValues : undefined;
  const rows = Array.isArray(typed.rows) ? typed.rows : undefined;
  const valuesPreview = Array.isArray(typed.valuesPreview) ? typed.valuesPreview : undefined;
  if (!encodedValues && !rows && !valuesPreview) {
    return undefined;
  }
  const exactRows = encodedValues ?? rows ?? valuesPreview;
  const exactCellCount = matrixCellCount(exactRows as CellMatrix);
  if (exactCellCount <= 0 || exactCellCount > 300) {
    return undefined;
  }
  const slimAnswer = stripUndefinedRecord({
    kind: typed.kind,
    source: typed.source,
    sheetName: typed.sheetName,
    range: typed.range,
    shape: typed.shape,
    metrics: typed.metrics,
    inlineRowCount: typed.inlineRowCount,
    totalRowCount: typed.totalRowCount,
    inlineIsComplete: typed.inlineIsComplete,
    missingInlineReason: typed.missingInlineReason,
    ...(encodedValues ? { encodedValues } : rows ? { rows } : {}),
    ...(encodedValues ? { valueEncoding: typed.valueEncoding } : {}),
    ...(valuesPreview && !encodedValues ? { valuesPreview } : {}),
    previewRange: typed.previewRange,
    previewTruncated: typed.previewTruncated,
    resultUri: typed.resultUri,
    fullResultUri: typed.fullResultUri,
    inlineRowsReason: "narrow_exact_read"
  });
  const { candidates: _candidates, changes: _changes, ...outputWithoutHeavyLists } = output;
  const compact = stripUndefinedOptionals({
    ...outputWithoutHeavyLists,
    answer: slimAnswer,
    proof: output.proof.slice(0, 2),
    resourceLinks: output.resourceLinks.slice(0, 2),
    warnings: output.warnings.filter((warning) => !/compacted|fullResultUri/i.test(warning)).slice(0, 3)
  });
  if (Buffer.byteLength(JSON.stringify(compact)) <= byteBudget) {
    return compact;
  }
  const tighter = stripUndefinedOptionals({
    ...compact,
    answer: stripUndefinedRecord({
      kind: slimAnswer.kind,
      source: slimAnswer.source,
      sheetName: slimAnswer.sheetName,
      range: slimAnswer.range,
      inlineRowCount: slimAnswer.inlineRowCount,
      totalRowCount: slimAnswer.totalRowCount,
      inlineIsComplete: slimAnswer.inlineIsComplete,
      missingInlineReason: slimAnswer.missingInlineReason,
      ...(encodedValues ? { encodedValues } : rows ? { rows } : {}),
      ...(encodedValues ? { valueEncoding: slimAnswer.valueEncoding } : {}),
      ...(valuesPreview && !encodedValues && !rows ? { valuesPreview } : {}),
      previewRange: slimAnswer.previewRange,
      previewTruncated: slimAnswer.previewTruncated,
      resultUri: slimAnswer.resultUri,
      fullResultUri: slimAnswer.fullResultUri,
      inlineRowsReason: "narrow_exact_read"
    }),
    proof: compact.proof.slice(0, 1),
    resourceLinks: compact.resourceLinks.slice(0, 1)
  });
  return Buffer.byteLength(JSON.stringify(tighter)) <= byteBudget ? tighter : undefined;
}

function truncateForBudget(value: string, maxLength: number): string {
  const normalized = value.replace(/\s+/g, " ").trim();
  return normalized.length <= maxLength ? normalized : `${normalized.slice(0, Math.max(0, maxLength - 1))}…`;
}

function countMatrixRows(value: unknown): number | undefined {
  return Array.isArray(value) ? value.length : undefined;
}

function resultFreshnessForOutput(output: Omit<AgentRunOutput, "telemetry">, metadataCache: WorkbookMetadataCache): AgentResultFreshness | undefined {
  const workbookContextId = output.workbookContextId === undefined ? undefined : String(output.workbookContextId);
  if (!workbookContextId) {
    return undefined;
  }
  const metadata = metadataCache.getByContextId(workbookContextId);
  if (!metadata) {
    return undefined;
  }
  return stripUndefinedRecord({
    workbookId: metadata.workbook.workbookId,
    workbookContentVersion: metadata.contentVersion,
    workbookStructureHash: metadata.fingerprint.structureHash,
    contextUpdatedAt: metadata.updatedAt
  }) as AgentResultFreshness;
}

function buildContextUsed(
  input: AgentRunInput,
  output: Omit<AgentRunOutput, "telemetry">,
  runMetrics: AgentRunMetrics,
  cacheHit: boolean,
  payloadBytes: number,
  metadataCache: WorkbookMetadataCache
): AgentRunOutput["contextUsed"] {
  const decision = runMetrics.route.contextDecision;
  const answer = output.answer && typeof output.answer === "object" ? output.answer as Record<string, unknown> : undefined;
  const stages = new Set<string>();
  const included = new Set<string>(decision.include);
  if (output.workbookContextId) {
    stages.add("metadata");
  }
  const kind = stringValue(answer?.kind);
  if (kind === "workbook_summary" || kind === "workbook_overview" || kind === "sheet_summary" || kind === "semantic_workbook_index") {
    stages.add("schema");
    included.add("schema");
  }
  if (kind === "table_compact_read" || kind === "range_profile") {
    stages.add("schema");
    stages.add("field_context");
    stages.add(proofImpliesSelectedData(output.proof) ? "focused_values" : "sample_values");
    included.add("schema");
    included.add("field_context");
    included.add("values");
  }
  if (kind === "data_validation_summary" || kind === "patch_validation_failed") {
    stages.add("field_context");
    stages.add("audit_facets");
    included.add("field_context");
    included.add("validation");
  }
  if (output.mode === "preview_update" || output.taskOutcome === "preview_ready") {
    stages.add("target_resolution");
    stages.add("preview_proof");
    included.add("target");
    included.add("preview");
  }
  if (Array.isArray(answer?.fieldContext)) {
    stages.add("field_context");
    included.add("field_context");
  }
  if (Array.isArray((answer as Record<string, unknown> | undefined)?.validationChecks)) {
    stages.add("field_context");
    stages.add("audit_facets");
    included.add("field_context");
    included.add("validation");
  }
  const rangesRead = uniqueContextRanges(output.proof);
  const rowsRead = contextRowsRead(output.proof);
  const truncated = answerTruncated(answer) || output.continuation?.fullResultUri !== undefined;
  const strategy = output.mode === "preview_update" || output.taskOutcome === "preview_ready" ? "mutation" as const : decision.strategy;
  const suggestedNext = contextSuggestedNext(output, input);
  const requiredCacheFacets = contextDecisionCacheFacets(decision.include);
  const refreshPlan = output.workbookContextId ? metadataCache.planContextRefresh(String(output.workbookContextId), requiredCacheFacets) : undefined;
  const source = cacheHit && refreshPlan?.requiresRead !== true
    ? "cache" as const
    : runMetrics.internalReadCount > 0 || runMetrics.fullReadCellCount > 0
      ? refreshPlan && refreshPlan.cacheFacets.length > 0 ? "mixed" as const : "live" as const
      : output.workbookContextId
        ? "mixed" as const
        : "none" as const;
  const stagesUsed = [...stages];
  const skippedStages = decision.plannedStages.filter((stage) => !stages.has(stage));
  const stopReason = strategy === "mutation"
    ? "stop after target, field, validation, and preview proof are sufficient for safe preview"
    : decision.stopWhen;
  return stripUndefinedRecord({
    strategy,
    scope: decision.scope,
    ...(input.context?.level !== undefined ? { levelRequested: input.context.level } : {}),
    levelUsed: decision.level,
    levelReason: decision.reason,
    stagesPlanned: decision.plannedStages,
    stagesUsed,
    ...(skippedStages.length > 0 ? { skippedStages } : {}),
    stopReason,
    included: [...included],
    ...(rangesRead.length > 0 ? { rangesRead } : {}),
    ...(refreshPlan ? {
      requiredFacets: refreshPlan.requiredFacets,
      cachedFacetsUsed: refreshPlan.cacheFacets,
      ...(refreshPlan.missingFacets.length > 0 ? { missingFacets: refreshPlan.missingFacets } : {}),
      ...(refreshPlan.staleFacets.length > 0 ? { staleFacets: refreshPlan.staleFacets } : {}),
      ...(refreshPlan.liveFacets.length > 0 ? { facetsToRefresh: refreshPlan.liveFacets } : {}),
      refreshReason: refreshPlan.reason,
      freshnessRequiresRead: refreshPlan.requiresRead
    } : {}),
    ...(rowsRead !== undefined ? { rowsRead } : {}),
    estimatedTokens: Math.ceil(payloadBytes / 4),
    ...(truncated !== undefined ? { truncated } : {}),
    confidence: contextAnswerConfidence(output, runMetrics),
    source,
    continuation: {
      available: output.continuation !== undefined || output.nextAction === "fetch_resource",
      ...(suggestedNext !== undefined ? { suggestedNext } : {})
    }
  });
}

function contextFreshnessForOutput(
  output: Omit<AgentRunOutput, "telemetry">,
  metadataCache: WorkbookMetadataCache
): Pick<AgentRunOutput, "contextFreshness"> {
  const workbookContextId = output.workbookContextId ? String(output.workbookContextId) : undefined;
  if (!workbookContextId) {
    return {};
  }
  const state = metadataCache.getContextState(workbookContextId);
  return state ? { contextFreshness: state.freshness } : {};
}

function buildContextRefreshTelemetry(
  output: Omit<AgentRunOutput, "telemetry">,
  runMetrics: AgentRunMetrics,
  metadataCache: WorkbookMetadataCache
): NonNullable<AgentRunOutput["telemetry"]["contextRefresh"]> | undefined {
  const workbookContextId = output.workbookContextId ? String(output.workbookContextId) : undefined;
  if (!workbookContextId) {
    return undefined;
  }
  const plan = metadataCache.planContextRefresh(workbookContextId, contextDecisionCacheFacets(runMetrics.route.contextDecision.include));
  return {
    requiredFacets: plan.requiredFacets,
    cachedFacets: plan.cacheFacets,
    facetsToRefresh: plan.liveFacets,
    ...(plan.missingFacets.length > 0 ? { missingFacets: plan.missingFacets } : {}),
    ...(plan.staleFacets.length > 0 ? { staleFacets: plan.staleFacets } : {}),
    readStrategy: plan.readStrategy,
    reason: plan.reason,
    requiresRead: plan.requiresRead,
    confidence: plan.confidence
  };
}

function contextDecisionCacheFacets(include: string[]): ContextFacet[] {
  const facets = new Set<ContextFacet>();
  for (const facet of include) {
    switch (facet) {
      case "metadata":
        facets.add("metadata");
        break;
      case "schema":
        facets.add("schema");
        facets.add("headers");
        break;
      case "tables":
        facets.add("tableDimensions");
        break;
      case "regions":
        facets.add("regions");
        break;
      case "values":
        facets.add("values");
        break;
      case "field_context":
        facets.add("fieldContext");
        break;
      case "validation":
        facets.add("validation");
        break;
      case "formats":
        facets.add("formats");
        break;
      case "formulas":
        facets.add("formulas");
        facets.add("formulaResults");
        break;
      case "filters":
        facets.add("filters");
        break;
      case "names":
        facets.add("names");
        break;
    }
  }
  return [...facets];
}

function uniqueContextRanges(proof: AgentProofReference[]): string[] {
  const seen = new Set<string>();
  const ranges: string[] = [];
  for (const entry of proof) {
    const range = entry.range ? `${entry.sheetName}!${entry.range}` : undefined;
    if (!range || seen.has(range)) {
      continue;
    }
    seen.add(range);
    ranges.push(range);
  }
  return ranges.slice(0, 12);
}

function contextRowsRead(proof: AgentProofReference[]): number | undefined {
  const rows = proof.reduce((total, entry) => {
    const shape = entry.range ? rangeShape(entry.range) : undefined;
    return total + (shape?.rows ?? 0);
  }, 0);
  return rows > 0 ? rows : undefined;
}

function answerTruncated(answer: Record<string, unknown> | undefined): boolean | undefined {
  if (!answer) {
    return undefined;
  }
  if (answer.truncated === true || answer.previewTruncated === true) {
    return true;
  }
  if (answer.inlineIsComplete === false || answer.sourceComplete === false) {
    return true;
  }
  return undefined;
}

function contextAnswerConfidence(output: Omit<AgentRunOutput, "telemetry">, runMetrics: AgentRunMetrics): number {
  if (output.status === "SUCCESS" || output.status === "PREVIEW_READY") {
    return Math.max(0.75, Math.min(0.98, runMetrics.route.confidence));
  }
  if (output.status === "NEEDS_INPUT" || output.status === "AMBIGUOUS_TARGET") {
    return Math.min(0.55, runMetrics.route.confidence);
  }
  return Math.min(0.7, runMetrics.route.confidence);
}

function contextSuggestedNext(output: Omit<AgentRunOutput, "telemetry">, input: AgentRunInput): string[] | undefined {
  const suggestions = new Set<string>();
  if (output.nextAction === "fetch_resource" || output.continuation?.fullResultUri) {
    suggestions.add("fetch_full_result");
  }
  if (output.nextAction === "call_with_target" || output.status === "AMBIGUOUS_TARGET") {
    suggestions.add("provide_narrower_target");
  }
  if (output.mode === "preview_update" || input.mode === "preview_update") {
    suggestions.add("apply_preview_after_confirmation");
  }
  return suggestions.size > 0 ? [...suggestions] : undefined;
}

function compactContextUsedForBudget(contextUsed: AgentRunOutput["contextUsed"], maxPayloadBytes: number | undefined): AgentRunOutput["contextUsed"] {
  if (!contextUsed || maxPayloadBytes === undefined || maxPayloadBytes > 1_200) {
    return contextUsed;
  }
  if (maxPayloadBytes <= 900) {
    return {
      strategy: contextUsed.strategy,
      scope: contextUsed.scope,
      ...(contextUsed.levelRequested !== undefined ? { levelRequested: contextUsed.levelRequested } : {}),
      levelUsed: contextUsed.levelUsed,
      stagesUsed: contextUsed.stagesUsed.slice(0, 1),
      included: contextUsed.included.slice(0, 1)
    };
  }
  return stripUndefinedRecord({
    strategy: contextUsed.strategy,
    scope: contextUsed.scope,
    ...(contextUsed.levelRequested !== undefined ? { levelRequested: contextUsed.levelRequested } : {}),
    levelUsed: contextUsed.levelUsed,
    stagesUsed: contextUsed.stagesUsed.slice(0, 4),
    included: contextUsed.included.slice(0, 5),
    ...(contextUsed.facetsToRefresh !== undefined ? { facetsToRefresh: contextUsed.facetsToRefresh.slice(0, 4) } : {}),
    ...(contextUsed.freshnessRequiresRead !== undefined ? { freshnessRequiresRead: contextUsed.freshnessRequiresRead } : {}),
    ...(contextUsed.truncated !== undefined ? { truncated: contextUsed.truncated } : {}),
    ...(contextUsed.confidence !== undefined ? { confidence: contextUsed.confidence } : {}),
    ...(contextUsed.source !== undefined ? { source: contextUsed.source } : {})
  });
}

function compactPostContextForTinyBudget(output: Omit<AgentRunOutput, "telemetry">, maxPayloadBytes: number | undefined): Omit<AgentRunOutput, "telemetry"> {
  if (maxPayloadBytes === undefined || maxPayloadBytes > 900 || Buffer.byteLength(JSON.stringify(output)) <= maxPayloadBytes + 350) {
    return output;
  }
  return stripUndefinedOptionals({
    ...output,
    ...(output.candidates ? { candidates: output.candidates.slice(0, 1) } : {}),
    ...(output.contextFreshness ? { contextFreshness: compactContextFreshness(output.contextFreshness) } : {}),
    proof: output.proof.slice(0, 1),
    resourceLinks: output.resourceLinks.slice(0, 1),
    warnings: output.warnings.slice(0, 1)
  });
}

function compactContextFreshness(freshness: NonNullable<AgentRunOutput["contextFreshness"]>): NonNullable<AgentRunOutput["contextFreshness"]> {
  return {
    status: freshness.status,
    freshFacets: freshness.freshFacets.slice(0, 4),
    staleFacets: freshness.staleFacets.slice(0, 4),
    ...(freshness.staleRanges ? { staleRanges: freshness.staleRanges.slice(0, 2) } : {}),
    confidence: freshness.confidence,
    updatedAt: freshness.updatedAt
  };
}

function compactContextFreshnessForBudget(
  freshness: AgentRunOutput["contextFreshness"],
  maxPayloadBytes: number | undefined
): AgentRunOutput["contextFreshness"] {
  if (!freshness || maxPayloadBytes === undefined || maxPayloadBytes > 1_200) {
    return freshness;
  }
  if (maxPayloadBytes <= 900) {
    return {
      status: freshness.status,
      freshFacets: [],
      staleFacets: freshness.staleFacets.slice(0, 2),
      confidence: freshness.confidence,
      updatedAt: freshness.updatedAt
    };
  }
  return {
    status: freshness.status,
    freshFacets: [],
    staleFacets: freshness.staleFacets.slice(0, 3),
    confidence: freshness.confidence,
    updatedAt: freshness.updatedAt
  };
}

function uniqueChangeRanges(changes: NonNullable<AgentRunOutput["changes"]>): string[] {
  const seen = new Set<string>();
  const ranges: string[] = [];
  for (const change of changes) {
    const range = change.range ?? change.cell;
    if (!range) {
      continue;
    }
    const qualified = `${change.sheetName}!${range}`;
    if (seen.has(qualified)) {
      continue;
    }
    seen.add(qualified);
    ranges.push(qualified);
  }
  return ranges;
}

function isContextFacet(value: string): value is ContextFacet {
  return [
    "metadata",
    "schema",
    "headers",
    "tableDimensions",
    "regions",
    "fieldContext",
    "validation",
    "formats",
    "formulas",
    "formulaResults",
    "filters",
    "values",
    "aggregates",
    "rowPositions",
    "selection",
    "names"
  ].includes(value);
}

function minimalAnswerForBudget(answer: unknown, continuation: AgentRunOutput["continuation"] | undefined, workbookContextId: AgentRunOutput["workbookContextId"] | undefined): Record<string, unknown> {
  const typed = answer && typeof answer === "object" ? answer as Record<string, unknown> : {};
  if (typed.kind === "workbook_summary" || typed.kind === "workbook_overview") {
    return stripUndefinedRecord({
      kind: typed.kind,
      source: typed.source,
      sheetCount: typed.sheetCount,
      tableCount: typed.tableCount,
      namedRangeCount: typed.namedRangeCount,
      sheets: Array.isArray(typed.sheets) ? typed.sheets.slice(0, 8).map((sheet) => {
        const record = sheet as Record<string, unknown>;
        return stripUndefinedRecord({
          name: record.name,
          kind: record.kind,
          usedRange: record.usedRange,
          contextHints: Array.isArray(record.contextHints) ? record.contextHints.slice(0, 3) : undefined
        });
      }) : undefined,
      tables: Array.isArray(typed.tables) ? typed.tables.slice(0, 12) : undefined,
      namedRanges: Array.isArray(typed.namedRanges) ? typed.namedRanges.slice(0, 12) : undefined,
      resultUri: typed.resultUri ?? continuation?.resultUri,
      fullResultUri: typed.fullResultUri ?? continuation?.fullResultUri,
      resource: continuation?.resultUri ?? (workbookContextId ? contextResource(String(workbookContextId)).uri : undefined)
    });
  }
  if (typed.kind === "sheet_summary") {
    const sheet = typed.sheet && typeof typed.sheet === "object" ? typed.sheet as Record<string, unknown> : undefined;
    return stripUndefinedRecord({
      kind: typed.kind,
      source: typed.source,
      sheet: sheet ? compactSheetForBudget(sheet) : undefined,
      contextHints: Array.isArray(typed.contextHints) ? typed.contextHints.slice(0, 4) : undefined,
      tables: Array.isArray(typed.tables) ? typed.tables.slice(0, 6).flatMap((table) => compactTableForBudget(table) ?? []) : undefined,
      namedRangeCount: Array.isArray(typed.namedRanges) ? typed.namedRanges.length : undefined,
      sectionCount: Array.isArray(typed.sections) ? typed.sections.length : undefined,
      formulaRegionCount: Array.isArray(typed.formulaRegions) ? typed.formulaRegions.length : undefined,
      resultUri: typed.resultUri ?? continuation?.resultUri,
      fullResultUri: typed.fullResultUri ?? continuation?.fullResultUri,
      resource: continuation?.resultUri ?? (workbookContextId ? contextResource(String(workbookContextId)).uri : undefined)
    });
  }
  if (typed.kind === "semantic_workbook_index") {
    return compactSemanticIndexAnswer(typed, continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "table_compact_read") {
    return stripUndefinedRecord({
      kind: typed.kind,
      source: typed.source,
      sheetName: typed.sheetName,
      tableName: typed.tableName,
      range: typed.range,
      dataRange: typed.dataRange,
      rowOffset: typed.rowOffset,
      rowLimit: typed.rowLimit,
      inlineRowCount: typed.inlineRowCount,
      totalRowCount: typed.totalRowCount,
      inlineIsComplete: typed.inlineIsComplete,
      missingInlineReason: typed.missingInlineReason,
      projectedColumns: Array.isArray(typed.projectedColumns) ? typed.projectedColumns.slice(0, 16) : undefined,
      fieldContext: Array.isArray(typed.fieldContext) ? typed.fieldContext.slice(0, 12) : undefined,
      schemaSummary: typed.schemaSummary,
      shape: typed.shape,
      metrics: typed.metrics,
      headers: compactMatrixLike(typed.headers, 1, 16),
      encodedValues: compactMatrixLike(typed.encodedValues, 20, 16),
      valueEncoding: typed.valueEncoding,
      valuesPreview: compactMatrixLike(typed.valuesPreview, 5, 16),
      previewRange: typed.previewRange,
      previewTruncated: typed.previewTruncated,
      truncated: typed.truncated,
      resultUri: typed.resultUri ?? continuation?.resultUri,
      fullResultUri: typed.fullResultUri ?? continuation?.fullResultUri,
      resource: continuation?.resultUri ?? (workbookContextId ? contextResource(String(workbookContextId)).uri : undefined)
    });
  }
  if (typed.kind === "formula_read") {
    return compactFormulaReadAnswer(typed, "brief", continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "formula_patterns") {
    return compactFormulaPatternsAnswer(typed, "brief", continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "similar_rows") {
    return compactSimilarRowsAnswer(typed, continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "style_reference_candidates") {
    return compactStyleReferenceCandidatesAnswer(typed, continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "style_overview") {
    return compactStyleOverviewAnswer(typed, continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "grouped_header_summary") {
    return stripUndefinedRecord({
      kind: typed.kind,
      sheetName: typed.sheetName,
      range: typed.range,
      mergedRangeCount: typed.mergedRangeCount,
      mergeStatus: typed.mergeStatus,
      spans: Array.isArray(typed.spans) ? typed.spans.slice(0, 12) : undefined,
      unmergedLabels: Array.isArray(typed.unmergedLabels) ? typed.unmergedLabels.slice(0, 12) : undefined,
      resultUri: typed.resultUri ?? continuation?.resultUri,
      fullResultUri: typed.fullResultUri ?? continuation?.fullResultUri,
      resource: continuation?.resultUri ?? (workbookContextId ? contextResource(String(workbookContextId)).uri : undefined)
    });
  }
  if (typed.kind === "workbook_design_overview") {
    return compactWorkbookDesignOverviewAnswer(typed, continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "visual_readability_preview") {
    return compactVisualReadabilityPreviewAnswer(typed, continuation?.resultUri, continuation?.fullResultUri);
  }
  if (typed.kind === "reference_sheet_analysis") {
    return stripUndefinedRecord({
      kind: typed.kind,
      source: typed.source,
      reference: typed.reference,
      objectives: Array.isArray(typed.objectives) ? typed.objectives.slice(0, 8) : undefined,
      structure: typed.structure,
      columnProfiles: Array.isArray(typed.columnProfiles) ? typed.columnProfiles.slice(0, 12) : undefined,
      relationships: Array.isArray(typed.relationships) ? typed.relationships.slice(0, 8) : undefined,
      formulaPatterns: Array.isArray(typed.formulaPatterns) ? typed.formulaPatterns.slice(0, 6) : undefined,
      stylePatterns: Array.isArray(typed.stylePatterns) ? typed.stylePatterns.slice(0, 4) : undefined,
      rowConditionPatterns: Array.isArray(typed.rowConditionPatterns) ? typed.rowConditionPatterns.slice(0, 6) : undefined,
      evidenceRows: Array.isArray(typed.evidenceRows) ? typed.evidenceRows.slice(0, 4) : undefined,
      nextRequestHints: Array.isArray(typed.nextRequestHints) ? typed.nextRequestHints.slice(0, 2) : undefined,
      resultUri: typed.resultUri ?? continuation?.resultUri,
      fullResultUri: typed.fullResultUri ?? continuation?.fullResultUri,
      resource: continuation?.resultUri ?? (workbookContextId ? contextResource(String(workbookContextId)).uri : undefined)
    });
  }
  return stripUndefinedRecord({
    kind: typed.kind ?? "compacted_answer",
    source: typed.source,
    sheetName: typed.sheetName,
    tableName: typed.tableName,
    range: typed.range,
    dataRange: typed.dataRange,
    resultUri: typed.resultUri ?? continuation?.resultUri,
    fullResultUri: typed.fullResultUri ?? continuation?.fullResultUri,
    resource: continuation?.resultUri ?? (workbookContextId ? contextResource(String(workbookContextId)).uri : undefined)
  });
}

function compactMatrixLike(value: unknown, maxRows: number, maxColumns: number): unknown {
  if (!Array.isArray(value)) {
    return undefined;
  }
  if (value.length > 0 && Array.isArray(value[0])) {
    return (value as unknown[][]).slice(0, maxRows).map((row) => row.slice(0, maxColumns));
  }
  return maxRows > 0 ? value.slice(0, maxColumns) : undefined;
}

function compactSheetForBudget(sheet: Record<string, unknown>): Record<string, unknown> {
  return stripUndefinedRecord({
    name: sheet.name,
    kind: sheet.kind,
    usedRange: sheet.usedRange,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount
  });
}

function compactTableForBudget(table: unknown): Record<string, unknown> | undefined {
  if (!table || typeof table !== "object") {
    return undefined;
  }
  const typed = table as Record<string, unknown>;
  const columns = Array.isArray(typed.columns)
    ? typed.columns.slice(0, 16).map((column) => {
        const record = column && typeof column === "object" ? column as Record<string, unknown> : {};
        return stripUndefinedRecord({
          name: record.name,
          letter: record.letter,
          role: record.role,
          inferredType: record.inferredType
        });
      })
    : undefined;
  return stripUndefinedRecord({
    name: typed.name,
    sheetName: typed.sheetName,
    range: typed.range,
    dataRange: typed.dataRange,
    columnCount: typed.columnCount ?? columns?.length,
    columns
  });
}

function inlinePreviewForProfile(input: AgentRunInput, profile: ReturnType<typeof profileValues>, range: string): Record<string, unknown> | undefined {
  if (!shouldIncludeInlineValuesPreview(input, profile.shape.rows * profile.shape.columns)) {
    return undefined;
  }
  const rows = Array.isArray(profile.rows) ? profile.rows as CellMatrix : undefined;
  if (!rows || rows.length === 0) {
    return undefined;
  }
  return inlinePreviewForMatrix(input, rows, range);
}

function inlinePreviewForMatrix(input: AgentRunInput, matrix: CellMatrix, range?: string): Record<string, unknown> | undefined {
  const cellCount = matrixCellCount(matrix);
  if (!shouldIncludeInlineValuesPreview(input, cellCount) || cellCount === 0) {
    return undefined;
  }
  const maxCells = Math.min(500, input.budget?.maxPayloadBytes ? Math.max(25, Math.floor(input.budget.maxPayloadBytes / 80)) : 500);
  const preview: CellMatrix = [];
  let usedCells = 0;
  let truncated = false;
  for (const row of matrix) {
    const width = row.length;
    if (usedCells + width > maxCells) {
      truncated = true;
      break;
    }
    preview.push(row);
    usedCells += width;
  }
  if (preview.length === 0 && matrix[0]) {
    preview.push(matrix[0].slice(0, maxCells));
    truncated = matrix[0].length > maxCells || matrix.length > 1;
  }
  return stripUndefinedRecord({
    valuesPreview: preview,
    previewRange: range,
    previewTruncated: truncated || preview.length < matrix.length
  });
}

function inlineCompletenessMetadata(inlinePreview: Record<string, unknown> | undefined, matrix: CellMatrix, totalRows: number): Record<string, unknown> {
  const valuesPreview = Array.isArray(inlinePreview?.valuesPreview) ? inlinePreview.valuesPreview as CellMatrix : undefined;
  const inlineRowCount = valuesPreview?.length ?? 0;
  const previewTruncated = inlinePreview?.previewTruncated === true;
  const inlineIsComplete = inlineRowCount >= totalRows && !previewTruncated;
  const missingInlineReason = inlineIsComplete
    ? undefined
    : inlineRowCount === 0 && matrix.length === 0
      ? "empty_result"
      : previewTruncated || inlineRowCount < totalRows
        ? "payload_budget_preview"
        : undefined;
  return stripUndefinedRecord({
    inlineRowCount,
    totalRowCount: totalRows,
    inlineIsComplete,
    missingInlineReason
  });
}

function tinyInlinePreviewForMatrix(matrix: CellMatrix, range?: string): Record<string, unknown> | undefined {
  if (matrix.length === 0 || matrixCellCount(matrix) === 0) {
    return undefined;
  }
  const maxRows = 3;
  const maxColumns = 12;
  return stripUndefinedRecord({
    valuesPreview: matrix.slice(0, maxRows).map((row) => row.slice(0, maxColumns)),
    previewRange: range,
    previewTruncated: matrix.length > maxRows || maxMatrixColumns(matrix) > maxColumns,
    previewPurpose: "tiny_exact_sample_to_answer_without_followup"
  });
}

function shouldIncludeInlineValuesPreview(input: AgentRunInput, cellCount: number): boolean {
  if (cellCount > 500 && input.budget?.maxPayloadBytes === undefined) {
    return false;
  }
  const action = intentAction(input);
  return action === "read_values"
    || input.detailLevel === "table_sample"
    || input.detailLevel === "full_table"
    || (cellCount > 0 && cellCount <= 50 && Boolean(input.target?.range))
    || /\b(?:actual|raw|all|full|every|show|read|print)\b.{0,40}\b(?:values?|rows?|data|headers?)\b/i.test(input.request);
}

function shouldPreserveExactInlineData(input: AgentRunInput, responseMode: AgentResponseMode, cellCount: number, proof?: AgentProofReference[]): boolean {
  const maxCells = exactInlineCellLimit(input, responseMode);
  if (cellCount <= 0 || cellCount > maxCells) {
    return false;
  }
  const action = intentAction(input);
  return action === "read_values"
    || action === "read_range_compact"
    || input.detailLevel === "table_sample"
    || input.detailLevel === "full_table"
    || requestNeedsExactWorkbookData(input.request)
    || proofImpliesSelectedData(proof);
}

function exactInlineCellLimit(input: AgentRunInput, responseMode: AgentResponseMode): number {
  const modeLimit = responseMode === "standard" ? 500 : 300;
  if (input.budget?.maxPayloadBytes === undefined) {
    return modeLimit;
  }
  return Math.max(25, Math.min(modeLimit, Math.floor(input.budget.maxPayloadBytes / 100)));
}

function requestNeedsExactWorkbookData(request: string): boolean {
  return /\b(?:actual|raw|exact|show|read|print|list|analy[sz]e|inspect|look(?:\s+at| into)?|validate|transform|clean|fix|update|summari[sz]e)\b.{0,60}\b(?:values?|rows?|cells?|range|data|headers?|selection|selected|this|here|current)\b/i.test(request)
    || /\b(?:selected|selection|this|here|current)\b.{0,60}\b(?:values?|rows?|cells?|range|data)\b/i.test(request)
    || /\b(?:read|show|inspect|look(?:\s+at| into)?|analy[sz]e)\b.{0,80}\b[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?\b/i.test(request);
}

function proofImpliesSelectedData(proof: AgentProofReference[] | undefined): boolean {
  return (proof ?? []).some((entry) => /\b(?:selected|selection|active cell|active table row)\b/i.test(entry.label ?? ""));
}

function sparseRowsCellCount(rows: unknown[] | undefined): number {
  if (!rows) {
    return 0;
  }
  return rows.reduce<number>((total, row) => {
    const cells = row && typeof row === "object" && Array.isArray((row as { cells?: unknown }).cells)
      ? (row as { cells: unknown[] }).cells
      : [];
    return total + cells.length;
  }, 0);
}

function responseModeFromInput(input: AgentRunInput): AgentResponseMode {
  return input.responseMode ?? "brief";
}

function targetFreshnessRangesFromInput(input: AgentRunInput, activeWorkbookId: WorkbookId | string | undefined): A1Range[] | undefined {
  const target = input.target;
  const workbookId = target?.workbookId ?? activeWorkbookId;
  if (!workbookId || !target?.sheetName || !target.range) {
    return undefined;
  }
  return [{
    workbookId: workbookId as WorkbookId,
    sheetName: target.sheetName,
    address: stripSheetName(target.range)
  }];
}

type AgentDetectedResourceHandle =
  | { kind: "context"; id: string }
  | { kind: "semantic_index"; id: string }
  | { kind: "operation"; id: string }
  | { kind: "result"; id: string; view?: "summary" | "full" }
  | { kind: "compact"; id: string; view?: "summary" | "full" };

function detectAgentResourceHandle(input: AgentRunInput): AgentDetectedResourceHandle | undefined {
  if (input.operationId) {
    return { kind: "operation", id: String(input.operationId) };
  }
  const continuationHandle = continuationResultHandle(input);
  if ((continuationHandle?.kind === "result" || continuationHandle?.kind === "compact")
    && (isResultHandleContinuationRequest(input) || shouldReturnFullResource(input, continuationHandle.view))) {
    return continuationHandle;
  }
  const requestHandle = parseAgentResourceUri(input.request);
  if (requestHandle) {
    return requestHandle;
  }
  if (isContextOnlyContinuationRequest(input) && input.workbookContextId) {
    return { kind: "context", id: String(input.workbookContextId) };
  }
  return undefined;
}

function continuationResultHandle(input: AgentRunInput): AgentDetectedResourceHandle | undefined {
  const resultUri = input.continuation?.resultUri;
  const fullResultUri = input.continuation?.fullResultUri;
  const wantsFull = input.responseMode === "verbose" || shouldReturnFullResource(input);
  const selected = wantsFull && fullResultUri ? fullResultUri : resultUri ?? fullResultUri;
  return selected ? parseAgentResourceUri(selected) : undefined;
}

function parseAgentResourceUri(text: string): AgentDetectedResourceHandle | undefined {
  const compact = /excel:\/\/compact\/([A-Za-z0-9_-]+)(?:\?([^\s"'<>]+))?/i.exec(text);
  if (compact?.[1]) {
    return { kind: "compact", id: compact[1], ...resourceViewFromQuery(compact[2]) };
  }
  const result = /excel:\/\/agent\/results\/([A-Za-z0-9_-]+)(?:\?([^\s"'<>]+))?/i.exec(text);
  if (result?.[1]) {
    return { kind: "result", id: result[1], ...resourceViewFromQuery(result[2]) };
  }
  const operation = /excel:\/\/agent\/operations\/([A-Za-z0-9_-]+)/i.exec(text);
  if (operation?.[1]) {
    return { kind: "operation", id: operation[1] };
  }
  const semanticIndex = /excel:\/\/agent\/contexts\/([A-Za-z0-9_-]+)\/semantic-index/i.exec(text);
  if (semanticIndex?.[1]) {
    return { kind: "semantic_index", id: semanticIndex[1] };
  }
  const context = /excel:\/\/agent\/contexts\/([A-Za-z0-9_-]+)/i.exec(text);
  if (context?.[1]) {
    return { kind: "context", id: context[1] };
  }
  const contextText = /\b(?:continue|reuse|using|use)\s+(?:context\s+)?(wbctx_[A-Za-z0-9_-]+|ctx_[A-Za-z0-9_-]+)/i.exec(text);
  if (contextText?.[1]) {
    return { kind: "context", id: contextText[1] };
  }
  return undefined;
}

function resourceViewFromQuery(query: string | undefined): { view?: "summary" | "full" } {
  return query && /(?:^|&)view=full(?:&|$)/i.test(query) ? { view: "full" } : {};
}

function isContextOnlyContinuationRequest(input: AgentRunInput): boolean {
  const request = input.request.trim().toLowerCase();
  return Boolean(input.workbookContextId)
    && !input.target
    && !input.values
    && /^(continue|reuse|use|look at|show|inspect|summarize|summary)\b/.test(request)
    && /\b(context|workbook context|same workbook|previous context)\b/.test(request);
}

function isResultHandleContinuationRequest(input: AgentRunInput): boolean {
  return /\b(continue|reuse|use|show|inspect|read|fetch|open)\b/i.test(input.request)
    && /\b(stored result|previous result|result handle|result resource|resource handle|full result|stored detail|stored details)\b/i.test(input.request);
}

function shouldReturnFullResource(input: AgentRunInput, requestedView?: "summary" | "full"): boolean {
  if (requestedView === "full" || input.responseMode === "verbose") {
    return true;
  }
  return /\b(full|all rows|all values|raw values|complete|entire|everything|detail|details|audit)\b/i.test(input.request);
}

function applyContinuationInput(input: AgentRunInput): AgentRunInput {
  const continuation = input.continuation;
  if (!continuation) {
    return input;
  }
  return stripUndefinedInput({
    ...input,
    ...(input.workbookContextId !== undefined || continuation.workbookContextId !== undefined ? { workbookContextId: input.workbookContextId ?? continuation.workbookContextId } : {}),
    ...(input.operationId !== undefined || continuation.operationId !== undefined ? { operationId: input.operationId ?? continuation.operationId } : {}),
    ...(input.transactionId !== undefined || continuation.transactionId !== undefined ? { transactionId: input.transactionId ?? continuation.transactionId } : {}),
    ...(input.responseMode !== undefined || continuation.responseMode !== undefined ? { responseMode: input.responseMode ?? continuation.responseMode } : {})
  });
}

function stripUndefinedInput(input: AgentRunInput): AgentRunInput {
  const next = { ...input };
  if (next.workbookContextId === undefined) delete next.workbookContextId;
  if (next.operationId === undefined) delete next.operationId;
  if (next.transactionId === undefined) delete next.transactionId;
  if (next.responseMode === undefined) delete next.responseMode;
  return next;
}

function defaultResponseBudget(responseMode: AgentResponseMode): { maxExamples: number; maxPayloadBytes?: number; maxEstimatedTokens?: number } {
  if (responseMode === "verbose") {
    return { maxExamples: 10 };
  }
  if (responseMode === "standard") {
    return { maxExamples: 5, maxPayloadBytes: 12_000, maxEstimatedTokens: 3_000 };
  }
  return { maxExamples: 3, maxPayloadBytes: 6_000, maxEstimatedTokens: 1_500 };
}

function appendUniqueResource(resources: AgentRunOutput["resourceLinks"], resource: AgentRunOutput["resourceLinks"][number]): AgentRunOutput["resourceLinks"] {
  return resources.some((entry) => entry.uri === resource.uri) ? resources : [...resources, resource];
}

function answerNeedsResultResource(answer: unknown): boolean {
  if (!answer || typeof answer !== "object") {
    return false;
  }
  const text = JSON.stringify(answer);
  const typed = answer as Record<string, unknown>;
  if (typed.kind === "match_update_preview" || typed.kind === "match_update_result" || typed.kind === "match_update_no_match" || typed.kind === "exact_search_rows" || typed.kind === "exact_search_no_match") {
    return false;
  }
  if (typed.kind === "workbook_summary" || typed.kind === "workbook_overview" || typed.kind === "sheet_summary" || typed.kind === "data_validation_summary") {
    return false;
  }
  if (typed.kind === "range_profile" && typed.inlineIsComplete === true && Array.isArray(typed.valuesPreview) && matrixCellCount(typed.valuesPreview as CellMatrix) <= 50) {
    return false;
  }
  return text.length > 2_000
    || ["headers", "values", "formulas", "text", "numberFormat", "sample", "sparseRows", "rows", "valueCounts"].some((key) => typed[key] !== undefined)
    || Object.values(typed).some((value) => value && typeof value === "object" && ["headers", "values", "sample", "sparseRows", "rows", "valueCounts"].some((key) => (value as Record<string, unknown>)[key] !== undefined));
}

function compactAnswerForResponseMode(answer: unknown, responseMode: AgentResponseMode, input: AgentRunInput, resultUri?: string, fullResultUri?: string, proof?: AgentProofReference[]): unknown {
  if (!answer || typeof answer !== "object") {
    return answer;
  }
  if (responseMode === "verbose") {
    return answer;
  }
  const typed = answer as Record<string, unknown>;
  const kind = typeof typed.kind === "string" ? typed.kind : undefined;
  if (kind === "range_schema") {
    return compactRangeSchemaAnswer(typed, responseMode, resultUri, fullResultUri);
  }
  if (kind === "table_schema") {
    return compactTableSchemaAnswer(typed, responseMode, resultUri, fullResultUri);
  }
  if (kind === "range_profile") {
    return compactRangeProfileAnswer(typed, responseMode, input, resultUri, fullResultUri, proof);
  }
  if (kind === "range_value_counts") {
    return compactRangeValueCountsAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "table_compact_read") {
    return compactTableReadAnswer(typed, responseMode, input, resultUri, fullResultUri, proof);
  }
  if (kind === "comparison_profile") {
    return compactComparisonAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "semantic_workbook_index") {
    return compactSemanticIndexAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "workbook_overview") {
    return compactWorkbookOverviewAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "workbook_summary") {
    return compactWorkbookSummaryAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "sheet_summary") {
    return compactSheetSummaryAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "style_summary") {
    return compactStyleSummaryAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "style_overview") {
    return compactStyleOverviewAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "workbook_design_overview") {
    return compactWorkbookDesignOverviewAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "similar_rows") {
    return compactSimilarRowsAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "match_update_preview" || kind === "match_update_result" || kind === "match_update_no_match" || kind === "exact_search_rows" || kind === "exact_search_no_match") {
    return typed;
  }
  if (kind === "style_reference_candidates") {
    return compactStyleReferenceCandidatesAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "formula_read") {
    return compactFormulaReadAnswer(typed, responseMode, resultUri, fullResultUri);
  }
  if (kind === "formula_patterns") {
    return compactFormulaPatternsAnswer(typed, responseMode, resultUri, fullResultUri);
  }
  if (kind === "visual_readability_preview") {
    return compactVisualReadabilityPreviewAnswer(typed, resultUri, fullResultUri);
  }
  if (kind === "data_validation_summary") {
    return compactDataValidationSummaryAnswer(typed, resultUri, fullResultUri);
  }
  return compactGenericAnswer(typed, resultUri, fullResultUri);
}

function compactFormulaReadAnswer(answer: Record<string, unknown>, responseMode: AgentResponseMode, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const maxCells = responseMode === "standard" ? 12 : 8;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    sheetName: answer.sheetName,
    range: answer.range,
    shape: answer.shape,
    formulaCount: answer.formulaCount,
    hardcodedCount: answer.hardcodedCount,
    blankCount: answer.blankCount,
    formulaResultBlankCount: answer.formulaResultBlankCount,
    hardcodedInFormulaColumnCount: answer.hardcodedInFormulaColumnCount,
    missingFormulaGaps: Array.isArray(answer.missingFormulaGaps) ? answer.missingFormulaGaps.slice(0, 5) : undefined,
    cells: Array.isArray(answer.cells) ? answer.cells.slice(0, maxCells) : undefined,
    formulas: compactMatrixLike(answer.formulas, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    formulasR1C1: compactMatrixLike(answer.formulasR1C1, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    values: compactMatrixLike(answer.values, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    text: compactMatrixLike(answer.text, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    patternMatrix: compactMatrixLike(answer.patternMatrix, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    resultUri,
    fullResultUri
  });
}

function compactFormulaPatternsAnswer(answer: Record<string, unknown>, responseMode: AgentResponseMode, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const patterns = isRecord(answer.patterns) ? answer.patterns : {};
  return stripUndefinedRecord({
    kind: answer.kind,
    workbookId: patterns.workbookId,
    sheetName: patterns.sheetName,
    address: patterns.address,
    formulaCount: patterns.formulaCount,
    patterns: Array.isArray(patterns.patterns) ? patterns.patterns.slice(0, responseMode === "standard" ? 8 : 5) : undefined,
    cells: Array.isArray(patterns.cells) ? patterns.cells.slice(0, responseMode === "standard" ? 10 : 5) : undefined,
    formulas: compactMatrixLike(patterns.formulas, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    formulasR1C1: compactMatrixLike(patterns.formulasR1C1, responseMode === "standard" ? 5 : 3, responseMode === "standard" ? 8 : 5),
    resultUri,
    fullResultUri
  });
}

function compactRangeSchemaAnswer(answer: Record<string, unknown>, responseMode: AgentResponseMode, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const headers = Array.isArray(answer.headers) ? answer.headers as HeaderMetadata[] : [];
  const selected = headers.slice(0, responseMode === "standard" ? 3 : 1);
  const columns = selected.flatMap((header) => Array.isArray(header.columns) ? header.columns : []);
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    sheetName: answer.sheetName,
    range: answer.range,
    schemaSummary: schemaSummary(columns, responseMode === "standard" ? 16 : 10),
    headerCount: headers.length,
    resultUri,
    fullResultUri
  });
}

function compactTableSchemaAnswer(answer: Record<string, unknown>, responseMode: AgentResponseMode, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const columns = Array.isArray(answer.columns) ? answer.columns : [];
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    tableName: answer.tableName,
    sheetName: answer.sheetName,
    range: answer.range,
    headerRange: answer.headerRange,
    dataRange: answer.dataRange,
    schemaSummary: schemaSummary(columns, responseMode === "standard" ? 16 : 10),
    resultUri,
    fullResultUri
  });
}

function compactRangeProfileAnswer(answer: Record<string, unknown>, responseMode: AgentResponseMode, input: AgentRunInput, resultUri?: string, fullResultUri?: string, proof?: AgentProofReference[]): Record<string, unknown> {
  const rows = Array.isArray(answer.rows) ? answer.rows as CellMatrix : undefined;
  const sparseRows = Array.isArray(answer.sparseRows) ? answer.sparseRows : undefined;
  const exactCellCount = rows ? matrixCellCount(rows) : sparseRowsCellCount(sparseRows);
  const preserveExact = shouldPreserveExactInlineData(input, responseMode, exactCellCount, proof);
  const inlineRows = preserveExact && rows ? rows.length : countMatrixRows(answer.valuesPreview) ?? 0;
  const totalRows = typeof answer.totalRowCount === "number" ? answer.totalRowCount : typeof (answer.shape as Record<string, unknown> | undefined)?.rows === "number" ? (answer.shape as Record<string, number>).rows : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    sheetName: answer.sheetName,
    range: answer.range,
    shape: answer.shape,
    metrics: compactProfileMetrics(answer.metrics),
    inlineRowCount: inlineRows,
    totalRowCount: totalRows,
    inlineIsComplete: totalRows !== undefined ? inlineRows >= totalRows && answer.previewTruncated !== true : answer.inlineIsComplete,
    missingInlineReason: totalRows !== undefined && inlineRows < totalRows ? "payload_budget_preview" : answer.missingInlineReason,
    emptyResultKind: answer.emptyResultKind,
    rowMetadata: answer.rowMetadata,
    ...(preserveExact && rows ? { rows } : {}),
    ...(preserveExact && !rows && sparseRows ? { sparseRows } : {}),
    ...(preserveExact ? { emptySummary: answer.emptySummary } : {}),
    ...(preserveExact ? { inlineRowsReason: "narrow_exact_read" } : {}),
    valuesPreview: answer.valuesPreview,
    previewRange: answer.previewRange,
    previewTruncated: answer.previewTruncated,
    resultUri,
    fullResultUri
  });
}

function compactRangeValueCountsAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    column: answer.column,
    uniqueCount: answer.uniqueCount,
    valueCount: answer.valueCount,
    topValues: Array.isArray(answer.valueCounts) ? answer.valueCounts.slice(0, 10) : undefined,
    resultUri,
    fullResultUri
  });
}

function compactTableReadAnswer(answer: Record<string, unknown>, responseMode: AgentResponseMode, input: AgentRunInput, resultUri?: string, fullResultUri?: string, proof?: AgentProofReference[]): Record<string, unknown> {
  const schema = Array.isArray(answer.schema) ? answer.schema : [];
  const projectedColumns = Array.isArray(answer.projectedColumns) ? answer.projectedColumns : [];
  const profile = answer.profile && typeof answer.profile === "object" ? answer.profile as Record<string, unknown> : undefined;
  const values = Array.isArray(answer.values) ? answer.values as CellMatrix : undefined;
  const formulas = Array.isArray(answer.formulas) ? answer.formulas as CellMatrix : undefined;
  const text = Array.isArray(answer.text) ? answer.text as CellMatrix : undefined;
  const numberFormat = Array.isArray(answer.numberFormat) ? answer.numberFormat as CellMatrix : undefined;
  const exactMatrix = values ?? text ?? formulas;
  const preserveExact = shouldPreserveExactInlineData(input, responseMode, exactMatrix ? matrixCellCount(exactMatrix) : 0, proof);
  const roleProjected = preserveExact && exactMatrix
    ? roleAwareMatrixProjection(exactMatrix, schema, input, responseMode)
    : undefined;
  const inlineColumns = roleProjected?.columns ?? projectedColumns;
  const exactHeaders = Array.isArray(answer.headers)
    ? roleProjected ? projectVector(answer.headers, roleProjected.indexes) : answer.headers
    : undefined;
  const compactHeaders = exactHeaders ?? (Array.isArray(answer.headers) ? (answer.headers as unknown[]).slice(0, responseMode === "standard" ? 16 : 12) : undefined);
  const projectedValues = preserveExact ? (values && roleProjected ? projectMatrixColumns(values, roleProjected.indexes) : values) : undefined;
  const projectedFormulas = preserveExact ? (formulas && roleProjected ? projectMatrixColumns(formulas, roleProjected.indexes) : formulas) : undefined;
  const projectedText = preserveExact ? (text && roleProjected ? projectMatrixColumns(text, roleProjected.indexes) : text) : undefined;
  const projectedNumberFormat = preserveExact ? (numberFormat && roleProjected ? projectMatrixColumns(numberFormat, roleProjected.indexes) : numberFormat) : undefined;
  const compactInlineRows = countMatrixRows(projectedValues) ?? countMatrixRows(projectedText) ?? countMatrixRows(projectedFormulas) ?? countMatrixRows(answer.valuesPreview) ?? 0;
  const totalRows = typeof answer.totalRowCount === "number" ? answer.totalRowCount : undefined;
  const domainEncoding = projectedValues
    ? domainEncodeMatrix(projectedValues, inlineColumns, {
      formulas: projectedFormulas,
      numberFormat: projectedNumberFormat,
      responseMode
    })
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    tableName: answer.tableName,
    sheetName: answer.sheetName,
    range: answer.range,
    dataRange: answer.dataRange,
    rowOffset: answer.rowOffset,
    rowLimit: answer.rowLimit,
    inlineRowCount: compactInlineRows,
    totalRowCount: totalRows,
    inlineIsComplete: totalRows !== undefined ? compactInlineRows >= totalRows && answer.previewTruncated !== true : answer.inlineIsComplete,
    missingInlineReason: totalRows !== undefined && compactInlineRows < totalRows ? "payload_budget_preview" : answer.missingInlineReason,
    projectedColumns,
    fieldContext: Array.isArray(answer.fieldContext) ? answer.fieldContext.slice(0, responseMode === "standard" ? 16 : 12) : undefined,
    truncated: answer.truncated,
    nextPage: answer.nextPage,
    schemaSummary: schemaSummary(schema, responseMode === "standard" ? 16 : 10),
    shape: profile?.shape,
    metrics: compactProfileMetrics(profile?.metrics),
    rowMetadata: answer.rowMetadata,
    headers: compactHeaders,
    values: domainEncoding ? undefined : projectedValues,
    encodedValues: domainEncoding?.encodedValues,
    valueEncoding: domainEncoding?.encoding,
    formulas: projectedFormulas,
    text: projectedText,
    numberFormat: projectedNumberFormat,
    inlineColumnProjection: roleProjected ? {
      reason: "role_aware_wide_row_projection",
      selectedColumnIndexes: roleProjected.indexes,
      selectedColumns: roleProjected.columns,
      omittedColumnCount: roleProjected.omittedColumnCount
    } : undefined,
    valuesPreview: answer.valuesPreview,
    previewRange: answer.previewRange,
    previewTruncated: answer.previewTruncated,
    resultUri,
    fullResultUri
  });
}

function roleAwareMatrixProjection(
  matrix: CellMatrix,
  schema: unknown[],
  input: AgentRunInput,
  responseMode: AgentResponseMode
): { indexes: number[]; columns: unknown[]; omittedColumnCount: number } | undefined {
  const width = maxMatrixColumns(matrix);
  const limit = responseMode === "standard" ? 16 : 12;
  if (width <= limit || schema.length === 0) {
    return undefined;
  }
  const requested = new Set(tableReadColumnsFromInput(input).map((column) => typeof column === "number" ? column : normalizeHeaderName(column)));
  const scored = schema
    .filter((column): column is Record<string, unknown> => Boolean(column && typeof column === "object"))
    .map((column, fallbackIndex) => {
      const index = typeof column.index === "number" ? column.index : fallbackIndex;
      const normalized = typeof column.normalizedName === "string"
        ? column.normalizedName
        : typeof column.name === "string" ? normalizeHeaderName(column.name) : "";
      const role = typeof column.role === "string" ? column.role : "";
      const importance = typeof column.importance === "number" ? column.importance : roleImportance(role);
      const requestedBoost = requested.has(index) || requested.has(normalized) ? 1 : 0;
      const proximityBoost = index <= 2 ? 0.08 : 0;
      return { column, index, requested: requestedBoost > 0, score: importance + requestedBoost + proximityBoost };
    })
    .filter((entry) => entry.index >= 0 && entry.index < width)
    .sort((left, right) => right.score - left.score || left.index - right.index);
  const selectedSet = new Set<number>();
  for (const entry of scored.filter((candidate) => candidate.requested)) {
    selectedSet.add(entry.index);
  }
  for (const entry of scored) {
    if (selectedSet.size >= limit) {
      break;
    }
    selectedSet.add(entry.index);
  }
  const selected = [...selectedSet].sort((left, right) => left - right);
  if (selected.length === 0 || selected.length >= width) {
    return undefined;
  }
  const schemaByIndex = new Map(scored.map((entry) => [entry.index, entry.column]));
  return {
    indexes: selected,
    columns: selected.map((index) => schemaByIndex.get(index) ?? { index }),
    omittedColumnCount: Math.max(0, width - selected.length)
  };
}

function projectMatrixColumns<T>(matrix: T[][], indexes: number[]): T[][] {
  return matrix.map((row) => indexes.map((index) => row[index] as T));
}

function projectVector<T>(values: T[], indexes: number[]): T[] {
  return indexes.map((index) => values[index] as T);
}

function domainEncodeMatrix(
  matrix: CellMatrix,
  columns: unknown[],
  options: { formulas?: CellMatrix | undefined; numberFormat?: CellMatrix | undefined; responseMode: AgentResponseMode }
): { encodedValues: CellMatrix; encoding: Record<string, unknown> } | undefined {
  if (options.responseMode === "verbose" || matrix.length < 6) {
    return undefined;
  }
  const width = maxMatrixColumns(matrix);
  const encodableColumns: Array<{
    position: number;
    name?: unknown;
    role?: unknown;
    inferredType?: unknown;
    domain: unknown[];
    byKey: Map<string, number>;
  }> = [];
  for (let position = 0; position < width; position += 1) {
    const column = columns[position] && typeof columns[position] === "object" ? columns[position] as Record<string, unknown> : {};
    if (!isDomainEncodingCandidate(column, matrix, position, options)) {
      continue;
    }
    const domain = orderedColumnDomain(matrix, position);
    if (!shouldEncodeDomain(matrix, position, domain)) {
      continue;
    }
    encodableColumns.push({
      position,
      name: column.name,
      role: column.role,
      inferredType: column.inferredType,
      domain,
      byKey: new Map(domain.map((value, index) => [domainKey(value), index]))
    });
  }
  if (encodableColumns.length === 0) {
    return undefined;
  }
  const byPosition = new Map(encodableColumns.map((column) => [column.position, column]));
  const encodedValues = matrix.map((row) => row.map((value, position) => {
    const encoding = byPosition.get(position);
    return encoding ? encoding.byKey.get(domainKey(value)) ?? value : value;
  }));
  const rawBytes = Buffer.byteLength(JSON.stringify(matrix));
  const encodedBytes = Buffer.byteLength(JSON.stringify(encodedValues)) + Buffer.byteLength(JSON.stringify(encodableColumns.map((column) => column.domain)));
  if (rawBytes - encodedBytes < 80 || encodedBytes >= rawBytes * 0.9) {
    return undefined;
  }
  return {
    encodedValues,
    encoding: {
      kind: "domain_dictionary_by_column",
      basis: "column_role_cardinality_value_pattern",
      decodeInstruction: "For encodedValues, replace integer codes using valueEncoding.columns[position].domain[code]. Full raw values remain available through fullResultUri.",
      columns: encodableColumns.map((column) => stripUndefinedRecord({
        position: column.position,
        name: column.name,
        role: column.role,
        inferredType: column.inferredType,
        domain: column.domain
      }))
    }
  };
}

function isDomainEncodingCandidate(
  column: Record<string, unknown>,
  matrix: CellMatrix,
  position: number,
  options: { formulas?: CellMatrix | undefined; numberFormat?: CellMatrix | undefined }
): boolean {
  const role = typeof column.role === "string" ? column.role : "";
  const inferredType = typeof column.inferredType === "string" ? column.inferredType : "";
  if (["amount", "measure", "formula"].includes(role) || ["currency", "number", "formula"].includes(inferredType)) {
    return false;
  }
  if (columnHasFormula(options.formulas, position)) {
    return false;
  }
  const values = matrix.map((row) => row[position]).filter((value) => value !== undefined && value !== null && value !== "");
  if (values.length < 6 || values.some((value) => typeof value !== "string" && typeof value !== "boolean")) {
    return false;
  }
  if (["date", "status", "category", "vendor", "account", "description", "dimension", "identifier", "note"].includes(role)) {
    return true;
  }
  if (["date", "status", "text", "unknown"].includes(inferredType)) {
    return true;
  }
  const formats = options.numberFormat ? orderedColumnDomain(options.numberFormat, position).map(String).join(" ") : "";
  return /date|text|general|@/i.test(formats);
}

function columnHasFormula(formulas: CellMatrix | undefined, position: number): boolean {
  return Boolean(formulas?.some((row) => typeof row[position] === "string" && String(row[position]).startsWith("=")));
}

function orderedColumnDomain(matrix: CellMatrix, position: number): unknown[] {
  const seen = new Map<string, unknown>();
  for (const row of matrix) {
    const value = row[position];
    if (value === undefined || value === null || value === "") {
      continue;
    }
    const key = domainKey(value);
    if (!seen.has(key)) {
      seen.set(key, value);
    }
  }
  return [...seen.values()];
}

function shouldEncodeDomain(matrix: CellMatrix, position: number, domain: unknown[]): boolean {
  const nonEmptyCount = matrix.reduce((count, row) => {
    const value = row[position];
    return value === undefined || value === null || value === "" ? count : count + 1;
  }, 0);
  if (domain.length < 2 || domain.length > 32 || nonEmptyCount < 6) {
    return false;
  }
  if (domain.length / nonEmptyCount > 0.7) {
    return false;
  }
  return domain.some((value) => typeof value === "string" && value.length >= 4);
}

function domainKey(value: unknown): string {
  return `${typeof value}:${String(value)}`;
}

function roleImportance(role: string): number {
  switch (role) {
    case "description":
    case "amount":
    case "date":
    case "category":
    case "status":
    case "vendor":
      return 0.9;
    case "account":
    case "measure":
    case "formula":
      return 0.78;
    case "identifier":
    case "note":
      return 0.68;
    default:
      return 0.45;
  }
}

function compactComparisonAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const sheets = Array.isArray(answer.sheets)
    ? answer.sheets.map((sheet) => {
        const typed = sheet as Record<string, unknown>;
        return stripUndefinedRecord({
          sheetName: typed.sheetName,
          range: typed.range,
          shape: typed.shape,
          metrics: compactProfileMetrics(typed.metrics)
        });
      })
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    request: answer.request,
    sheets,
    numericComparison: answer.numericComparison,
    resultUri,
    fullResultUri
  });
}

function compactSemanticIndexAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const entries = Array.isArray(answer.entries)
    ? answer.entries.slice(0, 8).map((entry) => {
        const typed = entry as Record<string, unknown>;
        return stripUndefinedRecord({
          id: typed.id,
          label: typed.label,
          role: typed.role,
          sourceKind: typed.sourceKind,
          sheetName: typed.sheetName,
          tableName: typed.tableName,
          range: typed.range,
          confidence: typed.confidence,
          evidence: Array.isArray(typed.evidence) ? typed.evidence.slice(0, 4) : undefined,
          nextRequestHints: Array.isArray(typed.nextRequestHints) ? typed.nextRequestHints.slice(0, 3) : undefined
        });
      })
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    workbook: answer.workbook,
    detailLevel: answer.detailLevel,
    entryCount: answer.entryCount,
    entries,
    resultUri,
    fullResultUri
  });
}

function compactSimilarRowsAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const rows = Array.isArray(answer.rows)
    ? answer.rows.slice(0, 10).map((row) => {
        const typed = row as Record<string, unknown>;
        return stripUndefinedRecord({
          sheetName: typed.sheetName,
          range: typed.range,
          sheetRowNumber: typed.sheetRowNumber,
          values: typed.values,
          columns: Array.isArray(typed.columns) ? typed.columns.slice(0, 12) : undefined,
          score: typed.score,
          matchedSignals: typed.matchedSignals,
          matchedColumns: Array.isArray(typed.matchedColumns) ? typed.matchedColumns.slice(0, 8) : undefined,
          whyMatched: typed.whyMatched
        });
      })
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    sourceMode: answer.sourceMode,
    signals: Array.isArray(answer.signals) ? answer.signals.slice(0, 12) : undefined,
    predicates: Array.isArray(answer.predicates) ? answer.predicates.slice(0, 10) : undefined,
    comparedRanges: Array.isArray(answer.comparedRanges) ? answer.comparedRanges.slice(0, 6) : undefined,
    rows,
    resultUri,
    fullResultUri
  });
}

function compactStyleReferenceCandidatesAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const candidates = Array.isArray(answer.candidates)
    ? answer.candidates.slice(0, 8).map((candidate) => {
        const typed = candidate as Record<string, unknown>;
        return stripUndefinedRecord({
          sheetName: typed.sheetName,
          range: typed.range,
          label: typed.label,
          sourceKind: typed.sourceKind,
          confidence: typed.confidence,
          reason: typed.reason,
          styleSummary: typed.styleSummary,
          nextAction: typed.nextAction
        });
      })
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    candidates,
    resultUri,
    fullResultUri
  });
}

function compactWorkbookSummaryAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    workbook: answer.workbook,
    sheetCount: answer.sheetCount,
    tableCount: answer.tableCount,
    namedRangeCount: answer.namedRangeCount,
    semanticIndex: answer.semanticIndex && typeof answer.semanticIndex === "object" ? compactSemanticIndexAnswer(answer.semanticIndex as Record<string, unknown>) : undefined,
    sheets: Array.isArray(answer.sheets) ? answer.sheets.slice(0, 20) : undefined,
    resultUri,
    fullResultUri
  });
}

function compactSheetSummaryAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    sheet: answer.sheet,
    contextHints: answer.contextHints,
    tables: Array.isArray(answer.tables) ? answer.tables.slice(0, 8) : undefined,
    namedRanges: Array.isArray(answer.namedRanges) ? answer.namedRanges.slice(0, 8) : undefined,
    sections: Array.isArray(answer.sections) ? answer.sections.slice(0, 8) : undefined,
    summaryBlocks: Array.isArray(answer.summaryBlocks) ? answer.summaryBlocks.slice(0, 8) : undefined,
    formulaRegions: Array.isArray(answer.formulaRegions) ? answer.formulaRegions.slice(0, 8) : undefined,
    resultUri,
    fullResultUri
  });
}

function compactWorkbookOverviewAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const semanticIndex = answer.semanticIndex && typeof answer.semanticIndex === "object"
    ? compactSemanticIndexAnswer(answer.semanticIndex as Record<string, unknown>)
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    workbook: answer.workbook,
    selection: answer.selection,
    sheetCount: answer.sheetCount,
    tableCount: answer.tableCount,
    namedRangeCount: answer.namedRangeCount,
    sectionCount: answer.sectionCount,
    semanticIndex,
    sheets: answer.sheets,
    tables: answer.tables,
    namedRanges: answer.namedRanges,
    blankSheets: answer.blankSheets,
    resultUri,
    fullResultUri
  });
}

function compactStyleSummaryAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  return stripUndefinedRecord({
    kind: answer.kind,
    sheetName: answer.sheetName,
    range: answer.range,
    rowCount: answer.rowCount,
    columnCount: answer.columnCount,
    truncated: answer.truncated,
    fills: compactStyleDimension(answer.fills),
    fonts: compactStyleDimension(answer.fonts),
    borders: compactStyleDimension(answer.borders),
    alignment: compactStyleDimension(answer.alignment),
    numberFormats: compactStyleDimension(answer.numberFormats),
    conditionalFormatting: compactStyleDimension(answer.conditionalFormatting),
    rowHeights: compactStyleDimension(answer.rowHeights),
    columnWidths: compactStyleDimension(answer.columnWidths),
    dataValidation: compactStyleDimension(answer.dataValidation),
    resultUri,
    fullResultUri
  });
}

function compactStyleOverviewAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const groupedHeaderSuggestion = answer.groupedHeaderSuggestion && typeof answer.groupedHeaderSuggestion === "object"
    ? answer.groupedHeaderSuggestion as Record<string, unknown>
    : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    sheetName: answer.sheetName,
    range: answer.range,
    freezePanes: answer.freezePanes,
    table: answer.table,
    detected: answer.detected,
    currentStyle: answer.currentStyle,
    columnRoles: Array.isArray(answer.columnRoles) ? answer.columnRoles.slice(0, 16) : undefined,
    columnGroupSuggestions: Array.isArray(answer.columnGroupSuggestions) ? answer.columnGroupSuggestions.slice(0, 8) : undefined,
    groupedHeaderSuggestion: groupedHeaderSuggestion ? stripUndefinedRecord({
      kind: groupedHeaderSuggestion.kind,
      targetHeaderRange: groupedHeaderSuggestion.targetHeaderRange,
      levels: groupedHeaderSuggestion.levels,
      insertAboveHeader: groupedHeaderSuggestion.insertAboveHeader,
      styleExistingHeader: groupedHeaderSuggestion.styleExistingHeader,
      requiresStructuralPreview: groupedHeaderSuggestion.requiresStructuralPreview,
      defaultApplyBehavior: groupedHeaderSuggestion.defaultApplyBehavior,
      groups: Array.isArray(groupedHeaderSuggestion.groups) ? groupedHeaderSuggestion.groups.slice(0, 8) : undefined,
      operationsNeeded: groupedHeaderSuggestion.operationsNeeded
    }) : undefined,
    recommendations: Array.isArray(answer.recommendations) ? answer.recommendations.slice(0, 8) : undefined,
    recommendedWorkflow: answer.recommendedWorkflow,
    resultUri,
    fullResultUri
  });
}

function compactWorkbookDesignOverviewAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    workbook: answer.workbook,
    sheet: answer.sheet,
    table: answer.table,
    target: answer.target,
    dataState: answer.dataState,
    inspectionPolicy: answer.inspectionPolicy,
    relatedSheets: Array.isArray(answer.relatedSheets) ? answer.relatedSheets.slice(0, 6) : undefined,
    columnRecommendations: Array.isArray(answer.columnRecommendations) ? answer.columnRecommendations.slice(0, 40) : undefined,
    groupSuggestions: Array.isArray(answer.groupSuggestions) ? answer.groupSuggestions.slice(0, 8) : undefined,
    summary: answer.summary,
    nextWorkflows: Array.isArray(answer.nextWorkflows) ? answer.nextWorkflows.slice(0, 5) : undefined,
    resultUri,
    fullResultUri
  });
}

function compactDataValidationSummaryAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  return stripUndefinedRecord({
    kind: answer.kind,
    source: answer.source,
    method: answer.method,
    sheetName: answer.sheetName,
    range: answer.range,
    ruleCount: answer.ruleCount,
    type: answer.type,
    inCellDropDown: answer.inCellDropDown,
    sourceFormula: answer.sourceFormula,
    options: Array.isArray(answer.options) ? answer.options.slice(0, 100) : undefined,
    optionCount: answer.optionCount,
    sourceComplete: answer.sourceComplete,
    sourceRange: answer.sourceRange,
    validationRangeStatus: answer.validationRangeStatus,
    rules: Array.isArray(answer.rules) ? answer.rules.slice(0, 8) : undefined,
    fieldContext: Array.isArray(answer.fieldContext) ? answer.fieldContext.slice(0, 12) : undefined,
    guidance: answer.guidance,
    resultUri,
    fullResultUri
  });
}

function compactStyleDimension(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return value;
  }
  const full = JSON.stringify(value);
  if (full.length <= 800) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.slice(0, 5);
  }
  const record = value as Record<string, unknown>;
  const compact: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(record)) {
    if (entry === undefined) {
      continue;
    }
    if (entry === null || typeof entry !== "object") {
      compact[key] = entry;
    } else if (Array.isArray(entry)) {
      compact[key] = entry.slice(0, 5);
      if (entry.length > 5) {
        compact[`${key}Truncated`] = true;
      }
    } else if (["hash", "summary", "dominant", "default", "uniqueCount", "count", "truncated"].includes(key)) {
      compact[key] = entry;
    }
  }
  return Object.keys(compact).length > 0 ? compact : { truncated: true, fullBytes: Buffer.byteLength(full) };
}

function compactVisualReadabilityPreviewAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const visualPlan = answer.visualPlan && typeof answer.visualPlan === "object" ? answer.visualPlan as Record<string, unknown> : undefined;
  const detected = answer.detected && typeof answer.detected === "object" ? answer.detected as Record<string, unknown> : undefined;
  const rules = Array.isArray(visualPlan?.rules) ? visualPlan.rules as Array<Record<string, unknown>> : [];
  const ruleIds = Array.isArray(visualPlan?.ruleIds)
    ? visualPlan.ruleIds.filter((id) => typeof id === "string").slice(0, 32)
    : rules.slice(0, 32).map((rule) => rule.id).filter((id) => typeof id === "string");
  const groupedHeaderSuggestion = answer.groupedHeaderSuggestion && typeof answer.groupedHeaderSuggestion === "object" ? answer.groupedHeaderSuggestion as Record<string, unknown> : undefined;
  return stripUndefinedRecord({
    kind: answer.kind,
    action: answer.action,
    sheetName: answer.sheetName,
    range: answer.range,
    defaults: answer.defaults,
    sheetType: answer.sheetType,
    detected: detected ? stripUndefinedRecord({
      sheetName: detected.sheetName,
      usedRange: detected.usedRange,
      headerRow: detected.headerRow,
      headerRange: detected.headerRange,
      dataRange: detected.dataRange,
      tableRanges: detected.tableRanges,
      hasFilter: detected.hasFilter,
      hasFreezePane: detected.hasFreezePane,
      formulaColumns: detected.formulaColumns,
      existingStyleRanges: detected.existingStyleRanges,
      detectionSource: detected.detectionSource,
      confidence: detected.confidence
    }) : undefined,
    columnRoles: Array.isArray(answer.columnRoles) ? answer.columnRoles.slice(0, 12) : undefined,
    groupedHeaderSuggestion: groupedHeaderSuggestion ? stripUndefinedRecord({
      kind: groupedHeaderSuggestion.kind,
      targetHeaderRange: groupedHeaderSuggestion.targetHeaderRange,
      requiresStructuralPreview: groupedHeaderSuggestion.requiresStructuralPreview,
      defaultApplyBehavior: groupedHeaderSuggestion.defaultApplyBehavior,
      groups: Array.isArray(groupedHeaderSuggestion.groups) ? groupedHeaderSuggestion.groups.slice(0, 8) : undefined,
      operationsNeeded: groupedHeaderSuggestion.operationsNeeded
    }) : undefined,
    visualPlan: visualPlan ? stripUndefinedRecord({
      compilerStatus: visualPlan.compilerStatus,
      summary: visualPlan.summary,
      counts: visualPlan.counts,
      ruleScopes: visualPlan.ruleScopes,
      ruleIds,
      validationSuggestions: Array.isArray(visualPlan.validationSuggestions) ? visualPlan.validationSuggestions.slice(0, 6) : undefined,
      formulaSuggestions: Array.isArray(visualPlan.formulaSuggestions) ? visualPlan.formulaSuggestions.slice(0, 6) : undefined,
      referenceStyleSuggestions: Array.isArray(visualPlan.referenceStyleSuggestions) ? visualPlan.referenceStyleSuggestions.slice(0, 6) : undefined,
      printSuggestions: Array.isArray(visualPlan.printSuggestions) ? visualPlan.printSuggestions.slice(0, 6) : undefined,
      previewExamples: Array.isArray(visualPlan.previewExamples) ? visualPlan.previewExamples.slice(0, 5) : undefined,
      theme: visualPlan.theme,
      operationId: visualPlan.operationId,
      operationCount: visualPlan.operationCount,
      skipped: Array.isArray(visualPlan.skipped) ? visualPlan.skipped.slice(0, 8) : undefined,
      risk: visualPlan.risk,
      preservation: visualPlan.preservation
    }) : undefined,
    resultUri,
    fullResultUri
  });
}

function compactGenericAnswer(answer: Record<string, unknown>, resultUri?: string, fullResultUri?: string): Record<string, unknown> {
  const next = { ...answer };
  for (const key of ["headers", "values", "formulas", "text", "numberFormat", "sample", "sparseRows", "rows", "emptySummary", "schema", "profile"]) {
    delete next[key];
  }
  if (next.semanticIndex && typeof next.semanticIndex === "object") {
    next.semanticIndex = compactSemanticIndexAnswer(next.semanticIndex as Record<string, unknown>);
  }
  if (resultUri) {
    next.resultUri = resultUri;
  }
  if (fullResultUri) {
    next.fullResultUri = fullResultUri;
  }
  return stripUndefinedRecord(next);
}

function continuationForOutput(
  output: Omit<AgentRunOutput, "telemetry">,
  responseMode: AgentResponseMode,
  stored?: StoredAgentResult
): AgentRunOutput["continuation"] | undefined {
  const nextRequest = nextContinuationRequest(output, stored);
  const continuation = stripUndefinedRecord({
    ...(output.workbookContextId !== undefined ? { workbookContextId: output.workbookContextId } : {}),
    ...(output.operationId !== undefined ? { operationId: output.operationId } : {}),
    ...(output.transactionId !== undefined ? { transactionId: output.transactionId } : {}),
    ...(stored?.resourceUri !== undefined ? { resultUri: stored.resourceUri } : {}),
    ...(stored?.fullResourceUri !== undefined ? { fullResultUri: stored.fullResourceUri } : {}),
    ...(stored?.freshness !== undefined ? { freshness: stored.freshness } : {}),
    responseMode: responseMode === "verbose" ? "brief" : responseMode,
    ...(nextRequest !== undefined ? { nextRequest } : {})
  });
  return Object.keys(continuation).length > 0 ? continuation : undefined;
}

function nextContinuationRequest(output: Omit<AgentRunOutput, "telemetry">, stored?: StoredAgentResult): string | undefined {
  if (output.operationId && output.confirmationToken) {
    return "Continue with mode apply_update, operationId, and confirmationToken after user confirmation.";
  }
  if (stored?.resourceUri) {
    return "Reuse workbookContextId. For exact rows, raw values, audit detail, or transformation input, call excel.agent.run once with continuation.fullResultUri. excel:// handles are not web URLs; never use Webfetch/browser.";
  }
  if (output.workbookContextId) {
    return "Reuse workbookContextId on the next excel.agent.run call.";
  }
  return undefined;
}

function compactContinuationForBudget(
  continuation: AgentRunOutput["continuation"] | undefined,
  omitContextOnly = false
): AgentRunOutput["continuation"] | undefined {
  if (!continuation) {
    return undefined;
  }
  if (omitContextOnly && !continuation.operationId && !continuation.resultUri && !continuation.fullResultUri) {
    return undefined;
  }
  const compact = stripUndefinedRecord({
    ...(continuation.operationId !== undefined ? { operationId: continuation.operationId } : {}),
    ...(continuation.resultUri !== undefined ? { resultUri: continuation.resultUri } : {}),
    ...(continuation.fullResultUri !== undefined ? { fullResultUri: continuation.fullResultUri } : {})
  });
  return Object.keys(compact).length > 0 ? compact : undefined;
}

function schemaSummary(columns: unknown[], maxColumns: number): Record<string, unknown> {
  const compactColumns = columns
    .filter((column): column is Record<string, unknown> => Boolean(column && typeof column === "object"))
    .slice(0, maxColumns)
    .map((column) => stripUndefinedRecord({
      name: column.name,
      letter: column.letter,
      inferredType: column.inferredType,
      role: column.role,
      importance: column.importance
    }));
  return {
    columnCount: columns.length,
    columns: compactColumns,
    ...(columns.length > compactColumns.length ? { truncated: true } : {})
  };
}

function compactProfileMetrics(metrics: unknown): unknown {
  if (!metrics || typeof metrics !== "object") {
    return metrics;
  }
  const typed = metrics as Record<string, unknown>;
  return stripUndefinedRecord({
    nonEmptyCount: typed.nonEmptyCount,
    numericCount: typed.numericCount,
    measureColumns: typed.measureColumns
  });
}

function stripUndefinedRecord<T extends Record<string, unknown>>(value: T): T {
  const next = { ...value };
  for (const key of Object.keys(next)) {
    if (next[key] === undefined) {
      delete next[key];
    }
  }
  return next;
}

function compactCandidate(candidate: AgentCandidate): AgentCandidate {
  const next = { ...candidate };
  delete next.reason;
  delete next.nextRequestHint;
  delete next.aliases;
  return next;
}

function outputUsedCallerTargetHint(output: Omit<AgentRunOutput, "telemetry">): boolean {
  const hinted = output.candidates?.filter((candidate) => candidate.reason?.includes("caller target hint")) ?? [];
  if (hinted.length === 0) {
    return false;
  }
  if (output.proof.length === 0) {
    return output.status === "SUCCESS" && output.nextAction === "answer_now" && hinted[0]?.id === output.candidates?.[0]?.id;
  }
  return hinted.some((candidate) => output.proof.some((proof) =>
    proof.sheetName === candidate.sheetName &&
    (!proof.range || !candidate.range || proof.range === candidate.range)
  ));
}

function outputUsedSemanticCandidate(output: Omit<AgentRunOutput, "telemetry">): boolean {
  const semantic = output.candidates?.filter((candidate) => candidate.semanticRole) ?? [];
  if (semantic.length === 0) {
    return false;
  }
  if (output.proof.length === 0) {
    return output.status === "SUCCESS" && output.nextAction === "answer_now" && semantic[0]?.id === output.candidates?.[0]?.id;
  }
  return semantic.some((candidate) => output.proof.some((proof) =>
    proof.sheetName === candidate.sheetName &&
    (!proof.range || !candidate.range || proof.range === candidate.range)
  ));
}

function withActionHandlerMetric(output: Omit<AgentRunOutput, "telemetry">, actionHandlerId: AgentActionHandlerId): Omit<AgentRunOutput, "telemetry"> {
  return {
    ...output,
    metrics: {
      ...(output.metrics ?? {}),
      actionHandlerId
    }
  };
}

function stripUndefinedOptionals(output: Omit<AgentRunOutput, "telemetry">): Omit<AgentRunOutput, "telemetry"> {
  const next = { ...output };
  if (next.answer === undefined) delete next.answer;
  if (next.metrics === undefined) delete next.metrics;
  if (next.changes === undefined) delete next.changes;
  if (next.candidates === undefined) delete next.candidates;
  if (next.continuation === undefined) delete next.continuation;
  if (next.confirmationToken === undefined) delete next.confirmationToken;
  if (next.operationId === undefined) delete next.operationId;
  if (next.workbookContextId === undefined) delete next.workbookContextId;
  if (next.finalAnswer === undefined) delete next.finalAnswer;
  if (next.agentInstruction === undefined) delete next.agentInstruction;
  if (next.maxRecommendedFollowupCalls === undefined) delete next.maxRecommendedFollowupCalls;
  if (next.requiredFollowup === undefined) delete next.requiredFollowup;
  return next;
}

function withTaskOutcomeContract(output: Omit<AgentRunOutput, "telemetry">): Omit<AgentRunOutput, "telemetry"> {
  const existing = output.taskOutcome;
  const taskOutcome = existing ?? deriveTaskOutcome(output);
  const finalAnswer = output.finalAnswer ?? deriveFinalAnswer(output, taskOutcome);
  const agentInstruction = output.agentInstruction ?? deriveAgentInstruction(output, taskOutcome);
  const maxRecommendedFollowupCalls = output.maxRecommendedFollowupCalls ?? recommendedFollowupCount(output, taskOutcome);
  const requiredFollowup = output.requiredFollowup ?? deriveRequiredFollowup(output, taskOutcome);
  return stripUndefinedOptionals({
    ...output,
    taskOutcome,
    agentInstruction,
    maxRecommendedFollowupCalls,
    ...(finalAnswer !== undefined ? { finalAnswer } : {}),
    ...(requiredFollowup !== undefined ? { requiredFollowup } : {})
  });
}

function deriveTaskOutcome(output: Omit<AgentRunOutput, "telemetry">): NonNullable<AgentRunOutput["taskOutcome"]> {
  if (output.status === "PREVIEW_READY" || output.nextAction === "call_apply_update") {
    return "preview_ready";
  }
  if (output.status === "SUCCESS" && output.mode === "apply_update") {
    return "apply_complete";
  }
  if (output.nextAction === "ask_user" || output.nextAction === "call_with_target") {
    return "needs_user_input";
  }
  if (output.status === "SUCCESS" && output.nextAction === "answer_now") {
    return "final_answer";
  }
  return "cannot_complete";
}

function deriveFinalAnswer(output: Omit<AgentRunOutput, "telemetry">, taskOutcome: NonNullable<AgentRunOutput["taskOutcome"]>): string | undefined {
  if (taskOutcome === "preview_ready") {
    return `${output.summary} Review the preview, then call apply_update with the returned operationId and confirmationToken if the user confirms.`;
  }
  if (taskOutcome === "needs_user_input" || taskOutcome === "cannot_complete" || taskOutcome === "apply_complete" || taskOutcome === "final_answer") {
    return conciseFinalAnswer(output) ?? output.summary;
  }
  return undefined;
}

function conciseFinalAnswer(output: Omit<AgentRunOutput, "telemetry">): string | undefined {
  const answer = output.answer && typeof output.answer === "object" ? output.answer as Record<string, unknown> : undefined;
  if (!answer) {
    return undefined;
  }
  if (answer.kind === "sheet_summary") {
    const sheet = answer.sheet && typeof answer.sheet === "object" ? answer.sheet as Record<string, unknown> : undefined;
    const sheetName = typeof sheet?.name === "string" ? sheet.name : "the requested sheet";
    const usedRange = typeof sheet?.usedRange === "string" ? ` used range ${sheet.usedRange}` : "";
    const tables = Array.isArray(answer.tables) ? answer.tables : [];
    const sections = Array.isArray(answer.sections) ? answer.sections : [];
    const tableSummary = tables.length > 0
      ? ` Tables: ${tables.slice(0, 4).map((table) => {
          const typed = table && typeof table === "object" ? table as Record<string, unknown> : {};
          const name = typeof typed.name === "string" ? typed.name : "unnamed";
          const range = typeof typed.range === "string" ? ` ${typed.range}` : "";
          const columns = Array.isArray(typed.columns)
            ? ` (${typed.columns.slice(0, 6).map((column) => {
                const record = column && typeof column === "object" ? column as Record<string, unknown> : {};
                return String(record.name ?? "").trim();
              }).filter(Boolean).join(", ")}${typed.columns.length > 6 ? ", ..." : ""})`
            : "";
          return `${name}${range}${columns}`;
        }).join("; ")}.`
      : " No Excel tables were found on this sheet.";
    const sectionSummary = sections.length > 0
      ? ` Sections: ${sections.slice(0, 4).map((section) => {
          const typed = section && typeof section === "object" ? section as Record<string, unknown> : {};
          const label = String(typed.label ?? typed.id ?? "section").trim();
          const range = typeof typed.range === "string" ? ` ${typed.range}` : "";
          const headerRange = typeof typed.headerRange === "string" ? ` header ${typed.headerRange}` : "";
          const dataRange = typeof typed.dataRange === "string" ? ` data ${typed.dataRange}` : "";
          return `${label}${range}${headerRange}${dataRange}`.trim();
        }).join("; ")}. Use section/header/row anchors for edits before raw coordinates.`
      : "";
    return `${sheetName}${usedRange}.${tableSummary}${sectionSummary}`;
  }
  if (answer.kind === "table_compact_read") {
    const tableName = stringValue(answer.tableName) ?? "the requested table";
    const range = stringValue(answer.range);
    const shape = answer.shape && typeof answer.shape === "object" ? answer.shape as Record<string, unknown> : undefined;
    const rows = typeof shape?.rows === "number" ? shape.rows : undefined;
    const columns = typeof shape?.columns === "number" ? shape.columns : undefined;
    const schemaSummaryAnswer = answer.schemaSummary && typeof answer.schemaSummary === "object" ? answer.schemaSummary as Record<string, unknown> : undefined;
    const schemaColumns = Array.isArray(schemaSummaryAnswer?.columns) ? schemaSummaryAnswer.columns : [];
    const headers = schemaColumns.length > 0
      ? schemaColumns.slice(0, 8).map((column) => {
          const record = column && typeof column === "object" ? column as Record<string, unknown> : {};
          return String(record.name ?? "").trim();
        }).filter(Boolean)
      : Array.isArray(answer.projectedColumns)
        ? answer.projectedColumns.slice(0, 8).map((column) => {
            const record = column && typeof column === "object" ? column as Record<string, unknown> : {};
            return String(record.name ?? "").trim();
          }).filter(Boolean)
        : [];
    const previewRows = Array.isArray(answer.valuesPreview) ? answer.valuesPreview.length : undefined;
    return `Read ${tableName}${range ? ` at ${range}` : ""}${rows && columns ? ` (${rows} rows x ${columns} columns in this page)` : ""}.${headers.length > 0 ? ` Columns include: ${headers.join(", ")}.` : ""}${previewRows ? ` ${previewRows} exact preview rows are inline; answer now unless the user asked for more rows.` : ""}`;
  }
  if (answer.kind === "data_validation_summary") {
    const sheetName = stringValue(answer.sheetName) ?? "the requested sheet";
    const range = stringValue(answer.range) ?? "the requested range";
    if (answer.validationRangeStatus === "mixed_or_inconsistent_range") {
      return `Read dropdown validation for ${sheetName}!${range}. The range has mixed/inconsistent validation, so this result is inconclusive for dropdown options; do not say the dropdown is broken or missing options from this range summary. Inspect one representative data cell with the dropdown, or update the known source-list cell directly if the user asked to add an option.`;
    }
    const optionCount = typeof answer.optionCount === "number" ? answer.optionCount : 0;
    const options = Array.isArray(answer.options)
      ? answer.options.slice(0, 12).map((option) => String(option)).filter(Boolean)
      : [];
    const sourceRange = stringValue(answer.sourceRange);
    return `Read dropdown validation for ${sheetName}!${range}. ${optionCount} option${optionCount === 1 ? "" : "s"} found${options.length > 0 ? `: ${options.join(", ")}${optionCount > options.length ? ", ..." : ""}` : ""}.${sourceRange ? ` Source list: ${sourceRange}.` : ""} Answer now; do not fetch fullResultUri for dropdown options.`;
  }
  if (answer.kind === "workbook_summary" || answer.kind === "workbook_overview") {
    const sheetCount = typeof answer.sheetCount === "number" ? answer.sheetCount : undefined;
    const tableCount = typeof answer.tableCount === "number" ? answer.tableCount : undefined;
    const sheets = Array.isArray(answer.sheets)
      ? answer.sheets.slice(0, 6).map((sheet) => {
          const record = sheet && typeof sheet === "object" ? sheet as Record<string, unknown> : {};
          const name = String(record.name ?? "").trim();
          const usedRange = typeof record.usedRange === "string" ? ` ${record.usedRange}` : "";
          return `${name}${usedRange}`.trim();
        }).filter(Boolean)
      : [];
    return `Workbook summary: ${sheetCount ?? "multiple"} sheets, ${tableCount ?? 0} tables.${sheets.length > 0 ? ` Sheets include: ${sheets.join("; ")}.` : ""}`;
  }
  return undefined;
}

function deriveAgentInstruction(output: Omit<AgentRunOutput, "telemetry">, taskOutcome: NonNullable<AgentRunOutput["taskOutcome"]>): string {
  if (taskOutcome === "preview_ready") {
    return "Do not rediscover workbook context. Ask for user confirmation if needed, then call excel.agent.run with mode apply_update, operationId, and confirmationToken.";
  }
  if (taskOutcome === "needs_user_input") {
    return "Ask the user for the missing information from finalAnswer; do not call workbook tools again until the user responds.";
  }
  if (taskOutcome === "cannot_complete") {
    if (output.nextAction === "fetch_resource") {
      return "Call excel.agent.run once with the returned resultUri/fullResultUri in request or continuation. Do not use Webfetch/browser for excel:// handles.";
    }
    return "Report finalAnswer and warnings to the user; do not loop through more workbook discovery unless the user changes the request.";
  }
  if (taskOutcome === "apply_complete") {
    return "Answer the user now from finalAnswer, proof, compactProof, and warnings; do not call workbook tools again for this task.";
  }
  return "Answer the user now from finalAnswer, proof, and inline structuredContent; do not call workbook tools again for this task.";
}

function recommendedFollowupCount(output: Omit<AgentRunOutput, "telemetry">, taskOutcome: NonNullable<AgentRunOutput["taskOutcome"]>): number {
  if (taskOutcome === "preview_ready") {
    return 1;
  }
  if (output.nextAction === "fetch_resource" || output.nextAction === "retry_after_refresh" || output.nextAction === "call_preview_update") {
    return 1;
  }
  return 0;
}

function deriveRequiredFollowup(
  output: Omit<AgentRunOutput, "telemetry">,
  taskOutcome: NonNullable<AgentRunOutput["taskOutcome"]>
): AgentRunOutput["requiredFollowup"] | undefined {
  if (taskOutcome === "preview_ready" && output.operationId && output.confirmationToken) {
    return {
      mode: "apply_update",
      nextAction: "call_apply_update",
      operationId: output.operationId,
      confirmationToken: output.confirmationToken,
      instruction: "Call excel.agent.run once with mode apply_update, operationId, and confirmationToken after user confirmation."
    };
  }
  if (output.nextAction === "call_preview_update") {
    return {
      mode: "preview_update",
      nextAction: "call_preview_update",
      instruction: "Call excel.agent.run with mode preview_update using the grouped workflow suggested in finalAnswer."
    };
  }
  if (output.nextAction === "fetch_resource") {
    return {
      mode: "answer",
      nextAction: "fetch_resource",
      instruction: "Call excel.agent.run once with continuation.fullResultUri or the returned excel:// handle in request. Never use Webfetch/browser for excel:// handles."
    };
  }
  return undefined;
}

function isMatchUpdatePreviewAnswer(answer: unknown): answer is Record<string, unknown> & { kind: "match_update_preview" } {
  return typeof answer === "object" && answer !== null && (answer as { kind?: unknown }).kind === "match_update_preview";
}

function resolveUpdateTarget(metadata: WorkbookMetadata, input: AgentRunInput):
  | Extract<AgentTargetResolution, { ok: true }>
  | { ok: false; status: AgentRunOutput["status"]; summary: string; candidates?: AgentCandidate[]; nextAction: AgentRunOutput["nextAction"]; warnings: string[] } {
  if (input.values === undefined) {
    const candidates = findAgentCandidates(metadata, input).slice(0, 5);
    return {
      ok: false,
      status: "NEEDS_INPUT",
      summary: "Preview needs structured values; rows embedded in request text are not used for safe writes.",
      ...(candidates.length > 0 ? { candidates } : {}),
      nextAction: "ask_user",
      warnings: ["Direct cell/range updates must be supplied as values.patches. Specialized workflows such as table appends may use their dedicated structured values."]
    };
  }
  const resolved = resolveAgentUpdateTarget(metadata, input);
  if (resolved.ok && !input.target?.range && (
    resolved.candidate.kind === "sheet" ||
    (resolved.candidate.kind === "table" && !isTableAppendIntent(input.request)) ||
    resolved.candidate.kind === "region" ||
    (resolved.candidate.kind === "range" && !input.target?.candidateId)
  )) {
    return {
      ok: false,
      status: "AMBIGUOUS_TARGET",
      summary: "Preview needs a narrower range, row, column, named range, or registered region before updating a sheet/table target.",
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      nextAction: "call_with_target",
      warnings: ["Sheet/table-wide natural-language updates are not previewed without a narrower target."]
    };
  }
  return resolved;
}

function validateDirectPatchInput(input: AgentRunInput): { ok: true } | { ok: false; summary: string; warnings: string[] } {
  const patches = input.values?.patches;
  if (!Array.isArray(patches)) {
    return { ok: true };
  }
  for (const [index, patch] of patches.entries()) {
    if (!patch || typeof patch !== "object") {
      return {
        ok: false,
        summary: `Patch ${index + 1} is not a valid object.`,
        warnings: ["Each direct update patch must include target and one supported payload field."]
      };
    }
    if (!patch.target) {
      return {
        ok: false,
        summary: `Patch ${index + 1} is missing target.`,
        warnings: ["Each direct update patch must include patch.target with sheetName/range or an equivalent target."]
      };
    }
    const typed = patch as Record<string, unknown>;
    const payloadKeys = ["values", "rows", "formulas", "style", "numberFormat", "numberFormats", "validation", "options", "allowedValues", "conditionalFormatting", "note", "comment"];
    if (!payloadKeys.some((key) => typed[key] !== undefined)) {
      return {
        ok: false,
        summary: `Patch ${index + 1} is missing a mutation payload.`,
        warnings: ["Each direct update patch must include values, rows, formulas, style, numberFormat, validation, conditionalFormatting, note, or comment."]
      };
    }
    const hasValuePayload = typed.values !== undefined || typed.rows !== undefined;
    const nonValuePayloadCount = ["formulas", "style", "validation", "options", "allowedValues", "conditionalFormatting", "note", "comment"]
      .filter((key) => typed[key] !== undefined).length;
    if (hasValuePayload && nonValuePayloadCount > 0) {
      return {
        ok: false,
        summary: `Patch ${index + 1} mixes value payloads with another mutation family.`,
        warnings: ["Split values/formulas/styles/validation/comments into separate patches so preview can compile the correct operation family."]
      };
    }
  }
  return { ok: true };
}

function usesLegacyDirectValueWrite(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  if (!Array.isArray(values?.values) || Array.isArray(values?.patches)) {
    return false;
  }
  const action = intentAction(input);
  if (action && action !== "write_values") {
    return false;
  }
  if (isTableAppendIntent(input.request)) {
    return false;
  }
  return Boolean(input.target?.range || action === "write_values" || /\b(update|write|set|change|edit)\b/i.test(input.request));
}

type CanonicalDirectMutationAction =
  | "write_formulas"
  | "write_number_formats"
  | "format_range"
  | "write_data_validation"
  | "write_conditional_formatting";

function legacyDirectMutationPatchIssue(input: AgentRunInput): { summary: string; warnings: string[] } | undefined {
  const action = intentAction(input);
  if (!isCanonicalDirectMutationAction(action)) {
    return undefined;
  }
  const values = input.values as Record<string, unknown> | undefined;
  if (!values || Array.isArray(values.patches)) {
    return undefined;
  }
  const payloadKeys = canonicalDirectMutationPayloadKeys(action);
  if (!payloadKeys.some((key) => values[key] !== undefined)) {
    return undefined;
  }
  const label = canonicalDirectMutationLabel(action);
  return {
    summary: `${label} updates must use values.patches. Top-level structured ${label} payloads are no longer accepted for agent-facing direct mutations.`,
    warnings: [`Use values.patches[].${payloadKeys[0]} with patch.target for ${label} updates.`]
  };
}

function isCanonicalDirectMutationAction(action: string | undefined): action is CanonicalDirectMutationAction {
  return action === "write_formulas"
    || action === "write_number_formats"
    || action === "format_range"
    || action === "write_data_validation"
    || action === "write_conditional_formatting";
}

function canonicalDirectMutationPayloadKeys(action: CanonicalDirectMutationAction): string[] {
  switch (action) {
    case "write_formulas":
      return ["formulas", "formula", "values", "rows"];
    case "write_number_formats":
      return ["numberFormats", "numberFormat", "formats"];
    case "format_range":
      return ["style", "fillColor", "fontColor", "fontBold", "fontItalic", "horizontalAlignment", "verticalAlignment", "borders"];
    case "write_data_validation":
      return ["validation", "source", "options", "allowedValues"];
    case "write_conditional_formatting":
      return ["conditionalFormatting", "rule", "formula", "style"];
  }
}

function canonicalDirectMutationLabel(action: CanonicalDirectMutationAction): string {
  switch (action) {
    case "write_formulas":
      return "formula";
    case "write_number_formats":
      return "number-format";
    case "format_range":
      return "style";
    case "write_data_validation":
      return "data-validation";
    case "write_conditional_formatting":
      return "conditional-formatting";
  }
}

function canonicalDirectMutationInput(input: AgentRunInput, action: CanonicalDirectMutationAction): AgentRunInput {
  const patches = input.values?.patches;
  if (!Array.isArray(patches) || patches.length === 0) {
    return input;
  }
  const keys = canonicalDirectMutationPayloadKeys(action);
  const patch = patches.find((candidate) => {
    if (!candidate || typeof candidate !== "object") {
      return false;
    }
    const typed = candidate as Record<string, unknown>;
    return keys.some((key) => typed[key] !== undefined);
  });
  if (!patch || !patch.target) {
    return input;
  }
  const target = inheritPatchTargetSheet(input, patch.target);
  const patchRecord = patch as Record<string, unknown>;
  const nextValues: Record<string, unknown> = { ...patchRecord };
  delete nextValues.target;
  if (action === "write_conditional_formatting" && nextValues.conditionalFormatting && nextValues.rule === undefined) {
    nextValues.rule = nextValues.conditionalFormatting;
  }
  return {
    ...input,
    target: { ...input.target, ...target },
    values: {
      ...nextValues,
      patches
    }
  };
}

function canonicalPatchRequiredOutput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  summary: string,
  warnings: string[]
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    answer: {
      kind: "canonical_patch_required",
      code: "CANONICAL_PATCH_REQUIRED",
      example: {
        mode: "preview_update",
        intent: { action: "write_values" },
        values: {
          patches: [
            {
              target: { sheetName: "Sales", range: "E2" },
              values: [["Reviewed"]]
            }
          ]
        }
      }
    },
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: ["CANONICAL_PATCH_REQUIRED", ...warnings]
  };
}

function safetyArtifactNeedsInput(metadata: WorkbookMetadata, requestedMode: AgentRunMode, summary: string): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: []
  };
}

function workbookLevelNeedsInput(metadata: WorkbookMetadata, requestedMode: AgentRunMode, summary: string): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: []
  };
}

function validationNeedsInput(metadata: WorkbookMetadata, requestedMode: AgentRunMode, summary: string): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: []
  };
}

function isValidationIntentAction(action: AgentIntentAction | undefined): action is Extract<AgentIntentAction, `validate_${string}`> {
  return typeof action === "string" && action.startsWith("validate_");
}

function isFormulaReadIntentAction(
  action: AgentIntentAction | undefined
): action is "read_formulas" | "read_formula_patterns" | "get_formula_dependency_graph" | "trace_formula_precedents" | "trace_formula_dependents" | "find_formula_errors" | "explain_formula" {
  return action === "read_formulas"
    || action === "read_formula_patterns"
    || action === "get_formula_dependency_graph"
    || action === "trace_formula_precedents"
    || action === "trace_formula_dependents"
    || action === "find_formula_errors"
    || action === "explain_formula";
}

function validationRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput) {
  const values = input.values as Record<string, unknown> | undefined;
  const sheetName = stringValue(input.target?.sheetName ?? values?.sheetName);
  const tableName = stringValue(input.target?.tableName ?? values?.tableName);
  const address = stringValue(input.target?.range ?? values?.address ?? values?.range);
  const templateId = stringValue(values?.templateId);
  const targetSheetName = stringValue(values?.targetSheetName ?? sheetName);
  const pair = snapshotPairFromInput(input);
  return {
    ...(sheetName ? { sheetName } : {}),
    ...(tableName ? { tableName } : {}),
    ...(address ? { address: normalizeOperationRange(metadata, sheetName ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", address) } : {}),
    ...(templateId ? { templateId: templateId as TemplateId } : {}),
    ...(targetSheetName ? { targetSheetName } : {}),
    ...(snapshotIdFromInput(input) ? { snapshotId: snapshotIdFromInput(input) } : {}),
    ...(pair ? { leftSnapshotId: pair.left, rightSnapshotId: pair.right } : {})
  };
}

function validationSummary(action: AgentIntentAction, result: unknown): string {
  const issues = Array.isArray((result as { issues?: unknown }).issues) ? (result as { issues: unknown[] }).issues.length : 0;
  const label = action.replace(/^validate_/, "").replaceAll("_", " ");
  return issues > 0 ? `Validated ${label}; found ${issues} issue(s).` : `Validated ${label}; no issues found.`;
}

function compactValidationReport(report: unknown) {
  const issues = Array.isArray((report as { issues?: unknown }).issues) ? (report as { issues: Array<Record<string, unknown>> }).issues : [];
  const severityCounts = countBy(issues, (issue) => typeof issue.severity === "string" ? issue.severity : "unknown");
  const categoryCounts = countBy(issues, (issue) => typeof issue.category === "string" ? issue.category : "unknown");
  return {
    ok: (report as { ok?: unknown }).ok !== false,
    issueCount: issues.length,
    severityCounts,
    categoryCounts,
    examples: issues.slice(0, 5),
    fullReportTruncated: issues.length > 5
  };
}

function countBy<T>(values: T[], keyFor: (value: T) => string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const value of values) {
    const key = keyFor(value);
    counts[key] = (counts[key] ?? 0) + 1;
  }
  return counts;
}

function validationIssueMessages(result: unknown): string[] {
  const issues = (result as { issues?: unknown }).issues;
  return Array.isArray(issues)
    ? issues.slice(0, 5).flatMap((issue) => typeof (issue as { message?: unknown }).message === "string" ? [(issue as { message: string }).message] : [])
    : [];
}

function explainFormulaString(formula: string) {
  const normalized = formula.trim().startsWith("=") ? formula.trim() : `=${formula.trim()}`;
  const functions = [...normalized.matchAll(/\b([A-Z][A-Z0-9_.]*)\s*\(/gi)].map((match) => match[1]!.toUpperCase());
  const references = [...normalized.matchAll(/(?:'[^']+'|[A-Z_][A-Z0-9_ ]*!|\b)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/gi)].map((match) => match[0]);
  return {
    ok: true,
    formula: normalized,
    summary: {
      functions: [...new Set(functions)],
      references: [...new Set(references)],
      hasExternalReference: /\[[^\]]+\]/.test(normalized),
      hasStructuredReference: /\[[#@\w ,:[\]]+\]/.test(normalized),
      hasVolatileFunction: functions.some((fn) => ["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"].includes(fn))
    }
  };
}

function safetyArtifactAction(
  kind: "snapshot.refresh" | "snapshot.invalidate" | "snapshot.delete" | "backup.pin" | "backup.unpin" | "backup.delete",
  id: string
) {
  switch (kind) {
    case "snapshot.refresh":
      return { kind, snapshotId: id as SnapshotId };
    case "snapshot.invalidate":
      return { kind, snapshotId: id as SnapshotId };
    case "snapshot.delete":
      return { kind, snapshotId: id as SnapshotId };
    case "backup.pin":
      return { kind, backupId: id as BackupId };
    case "backup.unpin":
      return { kind, backupId: id as BackupId };
    case "backup.delete":
      return { kind, backupId: id as BackupId };
  }
}

function snapshotIdFromInput(input: AgentRunInput): string | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  return stringValue(input.target?.entity)
    ?? stringValue(values?.snapshotId)
    ?? stringValue(values?.id)
    ?? stringValue(input.operationId);
}

function snapshotPairFromInput(input: AgentRunInput): { left: string; right: string } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const left = stringValue(values?.leftSnapshotId ?? values?.baseSnapshotId ?? values?.beforeSnapshotId);
  const right = stringValue(values?.rightSnapshotId ?? values?.currentSnapshotId ?? values?.afterSnapshotId);
  return left && right ? { left, right } : undefined;
}

function backupIdFromInput(input: AgentRunInput): string | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  return stringValue(input.target?.entity)
    ?? stringValue(values?.backupId)
    ?? stringValue(values?.id)
    ?? stringValue(input.operationId);
}

function fileBackupModeFromInput(value: unknown): WorkbookCreateFileBackupRequest["mode"] | undefined {
  const mode = stringValue(value);
  return mode === "export-copy" || mode === "save-copy-as" ? mode : undefined;
}

function fileRestoreModeFromInput(value: unknown): NonNullable<WorkbookRestoreFileBackupRequest["mode"]> {
  const mode = stringValue(value);
  return mode === "replace-open-workbook" || mode === "restore-into-open-workbook" ? mode : "open-as-new";
}

function workbookLocalConfigOptionsFromInput(input: AgentRunInput): { includePermissions?: boolean } {
  const values = input.values as Record<string, unknown> | undefined;
  return typeof values?.includePermissions === "boolean" ? { includePermissions: values.includePermissions } : {};
}

function permissionUpdateFromInput(input: AgentRunInput, workbookId: WorkbookId): Partial<PermissionState> {
  const values = input.values ?? {};
  const nested = typeof values.permissions === "object" && values.permissions !== null && !Array.isArray(values.permissions)
    ? values.permissions as Record<string, unknown>
    : {};
  const valueFor = (key: keyof PermissionState | "scopeToWorkbook") => nested[key] ?? values[key];
  const update: Partial<PermissionState> = {};
  const allowWrites = booleanValue(valueFor("allowWrites"));
  if (allowWrites !== undefined) update.allowWrites = allowWrites;
  const allowDestructiveActions = booleanValue(valueFor("allowDestructiveActions"));
  if (allowDestructiveActions !== undefined) update.allowDestructiveActions = allowDestructiveActions;
  const allowWorkbookActions = booleanValue(valueFor("allowWorkbookActions"));
  if (allowWorkbookActions !== undefined) update.allowWorkbookActions = allowWorkbookActions;
  const allowMacroExecution = booleanValue(valueFor("allowMacroExecution"));
  if (allowMacroExecution !== undefined) update.allowMacroExecution = allowMacroExecution;
  const requireConfirmationFor = destructiveLevelsFromUnknown(valueFor("requireConfirmationFor"));
  if (requireConfirmationFor) update.requireConfirmationFor = requireConfirmationFor;
  const scope = permissionScopeFromUnknown(valueFor("scope"));
  if (scope) {
    update.scope = scope;
  } else if (booleanValue(valueFor("scopeToWorkbook")) === true) {
    update.scope = { workbookId };
  }
  return update;
}

function destructiveLevelsFromUnknown(value: unknown): PermissionState["requireConfirmationFor"] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }
  const levels = value.filter((item): item is PermissionState["requireConfirmationFor"][number] =>
    item === "none" || item === "values" || item === "format" || item === "structure" || item === "workbook"
  );
  return [...new Set(levels)];
}

function permissionScopeFromUnknown(value: unknown): PermissionState["scope"] | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  const scope: PermissionState["scope"] = {};
  const scopedWorkbookId = stringValue(record.workbookId);
  if (scopedWorkbookId) scope.workbookId = scopedWorkbookId as WorkbookId;
  if (Array.isArray(record.sheetNames) && record.sheetNames.every((item) => typeof item === "string")) {
    scope.sheetNames = [...record.sheetNames];
  }
  if (Array.isArray(record.regionNames) && record.regionNames.every((item) => typeof item === "string")) {
    scope.regionNames = [...record.regionNames];
  }
  return Object.keys(scope).length > 0 ? scope : undefined;
}

function workbookLocalConfigImportRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): WorkbookLocalConfigImportRequest | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const config = values?.config;
  if (!config || typeof config !== "object") {
    return undefined;
  }
  const request: WorkbookLocalConfigImportRequest = {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    config: config as WorkbookLocalConfigImportRequest["config"]
  };
  for (const key of ["includeTemplates", "includeRegions", "includePermissions", "overwrite"] as const) {
    if (typeof values?.[key] === "boolean") {
      request[key] = values[key];
    }
  }
  return request;
}

function workbookEmbeddedLocalConfigImportRequestFromInput(
  workbookId: WorkbookId,
  input: AgentRunInput
): { workbookId: WorkbookId; includeTemplates?: boolean; includeRegions?: boolean; includePermissions?: boolean; overwrite?: boolean } {
  const values = input.values as Record<string, unknown> | undefined;
  const request: { workbookId: WorkbookId; includeTemplates?: boolean; includeRegions?: boolean; includePermissions?: boolean; overwrite?: boolean } = { workbookId };
  for (const key of ["includeTemplates", "includeRegions", "includePermissions", "overwrite"] as const) {
    if (typeof values?.[key] === "boolean") {
      request[key] = values[key];
    }
  }
  return request;
}

function workbookCloseBehaviorFromInput(input: AgentRunInput): "Save" | "SkipSave" | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const behavior = stringValue(values?.closeBehavior ?? values?.behavior);
  return behavior === "Save" || behavior === "SkipSave" ? behavior : undefined;
}

function backupRetentionRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput, dryRun: boolean): WorkbookBackupRetentionRequest {
  const values = input.values as Record<string, unknown> | undefined;
  const kind = backupRetentionKindFromInput(values?.kind);
  const maxAgeDays = positiveNumber(values?.maxAgeDays);
  const maxBackupsPerWorkbook = nonNegativeInteger(values?.maxBackupsPerWorkbook);
  const maxTotalBytes = nonNegativeInteger(values?.maxTotalBytes);
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    ...(kind ? { kind } : {}),
    ...(maxAgeDays !== undefined ? { maxAgeDays } : {}),
    ...(maxBackupsPerWorkbook !== undefined ? { maxBackupsPerWorkbook } : {}),
    ...(maxTotalBytes !== undefined ? { maxTotalBytes } : {}),
    dryRun
  };
}

function backupRetentionKindFromInput(value: unknown): WorkbookBackupRetentionRequest["kind"] | undefined {
  const kind = stringValue(value);
  return kind === "file-copy" || kind === "snapshot-json" || kind === "all" ? kind : undefined;
}

function sheetClearApplyToFromInput(value: unknown): "all" | "contents" | "formats" {
  const applyTo = stringValue(value);
  return applyTo === "contents" || applyTo === "formats" ? applyTo : "all";
}

function sheetProtectionOptionsFromInput(values: Record<string, unknown> | undefined): Extract<ExcelOperation, { kind: "sheet.protect" }>["options"] | undefined {
  const rawOptions = values?.options && typeof values.options === "object" ? values.options as Record<string, unknown> : {};
  const valueFor = (...keys: string[]) => {
    for (const key of keys) {
      if (rawOptions[key] !== undefined) return rawOptions[key];
      if (values?.[key] !== undefined) return values[key];
    }
    return undefined;
  };
  const options: NonNullable<Extract<ExcelOperation, { kind: "sheet.protect" }>["options"]> = {};
  const setBoolean = (field: Exclude<keyof typeof options, "selectionMode">, ...keys: string[]) => {
    const value = booleanValue(valueFor(...keys));
    if (value !== undefined) {
      options[field] = value;
    }
  };
  setBoolean("allowFormatCells", "allowFormatCells", "allowFormat", "formatCells");
  setBoolean("allowFormatColumns", "allowFormatColumns", "formatColumns");
  setBoolean("allowFormatRows", "allowFormatRows", "formatRows");
  setBoolean("allowInsertColumns", "allowInsertColumns", "insertColumns");
  setBoolean("allowInsertRows", "allowInsertRows", "insertRows");
  setBoolean("allowDeleteColumns", "allowDeleteColumns", "deleteColumns");
  setBoolean("allowDeleteRows", "allowDeleteRows", "deleteRows");
  setBoolean("allowSort", "allowSort", "sort");
  setBoolean("allowAutoFilter", "allowAutoFilter", "allowFilter", "filter", "autoFilter");
  setBoolean("allowPivotTables", "allowPivotTables", "pivotTables");
  setBoolean("protectDrawingObjects", "protectDrawingObjects", "objects");
  setBoolean("protectScenarios", "protectScenarios", "scenarios");
  const selectionMode = stringValue(valueFor("selectionMode"));
  if (selectionMode === "normal" || selectionMode === "unlocked" || selectionMode === "none") {
    options.selectionMode = selectionMode;
  }
  return Object.keys(options).length > 0 ? options : undefined;
}

function positiveNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : undefined;
}

function nonNegativeInteger(value: unknown): number | undefined {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : undefined;
}

function compactSnapshotListResult(result: unknown) {
  const snapshots = Array.isArray((result as { snapshots?: unknown }).snapshots)
    ? (result as { snapshots: unknown[] }).snapshots.map(compactSnapshot)
    : [];
  return {
    ...(typeof result === "object" && result !== null && "ok" in result ? { ok: (result as { ok?: unknown }).ok } : {}),
    snapshots
  };
}

function compactSnapshot(snapshot: unknown) {
  const record = snapshot as {
    snapshotId?: unknown;
    workbookId?: unknown;
    createdAt?: unknown;
    reason?: unknown;
    affectedRanges?: unknown;
    invalidatedAt?: unknown;
    payload?: { rangeSnapshots?: unknown[] };
  } | undefined;
  const affectedRanges = Array.isArray(record?.affectedRanges)
    ? record.affectedRanges.filter(isA1RangeLike)
    : [];
  const rangeSnapshots = Array.isArray(record?.payload?.rangeSnapshots) ? record.payload.rangeSnapshots : [];
  return {
    snapshotId: record?.snapshotId,
    workbookId: record?.workbookId,
    createdAt: record?.createdAt,
    reason: record?.reason,
    affectedRanges,
    ...(record?.invalidatedAt !== undefined ? { invalidatedAt: record.invalidatedAt } : {}),
    payloadSummary: {
      rangeSnapshotCount: rangeSnapshots.length,
      cellCount: rangeSnapshots.reduce<number>((total, item) => {
        const count = (item as { fingerprint?: { cellCount?: unknown } })?.fingerprint?.cellCount;
        return total + (typeof count === "number" && Number.isFinite(count) ? count : 0);
      }, 0)
    }
  };
}

function isA1RangeLike(value: unknown): value is A1Range {
  return typeof value === "object"
    && value !== null
    && typeof (value as { sheetName?: unknown }).sheetName === "string"
    && typeof (value as { address?: unknown }).address === "string";
}

function intentAction(input: AgentRunInput): AgentIntentAction | undefined {
  const action = (input.intent as { action?: unknown } | undefined)?.action;
  return isAgentIntentAction(action) ? action : undefined;
}

function answerIntent(input: AgentRunInput): "schema" | "values" {
  const action = intentAction(input);
  if (action === "read_schema") {
    return "schema";
  }
  if (action === "read_values") {
    return "values";
  }
  const request = input.request;
  const valueIntent = /\b(actual\s+values?|values?|rows?|records?|data|sample|examples?|first\s+\d+|last\s+\d+|preview|contents?)\b/i.test(request)
    || /[A-Z]{1,3}\d+\s*:\s*[A-Z]{1,3}\d+/i.test(request)
    || /\b(selection|highlighted|active cell|current cell|this cell|this range|this column|selected (?:cell|range|col(?:umn)?)|active column|current column)\b/i.test(request);
  if (valueIntent) {
    return "values";
  }
  return /\b(schema|columns?|headers?|fields?|structure)\b/i.test(request) ? "schema" : "values";
}

function isTableAppendIntent(request: string): boolean {
  return /\b(append|add|insert|fill|mock|seed|generate)\b/i.test(request)
    && /\b(rows?|records?|logs?|data|transactions?|table)\b/i.test(request);
}

function resolveSchemaTable(metadata: WorkbookMetadata, input: AgentRunInput, resolved: Extract<AgentTargetResolution, { ok: true }>): TableMetadata | undefined {
  const targetTableName = input.target?.tableName;
  if (targetTableName) {
    const normalized = normalizeAgentLookup(targetTableName);
    const exact = metadata.tables.filter((table) => normalizeAgentLookup(table.name ?? table.id) === normalized);
    if (input.target?.sheetName) {
      const normalizedSheet = normalizeAgentLookup(input.target.sheetName);
      return exact.find((table) => normalizeAgentLookup(table.sheetName) === normalizedSheet);
    }
    if (exact.length === 1) {
      return exact[0];
    }
  }
  if (resolved.candidate.kind === "table") {
    return metadata.tables.find((table) => table.id === resolved.candidate.id)
      ?? metadata.tables.find((table) => table.sheetName === resolved.sheetName && table.range === resolved.range);
  }
  return metadata.tables.find((table) => table.sheetName === resolved.sheetName && table.range === resolved.range);
}

function resolveSchemaHeaders(metadata: WorkbookMetadata, resolved: Extract<AgentTargetResolution, { ok: true }>): HeaderMetadata[] {
  const sheet = metadata.sheets.find((candidate) => candidate.name === resolved.sheetName);
  if (!sheet) {
    return [];
  }
  const matching = sheet.headers.filter((header) => normalizeAddressForCompare(header.range) === normalizeAddressForCompare(resolved.range));
  if (matching.length > 0) {
    return matching;
  }
  const overlapping = sheet.headers
    .filter((header) => rangesOverlapAddresses(normalizeAddressForCompare(header.range), normalizeAddressForCompare(resolved.range)))
    .sort((left, right) => right.confidence - left.confidence || left.row - right.row);
  if (overlapping.length > 0) {
    return overlapping.slice(0, 3);
  }
  return sheet.headers
    .filter((header) => header.confidence >= 0.75 && header.row <= 3)
    .sort((left, right) => right.confidence - left.confidence || left.row - right.row)
    .slice(0, 3);
}

function normalizeAddressForCompare(address: string): string {
  return stripSheetName(address).replace(/\$/g, "");
}

function headerAnswer(header: HeaderMetadata) {
  return {
    sheetName: header.sheetName,
    row: header.row,
    range: header.range,
    confidence: header.confidence,
    columns: header.columns.map((column) => ({
      name: column.name,
      normalizedName: column.normalizedName,
      inferredType: column.inferredType,
      role: column.role,
      importance: column.importance,
      index: column.index,
      letter: column.letter
    }))
  };
}

function normalizeAgentLookup(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, "_").replace(/[^\w]/g, "");
}

type AdvancedMutationKind = "formula" | "template" | "style" | "other";

function advancedMutationDecision(input: AgentRunInput): { kind: AdvancedMutationKind; safetyDecision: string; summary: string; warning: string } | undefined {
  const request = input.request.toLowerCase();
  if (isSmallExplicitValueWrite(input) && !hasAdvancedMutationKeyword(request)) {
    return undefined;
  }
  const kind: AdvancedMutationKind = /\b(formula|formulas|calculate|calculation|total\s+row)\b/.test(request)
    ? "formula"
    : /\b(template|duplicate|copy)\b/.test(request)
      ? "template"
      : /\b(style|format|formatting|header\s+row)\b/.test(request)
        ? "style"
        : "other";
  if (kind !== "other" || /\b(repair|pivot|chart)\b/.test(request)) {
    return {
      kind,
      safetyDecision: "manual_review:advanced_workflow",
      summary: "This request needs a formula/template/style/report-aware workflow, not a generic value update. Use a dedicated preview or advanced workflow path.",
      warning: "Auto mode blocked advanced mutation intent so formulas, templates, styles, pivots, or charts are not modified as plain values."
    };
  }
  return undefined;
}

function hasAdvancedMutationKeyword(request: string): boolean {
  return /\b(formula|formulas|calculate|calculation|total\s+row|template|duplicate|copy|style|format|formatting|header\s+row|repair|pivot|chart)\b/.test(request);
}

function autoApplyDecision(input: AgentRunInput, preview: Omit<AgentRunOutput, "telemetry">): { allow: true; safetyDecision: string } | { allow: false; safetyDecision: string; reason: string } {
  if (preview.status !== "PREVIEW_READY") {
    return { allow: false, safetyDecision: "manual_review:preview_not_ready", reason: "preview did not produce an apply-ready operation" };
  }
  const action = intentAction(input);
  if (action === "transform_values" || action === "derive_values") {
    return { allow: false, safetyDecision: "manual_review:broad_compiled_transform", reason: "backend-compiled broad transforms and derivations require preview before apply" };
  }
  if (isMatchUpdatePreviewAnswer(preview.answer)) {
    if ((preview.changes?.length ?? 0) === 0 || (preview.changes?.length ?? 0) > 16) {
      return { allow: false, safetyDecision: "manual_review:match_update_change_count", reason: "the match/update change count is not safe for auto-apply" };
    }
    return { allow: true, safetyDecision: "auto_apply:match_update" };
  }
  if (!input.values) {
    return { allow: false, safetyDecision: "manual_review:missing_values", reason: "values were not provided explicitly" };
  }
  if (!hasExplicitAutoApplyTarget(input)) {
    return { allow: false, safetyDecision: "manual_review:implicit_range", reason: "the target range was not explicit enough for auto-apply" };
  }
  if (isRiskyMutationRequest(input.request)) {
    return { allow: false, safetyDecision: "manual_review:risky_request", reason: "the request may be structural, destructive, or formula-sensitive" };
  }
  const cellCount = autoApplyValueCellCount(input.values);
  if (cellCount <= 0) {
    return { allow: false, safetyDecision: "manual_review:missing_values", reason: "values were not provided explicitly" };
  }
  if (cellCount > 16) {
    return { allow: false, safetyDecision: "manual_review:large_scope", reason: "the edit touches more than 16 cells" };
  }
  if (containsFormulaLikeAutoApplyValue(input.values)) {
    return { allow: false, safetyDecision: "manual_review:formula_values", reason: "formula writes require a formula-aware workflow" };
  }
  if ((preview.changes?.length ?? 0) === 0 || (preview.changes?.length ?? 0) > 16) {
    return { allow: false, safetyDecision: "manual_review:change_count", reason: "the previewed change count is not safe for auto-apply" };
  }
  const unsafeChange = unsafeAutoApplyChangeReason(input, preview.changes ?? []);
  if (unsafeChange) {
    return { allow: false, safetyDecision: "manual_review:target_looks_like_header_or_clear", reason: unsafeChange };
  }
  return { allow: true, safetyDecision: "auto_apply:scoped_value_edit" };
}

function isSmallExplicitValueWrite(input: AgentRunInput): boolean {
  if (!input.values || !hasExplicitAutoApplyTarget(input) || containsFormulaLikeAutoApplyValue(input.values)) {
    return false;
  }
  const cellCount = autoApplyValueCellCount(input.values);
  return cellCount > 0 && cellCount <= 16;
}

function hasExplicitAutoApplyTarget(input: AgentRunInput): boolean {
  if (input.target?.range) {
    return true;
  }
  const patches = input.values?.patches;
  if (Array.isArray(patches) && patches.length > 0 && patches.every((patch) => Boolean(patch?.target?.range))) {
    return true;
  }
  return semanticValuePatchesFromInput(input).length > 0;
}

function autoApplyValueCellCount(values: AgentRunInput["values"]): number {
  if (!values) {
    return 0;
  }
  const semanticPatches = semanticValuePatchesFromInput({ request: "", values });
  if (semanticPatches.length > 0) {
    return semanticPatches.length;
  }
  if (Array.isArray(values.patches)) {
    return values.patches.reduce((total, patch) => {
      const matrix = Array.isArray(patch.values)
        ? patch.values
        : Array.isArray(patch.rows)
          ? patch.rows
          : [];
      return total + matrixCellCount(matrix as CellMatrix);
    }, 0);
  }
  return matrixCellCount(objectToCellMatrix(values));
}

function containsFormulaLikeAutoApplyValue(values: AgentRunInput["values"]): boolean {
  if (!values) {
    return false;
  }
  const semanticPatches = semanticValuePatchesFromInput({ request: "", values });
  if (semanticPatches.some((patch) => typeof patch.value === "string" && patch.value.trim().startsWith("="))) {
    return true;
  }
  if (Array.isArray(values.patches)) {
    return values.patches.some((patch) => {
      const matrix = Array.isArray(patch.values)
        ? patch.values
        : Array.isArray(patch.rows)
          ? patch.rows
          : [];
      return containsFormulaLikeValue(matrix as CellMatrix);
    });
  }
  return containsFormulaLikeValue(objectToCellMatrix(values));
}

function isRiskyMutationRequest(request: string): boolean {
  return /\b(delete|remove|clear|wipe|reset|drop|resize|rename|move|copy|convert|formula|formulas|template|pivot|chart|style|format|merge|unmerge|sort|filter|append|insert)\b/i.test(request);
}

function containsFormulaLikeValue(values: CellMatrix): boolean {
  return values.flat().some((value) => typeof value === "string" && value.trim().startsWith("="));
}

function hasFormulaWriteInput(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  return Boolean(values && (Array.isArray(values.formulas) || typeof values.formula === "string" || intentAction(input) === "write_formulas"));
}

function isConditionalFormattingMutationRequest(request: string): boolean {
  return /\b(conditional\s+format|conditional\s+formatting|formula\s+(format|color|colour|fill)|if\b.*\b(fill|color|colour|format))\b/i.test(request);
}

function unsafeAutoApplyChangeReason(input: AgentRunInput, changes: NonNullable<AgentRunOutput["changes"]>): string | undefined {
  const request = input.request.toLowerCase();
  const explicitlyHeaderEdit = /\b(header|heading|title|label|reference|caption)\b/i.test(input.request);
  const explicitlyClearing = /\b(clear|remove|blank|empty|delete|wipe)\b/i.test(input.request);
  for (const change of changes) {
    const before = (change as { before?: unknown }).before;
    const after = (change as { after?: unknown }).after;
    if (isBlankishAutoApplyValue(after) && !explicitlyClearing) {
      return "blank/null writes require an explicit clear request";
    }
    if (!explicitlyHeaderEdit && looksLikeProtectedLabelValue(before) && looksLikeHeaderOverwriteValue(after)) {
      const cell = typeof (change as { cell?: unknown }).cell === "string" ? ` ${(change as { cell: string }).cell}` : "";
      return `target${cell} appears to contain a header/title/reference label`;
    }
    if (!explicitlyHeaderEdit && looksLikeProtectedLabelValue(before) && typeof after === "string" && request.includes(String(before).toLowerCase())) {
      return "target appears to overwrite a referenced label/header";
    }
  }
  return undefined;
}

function isBlankishAutoApplyValue(value: unknown): boolean {
  return value === null || value === undefined || value === "";
}

function looksLikeHeaderOverwriteValue(value: unknown): boolean {
  return value === null || value === undefined || value === "" || typeof value === "number" || typeof value === "boolean";
}

function looksLikeProtectedLabelValue(value: unknown): boolean {
  if (typeof value !== "string") {
    return false;
  }
  const text = value.replace(/\s+/g, " ").trim();
  if (text.length < 3) {
    return false;
  }
  if (/\b(header|title|reference|diesel|rate|vendor\s+propose|truck\s+available|available|thb\/trip|trip|route|origin|destination|item\s+no|transport\s+mode)\b/i.test(text)) {
    return true;
  }
  if (text.includes("\n") && /[A-Za-z]/.test(text)) {
    return true;
  }
  return false;
}

function isFormulaMutationRequest(request: string): boolean {
  return /\b(formula|formulas|calculate|calculation|total\s+row)\b/i.test(request);
}

function hasCellValues(values: CellMatrix): boolean {
  return values.some((row) => row.some((value) => value !== null && value !== undefined && value !== ""));
}

function broadMutationNeedsScopeOutput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  requestedMode: AgentRunMode,
  resolved: Extract<AgentTargetResolution, { ok: true }>,
  matrix: CellMatrix
): Omit<AgentRunOutput, "telemetry"> | undefined {
  const targetCells = cellCountFromAddress(resolved.range) ?? 0;
  const writeCells = matrixCellCount(matrix);
  const backendSideAction = ["transform_values", "derive_values", "settle_reconciliation", "grouped_header"].includes(intentAction(input) ?? "");
  if (!isPotentialBroadMutationRequest(input) || backendSideAction || targetCells < 500 || writeCells >= targetCells) {
    return undefined;
  }
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `The requested update may affect ${targetCells.toLocaleString()} cells in ${resolved.sheetName}!${resolved.range}. Ask whether to apply it to all rows, matching rows, the selected range, or a specified range before building a large payload.`,
    answer: {
      kind: "broad_mutation_scope_guard",
      sheetName: resolved.sheetName,
      range: resolved.range,
      estimatedCellCount: targetCells,
      providedCellCount: writeCells,
      alternatives: ["all rows", "matching rows only", "selected range", "specified range"]
    },
    candidates: findAgentCandidates(metadata, input).slice(0, 5),
    proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "broad update target" }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: ["Broad update scope is unclear. Do not generate or apply a large write matrix until the user confirms scope."]
  };
}

function broadMutationExplicitTargetNeedsScopeOutput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  requestedMode: AgentRunMode
): Omit<AgentRunOutput, "telemetry"> | undefined {
  const sheetName = input.target?.sheetName;
  const range = input.target?.range;
  const targetCells = typeof range === "string" ? cellCountFromAddress(range) : undefined;
  const backendSideAction = ["transform_values", "derive_values", "settle_reconciliation", "grouped_header"].includes(intentAction(input) ?? "");
  if (!isPotentialBroadMutationRequest(input) || backendSideAction || !sheetName || !range || targetCells === undefined || targetCells < 500) {
    return undefined;
  }
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: `The requested update may affect ${targetCells.toLocaleString()} cells in ${sheetName}!${range}. Ask whether to apply it to all rows, matching rows, the selected range, or a specified range before building a large payload.`,
    answer: {
      kind: "broad_mutation_scope_guard",
      sheetName,
      range,
      estimatedCellCount: targetCells,
      providedCellCount: 0,
      alternatives: ["all rows", "matching rows only", "selected range", "specified range"]
    },
    candidates: findAgentCandidates(metadata, input).slice(0, 5),
    proof: [{ sheetName, range, label: "broad update target" }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings: ["Broad update scope is unclear. Do not generate or apply a large write matrix until the user confirms scope."]
  };
}

function isPotentialBroadMutationRequest(input: AgentRunInput): boolean {
  return /\b(?:all|whole|entire|every|these|matching|rows?|sheet|worksheet)\b/i.test(input.request)
    && /\b(?:color|colour|fill|highlight|format|update|fix|change|set|replace)\b/i.test(input.request);
}

function normalizeShortYearDatesForWrite(workbookId: WorkbookId, input: AgentRunInput, range: string, matrix: CellMatrix): {
  matrix: CellMatrix;
  numberFormatEntries: Array<{ target: A1Range; numberFormat: string[][]; preserveValues?: true }>;
  warnings: string[];
} {
  if (!/\b(?:date|booking|bookings|วันที่|วัน\/เวลา|เรือออก|คืนตู้|รับตู้|ปิดสาย)\b/i.test(input.request)) {
    return { matrix, numberFormatEntries: [], warnings: [] };
  }
  const parsedRange = tryParseA1Address(stripSheetName(range));
  if (!parsedRange) {
    return { matrix, numberFormatEntries: [], warnings: [] };
  }
  const sheetName = input.target?.sheetName;
  if (!sheetName) {
    return { matrix, numberFormatEntries: [], warnings: [] };
  }
  let changed = false;
  const entries: Array<{ target: A1Range; numberFormat: string[][]; preserveValues?: true }> = [];
  const normalized = matrix.map((row, rowIndex) => row.map((value, columnIndex) => {
    const date = parseShortYearDate(value);
    if (!date) {
      return value;
    }
    changed = true;
    entries.push({
      target: {
        workbookId,
        sheetName,
        address: `${numberToColumn(parsedRange.startColumn + columnIndex)}${parsedRange.startRow + rowIndex}`
      },
      numberFormat: [["dd/mm/yyyy"]],
      preserveValues: true
    });
    return excelSerialFromDate(date.year, date.month, date.day);
  }));
  return {
    matrix: normalized,
    numberFormatEntries: entries,
    warnings: changed ? ["Short-year date text was normalized to Excel date serials and formatted as dd/mm/yyyy."] : []
  };
}

function parseShortYearDate(value: unknown): { year: number; month: number; day: number } | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const match = /^\s*(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2})\s*$/.exec(value);
  if (!match?.[1] || !match[2] || !match[3]) {
    return undefined;
  }
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = 2000 + Number(match[3]);
  if (day < 1 || day > 31 || month < 1 || month > 12) {
    return undefined;
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    return undefined;
  }
  return { year, month, day };
}

function excelSerialFromDate(year: number, month: number, day: number): number {
  const excelEpoch = Date.UTC(1899, 11, 30);
  const date = Date.UTC(year, month - 1, day);
  return Math.round((date - excelEpoch) / 86_400_000);
}

function firstNonEmptyRows(values: CellMatrix, limit: number): CellMatrix {
  const nonEmpty = values.filter((row) => row.some((value) => value !== null && value !== undefined && value !== ""));
  return (nonEmpty.length > 0 ? nonEmpty : values).slice(0, limit);
}

function objectToCellMatrix(values: Record<string, unknown>): CellMatrix {
  if (Array.isArray(values.formulas)) return values.formulas as CellMatrix;
  if (typeof values.formula === "string") return [[values.formula]];
  if (Array.isArray(values.rows)) return values.rows as CellMatrix;
  if (Array.isArray(values.values)) return values.values as CellMatrix;
  return [Object.values(values) as CellMatrix[number]];
}

type FormulaMatrixNormalization = {
  ok: true;
  formulas: CellMatrix;
  dimensions: { rows: number; columns: number };
  formulaPattern?: string;
  warnings: string[];
} | {
  ok: false;
  summary: string;
  warnings: string[];
  suggestedValues?: Record<string, unknown>;
};

function formulaMatrixFromInput(input: AgentRunInput, address: string, fallbackMatrix: CellMatrix): FormulaMatrixNormalization {
  const values = input.values as Record<string, unknown> | undefined;
  const rawMatrix = matrixFromUnknown(values?.formulas)
    ?? (typeof values?.formula === "string" ? [[values.formula]] : undefined)
    ?? matrixFromUnknown(values?.values)
    ?? matrixFromUnknown(values?.rows)
    ?? fallbackMatrix;
  const formulas = normalizeFormulaOnlyMatrix(rawMatrix);
  const dimensions = dimensionsFromAddress(address) ?? {
    rows: Math.max(1, formulas.length),
    columns: Math.max(1, maxMatrixColumns(formulas))
  };
  const formulaCells = formulas.flat().filter(formulaLike);
  if (formulaCells.length === 0) {
    return {
      ok: false,
      summary: "Formula updates need values.formula, values.formulas, or a formula matrix in values.values.",
      warnings: ["Formula cells must be strings that start with '='."],
      suggestedValues: { formula: "=FORMULA_HERE" }
    };
  }
  const matrixRows = formulas.length;
  const matrixColumns = Math.max(1, maxMatrixColumns(formulas));
  if (matrixRows === dimensions.rows && matrixColumns === dimensions.columns) {
    const formulaPattern = formulaCells.length === 1 ? String(formulaCells[0]) : undefined;
    return {
      ok: true,
      formulas: padFormulaMatrix(formulas, dimensions.rows, dimensions.columns),
      dimensions,
      ...(formulaPattern !== undefined ? { formulaPattern } : {}),
      warnings: []
    };
  }
  if (formulaCells.length === 1 && matrixRows === 1 && matrixColumns === 1) {
    const formula = String(formulaCells[0]);
    const expanded = expandSingleFormulaAcrossRange(formula, dimensions);
    if (expanded.ok) {
      return {
        ok: true,
        formulas: expanded.formulas,
        dimensions,
        formulaPattern: formula,
        warnings: dimensions.rows * dimensions.columns > 1 ? [`Expanded one formula across ${dimensions.rows * dimensions.columns} cell(s) using relative A1 references.`] : []
      };
    }
    return {
      ok: false,
      summary: expanded.summary,
      warnings: expanded.warnings,
      suggestedValues: {
        source: { sheetName: input.target?.sheetName, range: cellAddressFor(address, 0, 0) },
        destination: { sheetName: input.target?.sheetName, range: address },
        formula
      }
    };
  }
  return {
    ok: false,
    summary: `Formula matrix shape ${matrixRows}x${matrixColumns} does not match target range ${address} (${dimensions.rows}x${dimensions.columns}).`,
    warnings: ["Provide values.formulas with one formula per target cell, or provide values.formula as one formula to expand across the target range."],
    suggestedValues: {
      formula: formulaCells[0],
      formulas: padFormulaMatrix(formulas, dimensions.rows, dimensions.columns)
    }
  };
}

function padFormulaMatrix(matrix: CellMatrix, rows: number, columns: number): CellMatrix {
  return Array.from({ length: rows }, (_row, rowIndex) =>
    Array.from({ length: columns }, (_column, columnIndex) => {
      const value = matrix[rowIndex]?.[columnIndex];
      return formulaLike(value) ? String(value) : null;
    })
  );
}

function expandSingleFormulaAcrossRange(formula: string, dimensions: { rows: number; columns: number }): { ok: true; formulas: CellMatrix } | { ok: false; summary: string; warnings: string[] } {
  if (dimensions.rows <= 0 || dimensions.columns <= 0) {
    return { ok: false, summary: "Formula target range has no cells.", warnings: ["Provide a non-empty A1 target range."] };
  }
  if (/[![\]#]|[@]/.test(formula)) {
    return {
      ok: false,
      summary: "Formula expansion needs a same-sheet A1 formula or an explicit formula matrix.",
      warnings: ["External-sheet, structured, spill, and implicit-intersection formulas should be copied from a source cell or supplied as values.formulas."]
    };
  }
  return {
    ok: true,
    formulas: Array.from({ length: dimensions.rows }, (_row, rowOffset) =>
      Array.from({ length: dimensions.columns }, (_column, columnOffset) => shiftFormulaA1References(formula, rowOffset, columnOffset))
    )
  };
}

function shiftFormulaA1References(formula: string, rowOffset: number, columnOffset: number): string {
  return formula.replace(/(?<![A-Z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Z0-9_])/gi, (token, absoluteColumn: string, column: string, absoluteRow: string, row: string, index: number) => {
    if (formulaIndexInsideString(formula, index)) {
      return token;
    }
    const nextColumn = absoluteColumn ? column.toUpperCase() : numberToColumn(columnToNumber(column) + columnOffset);
    const nextRow = absoluteRow ? row : String(Number(row) + rowOffset);
    return `${absoluteColumn}${nextColumn}${absoluteRow}${nextRow}`;
  });
}

function formulaIndexInsideString(formula: string, index: number): boolean {
  let quoteCount = 0;
  for (let cursor = 0; cursor < index; cursor += 1) {
    if (formula[cursor] === '"') {
      if (formula[cursor + 1] === '"') {
        cursor += 1;
      } else {
        quoteCount += 1;
      }
    }
  }
  return quoteCount % 2 === 1;
}

function formulaPreviewExamples(range: string, formulas: CellMatrix): Array<{ cell: string; formula: string }> {
  const examples: Array<{ cell: string; formula: string }> = [];
  for (let rowIndex = 0; rowIndex < formulas.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < (formulas[rowIndex]?.length ?? 0); columnIndex += 1) {
      const formula = formulas[rowIndex]?.[columnIndex];
      if (!formulaLike(formula)) {
        continue;
      }
      examples.push({ cell: cellAddressFor(range, rowIndex, columnIndex), formula: String(formula) });
      if (examples.length >= 8) {
        return examples;
      }
    }
  }
  return examples;
}

function formulaNeedsInputOutput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  resolved: Extract<AgentTargetResolution, { ok: true }>,
  summary: string,
  suggestedValues?: Record<string, unknown>,
  warnings: string[] = []
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    answer: stripUndefinedRecord({
      kind: "formula_update_needs_input",
      target: { sheetName: resolved.sheetName, range: resolved.range },
      suggestedIntentAction: "write_formulas",
      suggestedValues
    }),
    proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula target" }],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    agentInstruction: "Use the suggested formula payload shape if present; do not split this into repeated single-cell formula writes.",
    maxRecommendedFollowupCalls: 1,
    warnings
  };
}

function numberFormatMatrixFromInput(input: AgentRunInput, address: string): string[][] | undefined {
  const values = input.values as (Record<string, unknown> | undefined);
  const raw = values?.numberFormats ?? values?.numberFormat ?? values?.formats;
  if (Array.isArray(raw) && raw.every((row) => Array.isArray(row))) {
    return raw.map((row) => row.map((value) => String(value ?? "")));
  }
  if (typeof raw === "string") {
    const dimensions = dimensionsFromAddress(address);
    if (!dimensions) {
      return [[raw]];
    }
    return Array.from({ length: dimensions.rows }, () => Array.from({ length: dimensions.columns }, () => raw));
  }
  return undefined;
}

function rangeTransferEndpointsFromInput(metadata: WorkbookMetadata, input: AgentRunInput): { source: A1Range; destination: A1Range } | undefined {
  const values = input.values as (Record<string, unknown> | undefined);
  const fallbackSheetName = input.target?.sheetName;
  const source = explicitRangeTarget(metadata, values?.source, fallbackSheetName);
  const destination = explicitRangeTarget(metadata, values?.destination ?? values?.target ?? input.target, fallbackSheetName);
  return source && destination ? { source, destination } : undefined;
}

function formulaCopyPatternsRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): FormulaCopyPatternsRequest | undefined {
  const endpoints = rangeTransferEndpointsFromInput(metadata, input);
  if (!endpoints) {
    return undefined;
  }
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    sourceSheetName: endpoints.source.sheetName,
    targetSheetName: endpoints.destination.sheetName,
    sourceAddress: endpoints.source.address,
    targetAddress: endpoints.destination.address
  };
}

function formulaTemplateRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const templateId = stringValue(values?.templateId);
  const targetSheetName = stringValue(input.target?.sheetName ?? values?.targetSheetName ?? values?.sheetName);
  if (!templateId || !targetSheetName) {
    return undefined;
  }
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    templateId: templateId as TemplateId,
    targetSheetName
  };
}

function repairTemplateRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const templateId = stringValue(values?.templateId);
  const targetSheetName = stringValue(input.target?.sheetName ?? values?.targetSheetName ?? values?.sheetName);
  if (!templateId && !targetSheetName) {
    return undefined;
  }
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    ...(templateId ? { templateId: templateId as TemplateId } : {}),
    ...(targetSheetName ? { targetSheetName } : {})
  };
}

function styleFingerprintRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): { workbookId: WorkbookId; sheetName: string; address?: string; maxCellSamples?: number } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const sheetName = stringValue(input.target?.sheetName ?? values?.sheetName ?? metadata.workbook.activeSheet);
  if (!sheetName) {
    return undefined;
  }
  const address = stringValue(input.target?.range ?? values?.address ?? values?.range);
  const maxCellSamples = positiveNumberValue(values?.maxCellSamples);
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    sheetName,
    ...(address ? { address: normalizeOperationRange(metadata, sheetName, address) } : {}),
    ...(maxCellSamples !== undefined ? { maxCellSamples } : {})
  };
}

function styleCompareRequestFromInput(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { workbookId: WorkbookId; sourceSheetName: string; targetSheetName: string; sourceAddress?: string; targetAddress?: string; dimensions?: StyleDimension[]; maxCellSamples?: number } | undefined {
  const copy = styleCopyRequestFromInput(metadata, input);
  const values = input.values as Record<string, unknown> | undefined;
  if (!copy) {
    return undefined;
  }
  const maxCellSamples = positiveNumberValue(values?.maxCellSamples);
  return {
    workbookId: copy.workbookId,
    sourceSheetName: copy.sourceSheetName,
    targetSheetName: copy.targetSheetName,
    ...(copy.sourceAddress !== undefined ? { sourceAddress: copy.sourceAddress } : {}),
    ...(copy.targetAddress !== undefined ? { targetAddress: copy.targetAddress } : {}),
    ...(copy.dimensions.length > 0 ? { dimensions: copy.dimensions } : {}),
    ...(maxCellSamples !== undefined ? { maxCellSamples } : {})
  };
}

function styleCopyRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): StyleCopyRequest | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const endpoints = rangeTransferEndpointsFromInput(metadata, input);
  if (endpoints) {
    return {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      sourceSheetName: endpoints.source.sheetName,
      targetSheetName: endpoints.destination.sheetName,
      sourceAddress: endpoints.source.address,
      targetAddress: endpoints.destination.address,
      dimensions: styleDimensionsFromInput(values)
    };
  }
  const sourceSheetName = stringValue(values?.sourceSheetName ?? values?.templateSheetName);
  const targetSheetName = stringValue(values?.targetSheetName ?? input.target?.sheetName ?? values?.sheetName);
  if (!sourceSheetName || !targetSheetName) {
    return undefined;
  }
  const sourceAddress = stringValue(values?.sourceAddress);
  const targetAddress = stringValue(values?.targetAddress ?? input.target?.range ?? values?.address ?? values?.range);
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    sourceSheetName,
    targetSheetName,
    ...(sourceAddress ? { sourceAddress: normalizeOperationRange(metadata, sourceSheetName, sourceAddress) } : {}),
    ...(targetAddress ? { targetAddress: normalizeOperationRange(metadata, targetSheetName, targetAddress) } : {}),
    dimensions: styleDimensionsFromInput(values)
  };
}

function styleCopyRequestsFromInput(metadata: WorkbookMetadata, input: AgentRunInput): StyleCopyRequest[] {
  const values = input.values as Record<string, unknown> | undefined;
  const rawEntries = values?.styleCopies ?? values?.copies ?? values?.entries;
  if (!Array.isArray(rawEntries)) {
    const request = styleCopyRequestFromInput(metadata, input);
    return request ? [request] : [];
  }
  const requests: StyleCopyRequest[] = [];
  for (const rawEntry of rawEntries) {
    if (!rawEntry || typeof rawEntry !== "object") {
      continue;
    }
    const entryValues = rawEntry as Record<string, unknown>;
    const request = styleCopyRequestFromInput(metadata, { ...input, values: entryValues } as AgentRunInput);
    if (request) {
      requests.push(request);
    }
  }
  return requests;
}

function shouldPreviewReplaceStyledTable(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  return Boolean(
    values
    && Array.isArray(values.headers)
    && (Array.isArray(values.row) || Array.isArray(values.rows))
    && /\b(replace|rotate|headers?|table|field\/value|field\s+value|ocr|screenshot|extract(?:ed)?|form|invoice|shipment|manifest|booking|horizontal|template style|styled)\b/i.test(input.request)
  );
}

function shouldPreviewGroupedHeader(input: AgentRunInput): boolean {
  const values = input.values as Record<string, unknown> | undefined;
  return Boolean(values?.groupedHeader)
    || /\b(?:grouped?|multi[-\s]?level|two[-\s]?layer|2[-\s]?layer|higher[-\s]?level)\b/i.test(input.request)
      && /\b(?:headers?|heading|column groups?|merge|merged|band|bands?)\b/i.test(input.request);
}

function replaceStyledTablePlanFromInput(metadata: WorkbookMetadata, input: AgentRunInput): {
  sheetName: string;
  writeRange: string;
  matrix: CellMatrix;
  clearRanges: string[];
  operations: ExcelOperation[];
  styleCopies: StyleCopyRequest[];
  changes: NonNullable<AgentRunOutput["changes"]>;
} | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const headers = Array.isArray(values?.headers) ? values.headers.map((value) => String(value)) : [];
  const rawRows = Array.isArray(values?.rows) ? values.rows : Array.isArray(values?.row) ? [values.row] : [];
  const rows = rawRows.filter(Array.isArray) as unknown[][];
  const sheetName = stringValue(input.target?.sheetName ?? values?.targetSheetName ?? values?.sheetName);
  if (!sheetName || headers.length === 0 || rows.length === 0) {
    return undefined;
  }
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  const matrix = [headers, ...rows.map((row) => headers.map((_header, index) => row[index] ?? null))] as CellMatrix;
  const writeRange = normalizeOperationRange(metadata, sheetName, stringValue(input.target?.range ?? values?.targetAddress ?? values?.address ?? values?.range) ?? matrixRangeFromOrigin("A1", matrix.length, headers.length));
  const operations: ExcelOperation[] = [];
  const clearRanges = clearRangesFromReplaceInput(metadata, input, sheetName, writeRange);
  for (const clearRange of clearRanges) {
    operations.push({
      kind: "range.clear",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "structure",
      reason: input.request,
      target: { workbookId, sheetName, address: clearRange },
      applyTo: "all"
    });
  }
  operations.push({
    kind: "range.write_values",
    operationId: makeId<OperationId>("op"),
    workbookId,
    destructiveLevel: "values",
    reason: input.request,
    target: { workbookId, sheetName, address: writeRange },
    values: matrix,
    preserveFormats: true
  });
  if (shouldAutofitReplacementColumns(input, values)) {
    operations.push({
      kind: "range.autofit_columns",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId, sheetName, address: writeRange }
    });
  }
  const styleCopies = replaceStyledTableStyleCopies(metadata, input, sheetName, writeRange);
  const changes: NonNullable<AgentRunOutput["changes"]> = [
    ...clearRanges.map((range) => ({ sheetName, range, before: "existing values and stale formats", after: "cleared values and formats" })),
    ...matrix.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
      sheetName,
      range: writeRange,
      cell: cellAddressFor(writeRange, rowIndex, columnIndex),
      after: value
    }))),
    ...styleCopies.map((request) => ({ sheetName: request.targetSheetName, ...(request.targetAddress ? { range: request.targetAddress } : {}), after: `copied style dimensions from ${request.sourceSheetName}` }))
  ];
  return { sheetName, writeRange, matrix, clearRanges, operations, styleCopies, changes };
}

function shouldAutofitReplacementColumns(input: AgentRunInput, values: Record<string, unknown> | undefined): boolean {
  const explicit = booleanValue(values?.autofit ?? values?.autoFit ?? values?.autofitColumns ?? values?.autoFitColumns);
  if (explicit !== undefined) {
    return explicit;
  }
  return /\b(?:autofit|auto\s*fit)\b/i.test(input.request);
}

function clearRangesFromReplaceInput(metadata: WorkbookMetadata, input: AgentRunInput, sheetName: string, writeRange: string): string[] {
  const values = input.values as Record<string, unknown> | undefined;
  const rawRanges = values?.clearRanges;
  if (Array.isArray(rawRanges)) {
    return rawRanges.flatMap((rawRange) => {
      const range = typeof rawRange === "string"
        ? rawRange
        : rawRange && typeof rawRange === "object"
          ? stringValue((rawRange as Record<string, unknown>).address ?? (rawRange as Record<string, unknown>).range)
          : undefined;
      return range ? [normalizeOperationRange(metadata, sheetName, range)] : [];
    });
  }
  const clearRange = stringValue(values?.clearRange ?? values?.clearAddress);
  if (clearRange) {
    return [normalizeOperationRange(metadata, sheetName, clearRange)];
  }
  const used = usedRangeForSheet(metadata, sheetName);
  return used && used !== writeRange ? [used] : [];
}

function replaceStyledTableStyleCopies(metadata: WorkbookMetadata, input: AgentRunInput, sheetName: string, writeRange: string): StyleCopyRequest[] {
  const values = input.values as Record<string, unknown> | undefined;
  const explicit = styleCopyRequestsFromInput(metadata, input);
  if (explicit.length > 0) {
    return explicit;
  }
  const headerSource = explicitRangeTarget(metadata, values?.headerStyleSource ?? values?.headerSource);
  const bodySource = explicitRangeTarget(metadata, values?.bodyStyleSource ?? values?.bodySource);
  const parsedTarget = rangeShape(writeRange);
  if (!parsedTarget) {
    return [];
  }
  const requests: StyleCopyRequest[] = [];
  if (headerSource?.sheetName && headerSource.address) {
    requests.push(...chunkStyleCopyRequests(metadata, headerSource.sheetName, headerSource.address, sheetName, rowsRange(writeRange, 1, 1), styleDimensionsFromInput(values)));
  }
  if (bodySource?.sheetName && bodySource.address && parsedTarget.rows > 1) {
    requests.push(...chunkStyleCopyRequests(metadata, bodySource.sheetName, bodySource.address, sheetName, rowsRange(writeRange, 2, parsedTarget.rows - 1), styleDimensionsFromInput(values)));
  }
  return requests;
}

function chunkStyleCopyRequests(metadata: WorkbookMetadata, sourceSheetName: string, sourceAddress: string, targetSheetName: string, targetAddress: string, dimensions: StyleDimension[]): StyleCopyRequest[] {
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  const source = rangeShape(normalizeOperationRange(metadata, sourceSheetName, sourceAddress));
  const target = rangeShape(normalizeOperationRange(metadata, targetSheetName, targetAddress));
  if (!source || !target) {
    return [];
  }
  const sourceWidth = Math.max(1, source.columns);
  const sourceHeight = Math.max(1, source.rows);
  const requests: StyleCopyRequest[] = [];
  for (let rowOffset = 0; rowOffset < target.rows; rowOffset += sourceHeight) {
    const height = Math.min(sourceHeight, target.rows - rowOffset);
    for (let columnOffset = 0; columnOffset < target.columns; columnOffset += sourceWidth) {
      const width = Math.min(sourceWidth, target.columns - columnOffset);
      requests.push({
        workbookId,
        sourceSheetName,
        targetSheetName,
        sourceAddress: addressFromBounds(source.startRow, source.startColumn, height, width),
        targetAddress: addressFromBounds(target.startRow + rowOffset, target.startColumn + columnOffset, height, width),
        dimensions
      });
    }
  }
  return requests;
}

function styleCopyDimensionIssue(request: StyleCopyRequest): string | undefined {
  if (!request.sourceAddress || !request.targetAddress) {
    return undefined;
  }
  const source = rangeShape(request.sourceAddress);
  const target = rangeShape(request.targetAddress);
  if (!source || !target) {
    return undefined;
  }
  if (source.rows !== target.rows || source.columns !== target.columns) {
    return `Style copy source and target ranges must have the same dimensions. Got source ${request.sourceSheetName}!${request.sourceAddress} (${source.rows}x${source.columns}) and target ${request.targetSheetName}!${request.targetAddress} (${target.rows}x${target.columns}).`;
  }
  return undefined;
}

function rangeShape(address: string): { startRow: number; startColumn: number; endRow: number; endColumn: number; rows: number; columns: number } | undefined {
  const parsed = tryParseA1Address(stripSheetName(address));
  if (!parsed) {
    return undefined;
  }
  return {
    startRow: parsed.startRow,
    startColumn: parsed.startColumn,
    endRow: parsed.endRow,
    endColumn: parsed.endColumn,
    rows: parsed.endRow - parsed.startRow + 1,
    columns: parsed.endColumn - parsed.startColumn + 1
  };
}

function matrixRangeFromOrigin(origin: string, rows: number, columns: number): string {
  const parsed = tryParseA1Address(origin) ?? { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 };
  return addressFromBounds(parsed.startRow, parsed.startColumn, rows, columns);
}

function rowsRange(address: string, oneBasedRowOffset: number, rowCount: number): string {
  const parsed = rangeShape(address);
  if (!parsed) {
    return address;
  }
  const row = parsed.startRow + oneBasedRowOffset - 1;
  return addressFromBounds(row, parsed.startColumn, rowCount, parsed.columns);
}

function addressFromBounds(startRow: number, startColumn: number, rows: number, columns: number): string {
  const endRow = startRow + Math.max(1, rows) - 1;
  const endColumn = startColumn + Math.max(1, columns) - 1;
  return `${numberToColumn(startColumn)}${startRow}:${numberToColumn(endColumn)}${endRow}`;
}

type TransformValuesOperation = "add_prefix" | "add_suffix" | "replace_text" | "normalize_whitespace" | "case" | "fill_blank" | "map_values" | "conditional_replace";
type DeriveValuesOperation = "copy_from_source" | "copy_if_blank" | "extract_pattern" | "lookup_map" | "conditional_map" | "normalize_from_source" | "formula_like";

interface ValueColumnScope {
  sheetName: string;
  address: string;
  columnLetter: string;
  startRow: number;
  endRow: number;
  rowCount: number;
  headerName?: string;
  tableName?: string;
  sectionId?: string;
}

interface CompiledColumnPlan {
  afterValues: unknown[];
  changedRows: boolean[];
  scannedCount: number;
  changedCount: number;
  skipped: Record<string, number>;
  unmatchedCount?: number;
  examples: Array<Record<string, unknown>>;
  warnings: string[];
}

interface LookupMapScope {
  keyScope: ValueColumnScope;
  valueScope: ValueColumnScope;
  duplicateKeys: string[];
}

interface SettlementBundlePlan {
  ok: true;
  sheetName: string;
  tableName?: string;
  scopes: {
    cashAmount: ValueColumnScope;
    actualAmount: ValueColumnScope;
    paymentVariance: ValueColumnScope;
    reconciliationNote: ValueColumnScope;
    detailNotes: ValueColumnScope;
  };
  reference?: Record<string, unknown>;
  operations: ExcelOperation[];
  changes: NonNullable<AgentRunOutput["changes"]>;
  proof: AgentProofReference[];
  scannedCount: number;
  changedCount: number;
  skipped: Record<string, number>;
  examples: Array<Record<string, unknown>>;
  warnings: string[];
}

function transformOperationFromInput(input: AgentRunInput): { ok: true; operation: TransformValuesOperation } | { ok: false; summary: string; warnings: string[] } {
  const raw = keyedStringValue(input.values, "operation", "transform", "type");
  const normalized = normalizeOperationName(raw ?? "");
  if (isTransformValuesOperation(normalized)) {
    return { ok: true, operation: normalized };
  }
  const request = input.request.toLowerCase();
  if (/\bprefix\b|\bprepend\b|\badd .{0,20}\bfront\b/.test(request)) return { ok: true, operation: "add_prefix" };
  if (/\bsuffix\b|\bappend\b/.test(request)) return { ok: true, operation: "add_suffix" };
  if (/\breplace\b|\bchange\b/.test(request)) return { ok: true, operation: "replace_text" };
  if (/\bfill\b/.test(request) && /\bblank|empty|missing\b/.test(request)) return { ok: true, operation: "fill_blank" };
  if (/\btrim\b|\bnormalize\s+space|\bwhitespace\b/.test(request)) return { ok: true, operation: "normalize_whitespace" };
  return {
    ok: false,
    summary: "Value transform needs values.operation such as add_prefix, replace_text, fill_blank, map_values, or normalize_whitespace.",
    warnings: ["Use transform_values for deterministic column transforms; provide operation plus required values like prefix, find/replacement, value, or map."]
  };
}

function derivationOperationFromInput(input: AgentRunInput): { ok: true; operation: DeriveValuesOperation } | { ok: false; summary: string; warnings: string[] } {
  const raw = keyedStringValue(input.values, "derivation", "operation", "type");
  const normalized = normalizeOperationName(raw ?? "");
  if (isDeriveValuesOperation(normalized)) {
    return { ok: true, operation: normalized };
  }
  if (normalized === "formula_like") return { ok: true, operation: "formula_like" };
  const request = input.request.toLowerCase();
  if (/\bfill\b/.test(request) && /\bblank|empty|missing\b/.test(request)) return { ok: true, operation: "copy_if_blank" };
  if (/\bextract\b/.test(request)) return { ok: true, operation: "extract_pattern" };
  if (/\blookup|master|map\b/.test(request)) return { ok: true, operation: "lookup_map" };
  if (/\b(formula|calculate|calculation|diff|difference|variance|actual\s*-\s*cash|cash\s*-\s*actual)\b/.test(request)) return { ok: true, operation: "formula_like" };
  if (/\bbased on|conditional|if\b/.test(request)) return { ok: true, operation: "conditional_map" };
  if (/\bnormalize\b/.test(request)) return { ok: true, operation: "normalize_from_source" };
  return { ok: true, operation: "copy_from_source" };
}

async function compileSettlementBundle(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  readColumn: (scope: ValueColumnScope) => Promise<{ values: CellMatrix; formulas: CellMatrix }>
): Promise<SettlementBundlePlan | { ok: false; summary: string; warnings: string[]; candidates?: AgentCandidate[] }> {
  const values = input.values ?? {};
  const sheetName = input.target?.sheetName ?? keyedStringValue(values, "targetSheetName", "sheetName") ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name;
  if (!sheetName) {
    return { ok: false, summary: "Settlement workflow needs a target sheet.", warnings: ["Provide target.sheetName for the transaction sheet."] };
  }
  const tableName = input.target?.tableName ?? keyedStringValue(values, "tableName");
  const baseTarget = { ...input.target, sheetName, ...(tableName ? { tableName } : {}) };
  const resolveColumn = (column: string): { ok: true; scope: ValueColumnScope } | { ok: false; summary: string; warnings: string[]; candidates?: AgentCandidate[] } =>
    resolveValueColumnScope(metadata, {
      ...input,
      target: { ...baseTarget, column },
      values: { ...values, targetColumn: column }
    }, "target");
  const cashAmount = resolveColumn("Cash Amount");
  if (!cashAmount.ok) return { ...cashAmount, summary: `Settlement workflow could not resolve Cash Amount. ${cashAmount.summary}` };
  const actualAmount = resolveColumn("Actual Amount");
  if (!actualAmount.ok) return { ...actualAmount, summary: `Settlement workflow could not resolve Actual Amount. ${actualAmount.summary}` };
  const paymentVariance = resolveColumn("Payment Variance");
  if (!paymentVariance.ok) return { ...paymentVariance, summary: `Settlement workflow could not resolve Payment Variance. ${paymentVariance.summary}` };
  const reconciliationNote = resolveColumn("Reconciliation Note");
  if (!reconciliationNote.ok) return { ...reconciliationNote, summary: `Settlement workflow could not resolve Reconciliation Note. ${reconciliationNote.summary}` };
  const detailNotes = resolveColumn("Detail Notes");
  if (!detailNotes.ok) return { ...detailNotes, summary: `Settlement workflow could not resolve Detail Notes. ${detailNotes.summary}` };

  const scopes = {
    cashAmount: cashAmount.scope,
    actualAmount: actualAmount.scope,
    paymentVariance: paymentVariance.scope,
    reconciliationNote: reconciliationNote.scope,
    detailNotes: detailNotes.scope
  };
  const rowCount = Math.min(...Object.values(scopes).map((scope) => scope.rowCount));
  const misaligned = Object.values(scopes).find((scope) => scope.rowCount !== rowCount);
  if (misaligned) {
    return {
      ok: false,
      summary: "Settlement workflow columns do not align by row.",
      warnings: [`Expected matching row counts; ${misaligned.headerName ?? misaligned.columnLetter} has ${misaligned.rowCount} row(s), common row count is ${rowCount}.`]
    };
  }

  const [cashSnapshot, actualSnapshot, varianceSnapshot, reconciliationSnapshot, detailSnapshot] = await Promise.all([
    readColumn(scopes.cashAmount),
    readColumn(scopes.actualAmount),
    readColumn(scopes.paymentVariance),
    readColumn(scopes.reconciliationNote),
    readColumn(scopes.detailNotes)
  ]);
  const cashValues = columnVector(cashSnapshot.values);
  const actualValues = columnVector(actualSnapshot.values);
  const varianceValues = columnVector(varianceSnapshot.values);
  const varianceFormulas = columnVector(varianceSnapshot.formulas);
  const reconciliationValues = columnVector(reconciliationSnapshot.values);
  const detailValues = columnVector(detailSnapshot.values);
  const rowUpdates = settlementRowUpdates(values);
  const explicitRows = settlementExplicitRows(values, rowUpdates);
  const rowIndexes = explicitRows.length > 0
    ? explicitRows.map((row) => row - scopes.paymentVariance.startRow).filter((index) => index >= 0 && index < rowCount)
    : Array.from({ length: rowCount }, (_value, index) => index);
  const globalReconciliationNote = keyedSettlementString(values, "reconciliationNote", "settlementNote", "note");
  const globalDetailNotes = keyedSettlementString(values, "detailNotes", "detailNote", "details");
  const updateVariance = values.updatePaymentVariance !== false && values.variance !== false;
  const writeFormula = values.paymentVarianceMode !== "value" && values.varianceMode !== "value" && values.writeFormula !== false;
  const skipped: Record<string, number> = {};
  const formulaAfter = Array.from({ length: rowCount }, () => "");
  const formulaChanged = Array.from({ length: rowCount }, () => false);
  const varianceValueAfter = [...varianceValues];
  const varianceValueChanged = Array.from({ length: rowCount }, () => false);
  const reconciliationAfter = [...reconciliationValues];
  const reconciliationChanged = Array.from({ length: rowCount }, () => false);
  const detailAfter = [...detailValues];
  const detailChanged = Array.from({ length: rowCount }, () => false);
  const examples: Array<Record<string, unknown>> = [];
  let changedCount = 0;

  const updatesByRow = new Map<number, SettlementRowUpdate>();
  for (const update of rowUpdates) {
    if (update.row !== undefined) updatesByRow.set(update.row, update);
  }

  for (const index of rowIndexes) {
    const sheetRow = scopes.paymentVariance.startRow + index;
    const rowUpdate = updatesByRow.get(sheetRow);
    const cash = cashValues[index];
    const actual = actualValues[index];
    const currentVariance = varianceValues[index];
    let changed = false;
    const example: Record<string, unknown> = {
      row: sheetRow,
      source: {
        "Cash Amount": cash ?? "",
        "Actual Amount": actual ?? ""
      },
      before: {
        "Payment Variance": currentVariance ?? "",
        "Reconciliation Note": reconciliationValues[index] ?? "",
        "Detail Notes": detailValues[index] ?? ""
      },
      after: {}
    };
    if (updateVariance) {
      const cashNumber = numericCellValue(cash);
      const actualNumber = numericCellValue(actual);
      if (cashNumber === undefined || actualNumber === undefined) {
        incrementCount(skipped, "blankAmountSource");
      } else if (writeFormula) {
        const expectedFormula = `=${scopes.actualAmount.columnLetter}${sheetRow}-${scopes.cashAmount.columnLetter}${sheetRow}`;
        if (normalizeFormulaText(varianceFormulas[index]) !== normalizeFormulaText(expectedFormula)) {
          formulaAfter[index] = expectedFormula;
          formulaChanged[index] = true;
          changed = true;
          changedCount += 1;
          (example.after as Record<string, unknown>)["Payment Variance"] = expectedFormula;
        }
      } else {
        const expectedValue = actualNumber - cashNumber;
        if (!Object.is(currentVariance, expectedValue)) {
          varianceValueAfter[index] = expectedValue;
          varianceValueChanged[index] = true;
          changed = true;
          changedCount += 1;
          (example.after as Record<string, unknown>)["Payment Variance"] = expectedValue;
        }
      }
    }
    const nextReconciliation = rowUpdate?.reconciliationNote ?? globalReconciliationNote;
    if (nextReconciliation !== undefined && !Object.is(reconciliationValues[index], nextReconciliation)) {
      reconciliationAfter[index] = nextReconciliation;
      reconciliationChanged[index] = true;
      changed = true;
      changedCount += 1;
      (example.after as Record<string, unknown>)["Reconciliation Note"] = nextReconciliation;
    }
    const nextDetail = rowUpdate?.detailNotes ?? globalDetailNotes;
    if (nextDetail !== undefined && !Object.is(detailValues[index], nextDetail)) {
      detailAfter[index] = nextDetail;
      detailChanged[index] = true;
      changed = true;
      changedCount += 1;
      (example.after as Record<string, unknown>)["Detail Notes"] = nextDetail;
    }
    if (changed && examples.length < 10) {
      examples.push(example);
    }
  }

  const operations: ExcelOperation[] = [];
  const changes: NonNullable<AgentRunOutput["changes"]> = [];
  if (writeFormula) {
    const formulaRuns = changedColumnRuns(scopes.paymentVariance, formulaAfter, formulaChanged);
    operations.push(...formulaRuns.map((run) => ({
      kind: "range.write_formulas" as const,
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values" as const,
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: scopes.paymentVariance.sheetName, address: run.address },
      formulas: run.values.map((row) => [typeof row[0] === "string" ? row[0] : null]),
      preserveFormats: true as const
    })));
  } else {
    const varianceRuns = changedColumnRuns(scopes.paymentVariance, varianceValueAfter, varianceValueChanged);
    if (varianceRuns.length > 0) {
      operations.push({
        kind: "range.write_values_many",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "values",
        reason: input.request,
        entries: varianceRuns.map((run) => ({
          target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: scopes.paymentVariance.sheetName, address: run.address },
          values: run.values,
          preserveFormats: true
        }))
      });
    }
  }
  const noteEntries = [
    ...changedColumnRuns(scopes.reconciliationNote, reconciliationAfter, reconciliationChanged).map((run) => ({ scope: scopes.reconciliationNote, run })),
    ...changedColumnRuns(scopes.detailNotes, detailAfter, detailChanged).map((run) => ({ scope: scopes.detailNotes, run }))
  ];
  if (noteEntries.length > 0) {
    operations.push({
      kind: "range.write_values_many",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      entries: noteEntries.map(({ scope, run }) => ({
        target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: scope.sheetName, address: run.address },
        values: run.values,
        preserveFormats: true
      }))
    });
  }
  changes.push(
    { sheetName, range: scopes.paymentVariance.address, before: `${rowCount} scanned`, after: `${formulaChanged.filter(Boolean).length + varianceValueChanged.filter(Boolean).length} variance update(s)` },
    { sheetName, range: scopes.reconciliationNote.address, before: "existing reconciliation notes", after: `${reconciliationChanged.filter(Boolean).length} note update(s)` },
    { sheetName, range: scopes.detailNotes.address, before: "existing detail notes", after: `${detailChanged.filter(Boolean).length} detail note update(s)` },
    ...examples.map((example) => ({ sheetName, range: `${scopes.paymentVariance.columnLetter}${example.row}`, before: example.before, after: example.after }))
  );
  const reference = await settlementReferenceProof(metadata, input, readColumn);
  return {
    ok: true,
    sheetName,
    ...(tableName ? { tableName } : scopes.paymentVariance.tableName ? { tableName: scopes.paymentVariance.tableName } : {}),
    scopes,
    ...(reference ? { reference } : {}),
    operations,
    changes,
    proof: [
      { sheetName, range: scopes.cashAmount.address, label: "Cash Amount source" },
      { sheetName, range: scopes.actualAmount.address, label: "Actual Amount source" },
      { sheetName, range: scopes.paymentVariance.address, label: "Payment Variance target" },
      { sheetName, range: scopes.reconciliationNote.address, label: "Reconciliation Note target" },
      { sheetName, range: scopes.detailNotes.address, label: "Detail Notes target" }
    ],
    scannedCount: rowIndexes.length,
    changedCount,
    skipped,
    examples,
    warnings: []
  };
}

interface SettlementRowUpdate {
  row?: number;
  reconciliationNote?: string;
  detailNotes?: string;
}

function settlementRowUpdates(values: AgentRunInput["values"] | undefined): SettlementRowUpdate[] {
  const raw = values?.rowUpdates ?? values?.settlements ?? values?.updates;
  if (!Array.isArray(raw)) {
    return [];
  }
  return raw.flatMap((entry): SettlementRowUpdate[] => {
    if (typeof entry === "number" && Number.isInteger(entry)) {
      return [{ row: entry }];
    }
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
      return [];
    }
    const record = entry as Record<string, unknown>;
    const row = numberFromUnknown(record.row ?? record.rowNumber ?? record.sheetRow);
    const reconciliationNote = stringFromUnknown(record.reconciliationNote ?? record.settlementNote ?? record.note);
    const detailNotes = stringFromUnknown(record.detailNotes ?? record.detailNote ?? record.details);
    return [stripUndefinedRecord({ row, reconciliationNote, detailNotes }) as SettlementRowUpdate];
  });
}

function settlementExplicitRows(values: AgentRunInput["values"] | undefined, rowUpdates: SettlementRowUpdate[]): number[] {
  const rowNumbers = arrayNumbersFromUnknown(values?.rowNumbers ?? values?.sheetRows);
  const rows = arrayNumbersFromUnknown(values?.rows);
  const updateRows = rowUpdates.flatMap((update) => update.row !== undefined ? [update.row] : []);
  return Array.from(new Set([...rowNumbers, ...rows, ...updateRows])).filter((row) => Number.isInteger(row) && row > 0);
}

function keyedSettlementString(values: AgentRunInput["values"] | undefined, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = values?.[key];
    if (typeof value === "string") {
      return value;
    }
  }
  return undefined;
}

function numberFromUnknown(value: unknown): number | undefined {
  return typeof value === "number" && Number.isInteger(value) ? value : undefined;
}

function stringFromUnknown(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function arrayNumbersFromUnknown(value: unknown): number[] {
  return Array.isArray(value) && value.every((item) => typeof item === "number" && Number.isInteger(item))
    ? value
    : [];
}

function normalizeFormulaText(value: unknown): string {
  return typeof value === "string" ? value.replace(/\s+/g, "").toUpperCase() : "";
}

async function settlementReferenceProof(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  readColumn: (scope: ValueColumnScope) => Promise<{ values: CellMatrix; formulas: CellMatrix }>
): Promise<Record<string, unknown> | undefined> {
  const referenceSheetName = keyedStringValue(input.values, "referenceSheetName", "referenceSheet", "templateSheetName");
  if (!referenceSheetName) {
    return undefined;
  }
  const referenceTableName = keyedStringValue(input.values, "referenceTableName");
  const referenceScope = resolveValueColumnScope(metadata, {
    ...input,
    target: { sheetName: referenceSheetName, ...(referenceTableName ? { tableName: referenceTableName } : {}), column: "Payment Variance" },
    values: { ...(input.values ?? {}), targetColumn: "Payment Variance" }
  }, "target");
  if (!referenceScope.ok) {
    return {
      sheetName: referenceSheetName,
      warning: referenceScope.summary
    };
  }
  const snapshot = await readColumn(referenceScope.scope);
  const formulas = columnVector(snapshot.formulas);
  const firstFormulaIndex = formulas.findIndex((formula) => formulaLike(formula));
  return stripUndefinedRecord({
    sheetName: referenceSheetName,
    paymentVariance: {
      range: referenceScope.scope.address,
      header: referenceScope.scope.headerName,
      formulaExample: firstFormulaIndex >= 0 ? formulas[firstFormulaIndex] : undefined,
      formulaCell: firstFormulaIndex >= 0 ? `${referenceScope.scope.columnLetter}${referenceScope.scope.startRow + firstFormulaIndex}` : undefined
    }
  });
}

function deriveFormulaLikeValue(current: unknown, sourceValues: unknown[], values: AgentRunInput["values"]): unknown | typeof noChange | typeof skipBecauseBlankSource {
  const operation = normalizeOperationName(keyedStringValue(values, "formula", "formulaType", "expression") ?? "");
  const first = numericCellValue(sourceValues[0]);
  const second = numericCellValue(sourceValues[1]);
  if (first === undefined || second === undefined) return skipBecauseBlankSource;
  if (operation === "cash_minus_actual") {
    return first - second;
  }
  if (operation === "actual_minus_cash" || operation === "difference" || operation === "diff" || operation === "variance" || operation === "") {
    return second - first;
  }
  const requestFormula = keyedStringValue(values, "operationFormula", "calculation");
  if (requestFormula && /cash.*actual/i.test(requestFormula) && !/actual.*cash/i.test(requestFormula)) {
    return first - second;
  }
  if (Object.is(current, second - first)) return noChange;
  return second - first;
}

function shouldInspectFormulaInline(input: AgentRunInput): boolean {
  if (intentAction(input) === "read_formulas") return true;
  if (intentAction(input) !== undefined) return false;
  if (isDataValidationReadRequest(input.request.toLowerCase())) return false;
  if (!/\b(formula|formulas|raw formula|r1c1|calculation)\b/i.test(input.request)) return false;
  if (/\b(write|set|apply|fill|copy|repair|replace|update|convert)\b/i.test(input.request) && !/\b(read|check|show|inspect|is|has|whether)\b/i.test(input.request)) {
    return false;
  }
  return Boolean(input.target?.range || input.target?.sheetName || /\b(this|selected|current)\b/i.test(input.request));
}

function shouldPreviewSettlementBundle(input: AgentRunInput): boolean {
  if (intentAction(input) !== undefined) return false;
  return /\b(settle|settlement|reconcile|reconciliation)\b/i.test(input.request)
    && /\b(payment\s+variance|variance|reconciliation\s+note|detail\s+notes?)\b/i.test(input.request);
}

function normalizeOperationName(value: string): string {
  return value.trim().toLowerCase().replace(/[\s-]+/g, "_");
}

function isTransformValuesOperation(value: string): value is TransformValuesOperation {
  return ["add_prefix", "add_suffix", "replace_text", "normalize_whitespace", "case", "fill_blank", "map_values", "conditional_replace"].includes(value);
}

function isDeriveValuesOperation(value: string): value is DeriveValuesOperation {
  return ["copy_from_source", "copy_if_blank", "extract_pattern", "lookup_map", "conditional_map", "normalize_from_source"].includes(value);
}

function resolveLookupMapScope(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { ok: true } & LookupMapScope | { ok: false; summary: string; warnings: string[]; candidates?: AgentCandidate[] } {
  const values = input.values ?? {};
  const lookupSheetName = keyedStringValue(values, "lookupSheetName", "lookupSheet", "referenceSheetName", "referenceSheet", "sourceLookupSheet");
  const lookupTableName = keyedStringValue(values, "lookupTableName", "lookupTable", "referenceTableName", "referenceTable");
  const lookupKeyRange = keyedStringValue(values, "lookupKeyRange", "keyRange");
  const lookupValueRange = keyedStringValue(values, "lookupValueRange", "valueRange");
  const lookupKeyColumn = keyedStringValue(values, "lookupKeyColumn", "keyColumn", "lookupFromColumn");
  const lookupValueColumn = keyedStringValue(values, "lookupValueColumn", "valueColumn", "lookupToColumn");
  const sheetName = lookupSheetName
    ?? (lookupTableName ? metadata.tables.find((table) => sameText(table.name, lookupTableName) || sameText(table.id, lookupTableName))?.sheetName : undefined)
    ?? input.target?.sheetName
    ?? metadata.workbook.activeSheet
    ?? metadata.sheets[0]?.name;
  if (!sheetName) {
    return { ok: false, summary: "Lookup derivation needs a lookup sheet or table.", warnings: ["Provide values.lookupSheetName or values.lookupTableName with lookup key/value columns."] };
  }
  const keyInput: AgentRunInput = {
    ...input,
    target: { sheetName, ...(lookupTableName ? { tableName: lookupTableName } : {}) },
    values: {
      ...values,
      sourceSheetName: sheetName,
      ...(lookupKeyRange ? { sourceRange: lookupKeyRange } : {}),
      ...(lookupKeyColumn ? { sourceColumn: lookupKeyColumn } : {})
    }
  };
  const valueInput: AgentRunInput = {
    ...input,
    target: { sheetName, ...(lookupTableName ? { tableName: lookupTableName } : {}) },
    values: {
      ...values,
      sourceSheetName: sheetName,
      ...(lookupValueRange ? { sourceRange: lookupValueRange } : {}),
      ...(lookupValueColumn ? { sourceColumn: lookupValueColumn } : {})
    }
  };
  const keyScope = resolveValueColumnScope(metadata, keyInput, "source");
  if (!keyScope.ok) {
    return { ...keyScope, summary: `Lookup key could not be resolved. ${keyScope.summary}` };
  }
  const valueScope = resolveValueColumnScope(metadata, valueInput, "source");
  if (!valueScope.ok) {
    return { ...valueScope, summary: `Lookup value could not be resolved. ${valueScope.summary}` };
  }
  if (keyScope.scope.rowCount !== valueScope.scope.rowCount) {
    return {
      ok: false,
      summary: "Lookup key/value ranges do not align.",
      warnings: [`Lookup key rows: ${keyScope.scope.rowCount}; lookup value rows: ${valueScope.scope.rowCount}.`]
    };
  }
  return { ok: true, keyScope: keyScope.scope, valueScope: valueScope.scope, duplicateKeys: [] };
}

function resolveDeriveSourceScopes(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { ok: true; scopes: ValueColumnScope[] } | { ok: false; summary: string; warnings: string[]; candidates?: AgentCandidate[] } {
  const sourceColumns = keyedArrayStringValue(input.values, "sourceColumns", "sources", "source");
  const scopes: ValueColumnScope[] = [];
  if (sourceColumns.length > 0) {
    for (const column of sourceColumns) {
      const resolved = resolveValueColumnScope(metadata, {
        ...input,
        target: { ...input.target, column },
        values: { ...(input.values ?? {}), sourceColumn: column }
      }, "source");
      if (!resolved.ok) {
        return resolved;
      }
      scopes.push(resolved.scope);
    }
    return { ok: true, scopes };
  }
  const resolved = resolveValueColumnScope(metadata, input, "source");
  return resolved.ok ? { ok: true, scopes: [resolved.scope] } : resolved;
}

function resolveValueColumnScope(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  purpose: "target" | "source"
): { ok: true; scope: ValueColumnScope } | { ok: false; summary: string; warnings: string[]; candidates?: AgentCandidate[] } {
  const values = input.values ?? {};
  const explicitRange = purpose === "target"
    ? input.target?.range ?? keyedStringValue(values, "targetRange", "range")
    : keyedStringValue(values, "sourceRange");
  const sheetName = purpose === "source"
    ? keyedStringValue(values, "sourceSheetName") ?? input.target?.sheetName ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name
    : input.target?.sheetName ?? keyedStringValue(values, "targetSheetName", "sheetName") ?? metadata.workbook.activeSheet ?? metadata.sheets[0]?.name;
  if (!sheetName) {
    return { ok: false, summary: `Could not resolve ${purpose} sheet.`, warnings: ["Provide target.sheetName or a workbook context with an active sheet."] };
  }
  if (explicitRange) {
    const parsed = tryParseA1Address(stripSheetName(explicitRange));
    if (!parsed) {
      return { ok: false, summary: `Could not parse ${purpose} range ${explicitRange}.`, warnings: ["Use an A1-style range."] };
    }
    if (parsed.startColumn !== parsed.endColumn) {
      return { ok: false, summary: `${purpose} range ${explicitRange} spans multiple columns.`, warnings: ["Broad value transform/derive currently requires one target column."] };
    }
    return {
      ok: true,
      scope: {
        sheetName,
        address: stripSheetName(explicitRange),
        columnLetter: numberToColumn(parsed.startColumn),
        startRow: parsed.startRow,
        endRow: parsed.endRow,
        rowCount: parsed.endRow - parsed.startRow + 1
      }
    };
  }

  const columnHint = purpose === "target"
    ? input.target?.column ?? keyedStringValue(values, "targetColumn", "targetHeader", "column", "header")
    : keyedStringValue(values, "sourceColumn", "sourceHeader", "fromColumn", "fromHeader");
  if (!columnHint) {
    return {
      ok: false,
      summary: `${purpose} column is ambiguous.`,
      warnings: [`Provide ${purpose === "target" ? "target.column or values.targetColumn" : "values.sourceColumn"}.`],
      candidates: columnCandidates(metadata, sheetName)
    };
  }

  const tableName = purpose === "target" ? input.target?.tableName ?? keyedStringValue(values, "tableName") : keyedStringValue(values, "sourceTableName") ?? input.target?.tableName;
  const table = tableName
    ? metadata.tables.find((candidate) => sameText(candidate.name, tableName) || sameText(candidate.id, tableName))
    : metadata.tables.find((candidate) => candidate.sheetName === sheetName && candidate.columns.some((column) => columnMatchesHint(column, columnHint)));
  if (table?.dataRange) {
    const column = table.columns.find((candidate) => columnMatchesHint(candidate, columnHint));
    const data = tryParseA1Address(stripSheetName(table.dataRange));
    if (column && data) {
      const absoluteColumn = columnAbsoluteNumber(column, data.startColumn);
      return {
        ok: true,
        scope: {
          sheetName: table.sheetName,
          address: addressFromBounds(data.startRow, absoluteColumn, data.endRow - data.startRow + 1, 1),
          columnLetter: numberToColumn(absoluteColumn),
          startRow: data.startRow,
          endRow: data.endRow,
          rowCount: data.endRow - data.startRow + 1,
          headerName: column.name,
          ...(table.name ? { tableName: table.name } : {})
        }
      };
    }
  }

  const sheet = metadata.sheets.find((candidate) => candidate.name === sheetName);
  const header = sheet?.headers
    .filter((candidate) => candidate.columns.some((column) => columnMatchesHint(column, columnHint)))
    .sort((left, right) => right.confidence - left.confidence)[0];
  const column = header?.columns.find((candidate) => columnMatchesHint(candidate, columnHint));
  const used = sheet?.usedRange ? tryParseA1Address(stripSheetName(sheet.usedRange)) : undefined;
  if (header && column && used) {
    const absoluteColumn = columnAbsoluteNumber(column, used.startColumn);
    const startRow = header.row + 1;
    return {
      ok: true,
      scope: {
        sheetName,
        address: addressFromBounds(startRow, absoluteColumn, Math.max(1, used.endRow - startRow + 1), 1),
        columnLetter: numberToColumn(absoluteColumn),
        startRow,
        endRow: used.endRow,
        rowCount: Math.max(1, used.endRow - startRow + 1),
        headerName: column.name
      }
    };
  }

  const section = metadata.sections.find((candidate) =>
    candidate.sheetName === sheetName && candidate.columns.some((column) => columnMatchesHint(column, columnHint))
  );
  const dataRange = section ? dataRangeForSection(section) : undefined;
  const data = dataRange ? tryParseA1Address(stripSheetName(dataRange)) : undefined;
  const sectionColumn = section?.columns.find((candidate) => columnMatchesHint(candidate, columnHint));
  if (section && data && sectionColumn) {
    const absoluteColumn = columnAbsoluteNumber(sectionColumn, data.startColumn);
    return {
      ok: true,
      scope: {
        sheetName,
        address: addressFromBounds(data.startRow, absoluteColumn, data.endRow - data.startRow + 1, 1),
        columnLetter: numberToColumn(absoluteColumn),
        startRow: data.startRow,
        endRow: data.endRow,
        rowCount: data.endRow - data.startRow + 1,
        headerName: sectionColumn.name,
        sectionId: section.id
      }
    };
  }

  return {
    ok: false,
    summary: `Could not match ${purpose} column "${columnHint}" on ${sheetName}.`,
    warnings: [`Available columns: ${columnCandidates(metadata, sheetName).map((candidate) => candidate.label).slice(0, 12).join(", ") || "none"}.`],
    candidates: columnCandidates(metadata, sheetName)
  };
}

function columnAbsoluteNumber(column: ColumnMetadata, rangeStartColumn: number): number {
  const byLetter = /^[A-Z]+$/i.test(column.letter) ? columnToNumber(column.letter) : undefined;
  return byLetter && byLetter > 0 ? byLetter : rangeStartColumn + column.index;
}

function columnCandidates(metadata: WorkbookMetadata, sheetName: string): AgentCandidate[] {
  const tableCandidates = metadata.tables
    .filter((table) => table.sheetName === sheetName)
    .flatMap((table) => table.columns.map((column) => ({
      id: `${table.id}:${column.index}`,
      kind: "column" as const,
      label: column.name || column.letter,
      sheetName: table.sheetName,
      ...(table.name !== undefined ? { tableName: table.name } : {}),
      semanticRole: "data_table" as const,
      confidence: 0.8,
      reason: table.name ? `Column in table ${table.name}` : "Column in table"
    })));
  const headerCandidates = metadata.sheets
    .find((sheet) => sheet.name === sheetName)?.headers
    .flatMap((header) => header.columns.map((column) => ({
      id: `${header.id}:${column.index}`,
      kind: "column" as const,
      label: column.name || column.letter,
      sheetName,
      range: header.range,
      confidence: header.confidence,
      reason: `Header row ${header.row}`
    }))) ?? [];
  return [...tableCandidates, ...headerCandidates].slice(0, 20);
}

function columnMatchesHint(column: ColumnMetadata, hint: string): boolean {
  const normalized = normalizeHeaderName(hint);
  return normalizeHeaderName(column.name) === normalized
    || column.normalizedName === normalized
    || column.letter.toLowerCase() === hint.trim().toLowerCase()
    || normalizeHeaderName(column.name).includes(normalized);
}

function sameText(left: unknown, right: unknown): boolean {
  return typeof left === "string" && typeof right === "string" && left.trim().toLowerCase() === right.trim().toLowerCase();
}

type VisualReadabilityStyleDepth = "basic" | "standard" | "comprehensive";
type VisualReadabilityDensity = "compact" | "comfortable" | "presentation";
type VisualReadabilityPresentationMode = "working_sheet" | "print_ready" | "executive_report";
type VisualReadabilitySheetType = "generic_table" | "tabular_data" | "record_tracker" | "numeric_report" | "input_template" | "summary_report" | "mixed_template" | "unknown";
type VisualReadabilitySuggestionBucket = "layout" | "validation" | "reference_style" | "formula_helpers" | "structure" | "freeze_panes" | "print_settings";
type VisualReadabilityStylePreservationMode = "protected_regions" | "strict" | "none";

interface VisualReadabilityDetectedStructure {
  sheetName: string;
  usedRange?: string;
  headerRow?: number;
  headerRange?: string;
  dataRange?: string;
  tableRanges: string[];
  hasFilter: boolean;
  hasFreezePane: boolean;
  mergedRanges: string[];
  hiddenRows: number[];
  hiddenColumns: string[];
  protectedRanges: string[];
  existingStyleRanges: string[];
  protectedStyleRanges: string[];
  formulaColumns: string[];
  formulaRanges: string[];
  totalRows: number[];
  subtotalRows: number[];
  existingConditionalFormattingRanges: string[];
  existingDataValidationRanges: string[];
  detectionSource: "metadata";
  confidence: number;
}

interface VisualReadabilityColumnRole {
  column: string;
  header: string;
  role: string;
  inferredType: string;
  confidence: number;
  signals: string[];
}

type VisualRuleScope = "sheet" | "table" | "group" | "column" | "conditional_range" | "row" | "cell";
type VisualRuleKind = "font" | "fill" | "border" | "alignment" | "number_format" | "width" | "height" | "wrap" | "freeze" | "filter";

interface VisualReadabilityRule {
  id: string;
  scope: VisualRuleScope;
  target: string;
  kind: VisualRuleKind;
  value: unknown;
  risk: "low" | "medium" | "high";
  reason: string;
}

interface VisualReadabilityValidationSuggestion {
  id: string;
  target: string;
  source: string[];
  risk: "medium" | "high";
  reason: string;
  existingValidation: "preserved" | "not_detected";
}

interface VisualReadabilityFormulaSuggestion {
  id: string;
  suggestedColumnAfter: string;
  formulaName: string;
  formulaExample: string;
  risk: "medium" | "high";
  reason: string;
}

interface VisualReadabilityReferenceStyleOption {
  sheetName: string;
  adaptToTargetStructure: boolean;
  preserveTargetValues: boolean;
  preserveFormulas: boolean;
}

interface VisualReadabilityReferenceStyleSuggestion {
  id: string;
  referenceSheetName: string;
  target: string;
  pattern: string;
  risk: "medium" | "high";
  reason: string;
  preserveTargetValues: true;
  preserveFormulas: boolean;
}

interface VisualReadabilityPrintSuggestion {
  id: string;
  target: string;
  setting: string;
  value: string;
  presentationMode: VisualReadabilityPresentationMode;
  risk: "medium" | "high";
  reason: string;
}

interface VisualReadabilityPlanPreview {
  compilerStatus: "preview_compiled_apply_pending";
  summary: string[];
  counts: {
    totalRules: number;
    layoutChanges: number;
    groupRules: number;
    columnRules: number;
    conditionalRules: number;
    rowRules: number;
    cellRules: number;
    validationSuggestions: number;
    formulaSuggestions: number;
    referenceStyleSuggestions: number;
    printSuggestions: number;
  };
  ruleScopes: Record<VisualRuleScope, number>;
  rules: VisualReadabilityRule[];
  validationSuggestions: VisualReadabilityValidationSuggestion[];
  formulaSuggestions: VisualReadabilityFormulaSuggestion[];
  referenceStyleSuggestions: VisualReadabilityReferenceStyleSuggestion[];
  printSuggestions: VisualReadabilityPrintSuggestion[];
  previewExamples: Array<{ range: string; before: string; after: string; ruleId: string }>;
  theme: {
    name: string;
    font: { family: string; bodySize: number; headerSize: number };
    density: VisualReadabilityDensity;
  };
}

interface VisualReadabilityCompiledOperationSet {
  operations: ExcelOperation[];
  skipped: Array<{ ruleId: string; target: string; reason: string }>;
}

interface VisualReadabilityPreservationContext {
  protectedRanges: string[];
  mergedRanges: string[];
  hiddenColumns: string[];
  existingStyleRanges: string[];
  protectedStyleRanges: string[];
  existingConditionalFormattingRanges: string[];
}

function visualReadabilityOptionsFromInput(input: AgentRunInput): {
  styleDepth: VisualReadabilityStyleDepth;
  profile: string;
  density: VisualReadabilityDensity;
  preserveFormulas: boolean;
  preserveExistingStyle: boolean;
  stylePreservationMode: VisualReadabilityStylePreservationMode;
  allowValidationSuggestions: boolean;
  allowFormulaSuggestions: boolean;
  allowReplaceConditionalFormatting: boolean;
  allowReplaceDataValidation: boolean;
  allowInsertRowsOrColumns: boolean;
  applySuggestionBuckets: VisualReadabilitySuggestionBucket[];
  freezePanes?: { rows?: number; columns?: number };
  referenceStyle?: VisualReadabilityReferenceStyleOption;
  presentationMode?: VisualReadabilityPresentationMode;
} {
  const values = input.values ?? {};
  const nested = typeof values.visualReadability === "object" && values.visualReadability !== null && !Array.isArray(values.visualReadability)
    ? values.visualReadability as Record<string, unknown>
    : {};
  const optionValue = (key: string) => nested[key] ?? values[key];
  const referenceStyle = visualReadabilityReferenceStyle(optionValue("referenceStyle") ?? values.referenceStyle);
  const presentationMode = visualReadabilityPresentationMode(optionValue("presentationMode"));
  const applySuggestionBuckets = visualReadabilitySuggestionBuckets(optionValue("applySuggestionBuckets"));
  const freezePanes = visualReadabilityFreezePanes(optionValue("freezePanes"), optionValue("freezeRows"), optionValue("freezeColumns"), input.request);
  const preserveExistingStyle = optionValue("preserveExistingStyle");
  return {
    styleDepth: visualReadabilityStyleDepth(optionValue("styleDepth")),
    profile: stringValue(optionValue("profile")) ?? "auto",
    density: visualReadabilityDensity(optionValue("density")),
    preserveFormulas: booleanValue(optionValue("preserveFormulas")) ?? true,
    preserveExistingStyle: booleanValue(preserveExistingStyle) ?? true,
    stylePreservationMode: visualReadabilityStylePreservationMode(optionValue("stylePreservationMode"), preserveExistingStyle),
    allowValidationSuggestions: booleanValue(optionValue("allowValidationSuggestions")) ?? false,
    allowFormulaSuggestions: booleanValue(optionValue("allowFormulaSuggestions")) ?? false,
    allowReplaceConditionalFormatting: booleanValue(optionValue("allowReplaceConditionalFormatting")) ?? false,
    allowReplaceDataValidation: booleanValue(optionValue("allowReplaceDataValidation")) ?? false,
    allowInsertRowsOrColumns: booleanValue(optionValue("allowInsertRowsOrColumns")) ?? false,
    applySuggestionBuckets,
    ...(freezePanes ? { freezePanes } : {}),
    ...(referenceStyle ? { referenceStyle } : {}),
    ...(presentationMode ? { presentationMode } : {})
  };
}

function visualReadabilityDetectionFailure(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  sheetName: string,
  range: string | undefined,
  detected: VisualReadabilityDetectedStructure
): Omit<AgentRunOutput, "telemetry"> | undefined {
  if (detected.hiddenColumns.includes("sheet")) {
    return {
      status: "VALIDATION_FAILED",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Visual readability preview blocked ${sheetName} because the sheet is hidden.`,
      proof: range ? [{ sheetName, range, label: "hidden visual readability target" }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "ask_user",
      warnings: ["Unhide the sheet or explicitly target a visible sheet before applying visual readability styling."]
    };
  }
  const targetCellCount = range ? cellCountFromAddress(stripSheetName(range)) : undefined;
  if (targetCellCount !== undefined && targetCellCount > 200_000) {
    return {
      status: "VALIDATION_FAILED",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Visual readability preview blocked ${sheetName}${range ? `!${range}` : ""} because the target is too large for a safe first-pass style plan.`,
      proof: range ? [{ sheetName, range, label: "oversized visual readability target" }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "ask_user",
      warnings: ["Retry with a smaller table/range or select the header and data range to style."]
    };
  }
  if (!detected.headerRange || detected.headerRow === undefined) {
    return {
      status: "NEEDS_INPUT",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: "Could not confidently detect a header row for visual readability styling.",
      proof: range ? [{ sheetName, range, label: "ambiguous visual readability target" }] : [],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "ask_user",
      warnings: [
        "Select the table/header range and retry, or provide target.range with the header row included.",
        "No visual styling operations were prepared."
      ]
    };
  }
  return undefined;
}

function visualReadabilityFormulaCheckRanges(detected: VisualReadabilityDetectedStructure): string[] {
  const ranges = detected.formulaRanges.length > 0
    ? detected.formulaRanges
    : detected.formulaColumns.map((column) => columnTargetFromDetected(detected, column));
  return uniqueDefined(ranges).slice(0, 12);
}

function compileVisualReadabilityOperations(
  workbookId: WorkbookId,
  sheetName: string,
  rules: VisualReadabilityRule[],
  validationSuggestions: VisualReadabilityValidationSuggestion[],
  options: {
    preserveExistingStyle: boolean;
    stylePreservationMode: VisualReadabilityStylePreservationMode;
    allowReplaceConditionalFormatting: boolean;
    allowReplaceDataValidation: boolean;
    applySuggestionBuckets: VisualReadabilitySuggestionBucket[];
    preservation: VisualReadabilityPreservationContext;
  }
): VisualReadabilityCompiledOperationSet {
  const styleEntries: Array<{ target: A1Range; style: NonNullable<RangeSnapshot["style"]>; preserveValues: true }> = [];
  const numberFormatEntries: Array<{ target: A1Range; numberFormat: string[][]; preserveValues: true }> = [];
  const operations: ExcelOperation[] = [];
  const skipped: VisualReadabilityCompiledOperationSet["skipped"] = [];
  const targetFor = (rule: VisualReadabilityRule): A1Range => ({ workbookId, sheetName, address: rule.target });
  const bucketEnabled = (bucket: VisualReadabilitySuggestionBucket) => options.applySuggestionBuckets.includes(bucket);
  const freezePanes: { rows?: number; columns?: number; reasons: string[] } = { reasons: [] };
  for (const rule of rules) {
    const retargetedRule = visualRuleWithPreservationSafeTarget(rule, options);
    if (!retargetedRule) {
      skipped.push({ ruleId: rule.id, target: rule.target, reason: "Target is fully inside a protected summary/template style area and was skipped." });
      continue;
    }
    const preservationSkip = visualRulePreservationSkip(retargetedRule, options);
    if (preservationSkip) {
      skipped.push(preservationSkip);
      continue;
    }
    if (retargetedRule.kind === "width") {
      const width = visualRuleWidth(retargetedRule.value);
      if (width === undefined) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Width rule did not include a supported width value." });
      } else {
        styleEntries.push({ target: targetFor(retargetedRule), style: { columnWidth: width }, preserveValues: true });
      }
      continue;
    }
    if (retargetedRule.kind === "alignment") {
      const alignment = visualRuleAlignment(retargetedRule.value);
      if (Object.keys(alignment).length === 0) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Alignment rule did not include supported alignment values." });
      } else {
        styleEntries.push({ target: targetFor(retargetedRule), style: alignment, preserveValues: true });
      }
      continue;
    }
    if (retargetedRule.kind === "wrap") {
      if (!bucketEnabled("layout")) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Wrap text is actionable through the layout bucket; no layout bucket was requested." });
      } else {
        styleEntries.push({ target: targetFor(retargetedRule), style: { wrapText: booleanValue(retargetedRule.value) ?? true }, preserveValues: true });
      }
      continue;
    }
    if (retargetedRule.kind === "height") {
      const rowHeight = visualRuleHeight(retargetedRule.value);
      if (!bucketEnabled("layout")) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Row height is actionable through the layout bucket; no layout bucket was requested." });
      } else if (rowHeight === undefined) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Height rule did not include a supported row height value." });
      } else {
        styleEntries.push({ target: targetFor(retargetedRule), style: { rowHeight }, preserveValues: true });
      }
      continue;
    }
    if (retargetedRule.kind === "number_format") {
      const format = stringValue(retargetedRule.value);
      const matrix = format ? repeatedNumberFormatMatrix(retargetedRule.target, format) : undefined;
      if (!matrix) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Number-format rule did not include a supported format or bounded target." });
      } else {
        numberFormatEntries.push({ target: targetFor(retargetedRule), numberFormat: matrix, preserveValues: true });
      }
      continue;
    }
    if (retargetedRule.kind === "fill" || retargetedRule.kind === "font" || retargetedRule.kind === "border") {
      if (retargetedRule.scope === "conditional_range") {
        const conditionalRule = visualConditionalFormattingRule(retargetedRule);
        if (!conditionalRule) {
          skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Conditional rule could not be expressed as a custom formula." });
        } else if (options.preserveExistingStyle && !options.allowReplaceConditionalFormatting) {
          operations.push({
            kind: "range.write_conditional_formatting",
            operationId: makeId<OperationId>("op"),
            workbookId,
            destructiveLevel: "format",
            reason: retargetedRule.reason,
            target: targetFor(retargetedRule),
            rule: conditionalRule
          });
        } else {
          operations.push({
            kind: "range.write_conditional_formatting",
            operationId: makeId<OperationId>("op"),
            workbookId,
            destructiveLevel: "format",
            reason: retargetedRule.reason,
            target: targetFor(retargetedRule),
            rule: conditionalRule
          });
        }
      } else {
        const style = visualRuleStyle(retargetedRule);
        if (Object.keys(style).length === 0) {
          skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Style rule did not include supported style properties." });
        } else {
          styleEntries.push({ target: targetFor(retargetedRule), style, preserveValues: true });
        }
      }
      continue;
    }
    if (retargetedRule.kind === "filter") {
      skipped.push({
        ruleId: retargetedRule.id,
        target: retargetedRule.target,
        reason: "Filter affordance is already provided by the detected Excel table and was preserved without reapplying worksheet AutoFilter."
      });
      continue;
    }
    if (retargetedRule.kind === "freeze") {
      if (!bucketEnabled("freeze_panes")) {
        skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Freeze panes are actionable through the freeze_panes bucket; no freeze_panes bucket was requested." });
      } else {
        const freeze = visualRuleFreezePanes(retargetedRule.value);
        if (!freeze) {
          skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: "Freeze rule did not include supported row or column counts." });
        } else {
          if (freeze.rows !== undefined) freezePanes.rows = freeze.rows;
          if (freeze.columns !== undefined) freezePanes.columns = freeze.columns;
          freezePanes.reasons.push(retargetedRule.reason);
        }
      }
      continue;
    }
    skipped.push({ ruleId: retargetedRule.id, target: retargetedRule.target, reason: `${retargetedRule.kind} is preview-only until a matching workbook operation is available.` });
  }
  if (bucketEnabled("validation")) {
    for (const suggestion of validationSuggestions) {
      if (suggestion.existingValidation === "preserved" && !options.allowReplaceDataValidation) {
        skipped.push({ ruleId: suggestion.id, target: suggestion.target, reason: "Existing data validation on the target is preserved by default." });
        continue;
      }
      operations.push({
        kind: "range.write_data_validation",
        operationId: makeId<OperationId>("op"),
        workbookId,
        destructiveLevel: "format",
        reason: suggestion.reason,
        target: { workbookId, sheetName, address: suggestion.target },
        validation: {
          type: "list",
          source: suggestion.source,
          inCellDropDown: true,
          ignoreBlanks: true
        }
      });
    }
  }
  if (freezePanes.rows !== undefined || freezePanes.columns !== undefined) {
    operations.push({
      kind: "sheet.freeze_panes",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: freezePanes.reasons.join(" "),
      sheetName,
      ...(freezePanes.rows !== undefined ? { rows: freezePanes.rows } : {}),
      ...(freezePanes.columns !== undefined ? { columns: freezePanes.columns } : {})
    });
  }
  if (styleEntries.length > 0) {
    operations.unshift({
      kind: "range.write_styles_many",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: "Apply visual readability style rules.",
      entries: styleEntries
    });
  }
  if (numberFormatEntries.length > 0) {
    operations.push({
      kind: "range.write_number_formats_many",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: "Apply visual readability number formats.",
      entries: numberFormatEntries
    });
  }
  return { operations, skipped };
}

function visualRulePreservationSkip(rule: VisualReadabilityRule, options: { preserveExistingStyle: boolean; stylePreservationMode: VisualReadabilityStylePreservationMode; allowReplaceConditionalFormatting: boolean; preservation: VisualReadabilityPreservationContext }): { ruleId: string; target: string; reason: string } | undefined {
  if (visualTargetOverlapsAny(rule.target, options.preservation.protectedRanges)) {
    return { ruleId: rule.id, target: rule.target, reason: "Target overlaps a protected range and was skipped." };
  }
  if (visualTargetOverlapsAny(rule.target, options.preservation.mergedRanges)) {
    return { ruleId: rule.id, target: rule.target, reason: "Target overlaps merged cells and was skipped." };
  }
  if (options.preservation.hiddenColumns.some((column) => column !== "sheet" && visualRuleTouchesColumn(rule.target, column))) {
    return { ruleId: rule.id, target: rule.target, reason: "Target overlaps hidden columns and was skipped." };
  }
  if (rule.kind === "filter" || rule.kind === "freeze") {
    return undefined;
  }
  if (options.preserveExistingStyle && options.stylePreservationMode !== "none") {
    const preservedStyleRanges = options.stylePreservationMode === "strict"
      ? uniqueDefined([...options.preservation.existingStyleRanges, ...options.preservation.protectedStyleRanges])
      : options.preservation.protectedStyleRanges;
    if (visualTargetOverlapsAny(rule.target, preservedStyleRanges)) {
      return { ruleId: rule.id, target: rule.target, reason: "Target overlaps a protected summary/template style area and was skipped." };
    }
  }
  if (rule.scope === "conditional_range" && options.preserveExistingStyle && !options.allowReplaceConditionalFormatting && visualTargetOverlapsAny(rule.target, options.preservation.existingConditionalFormattingRanges)) {
    return { ruleId: rule.id, target: rule.target, reason: "Existing conditional formatting on the target is preserved by default." };
  }
  return undefined;
}

function visualRuleWithPreservationSafeTarget(
  rule: VisualReadabilityRule,
  options: {
    preserveExistingStyle: boolean;
    stylePreservationMode: VisualReadabilityStylePreservationMode;
    preservation: VisualReadabilityPreservationContext;
  }
): VisualReadabilityRule | undefined {
  if (!options.preserveExistingStyle || options.stylePreservationMode === "none") {
    return rule;
  }
  if (rule.kind === "filter" || rule.kind === "freeze") {
    return rule;
  }
  const protectedStyleRanges = options.stylePreservationMode === "strict"
    ? uniqueDefined([...options.preservation.existingStyleRanges, ...options.preservation.protectedStyleRanges])
    : options.preservation.protectedStyleRanges;
  const safeTarget = visualTargetExcludingLeadingProtectedRows(rule.target, protectedStyleRanges);
  return safeTarget ? { ...rule, target: safeTarget } : undefined;
}

function visualTargetExcludingLeadingProtectedRows(target: string, protectedRanges: string[]): string | undefined {
  const parsedTarget = tryParseA1Address(stripSheetName(target));
  if (!parsedTarget) {
    return target;
  }
  let startRow = parsedTarget.startRow;
  for (const protectedRange of protectedRanges) {
    const parsedProtected = tryParseA1Address(stripSheetName(protectedRange));
    if (!parsedProtected || !rangesOverlapAddresses(addressFromBounds(startRow, parsedTarget.startColumn, parsedTarget.endRow - startRow + 1, parsedTarget.endColumn - parsedTarget.startColumn + 1), stripSheetName(protectedRange))) {
      continue;
    }
    const coversTargetStartColumns = parsedProtected.startColumn <= parsedTarget.startColumn && parsedProtected.endColumn >= parsedTarget.endColumn;
    const startsBeforeOrAtTarget = parsedProtected.startRow <= startRow;
    if (coversTargetStartColumns && startsBeforeOrAtTarget && parsedProtected.endRow >= startRow) {
      startRow = parsedProtected.endRow + 1;
    }
  }
  if (startRow > parsedTarget.endRow) {
    return undefined;
  }
  return addressFromBounds(startRow, parsedTarget.startColumn, parsedTarget.endRow - startRow + 1, parsedTarget.endColumn - parsedTarget.startColumn + 1);
}

function visualReadabilityZeroOperationWarnings(skipped: VisualReadabilityCompiledOperationSet["skipped"]): string[] {
  if (skipped.length === 0) {
    return ["No apply-ready visual operations were produced."];
  }
  const reasons = skipped.map((entry) => entry.reason).join(" ");
  const warnings: string[] = ["No apply-ready visual operations were produced; do not apply this preview."];
  if (/protected/.test(reasons)) {
    warnings.push("Some rules were blocked by protected or template-style areas.");
  }
  if (/bucket/.test(reasons)) {
    warnings.push("Some rules require an explicit opt-in suggestion bucket such as layout, validation, or freeze_panes.");
  }
  if (/unsupported|preview-only|operation schema/.test(reasons)) {
    warnings.push("Some rules are preview-only because no supported workbook operation exists for them yet.");
  }
  return warnings;
}

function visualTargetOverlapsAny(target: string, ranges: string[]): boolean {
  const normalizedTarget = stripSheetName(target);
  if (!tryParseA1Address(normalizedTarget)) {
    return false;
  }
  return ranges.some((range) => {
    const normalizedRange = stripSheetName(range);
    return tryParseA1Address(normalizedRange) ? rangesOverlapAddresses(normalizedTarget, normalizedRange) : false;
  });
}

function visualRuleTouchesColumn(target: string, column: string): boolean {
  const parsed = tryParseA1Address(stripSheetName(target));
  const columnIndex = columnToNumber(column);
  return parsed ? parsed.startColumn <= columnIndex && parsed.endColumn >= columnIndex : new RegExp(`^${column}(?:\\d|:|$)`, "i").test(stripSheetName(target));
}

function visualRuleWidth(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  return numberValue(record.preferred ?? record.max ?? record.min);
}

function visualRuleHeight(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  return numberValue(record.preferred ?? record.height ?? record.max ?? record.min);
}

function visualRuleFreezePanes(value: unknown): { rows?: number; columns?: number } | undefined {
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  const rows = visualFreezeCount(record.rows ?? record.rowCount ?? record.row);
  const columns = visualFreezeCount(record.columns ?? record.columnCount ?? record.column);
  const freeze = stripUndefinedRecord({ rows, columns }) as { rows?: number; columns?: number };
  return freeze.rows !== undefined || freeze.columns !== undefined ? freeze : undefined;
}

function visualRuleAlignment(value: unknown): NonNullable<RangeSnapshot["style"]> {
  if (!value || typeof value !== "object") {
    return {};
  }
  const record = value as Record<string, unknown>;
  const style: NonNullable<RangeSnapshot["style"]> = {};
  const horizontalAlignment = stringValue(record.horizontalAlignment);
  if (horizontalAlignment) style.horizontalAlignment = horizontalAlignment;
  const verticalAlignment = stringValue(record.verticalAlignment);
  if (verticalAlignment) style.verticalAlignment = verticalAlignment;
  return style;
}

function visualRuleStyle(rule: VisualReadabilityRule): NonNullable<RangeSnapshot["style"]> {
  if (!rule.value || typeof rule.value !== "object") {
    return {};
  }
  const record = rule.value as Record<string, unknown>;
  const style: NonNullable<RangeSnapshot["style"]> = {};
  const fillColor = colorString(record.fillColor);
  if (fillColor) style.fillColor = fillColor;
  const fontColor = colorString(record.fontColor);
  if (fontColor) style.fontColor = fontColor;
  const fontBold = booleanValue(record.fontBold);
  if (fontBold !== undefined) style.fontBold = fontBold;
  const fontName = stringValue(record.fontName);
  if (fontName) style.fontName = fontName;
  const fontSize = numberValue(record.fontSize);
  if (fontSize !== undefined) style.fontSize = fontSize;
  const borders = visualRuleBorders(record);
  if (borders) style.borders = borders;
  return style;
}

function visualRuleBorders(record: Record<string, unknown>): NonNullable<RangeSnapshot["style"]>["borders"] | undefined {
  const borders: Record<string, unknown> = {};
  for (const edge of ["edgeTop", "edgeBottom", "edgeLeft", "edgeRight", "insideHorizontal", "insideVertical"]) {
    const value = record[edge];
    if (value && typeof value === "object") {
      borders[edge] = value;
    }
  }
  return Object.keys(borders).length > 0 ? borders as NonNullable<RangeSnapshot["style"]>["borders"] : undefined;
}

function repeatedNumberFormatMatrix(address: string, format: string): string[][] | undefined {
  const parsed = tryParseA1Address(stripSheetName(address));
  if (!parsed) {
    return undefined;
  }
  const rows = parsed.endRow - parsed.startRow + 1;
  const columns = parsed.endColumn - parsed.startColumn + 1;
  if (rows < 1 || columns < 1 || rows * columns > 20_000) {
    return undefined;
  }
  return Array.from({ length: rows }, () => Array.from({ length: columns }, () => format));
}

function visualConditionalFormattingRule(rule: VisualReadabilityRule): Extract<ExcelOperation, { kind: "range.write_conditional_formatting" }>["rule"] | undefined {
  const parsed = tryParseA1Address(stripSheetName(rule.target));
  if (!parsed) {
    return undefined;
  }
  const column = visualRuleColumnFromId(rule.id);
  if (!column) {
    return undefined;
  }
  const firstRow = parsed.startRow;
  const fillColor = rule.value && typeof rule.value === "object" ? colorString((rule.value as Record<string, unknown>).fillColor) : undefined;
  const style: NonNullable<RangeSnapshot["style"]> = { fillColor: fillColor ?? "#FFF2CC" };
  const formula = rule.id.includes("formula_error")
    ? `=ISERROR($${column}${firstRow})`
    : `=AND(COUNTA($${numberToColumn(parsed.startColumn)}${firstRow}:$${numberToColumn(parsed.endColumn)}${firstRow})>0,$${column}${firstRow}="")`;
  return { type: "custom", formula, style };
}

function visualRuleColumnFromId(ruleId: string): string | undefined {
  const match = /^(?:column|conditional)\.([A-Z]+)\./.exec(ruleId);
  return match?.[1];
}

function visualFormulaSnapshotFromBatchResult(formulaRanges: string[], result: unknown): Map<string, string> {
  const snapshot = new Map<string, string>();
  const record = result && typeof result === "object" ? result as Record<string, unknown> : {};
  const entries = Array.isArray(record.readData)
    ? record.readData
    : Array.isArray(record.data)
      ? record.data
      : [];
  for (const [rangeIndex, entry] of entries.entries()) {
    const entryRecord = entry && typeof entry === "object" ? entry as Record<string, unknown> : {};
    const formulas = entryRecord.snapshot && typeof entryRecord.snapshot === "object" && Array.isArray((entryRecord.snapshot as Record<string, unknown>).formulas)
      ? (entryRecord.snapshot as { formulas: unknown[][] }).formulas
      : [];
    const range = formulaRanges[rangeIndex] ?? `range_${rangeIndex}`;
    for (const [rowIndex, row] of formulas.entries()) {
      if (!Array.isArray(row)) {
        continue;
      }
      for (const [columnIndex, value] of row.entries()) {
        if (typeof value === "string" && value.length > 0) {
          snapshot.set(`${range}:${rowIndex}:${columnIndex}`, value);
        }
      }
    }
  }
  return snapshot;
}

function compareVisualFormulaSnapshots(before: Map<string, string>, after: Map<string, string>): { checkedCount: number; changedCount: number } {
  const keys = new Set([...before.keys(), ...after.keys()]);
  let changedCount = 0;
  for (const key of keys) {
    if (before.get(key) !== after.get(key)) {
      changedCount += 1;
    }
  }
  return { checkedCount: keys.size, changedCount };
}

function visualReadabilityStyleDepth(value: unknown): VisualReadabilityStyleDepth {
  const raw = stringValue(value)?.toLowerCase();
  return raw === "basic" || raw === "comprehensive" ? raw : "standard";
}

function visualReadabilityDensity(value: unknown): VisualReadabilityDensity {
  const raw = stringValue(value)?.toLowerCase();
  return raw === "compact" || raw === "presentation" ? raw : "comfortable";
}

function visualReadabilityPresentationMode(value: unknown): VisualReadabilityPresentationMode | undefined {
  const raw = stringValue(value)?.toLowerCase();
  if (raw === "working_sheet" || raw === "print_ready" || raw === "executive_report") {
    return raw;
  }
  return undefined;
}

function visualReadabilityStylePreservationMode(value: unknown, preserveExistingStyle: unknown): VisualReadabilityStylePreservationMode {
  const raw = stringValue(value)?.toLowerCase().replace(/[\s-]+/g, "_");
  if (raw === "protected_regions" || raw === "strict" || raw === "none") {
    return raw;
  }
  if (booleanValue(preserveExistingStyle) === false) {
    return "none";
  }
  return "protected_regions";
}

function visualReadabilitySuggestionBuckets(value: unknown): VisualReadabilitySuggestionBucket[] {
  const rawValues = Array.isArray(value) ? value : typeof value === "string" ? value.split(/[, ]+/) : [];
  const allowed = new Set<VisualReadabilitySuggestionBucket>(["layout", "validation", "reference_style", "formula_helpers", "structure", "freeze_panes", "print_settings"]);
  const buckets: VisualReadabilitySuggestionBucket[] = [];
  for (const raw of rawValues) {
    const normalized = stringValue(raw)?.trim().toLowerCase().replace(/[-\s]+/g, "_") as VisualReadabilitySuggestionBucket | undefined;
    if (normalized && allowed.has(normalized) && !buckets.includes(normalized)) {
      buckets.push(normalized);
    }
  }
  return buckets;
}

function visualReadabilityFreezePanes(value: unknown, rowValue: unknown, columnValue: unknown, request: string): { rows?: number; columns?: number } | undefined {
  const record = value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
  const rows = visualFreezeCount(record.rows ?? record.rowCount ?? record.row ?? rowValue);
  const columns = visualFreezeCount(record.columns ?? record.columnCount ?? record.column ?? columnValue);
  const inferredColumns = columns ?? visualFreezeColumnCountFromRequest(request);
  const inferredRows = rows ?? visualFreezeRowCountFromRequest(request);
  const freezePanes = stripUndefinedRecord({ rows: inferredRows, columns: inferredColumns }) as { rows?: number; columns?: number };
  return freezePanes.rows !== undefined || freezePanes.columns !== undefined ? freezePanes : undefined;
}

function freezePanesFromInput(input: AgentRunInput): { rows?: number; columns?: number } | undefined {
  if (/\bunfreeze\b|\bremove\b.*\bfreeze|\bclear\b.*\bfreeze/i.test(input.request)) {
    return { rows: 0, columns: 0 };
  }
  const values = input.values as Record<string, unknown> | undefined;
  const nested = values?.freezePanes && typeof values.freezePanes === "object" && !Array.isArray(values.freezePanes)
    ? values.freezePanes as Record<string, unknown>
    : {};
  return visualReadabilityFreezePanes(
    values?.freezePanes,
    nested.rows ?? nested.rowCount ?? nested.row ?? values?.freezeRows ?? values?.rows,
    nested.columns ?? nested.columnCount ?? nested.column ?? values?.freezeColumns ?? values?.columns,
    input.request
  );
}

function freezePanesSheetName(metadata: WorkbookMetadata, input: AgentRunInput): string | undefined {
  const requested = stringValue(input.target?.sheetName);
  if (requested && !sameText(requested, "active") && !sameText(requested, "active_sheet")) {
    return metadata.sheets.find((sheet) => sameText(sheet.name, requested))?.name ?? requested;
  }
  return metadata.workbook.activeSheet ?? metadata.sheets[0]?.name;
}

function visualFreezeCount(value: unknown): number | undefined {
  if (value === true) return 1;
  if (value === false) return 0;
  const count = numberValue(value);
  return count !== undefined && Number.isFinite(count) && count >= 0 ? Math.floor(count) : undefined;
}

function visualFreezeColumnCountFromRequest(request: string): number | undefined {
  const lower = request.toLowerCase();
  const firstColumn = /\bfreeze\b.*\b(?:first\s+)?(?:col|column)\b/.test(lower) || /\bfreeze\b.*\bcolumn\s+a\b/.test(lower);
  if (!firstColumn) {
    return undefined;
  }
  const count = /\bfreeze\b.*\bfirst\s+(\d+)\s+(?:cols?|columns?)\b/.exec(lower)?.[1];
  return count ? Number.parseInt(count, 10) : 1;
}

function visualFreezeRowCountFromRequest(request: string): number | undefined {
  const lower = request.toLowerCase();
  const firstRow = /\bfreeze\b.*\b(?:first\s+|top\s+)?row\b/.test(lower) || /\bfreeze\b.*\bheader\b/.test(lower);
  if (!firstRow) {
    return undefined;
  }
  const count = /\bfreeze\b.*\b(?:first|top)\s+(\d+)\s+rows?\b/.exec(lower)?.[1];
  return count ? Number.parseInt(count, 10) : 1;
}

function visualReadabilityReferenceStyle(value: unknown): VisualReadabilityReferenceStyleOption | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  const sheetName = stringValue(record.sheetName ?? record.sheet ?? record.referenceSheetName ?? record.referenceSheet);
  if (!sheetName) {
    return undefined;
  }
  return {
    sheetName,
    adaptToTargetStructure: booleanValue(record.adaptToTargetStructure) ?? true,
    preserveTargetValues: booleanValue(record.preserveTargetValues) ?? true,
    preserveFormulas: booleanValue(record.preserveFormulas) ?? true
  };
}

function compileVisualReadabilityPlan(
  metadata: WorkbookMetadata,
  sheet: WorkbookMetadata["sheets"][number],
  detected: VisualReadabilityDetectedStructure,
  columns: VisualReadabilityColumnRole[],
  options: ReturnType<typeof visualReadabilityOptionsFromInput>,
  sheetType: VisualReadabilitySheetType
): VisualReadabilityPlanPreview {
  const theme = visualReadabilityTheme(options.density);
  const rules: VisualReadabilityRule[] = [];
  if (detected.headerRange) {
    rules.push({
      id: "layout.header_style",
      scope: "row",
      target: detected.headerRange,
      kind: "fill",
      value: { fillColor: theme.fills.header, fontBold: true, fontColor: theme.colors.headerText, borderBottom: theme.borders.headerBottom },
      risk: "low",
      reason: "Apply calm header styling to create visual hierarchy."
    });
    rules.push({
      id: "layout.header_alignment",
      scope: "row",
      target: detected.headerRange,
      kind: "alignment",
      value: { verticalAlignment: "Center" },
      risk: "low",
      reason: "Keep header labels vertically centered."
    });
  }
  if (detected.dataRange && detected.hasFilter) {
    const filterTarget = visualFilterTarget(detected);
    rules.push({
      id: "layout.filter",
      scope: "table",
      target: filterTarget,
      kind: "filter",
      value: { enabled: true },
      risk: "low",
      reason: "Preserve or enable filter affordance for table scanning."
    });
  }
  if (detected.headerRow !== undefined) {
    const freezeRows = options.freezePanes?.rows ?? detected.headerRow;
    rules.push({
      id: "layout.freeze_header",
      scope: "sheet",
      target: sheet.name,
      kind: "freeze",
      value: { rows: freezeRows },
      risk: "low",
      reason: "Keep the detected header row visible while scrolling."
    });
  }
  if (options.freezePanes?.columns !== undefined) {
    rules.push({
      id: "layout.freeze_columns",
      scope: "sheet",
      target: sheet.name,
      kind: "freeze",
      value: { columns: options.freezePanes.columns },
      risk: "low",
      reason: `Keep the first ${options.freezePanes.columns} column(s) visible while scrolling.`
    });
  }
  if (options.freezePanes?.rows !== undefined && detected.headerRow === undefined) {
    rules.push({
      id: "layout.freeze_rows",
      scope: "sheet",
      target: sheet.name,
      kind: "freeze",
      value: { rows: options.freezePanes.rows },
      risk: "low",
      reason: `Keep the first ${options.freezePanes.rows} row(s) visible while scrolling.`
    });
  }

  for (const column of columns) {
    const target = columnTargetFromDetected(detected, column.column);
    const width = visualWidthForRole(column.role);
    rules.push({
      id: `column.${column.column}.width`,
      scope: "column",
      target,
      kind: "width",
      value: width,
      risk: "low",
      reason: `Set ${column.header || column.column} width from role ${column.role}.`
    });
    rules.push({
      id: `column.${column.column}.alignment`,
      scope: "column",
      target,
      kind: "alignment",
      value: visualAlignmentForRole(column.role),
      risk: "low",
      reason: `Align ${column.header || column.column} from role ${column.role}.`
    });
    const numberFormat = visualNumberFormatForRole(column.role);
    if (numberFormat) {
      rules.push({
        id: `column.${column.column}.number_format`,
        scope: "column",
        target,
        kind: "number_format",
        value: numberFormat,
        risk: "low",
        reason: `Normalize ${column.header || column.column} display format.`
      });
    }
    if (column.role === "notes") {
      rules.push({
        id: `column.${column.column}.wrap`,
        scope: "column",
        target,
        kind: "wrap",
        value: true,
        risk: "low",
        reason: "Wrap long notes or descriptions without changing values."
      });
    }
    if (column.role === "formula_output") {
      rules.push({
        id: `column.${column.column}.formula_fill`,
        scope: "column",
        target,
        kind: "fill",
        value: { fillColor: theme.fills.formula },
        risk: "low",
        reason: "Distinguish calculated formula output from input cells."
      });
    }
  }

  if (options.styleDepth !== "basic") {
    for (const group of visualColumnGroups(columns, detected)) {
      rules.push({
        id: `group.${normalizeOperationName(group.label)}.${group.startColumn}_${group.endColumn}`,
        scope: "group",
        target: group.target,
        kind: "border",
        value: { edgeLeft: theme.borders.groupSeparator, fillColor: theme.fills.group },
        risk: "low",
        reason: `Softly separate ${group.label} columns.`
      });
    }
    for (const column of columns.filter((candidate) => ["date", "entity", "status"].includes(candidate.role)).slice(0, 4)) {
      rules.push({
        id: `conditional.${column.column}.missing_required`,
        scope: "conditional_range",
        target: detected.dataRange ?? columnTargetFromDetected(detected, column.column),
        kind: "fill",
        value: { condition: `${column.column} is blank on nonblank record rows`, fillColor: theme.fills.warning },
        risk: "low",
        reason: `Highlight missing likely-required ${column.header || column.column} values.`
      });
    }
    for (const formulaColumn of detected.formulaColumns.slice(0, 4)) {
      rules.push({
        id: `conditional.${formulaColumn}.formula_error`,
        scope: "conditional_range",
        target: detected.dataRange ?? columnTargetFromDetected(detected, formulaColumn),
        kind: "fill",
        value: { condition: `${formulaColumn} contains an Excel error`, fillColor: theme.fills.danger },
        risk: "low",
        reason: "Highlight formula errors without replacing formulas."
      });
    }
  }

  const validationSuggestions = shouldCompileVisualValidationSuggestions(options)
    ? compileVisualValidationSuggestions(detected, columns)
    : [];
  const formulaSuggestions = shouldCompileVisualFormulaSuggestions(options)
    ? compileVisualFormulaSuggestions(detected, columns)
    : [];
  const referenceStyleSuggestions = compileVisualReferenceStyleSuggestions(metadata, sheet, detected, columns, options);
  const printSuggestions = compileVisualPrintSuggestions(detected, options, sheetType);
  const ruleScopes = countVisualRuleScopes(rules);
  const counts = {
    totalRules: rules.length,
    layoutChanges: rules.filter((rule) => rule.scope === "sheet" || rule.scope === "table" || rule.id.startsWith("layout.")).length,
    groupRules: ruleScopes.group,
    columnRules: ruleScopes.column,
    conditionalRules: ruleScopes.conditional_range,
    rowRules: ruleScopes.row,
    cellRules: ruleScopes.cell,
    validationSuggestions: validationSuggestions.length,
    formulaSuggestions: formulaSuggestions.length,
    referenceStyleSuggestions: referenceStyleSuggestions.length,
    printSuggestions: printSuggestions.length
  };
  return {
    compilerStatus: "preview_compiled_apply_pending",
    summary: visualPlanSummary(sheetType, options, counts),
    counts,
    ruleScopes,
    rules: rules.slice(0, 200),
    validationSuggestions,
    formulaSuggestions,
    referenceStyleSuggestions,
    printSuggestions,
    previewExamples: rules.slice(0, 8).map((rule) => ({
      range: rule.target,
      before: "current workbook formatting",
      after: visualRuleAfterSummary(rule),
      ruleId: rule.id
    })),
    theme: {
      name: "office_clean",
      font: theme.font,
      density: options.density
    }
  };
}

function visualReadabilityTheme(density: VisualReadabilityDensity) {
  return {
    font: { family: "Calibri", bodySize: density === "compact" ? 10 : 11, headerSize: density === "presentation" ? 12 : 11 },
    colors: { headerText: "#1F2937" },
    fills: {
      header: "#D9EAF7",
      group: "#F3F6FA",
      formula: "#EEF2FF",
      warning: "#FFF2CC",
      danger: "#FCE4D6"
    },
    borders: {
      headerBottom: { style: "continuous", weight: "thin", color: "#8EA9DB" },
      groupSeparator: { style: "continuous", weight: "thin", color: "#B7C9E2" }
    }
  };
}

function shouldCompileVisualValidationSuggestions(options: ReturnType<typeof visualReadabilityOptionsFromInput>): boolean {
  return options.styleDepth === "comprehensive" || options.allowValidationSuggestions || options.applySuggestionBuckets.includes("validation");
}

function shouldCompileVisualFormulaSuggestions(options: ReturnType<typeof visualReadabilityOptionsFromInput>): boolean {
  return options.styleDepth === "comprehensive" || options.allowFormulaSuggestions || options.applySuggestionBuckets.includes("formula_helpers");
}

function compileVisualValidationSuggestions(detected: VisualReadabilityDetectedStructure, columns: VisualReadabilityColumnRole[]): VisualReadabilityValidationSuggestion[] {
  return columns
    .filter((column) => column.role === "status" || column.role === "category")
    .slice(0, 6)
    .map((column) => ({
      id: `validation.${column.column}.dropdown`,
      target: columnTargetFromDetected(detected, column.column),
      source: visualValidationOptionsForRole(column.role, column.header),
      risk: "medium",
      reason: `Suggest a dropdown for ${column.header || column.column}; preview only and not applied by visual styling.`,
      existingValidation: detected.existingDataValidationRanges.some((range) => rangesOverlapAddresses(stripSheetName(range), stripSheetName(columnTargetFromDetected(detected, column.column)))) ? "preserved" : "not_detected"
    }));
}

function visualValidationOptionsForRole(role: string, header: string): string[] {
  const normalized = normalizeHeaderName(header);
  if (role === "status") {
    if (/\b(payment|invoice|billing)\b/i.test(normalized)) {
      return ["Open", "Paid", "Overdue", "Disputed"];
    }
    return ["Open", "In Progress", "Blocked", "Done"];
  }
  if (/\bpriority\b/i.test(normalized)) {
    return ["Low", "Medium", "High", "Critical"];
  }
  return ["Option 1", "Option 2", "Option 3"];
}

function compileVisualFormulaSuggestions(detected: VisualReadabilityDetectedStructure, columns: VisualReadabilityColumnRole[]): VisualReadabilityFormulaSuggestion[] {
  const suggestions: VisualReadabilityFormulaSuggestion[] = [];
  const dateColumn = columns.find((column) => column.role === "date");
  const statusColumn = columns.find((column) => column.role === "status");
  const lastColumn = columns[columns.length - 1]?.column ?? columnsFromAddress(detected.usedRange ?? detected.headerRange ?? "A1:A1").at(-1) ?? "A";
  if (dateColumn && statusColumn) {
    suggestions.push({
      id: "formula.overdue_flag",
      suggestedColumnAfter: lastColumn,
      formulaName: "Overdue flag",
      formulaExample: `=AND($${dateColumn.column}2<TODAY(),$${statusColumn.column}2<>"Done")`,
      risk: "medium",
      reason: "Could add an overdue helper column, but formula insertion requires separate confirmation."
    });
  }
  const moneyColumns = columns.filter((column) => column.role === "money" || column.role === "number");
  if (moneyColumns.length >= 2) {
    suggestions.push({
      id: "formula.variance",
      suggestedColumnAfter: lastColumn,
      formulaName: "Variance",
      formulaExample: `=$${moneyColumns[0]!.column}2-$${moneyColumns[1]!.column}2`,
      risk: "medium",
      reason: "Could add a variance calculation, but formulas are not inserted by visual styling."
    });
  }
  return suggestions.slice(0, 6);
}

function visualFilterTarget(detected: VisualReadabilityDetectedStructure): string {
  return detected.tableRanges[0] ?? detected.usedRange ?? detected.headerRange ?? detected.dataRange ?? detected.sheetName;
}

function compileVisualReferenceStyleSuggestions(
  metadata: WorkbookMetadata,
  sheet: WorkbookMetadata["sheets"][number],
  detected: VisualReadabilityDetectedStructure,
  columns: VisualReadabilityColumnRole[],
  options: ReturnType<typeof visualReadabilityOptionsFromInput>
): VisualReadabilityReferenceStyleSuggestion[] {
  const reference = options.referenceStyle;
  if (!reference) {
    return [];
  }
  const referenceSheet = metadata.sheets.find((candidate) => sameText(candidate.name, reference.sheetName));
  if (!referenceSheet) {
    return [{
      id: "reference_style.sheet_not_found",
      referenceSheetName: reference.sheetName,
      target: detected.usedRange ?? sheet.name,
      pattern: "reference sheet lookup",
      risk: "high",
      reason: "Reference style sheet was not found, so no style patterns will be adapted.",
      preserveTargetValues: true,
      preserveFormulas: reference.preserveFormulas
    }];
  }
  const referenceHeader = referenceSheet.headers.sort((left, right) => right.confidence - left.confidence)[0];
  const suggestions: VisualReadabilityReferenceStyleSuggestion[] = [];
  if (referenceHeader?.range && detected.headerRange) {
    suggestions.push({
      id: "reference_style.header",
      referenceSheetName: referenceSheet.name,
      target: detected.headerRange,
      pattern: "header fill, font, and bottom border",
      risk: "medium",
      reason: `Adapt header styling from ${referenceSheet.name} to the detected target header without copying values.`,
      preserveTargetValues: true,
      preserveFormulas: reference.preserveFormulas
    });
  }
  if (reference.adaptToTargetStructure && columns.length > 0) {
    suggestions.push({
      id: "reference_style.columns_by_role",
      referenceSheetName: referenceSheet.name,
      target: detected.usedRange ?? detected.headerRange ?? sheet.name,
      pattern: "column widths, alignment, and number formats by semantic role",
      risk: "medium",
      reason: `Map style patterns from ${referenceSheet.name} by column role so extra or missing target columns are handled safely.`,
      preserveTargetValues: true,
      preserveFormulas: reference.preserveFormulas
    });
  }
  if (referenceSheet.kind === sheet.kind || referenceSheet.columnCount === sheet.columnCount) {
    suggestions.push({
      id: "reference_style.layout",
      referenceSheetName: referenceSheet.name,
      target: detected.usedRange ?? sheet.name,
      pattern: "freeze panes, filters, and table layout cues",
      risk: "medium",
      reason: `Reuse compatible layout cues from ${referenceSheet.name}; this remains preview-only in visual readability apply.`,
      preserveTargetValues: true,
      preserveFormulas: reference.preserveFormulas
    });
  }
  return suggestions.slice(0, 6);
}

function compileVisualPrintSuggestions(
  detected: VisualReadabilityDetectedStructure,
  options: ReturnType<typeof visualReadabilityOptionsFromInput>,
  sheetType: VisualReadabilitySheetType
): VisualReadabilityPrintSuggestion[] {
  const mode = options.presentationMode;
  if (!mode || mode === "working_sheet") {
    return [];
  }
  const target = detected.usedRange ?? detected.headerRange ?? detected.sheetName;
  const orientation = sheetType === "numeric_report" || sheetType === "summary_report" || (detected.usedRange && columnsFromAddress(detected.usedRange).length > 8)
    ? "landscape"
    : "portrait";
  const suggestions: VisualReadabilityPrintSuggestion[] = [
    {
      id: "print.orientation",
      target,
      setting: "orientation",
      value: orientation,
      presentationMode: mode,
      risk: "medium",
      reason: "Suggest page orientation for print/export readability; visual apply does not change print settings."
    },
    {
      id: "print.fit_to_width",
      target,
      setting: "fitToWidth",
      value: "1 page wide",
      presentationMode: mode,
      risk: "medium",
      reason: "Suggest fitting the report to one page wide for sharing."
    }
  ];
  if (detected.headerRange) {
    suggestions.push({
      id: "print.repeat_header",
      target: detected.headerRange,
      setting: "repeatRows",
      value: `row ${detected.headerRow ?? 1}`,
      presentationMode: mode,
      risk: "medium",
      reason: "Suggest repeating the detected header row when printed or exported."
    });
  }
  if (mode === "executive_report") {
    suggestions.push({
      id: "print.hide_gridlines",
      target,
      setting: "showGridlines",
      value: "false",
      presentationMode: mode,
      risk: "medium",
      reason: "Suggest hiding gridlines for cleaner executive presentation output."
    });
  }
  return suggestions.slice(0, 6);
}

function visualWidthForRole(role: string): { min: number; max: number } {
  switch (role) {
    case "date": return { min: 12, max: 14 };
    case "entity": return { min: 18, max: 30 };
    case "money": return { min: 12, max: 16 };
    case "number": return { min: 10, max: 14 };
    case "status": return { min: 12, max: 18 };
    case "notes": return { min: 28, max: 50 };
    case "id": return { min: 10, max: 16 };
    case "category": return { min: 12, max: 20 };
    default: return { min: 12, max: 24 };
  }
}

function visualAlignmentForRole(role: string): { horizontalAlignment: string; verticalAlignment: string } {
  if (role === "money" || role === "number") return { horizontalAlignment: "Right", verticalAlignment: "Center" };
  if (role === "date" || role === "status" || role === "category") return { horizontalAlignment: "Center", verticalAlignment: "Center" };
  if (role === "notes") return { horizontalAlignment: "Left", verticalAlignment: "Top" };
  return { horizontalAlignment: "Left", verticalAlignment: "Center" };
}

function visualNumberFormatForRole(role: string): string | undefined {
  if (role === "date") return "dd/mm/yyyy";
  if (role === "money" || role === "number") return "#,##0.00";
  return undefined;
}

function columnTargetFromDetected(detected: VisualReadabilityDetectedStructure, column: string): string {
  const parsed = detected.dataRange ? tryParseA1Address(stripSheetName(detected.dataRange)) : detected.usedRange ? tryParseA1Address(stripSheetName(detected.usedRange)) : undefined;
  return parsed ? `${column}${parsed.startRow}:${column}${parsed.endRow}` : `${column}:${column}`;
}

function visualColumnGroups(columns: VisualReadabilityColumnRole[], detected: VisualReadabilityDetectedStructure): Array<{ label: string; startColumn: string; endColumn: string; target: string }> {
  const groups: Array<{ label: string; columns: VisualReadabilityColumnRole[] }> = [];
  for (const column of columns) {
    const label = visualGroupLabel(column.role);
    const last = groups[groups.length - 1];
    if (last?.label === label) {
      last.columns.push(column);
    } else {
      groups.push({ label, columns: [column] });
    }
  }
  const parsed = detected.dataRange ? tryParseA1Address(stripSheetName(detected.dataRange)) : detected.usedRange ? tryParseA1Address(stripSheetName(detected.usedRange)) : undefined;
  return groups
    .filter((group) => group.columns.length > 0)
    .map((group) => {
      const startColumn = group.columns[0]!.column;
      const endColumn = group.columns[group.columns.length - 1]!.column;
      return {
        label: group.label,
        startColumn,
        endColumn,
        target: parsed ? `${startColumn}${parsed.startRow}:${endColumn}${parsed.endRow}` : `${startColumn}:${endColumn}`
      };
    });
}

function visualGroupLabel(role: string): string {
  if (role === "date") return "Dates";
  if (role === "money" || role === "number" || role === "formula_output") return "Metrics";
  if (role === "status") return "Status";
  if (role === "notes") return "Notes";
  return "Record Info";
}

function countVisualRuleScopes(rules: VisualReadabilityRule[]): Record<VisualRuleScope, number> {
  return {
    sheet: rules.filter((rule) => rule.scope === "sheet").length,
    table: rules.filter((rule) => rule.scope === "table").length,
    group: rules.filter((rule) => rule.scope === "group").length,
    column: rules.filter((rule) => rule.scope === "column").length,
    conditional_range: rules.filter((rule) => rule.scope === "conditional_range").length,
    row: rules.filter((rule) => rule.scope === "row").length,
    cell: rules.filter((rule) => rule.scope === "cell").length
  };
}

function visualPlanSummary(sheetType: VisualReadabilitySheetType, options: ReturnType<typeof visualReadabilityOptionsFromInput>, counts: VisualReadabilityPlanPreview["counts"]): string[] {
  return [
    `Detected ${sheetType} and selected ${options.profile} profile.`,
    `Compiled ${counts.totalRules} ${options.styleDepth} visual readability rule(s).`,
    "Supported safe operations are apply-ready; risky or unsupported rules remain skipped or preview-only."
  ];
}

function visualRuleAfterSummary(rule: VisualReadabilityRule): string {
  switch (rule.kind) {
    case "width": return "role-based width";
    case "alignment": return "role-based alignment";
    case "number_format": return `number format ${String(rule.value)}`;
    case "filter": return "filters enabled or preserved";
    case "freeze": return "freeze header suggestion";
    case "border": return "soft grouping border/fill";
    case "wrap": return "wrapped long text";
    case "fill": return rule.scope === "conditional_range" ? "conditional highlight rule" : "calm fill styling";
    case "font": return "font styling";
    case "height": return "row height";
  }
}

function detectVisualReadabilityStructure(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number], targetRange?: string): VisualReadabilityDetectedStructure {
  const sheetTables = metadata.tables.filter((table) => table.sheetName === sheet.name);
  const firstTable = sheetTables[0];
  const firstTableRange = firstTable?.range ? tryParseA1Address(stripSheetName(firstTable.range)) : undefined;
  const firstTableHeaderRange = firstTable?.headerRange ? stripSheetName(firstTable.headerRange) : undefined;
  const firstTableHeader = firstTableHeaderRange ? tryParseA1Address(firstTableHeaderRange) : undefined;
  const inferredTableHeaderRange = firstTableRange ? addressFromBounds(firstTableRange.startRow, firstTableRange.startColumn, 1, firstTableRange.endColumn - firstTableRange.startColumn + 1) : undefined;
  const inferredTableDataRange = firstTableRange && firstTableRange.endRow > firstTableRange.startRow
    ? addressFromBounds(firstTableRange.startRow + 1, firstTableRange.startColumn, firstTableRange.endRow - firstTableRange.startRow, firstTableRange.endColumn - firstTableRange.startColumn + 1)
    : undefined;
  const bestHeader = [
    ...sheet.headers,
    ...sheetTables.flatMap((table) => table.headerRange ? [{
      id: `${table.id}:header`,
      sheetName: table.sheetName,
      row: tryParseA1Address(stripSheetName(table.headerRange))?.startRow ?? tryParseA1Address(stripSheetName(table.range))?.startRow ?? 1,
      range: table.headerRange,
      columns: table.columns,
      confidence: 0.95
    } satisfies HeaderMetadata] : [])
  ].sort((left, right) => right.confidence - left.confidence)[0];
  const usedRange = targetRange ?? sheet.usedRange;
  const used = usedRange ? tryParseA1Address(stripSheetName(usedRange)) : undefined;
  const bestHeaderParsed = bestHeader?.range ? tryParseA1Address(stripSheetName(bestHeader.range)) : undefined;
  const bestHeaderLooksLikeWholeTable = Boolean(bestHeaderParsed && firstTableRange && rangesOverlapAddresses(stripSheetName(bestHeader!.range), stripSheetName(firstTable!.range)) && (bestHeaderParsed.endRow - bestHeaderParsed.startRow) >= 1);
  const headerRow = firstTableHeader?.startRow
    ?? (firstTableRange ? firstTableRange.startRow : undefined)
    ?? (bestHeaderLooksLikeWholeTable ? bestHeaderParsed?.startRow : bestHeader?.row);
  const headerRange = firstTableHeaderRange
    ?? inferredTableHeaderRange
    ?? (bestHeaderLooksLikeWholeTable && bestHeaderParsed ? addressFromBounds(bestHeaderParsed.startRow, bestHeaderParsed.startColumn, 1, bestHeaderParsed.endColumn - bestHeaderParsed.startColumn + 1) : bestHeader?.range);
  const dataRange = inferredTableDataRange ?? firstTable?.dataRange ?? (used && headerRow !== undefined && used.endRow > headerRow
    ? addressFromBounds(headerRow + 1, used.startColumn, used.endRow - headerRow, used.endColumn - used.startColumn + 1)
    : usedRange);
  const formulaColumns = uniqueDefined([
    ...columnsForVisualReadability(metadata, sheet, { headerRange, dataRange }).filter((column) => column.inferredType === "formula" || column.role === "formula").map((column) => column.letter),
    ...metadata.formulaRegions.filter((region) => region.sheetName === sheet.name).flatMap((region) => columnsFromAddress(region.range))
  ]);
  const protectedStyleRanges = uniqueDefined([
    ...metadata.summaryBlocks.filter((block) => block.sheetName === sheet.name).map((block) => block.range),
    ...metadata.sections.filter((section) => section.sheetName === sheet.name && (section.kind === "summary" || section.kind === "metadata")).map((section) => section.range)
  ]);
  return {
    sheetName: sheet.name,
    ...(usedRange ? { usedRange } : {}),
    ...(headerRow !== undefined ? { headerRow } : {}),
    ...(headerRange ? { headerRange } : {}),
    ...(dataRange ? { dataRange } : {}),
    tableRanges: sheetTables.map((table) => table.range),
    hasFilter: sheetTables.length > 0,
    hasFreezePane: false,
    mergedRanges: [],
    hiddenRows: [],
    hiddenColumns: sheet.isHidden ? ["sheet"] : [],
    protectedRanges: [],
    existingStyleRanges: protectedStyleRanges,
    protectedStyleRanges,
    formulaColumns,
    formulaRanges: metadata.formulaRegions.filter((region) => region.sheetName === sheet.name).map((region) => region.range),
    totalRows: [],
    subtotalRows: [],
    existingConditionalFormattingRanges: [],
    existingDataValidationRanges: [],
    detectionSource: "metadata",
    confidence: bestHeader ? bestHeader.confidence : firstTable ? 0.9 : usedRange ? 0.65 : 0.35
  };
}

function inferVisualReadabilityColumns(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number], detected: VisualReadabilityDetectedStructure): VisualReadabilityColumnRole[] {
  return columnsForVisualReadability(metadata, sheet, detected).map((column) => ({
    column: column.letter,
    header: column.name,
    role: visualColumnRole(column),
    inferredType: column.inferredType,
    confidence: Math.max(0.45, Math.min(0.98, column.importance ?? 0.7)),
    signals: visualColumnSignals(column)
  }));
}

function columnsForVisualReadability(metadata: WorkbookMetadata, sheet: WorkbookMetadata["sheets"][number], detected: { headerRange?: string | undefined; dataRange?: string | undefined }): ColumnMetadata[] {
  const tableColumns = metadata.tables.find((table) => table.sheetName === sheet.name && (!detected.dataRange || table.dataRange === detected.dataRange || table.range === detected.dataRange))?.columns
    ?? metadata.tables.find((table) => table.sheetName === sheet.name)?.columns;
  if (tableColumns && tableColumns.length > 0) {
    return tableColumns;
  }
  const headerColumns = sheet.headers.sort((left, right) => right.confidence - left.confidence)[0]?.columns;
  return headerColumns ?? [];
}

function visualColumnRole(column: ColumnMetadata): string {
  const normalized = normalizeComparableText(column.name);
  if (column.inferredType === "formula" || column.role === "formula") return "formula_output";
  if (column.inferredType === "currency" || column.role === "amount" || /ราคา|ยอด|ค่า|รวม|สุทธิ|ภาษี|amount|price|cost|fee|total|net|tax|revenue|payment|paid|balance/.test(normalized)) return "money";
  if (column.inferredType === "date" || column.role === "date" || /วันที่|วัน|date/.test(normalized)) return "date";
  if (column.inferredType === "status" || column.role === "status" || /\b(status|state|stage)\b/i.test(column.name) || /สถานะ/.test(normalized)) return "status";
  if (column.role === "note" || /\b(note|comment|description|detail)\b/i.test(column.name) || /หมายเหตุ|รายละเอียด/.test(normalized)) return "notes";
  if (column.role === "identifier" || /\b(id|no\.?|number|code)\b/i.test(column.name) || /เลข|รหัส|ทะเบียน|โค้ด|booking|บุ๊ค|บุก|phone|โทร|tax/.test(normalized)) return "id";
  if (column.role === "vendor" || column.role === "account" || /ลูกค้า|customer|vendor|supplier|account|ผู้รับเหมา/.test(normalized)) return "entity";
  if (column.role === "category" || /\b(type|category|priority|owner)\b/i.test(column.name) || /ประเภท|หมวด/.test(normalized)) return "category";
  if (column.inferredType === "number") return "number";
  return "unknown";
}

function visualColumnSignals(column: ColumnMetadata): string[] {
  return [
    column.name ? "header_text" : undefined,
    column.inferredType !== "unknown" ? `type_${column.inferredType}` : undefined,
    column.role && column.role !== "unknown" ? `role_${column.role}` : undefined,
    column.importance !== undefined && column.importance >= 0.9 ? "high_importance" : undefined
  ].filter((signal): signal is string => signal !== undefined);
}

function inferVisualReadabilitySheetType(sheet: WorkbookMetadata["sheets"][number], columns: VisualReadabilityColumnRole[]): VisualReadabilitySheetType {
  const roles = new Set(columns.map((column) => column.role));
  const moneyCount = columns.filter((column) => column.role === "money" || column.role === "number").length;
  if (sheet.kind === "summary") return "summary_report";
  if (sheet.kind === "template") return "input_template";
  if (roles.has("status") && (roles.has("date") || roles.has("entity"))) return "record_tracker";
  if (moneyCount >= 2 || roles.has("formula_output")) return "numeric_report";
  if (columns.length >= 3) return "tabular_data";
  return sheet.usedRange ? "generic_table" : "unknown";
}

function profileForVisualSheetType(sheetType: VisualReadabilitySheetType): string {
  switch (sheetType) {
    case "tabular_data":
      return "tabular_data";
    case "record_tracker":
      return "record_tracker";
    case "numeric_report":
      return "numeric_report";
    case "input_template":
      return "input_template";
    case "summary_report":
      return "summary_report";
    default:
      return "office_clean";
  }
}

function columnsFromAddress(address: string): string[] {
  const parsed = tryParseA1Address(stripSheetName(address));
  if (!parsed) return [];
  const columns: string[] = [];
  for (let column = parsed.startColumn; column <= parsed.endColumn; column += 1) {
    columns.push(numberToColumn(column));
  }
  return columns;
}

function uniqueDefined(values: Array<string | undefined>): string[] {
  return [...new Set(values.filter((value): value is string => typeof value === "string" && value.length > 0))];
}

function shouldDeriveAsFormula(input: AgentRunInput): boolean {
  const values = input.values ?? {};
  const mode = normalizeOperationName(keyedStringValue(values, "outputMode", "mode", "writeMode") ?? "");
  if (mode === "formula" || mode === "formulas") {
    return true;
  }
  return /\b(should\s+be\s+(?:a\s+)?formula|write\s+formula|as\s+formula|auto[-\s]?calculate|formula\s+(?:for|to)\s+all|fill\s+formula)\b/i.test(input.request);
}

function withFormulaDeriveSourceDefaults(input: AgentRunInput, target: ValueColumnScope): AgentRunInput {
  const values = input.values ?? {};
  const sourceColumns = keyedArrayStringValue(values, "sourceColumns", "sources", "source");
  if (sourceColumns.length > 0) {
    return input;
  }
  if (headerTextMatches(target.headerName, "payment_variance") || /payment\s+variance/i.test(input.request)) {
    return {
      ...input,
      values: {
        ...values,
        sourceColumns: ["Cash Amount", "Actual Amount"],
        formula: keyedStringValue(values, "formula", "formulaType", "expression") ?? "actual_minus_cash"
      }
    };
  }
  return input;
}

function compileFormulaDeriveMatrix(
  target: ValueColumnScope,
  sources: ValueColumnScope[],
  sourceSnapshots: Array<{ values: CellMatrix; formulas: CellMatrix }>,
  values: AgentRunInput["values"]
): { ok: true; formulas: CellMatrix; changedCount: number; examples: Array<Record<string, unknown>> } | { ok: false; summary: string; warnings: string[] } {
  if (sources.length < 2) {
    return {
      ok: false,
      summary: "Formula derivation needs at least two source columns.",
      warnings: ["Provide values.sourceColumns, for example [\"Cash Amount\", \"Actual Amount\"]."]
    };
  }
  const expression = normalizeOperationName(keyedStringValue(values, "formula", "formulaType", "expression", "operationFormula", "calculation") ?? "");
  const cashSource = sources.find((source) => headerTextMatches(source.headerName, "cash_amount"));
  const actualSource = sources.find((source) => headerTextMatches(source.headerName, "actual_amount"));
  const left = expression === "cash_minus_actual"
    ? cashSource ?? sources[0]
    : actualSource ?? sources[1];
  const right = expression === "cash_minus_actual"
    ? actualSource ?? sources[1]
    : cashSource ?? sources[0];
  if (!left || !right) {
    return {
      ok: false,
      summary: "Could not resolve two source columns for the formula derivation.",
      warnings: ["Provide source columns explicitly and keep them aligned with the target rows."]
    };
  }
  const sourceColumns = sourceSnapshots.map((snapshot) => columnVector(snapshot.values));
  const formulas = Array.from({ length: target.rowCount }, (_row, index) => {
    const row = target.startRow + index;
    return [`=${left.columnLetter}${row}-${right.columnLetter}${row}`];
  });
  const examples: Array<Record<string, unknown>> = [];
  for (let index = 0; index < target.rowCount && examples.length < 10; index += 1) {
    const row = target.startRow + index;
    examples.push({
      row,
      source: Object.fromEntries(sources.map((source, sourceIndex) => [source.headerName ?? source.columnLetter, sourceColumns[sourceIndex]?.[index] ?? ""])),
      after: { [target.headerName ?? target.columnLetter]: formulas[index]?.[0] }
    });
  }
  return { ok: true, formulas, changedCount: formulas.length, examples };
}

function headerTextMatches(header: string | undefined, normalizedNeedle: string): boolean {
  const normalized = normalizeHeaderName(header ?? "");
  return normalized.includes(normalizedNeedle) || normalized.replace(/_/g, "").includes(normalizedNeedle.replace(/_/g, ""));
}

function compileTransformPlan(
  target: ValueColumnScope,
  snapshot: { values: CellMatrix; formulas: CellMatrix },
  operation: TransformValuesOperation,
  values: AgentRunInput["values"]
): { ok: true } & CompiledColumnPlan | { ok: false; summary: string; warnings: string[] } {
  const before = columnVector(snapshot.values);
  const formulas = columnVector(snapshot.formulas);
  const afterValues = [...before];
  const changedRows = before.map(() => false);
  const skipped: Record<string, number> = {};
  const examples: Array<Record<string, unknown>> = [];
  let changedCount = 0;
  let unmatchedCount = 0;
  for (let index = 0; index < before.length; index += 1) {
    const current = before[index];
    if (formulaLike(formulas[index])) {
      incrementCount(skipped, "formulaTarget");
      continue;
    }
    const transformed = transformCellValue(current, operation, values);
    if (transformed === noChange) {
      if (operation === "replace_text" || operation === "map_values" || operation === "conditional_replace") unmatchedCount += 1;
      continue;
    }
    if (Object.is(transformed, current)) continue;
    afterValues[index] = transformed;
    changedRows[index] = true;
    changedCount += 1;
    if (examples.length < 10) {
      examples.push({ row: target.startRow + index, before: current ?? "", after: transformed ?? "" });
    }
  }
  return { ok: true, afterValues, changedRows, scannedCount: before.length, changedCount, skipped, unmatchedCount, examples, warnings: [] };
}

function compileDerivePlan(
  target: ValueColumnScope,
  targetSnapshot: { values: CellMatrix; formulas: CellMatrix },
  sources: ValueColumnScope[],
  sourceSnapshots: Array<{ values: CellMatrix; formulas: CellMatrix }>,
  operation: DeriveValuesOperation,
  values: AgentRunInput["values"]
): { ok: true } & CompiledColumnPlan | { ok: false; summary: string; warnings: string[] } {
  const before = columnVector(targetSnapshot.values);
  const targetFormulas = columnVector(targetSnapshot.formulas);
  const sourceColumns = sourceSnapshots.map((snapshot) => columnVector(snapshot.values));
  const afterValues = [...before];
  const changedRows = before.map(() => false);
  const skipped: Record<string, number> = {};
  const examples: Array<Record<string, unknown>> = [];
  let changedCount = 0;
  let unmatchedCount = 0;

  for (let index = 0; index < before.length; index += 1) {
    const current = before[index];
    if (formulaLike(targetFormulas[index])) {
      incrementCount(skipped, "formulaTarget");
      continue;
    }
    const sourceValues = sourceColumns.map((column) => column[index]);
    const derived = deriveCellValue(current, sourceValues, operation, values);
    if (derived === skipBecauseTargetNotBlank) {
      incrementCount(skipped, "nonBlankTarget");
      continue;
    }
    if (derived === skipBecauseBlankSource) {
      incrementCount(skipped, "blankSource");
      continue;
    }
    if (derived === noChange) {
      unmatchedCount += operation === "lookup_map" || operation === "conditional_map" || operation === "extract_pattern" ? 1 : 0;
      continue;
    }
    if (Object.is(derived, current)) continue;
    afterValues[index] = derived;
    changedRows[index] = true;
    changedCount += 1;
    if (examples.length < 10) {
      examples.push({
        row: target.startRow + index,
        source: Object.fromEntries(sources.map((source, sourceIndex) => [source.headerName ?? source.columnLetter, sourceValues[sourceIndex] ?? ""])),
        before: { [target.headerName ?? target.columnLetter]: current ?? "" },
        after: { [target.headerName ?? target.columnLetter]: derived ?? "" }
      });
    }
  }
  return { ok: true, afterValues, changedRows, scannedCount: before.length, changedCount, skipped, unmatchedCount, examples, warnings: [] };
}

function compileLookupDerivePlan(
  target: ValueColumnScope,
  targetSnapshot: { values: CellMatrix; formulas: CellMatrix },
  sources: ValueColumnScope[],
  sourceSnapshots: Array<{ values: CellMatrix; formulas: CellMatrix }>,
  lookup: LookupMapScope,
  lookupKeySnapshot: { values: CellMatrix; formulas: CellMatrix },
  lookupValueSnapshot: { values: CellMatrix; formulas: CellMatrix }
): { ok: true } & CompiledColumnPlan | { ok: false; summary: string; warnings: string[] } {
  const lookupKeys = columnVector(lookupKeySnapshot.values);
  const lookupValues = columnVector(lookupValueSnapshot.values);
  const lookupMap = new Map<string, unknown>();
  const duplicates = new Set<string>();
  for (let index = 0; index < lookupKeys.length; index += 1) {
    const key = lookupKey(lookupKeys[index]);
    if (!key) continue;
    if (lookupMap.has(key)) {
      duplicates.add(key);
      continue;
    }
    lookupMap.set(key, lookupValues[index]);
  }
  const before = columnVector(targetSnapshot.values);
  const targetFormulas = columnVector(targetSnapshot.formulas);
  const sourceColumn = columnVector(sourceSnapshots[0]?.values ?? []);
  const afterValues = [...before];
  const changedRows = before.map(() => false);
  const skipped: Record<string, number> = {};
  const examples: Array<Record<string, unknown>> = [];
  let changedCount = 0;
  let unmatchedCount = 0;
  for (let index = 0; index < before.length; index += 1) {
    const current = before[index];
    if (formulaLike(targetFormulas[index])) {
      incrementCount(skipped, "formulaTarget");
      continue;
    }
    const sourceValue = sourceColumn[index];
    const key = lookupKey(sourceValue);
    if (!key) {
      incrementCount(skipped, "blankSource");
      continue;
    }
    if (!lookupMap.has(key)) {
      unmatchedCount += 1;
      continue;
    }
    const derived = lookupMap.get(key);
    if (Object.is(derived, current)) continue;
    afterValues[index] = derived;
    changedRows[index] = true;
    changedCount += 1;
    if (examples.length < 10) {
      examples.push({
        row: target.startRow + index,
        source: { [sources[0]?.headerName ?? sources[0]?.columnLetter ?? "source"]: sourceValue ?? "" },
        lookup: {
          keyColumn: lookup.keyScope.headerName ?? lookup.keyScope.columnLetter,
          valueColumn: lookup.valueScope.headerName ?? lookup.valueScope.columnLetter
        },
        before: { [target.headerName ?? target.columnLetter]: current ?? "" },
        after: { [target.headerName ?? target.columnLetter]: derived ?? "" }
      });
    }
  }
  const warnings = duplicates.size > 0
    ? [`Lookup table contains ${duplicates.size} duplicate key(s); first value was used for each duplicate key.`]
    : [];
  return {
    ok: true,
    afterValues,
    changedRows,
    scannedCount: before.length,
    changedCount,
    skipped,
    unmatchedCount,
    examples,
    warnings
  };
}

const noChange = Symbol("noChange");
const skipBecauseTargetNotBlank = Symbol("skipBecauseTargetNotBlank");
const skipBecauseBlankSource = Symbol("skipBecauseBlankSource");

function transformCellValue(current: unknown, operation: TransformValuesOperation, values: AgentRunInput["values"]): unknown | typeof noChange {
  const text = current === undefined || current === null ? "" : String(current);
  switch (operation) {
    case "add_prefix": {
      const prefix = keyedRawStringValue(values, "prefix", "value") ?? "";
      if (!prefix || isBlankishAutoApplyValue(current) || text.startsWith(prefix)) return noChange;
      return `${prefix}${text}`;
    }
    case "add_suffix": {
      const suffix = keyedRawStringValue(values, "suffix", "value") ?? "";
      if (!suffix || isBlankishAutoApplyValue(current) || text.endsWith(suffix)) return noChange;
      return `${text}${suffix}`;
    }
    case "replace_text": {
      const find = keyedRawStringValue(values, "find", "from", "oldValue", "old");
      const replacement = keyedRawStringValue(values, "replacement", "replaceWith", "to", "newValue", "new") ?? "";
      if (!find || !text.includes(find)) return noChange;
      return text.split(find).join(replacement);
    }
    case "normalize_whitespace": {
      if (typeof current !== "string") return noChange;
      const normalized = current.trim().replace(/\s+/g, " ");
      return normalized === current ? noChange : normalized;
    }
    case "case": {
      const caseMode = normalizeOperationName(keyedStringValue(values, "case", "caseMode") ?? "");
      if (typeof current !== "string") return noChange;
      if (caseMode === "upper" || caseMode === "uppercase") return current.toUpperCase();
      if (caseMode === "lower" || caseMode === "lowercase") return current.toLowerCase();
      return noChange;
    }
    case "fill_blank": {
      if (!isBlankishAutoApplyValue(current)) return noChange;
      return values?.value ?? values?.fillValue ?? "";
    }
    case "map_values":
    case "conditional_replace": {
      const mapped = lookupMappedValue(current, values);
      return mapped.matched ? mapped.value : noChange;
    }
  }
}

function deriveCellValue(current: unknown, sourceValues: unknown[], operation: DeriveValuesOperation, values: AgentRunInput["values"]): unknown | typeof noChange | typeof skipBecauseTargetNotBlank | typeof skipBecauseBlankSource {
  const firstSource = sourceValues[0];
  switch (operation) {
    case "copy_if_blank":
      if (!isBlankishAutoApplyValue(current)) return skipBecauseTargetNotBlank;
      if (isBlankishAutoApplyValue(firstSource)) return skipBecauseBlankSource;
      return firstSource;
    case "copy_from_source":
      if (isBlankishAutoApplyValue(firstSource)) return skipBecauseBlankSource;
      return firstSource;
    case "normalize_from_source":
      if (isBlankishAutoApplyValue(firstSource)) return skipBecauseBlankSource;
      return typeof firstSource === "string" ? firstSource.trim().replace(/\s+/g, " ") : firstSource;
    case "conditional_map":
    case "lookup_map": {
      const mapped = lookupMappedValue(firstSource, values);
      return mapped.matched ? mapped.value : noChange;
    }
    case "extract_pattern": {
      const pattern = keyedStringValue(values, "pattern", "regex");
      if (!pattern || isBlankishAutoApplyValue(firstSource)) return noChange;
      const match = String(firstSource).match(new RegExp(pattern));
      return match?.[1] ?? match?.[0] ?? noChange;
    }
    case "formula_like":
      return deriveFormulaLikeValue(current, sourceValues, values);
  }
}

function lookupMappedValue(value: unknown, values: AgentRunInput["values"]): { matched: true; value: unknown } | { matched: false } {
  const mapValue = values?.map ?? values?.mapping ?? values?.replacements;
  if (mapValue && typeof mapValue === "object" && !Array.isArray(mapValue)) {
    const record = mapValue as Record<string, unknown>;
    const direct = record[String(value ?? "")];
    if (direct !== undefined) return { matched: true, value: direct };
    const normalized = normalizeHeaderName(String(value ?? ""));
    const key = Object.keys(record).find((candidate) => normalizeHeaderName(candidate) === normalized);
    if (key !== undefined) return { matched: true, value: record[key] };
  }
  return { matched: false };
}

function lookupKey(value: unknown): string | undefined {
  if (value === undefined || value === null || value === "") return undefined;
  return normalizeHeaderName(String(value));
}

function compileSheetTransformPlan(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { ok: true; operation: string; renames: Array<{ from: string; to: string }>; skipped: Array<{ sheetName: string; reason: string }>; warnings: string[] } | { ok: false; summary: string; warnings: string[]; candidates?: AgentCandidate[] } {
  const values = input.values ?? {};
  const operation = normalizeOperationName(keyedStringValue(values, "operation", "transform", "type") ?? "");
  const prefix = keyedRawStringValue(values, "prefix");
  const suffix = keyedRawStringValue(values, "suffix");
  const find = keyedRawStringValue(values, "find", "from", "oldValue", "old");
  const replacement = keyedRawStringValue(values, "replacement", "replaceWith", "to", "newValue", "new") ?? "";
  const explicitSheets = keyedArrayStringValue(values, "sheets", "sheetNames", "targetSheets");
  const includeHidden = values.includeHidden === true;
  const candidateSheets = metadata.sheets
    .filter((sheet) => includeHidden || sheet.isHidden !== true)
    .filter((sheet) => explicitSheets.length === 0 || explicitSheets.some((name) => sameText(name, sheet.name)));
  if (candidateSheets.length === 0) {
    return {
      ok: false,
      summary: "No sheets matched the requested sheet transform.",
      warnings: ["Provide values.sheets/values.sheetNames, or omit them to target all visible sheets."],
      candidates: metadata.sheets.slice(0, 12).map((sheet) => ({ id: sheet.id, kind: "sheet", label: sheet.name, sheetName: sheet.name, confidence: 0.8 }))
    };
  }
  const renames: Array<{ from: string; to: string }> = [];
  const skipped: Array<{ sheetName: string; reason: string }> = [];
  const existing = new Set(metadata.sheets.map((sheet) => sheet.name));
  const planned = new Set<string>();
  for (const sheet of candidateSheets) {
    let nextName: string | undefined;
    if (operation === "add_prefix" || (!operation && prefix)) {
      if (!prefix) return { ok: false, summary: "Sheet prefix transform needs values.prefix.", warnings: ["Provide values.prefix."] };
      nextName = sheet.name.startsWith(prefix) ? sheet.name : `${prefix}${sheet.name}`;
    } else if (operation === "add_suffix" || (!operation && suffix)) {
      if (!suffix) return { ok: false, summary: "Sheet suffix transform needs values.suffix.", warnings: ["Provide values.suffix."] };
      nextName = sheet.name.endsWith(suffix) ? sheet.name : `${sheet.name}${suffix}`;
    } else if (operation === "replace_text") {
      if (!find) return { ok: false, summary: "Sheet replace transform needs values.find and values.replacement.", warnings: ["Provide values.find and values.replacement."] };
      nextName = sheet.name.includes(find) ? sheet.name.split(find).join(replacement) : sheet.name;
    } else {
      return { ok: false, summary: "Sheet transform needs operation add_prefix, add_suffix, or replace_text.", warnings: ["Use transform_sheets for bounded sheet rename plans."] };
    }
    if (!nextName || nextName === sheet.name) {
      skipped.push({ sheetName: sheet.name, reason: "unchanged" });
      continue;
    }
    if (existing.has(nextName) || planned.has(nextName)) {
      skipped.push({ sheetName: sheet.name, reason: `target name already exists: ${nextName}` });
      continue;
    }
    planned.add(nextName);
    renames.push({ from: sheet.name, to: nextName });
  }
  if (renames.length === 0) {
    return { ok: false, summary: "Sheet transform produced no safe renames.", warnings: skipped.map((item) => `${item.sheetName}: ${item.reason}`).slice(0, 8) };
  }
  return {
    ok: true,
    operation: operation || (prefix ? "add_prefix" : suffix ? "add_suffix" : "replace_text"),
    renames,
    skipped,
    warnings: skipped.length > 0 ? [`Skipped ${skipped.length} sheet(s).`] : []
  };
}

function changedColumnRuns(scope: ValueColumnScope, afterValues: unknown[], changedRows: boolean[]): Array<{ address: string; values: CellMatrix }> {
  const runs: Array<{ address: string; values: CellMatrix }> = [];
  let start: number | undefined;
  let values: CellMatrix = [];
  const flush = (exclusiveIndex: number) => {
    if (start === undefined) return;
    runs.push({
      address: addressFromBounds(scope.startRow + start, columnToNumber(scope.columnLetter), exclusiveIndex - start, 1),
      values
    });
    start = undefined;
    values = [];
  };
  for (let index = 0; index < afterValues.length; index += 1) {
    if (!changedRows[index]) {
      flush(index);
      continue;
    }
    if (start === undefined) start = index;
    values.push([afterValues[index] as CellMatrix[number][number]]);
  }
  flush(afterValues.length);
  return runs;
}

function columnVector(matrix: CellMatrix): unknown[] {
  return matrix.map((row) => row[0]);
}

function formulaLike(value: unknown): boolean {
  return typeof value === "string" && value.trim().startsWith("=");
}

function incrementCount(record: Record<string, number>, key: string): void {
  record[key] = (record[key] ?? 0) + 1;
}

function compactValueColumnScope(scope: ValueColumnScope): Record<string, unknown> {
  return stripUndefinedRecord({
    sheet: scope.sheetName,
    header: scope.headerName,
    table: scope.tableName,
    sectionId: scope.sectionId,
    range: scope.address,
    column: scope.columnLetter,
    rows: scope.rowCount
  });
}

function keyedStringValue(values: AgentRunInput["values"] | undefined, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = values?.[key];
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return undefined;
}

function keyedRawStringValue(values: AgentRunInput["values"] | undefined, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = values?.[key];
    if (typeof value === "string" && value.length > 0) {
      return value;
    }
  }
  return undefined;
}

function keyedArrayStringValue(values: AgentRunInput["values"] | undefined, ...keys: string[]): string[] {
  for (const key of keys) {
    const value = values?.[key];
    if (Array.isArray(value)) {
      return value.filter((item): item is string => typeof item === "string" && item.trim().length > 0).map((item) => item.trim());
    }
    if (typeof value === "string" && value.trim()) {
      return [value.trim()];
    }
  }
  return [];
}

function transformNeedsInput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  summary: string,
  warnings: string[],
  candidates?: AgentCandidate[]
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: candidates && candidates.length > 0 ? "AMBIGUOUS_TARGET" : "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    ...(candidates && candidates.length > 0 ? { candidates } : {}),
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings
  };
}

function combineApplyResults(results: unknown[]) {
  const objects = results.filter((result): result is Record<string, unknown> => Boolean(result) && typeof result === "object");
  const failed = objects.find((result) => result.ok === false);
  const backups = objects.flatMap((result) => {
    if (Array.isArray(result.backups)) {
      return result.backups.map(String);
    }
    const backupId = (result.backup as { backupId?: unknown } | undefined)?.backupId
      ?? (result.backup as { backup?: { backupId?: unknown } } | undefined)?.backup?.backupId;
    return backupId ? [String(backupId)] : [];
  });
  const warnings = objects.flatMap((result) => Array.isArray(result.warnings) ? result.warnings.map(operationWarningMessage) : []);
  const error = failed?.error;
  const errorWarning = applyErrorMessage(error);
  if (errorWarning && !warnings.includes(errorWarning)) {
    warnings.push(errorWarning);
  }
  const telemetry = {
    stepCount: objects.length,
    cellsWritten: objects.reduce((total, result) => total + (numberValue((result.telemetry as Record<string, unknown> | undefined)?.cellsWritten) ?? 0), 0),
    styleCopyCount: objects.reduce((total, result) => total + (numberValue((result.telemetry as Record<string, unknown> | undefined)?.styleCopyCount) ?? 0), 0)
  };
  return {
    ok: failed ? false : objects.every((result) => result.ok !== false),
    backups,
    rollbackAvailable: objects.some((result) => Boolean(result.rollbackAvailable)) || backups.length > 0,
    warnings,
    ...(error !== undefined ? { error } : {}),
    telemetry,
    results: objects
  };
}

function operationWarningMessage(warning: unknown): string {
  if (typeof warning === "string") {
    return warning;
  }
  if (warning && typeof warning === "object") {
    const record = warning as Record<string, unknown>;
    const code = typeof record.code === "string" ? record.code : undefined;
    const message = typeof record.message === "string" ? record.message : undefined;
    const target = record.target && typeof record.target === "object" ? record.target as Record<string, unknown> : undefined;
    const targetText = target && typeof target.sheetName === "string" && typeof target.address === "string"
      ? ` (${target.sheetName}!${target.address})`
      : "";
    if (code && message) {
      return `${code}: ${message}${targetText}`;
    }
    if (message) {
      return `${message}${targetText}`;
    }
    if (code) {
      return `${code}${targetText}`;
    }
    try {
      return JSON.stringify(warning);
    } catch {
      return Object.prototype.toString.call(warning);
    }
  }
  return String(warning);
}

function applyErrorMessage(error: unknown): string | undefined {
  if (!error) {
    return undefined;
  }
  return operationWarningMessage(error);
}

function cleanRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput, resolved?: Extract<AgentTargetResolution, { ok: true }>): AgentCleanRequest | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const sheetName = stringValue(resolved?.sheetName ?? input.target?.sheetName ?? values?.sheetName);
  const address = stringValue(resolved?.range ?? input.target?.range ?? values?.address ?? values?.range);
  if (!sheetName || !address) {
    return undefined;
  }
  const targetAddress = stringValue(values?.targetAddress);
  const headerRowIndex = positiveIntegerValue(values?.headerRowIndex);
  const keyColumns = numberArrayValue(values?.keyColumns);
  const strategy = cleanFillStrategy(values?.strategy);
  const columnIndex = numberValue(values?.columnIndex);
  const columnIndexes = numberArrayValue(values?.columnIndexes);
  const delimiter = stringValue(values?.delimiter);
  const separator = stringValue(values?.separator);
  const numberFormat = stringValue(values?.numberFormat);
  const request: AgentCleanRequest = {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    sheetName,
    address: normalizeOperationRange(metadata, sheetName, address),
    ...(headerRowIndex !== undefined ? { headerRowIndex } : {}),
    ...(typeof values?.hasHeader === "boolean" ? { hasHeader: values.hasHeader } : {}),
    ...(keyColumns.length > 0 ? { keyColumns } : {}),
    ...(strategy ? { strategy } : {}),
    ...(Object.prototype.hasOwnProperty.call(values ?? {}, "value") ? { value: values?.value } : {}),
    ...(columnIndex !== undefined ? { columnIndex } : {}),
    ...(columnIndexes.length > 0 ? { columnIndexes } : {}),
    ...(delimiter ? { delimiter } : {}),
    ...(separator ? { separator } : {}),
    ...(targetAddress ? { targetAddress: normalizeOperationRange(metadata, sheetName, targetAddress) } : {}),
    ...(numberFormat ? { numberFormat } : {})
  };
  return request;
}

function cleanPatchRequestsFromInput(metadata: WorkbookMetadata, input: AgentRunInput, resolved?: Extract<AgentTargetResolution, { ok: true }>): AgentCleanRequest[] {
  const values = input.values as Record<string, unknown> | undefined;
  const patches = Array.isArray(values?.patches) ? values.patches : [];
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  const requests: AgentCleanRequest[] = [];
  for (const patch of patches) {
    if (!patch || typeof patch !== "object") {
      continue;
    }
    const typed = patch as Record<string, unknown>;
    const target = typed.target && typeof typed.target === "object" ? typed.target as Record<string, unknown> : {};
    const sheetName = stringValue(target.sheetName ?? typed.sheetName ?? resolved?.sheetName ?? input.target?.sheetName);
    const range = stringValue(target.range ?? typed.range ?? typed.address);
    if (!sheetName || !range) {
      continue;
    }
    const numberFormat = stringValue(typed.numberFormat ?? values?.numberFormat);
    requests.push({
      workbookId,
      sheetName,
      address: normalizeOperationRange(metadata, sheetName, range),
      ...(numberFormat ? { numberFormat } : {})
    });
  }
  return requests;
}

function cleanOutputAddress(request: AgentCleanRequest): string {
  return request.targetAddress ?? request.address;
}

function templateCaptureRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput): TemplateCaptureRequest | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const name = stringValue(values?.name ?? values?.templateName);
  const sourceSheetName = stringValue(values?.sourceSheetName ?? input.target?.sheetName);
  if (!name || !sourceSheetName) {
    return undefined;
  }
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    name,
    scope: values?.scope === "local" ? "local" : "workbook",
    sourceSheetName,
    dataRegions: stringArrayValue(values?.dataRegions)
  };
}

function templateRepairRequestFromInput(
  metadata: WorkbookMetadata,
  input: AgentRunInput
): { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string; repair?: AddinTemplateRepairRequest["repair"] } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const templateId = stringValue(values?.templateId ?? input.target?.entity) as TemplateId | undefined;
  const targetSheetName = stringValue(values?.targetSheetName ?? input.target?.sheetName);
  if (!templateId || !targetSheetName) {
    return undefined;
  }
  const repair = templateRepairKindsFromInput(values?.repair);
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    templateId,
    targetSheetName,
    ...(repair ? { repair } : {})
  };
}

function templateRegionFillRequestFromInput(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  runtime: Pick<RuntimeService, "getTemplate">
): { templateId: TemplateId; targetSheetName: string; operations: Array<Extract<ExcelOperation, { kind: "range.write_values" }>> } | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  const templateId = stringValue(values?.templateId ?? input.target?.entity) as TemplateId | undefined;
  const targetSheetName = stringValue(values?.targetSheetName ?? input.target?.sheetName);
  if (!templateId || !targetSheetName) {
    return undefined;
  }
  const templateResult = runtime.getTemplate(templateId) as { ok?: boolean; template?: { dataRegions?: string[] } };
  const declaredRegions = new Set((templateResult.ok === false ? [] : templateResult.template?.dataRegions ?? []).map((region) => stripSheetName(region)));
  const regions = templateRegionValuesFromInput(values, declaredRegions);
  if (regions.length === 0) {
    return undefined;
  }
  const workbookId = metadata.workbook.workbookId as WorkbookId;
  return {
    templateId,
    targetSheetName,
    operations: regions.map((region) => ({
      kind: "range.write_values",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "values",
      reason: `Fill template data region ${targetSheetName}!${region.address}`,
      target: { workbookId, sheetName: targetSheetName, address: region.address },
      values: region.values,
      preserveFormats: true
    }))
  };
}

function templateRegionValuesFromInput(
  values: Record<string, unknown> | undefined,
  declaredRegions: Set<string>
): Array<{ address: string; values: CellMatrix }> {
  const regionEntries: Array<{ address: string; values: CellMatrix }> = [];
  const rawRegions = values?.regions;
  if (Array.isArray(rawRegions)) {
    for (const rawRegion of rawRegions) {
      if (!rawRegion || typeof rawRegion !== "object") {
        continue;
      }
      const record = rawRegion as Record<string, unknown>;
      const address = stringValue(record.address ?? record.range);
      const matrix = matrixFromUnknown(record.values);
      const normalized = address ? stripSheetName(address) : undefined;
      if (normalized && matrix && templateRegionIsAllowed(normalized, declaredRegions)) {
        regionEntries.push({ address: normalized, values: matrix });
      }
    }
  }
  const rawRegionValues = values?.regionValues;
  if (rawRegionValues && typeof rawRegionValues === "object" && !Array.isArray(rawRegionValues)) {
    for (const [address, rawMatrix] of Object.entries(rawRegionValues as Record<string, unknown>)) {
      const matrix = matrixFromUnknown(rawMatrix);
      const normalized = stripSheetName(address);
      if (matrix && templateRegionIsAllowed(normalized, declaredRegions)) {
        regionEntries.push({ address: normalized, values: matrix });
      }
    }
  }
  return regionEntries;
}

function templateRegionIsAllowed(address: string, declaredRegions: Set<string>): boolean {
  return declaredRegions.size === 0 || declaredRegions.has(stripSheetName(address));
}

function templateIdFromInput(input: AgentRunInput): TemplateId | undefined {
  const values = input.values as Record<string, unknown> | undefined;
  return stringValue(values?.templateId ?? input.target?.entity ?? input.operationId) as TemplateId | undefined;
}

function templateRepairKindsFromInput(value: unknown): AddinTemplateRepairRequest["repair"] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }
  const allowed = new Set(["styles", "formulas", "dataRegions", "layout"]);
  const repair = value.filter((item): item is AddinTemplateRepairRequest["repair"][number] => typeof item === "string" && allowed.has(item));
  return repair.length > 0 ? repair : undefined;
}

function stringArrayValue(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string" && item.trim().length > 0) : [];
}

function styleDimensionsFromInput(values: Record<string, unknown> | undefined): StyleDimension[] {
  const allowed = new Set<StyleDimension>(["columnWidths", "rowHeights", "borders", "fills", "fonts", "alignment", "numberFormats", "conditionalFormatting", "dataValidation"]);
  const raw = stringArrayValue(values?.dimensions);
  return raw.filter((item): item is StyleDimension => allowed.has(item as StyleDimension));
}

function usedRangeForSheet(metadata: WorkbookMetadata, sheetName: string): string {
  return metadata.sheets.find((sheet) => sheet.name === sheetName)?.usedRange ?? "A1";
}

function positiveNumberValue(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : undefined;
}

function numberValue(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function positiveIntegerValue(value: unknown): number | undefined {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : undefined;
}

function numberArrayValue(value: unknown): number[] {
  return Array.isArray(value) ? value.filter((item): item is number => typeof item === "number" && Number.isInteger(item) && item >= 0) : [];
}

function cleanFillStrategy(value: unknown): AgentCleanRequest["strategy"] | undefined {
  return value === "value" || value === "zero" || value === "previous" || value === "next" ? value : undefined;
}

function formulaFillRequestFromInput(metadata: WorkbookMetadata, input: AgentRunInput, direction: "down" | "right"): FormulaFillRequest | undefined {
  const endpoints = rangeTransferEndpointsFromInput(metadata, input);
  if (!endpoints || endpoints.source.sheetName !== endpoints.destination.sheetName) {
    return undefined;
  }
  return {
    workbookId: metadata.workbook.workbookId as WorkbookId,
    sheetName: endpoints.source.sheetName,
    sourceAddress: endpoints.source.address,
    targetAddress: endpoints.destination.address,
    direction
  };
}

function explicitRangeTarget(metadata: WorkbookMetadata, value: unknown, fallbackSheetName?: string): A1Range | undefined {
  if (typeof value === "string") {
    const parsed = parseSheetQualifiedAddress(value);
    const sheetName = parsed?.sheetName ?? fallbackSheetName;
    const range = parsed?.address ?? value;
    if (!sheetName || !range) {
      return undefined;
    }
    const sheet = metadata.sheets.find((candidate) => normalizeComparableText(candidate.name) === normalizeComparableText(sheetName));
    if (!sheet) {
      return undefined;
    }
    return { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: sheet.name, address: normalizeOperationRange(metadata, sheet.name, stripSheetName(range)) };
  }
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const target = value as Partial<AgentRunTarget> & { address?: unknown };
  const rawRange = typeof target.range === "string" ? target.range : typeof target.address === "string" ? target.address : undefined;
  const parsed = rawRange ? parseSheetQualifiedAddress(rawRange) : undefined;
  const sheetName = typeof target.sheetName === "string" ? target.sheetName : parsed?.sheetName ?? fallbackSheetName;
  const range = parsed?.address ?? rawRange;
  if (!sheetName || !range) {
    return undefined;
  }
  const sheet = metadata.sheets.find((candidate) => normalizeComparableText(candidate.name) === normalizeComparableText(sheetName));
  if (!sheet) {
    return undefined;
  }
  return { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: sheet.name, address: normalizeOperationRange(metadata, sheet.name, range) };
}

function parseSheetQualifiedAddress(value: string): { sheetName: string; address: string } | undefined {
  const trimmed = value.trim();
  const quoted = /^'([^']+)'!(.+)$/.exec(trimmed);
  if (quoted?.[1] && quoted[2]) {
    return { sheetName: quoted[1], address: quoted[2] };
  }
  const unquoted = /^([^!]+)!(.+)$/.exec(trimmed);
  if (unquoted?.[1] && unquoted[2]) {
    return { sheetName: unquoted[1], address: unquoted[2] };
  }
  return undefined;
}

function rangeCopyTypeFromInput(input: AgentRunInput): NonNullable<Extract<ExcelOperation, { kind: "range.copy" }>["copyType"]> {
  const raw = (input.values as Record<string, unknown> | undefined)?.copyType;
  return raw === "all" || raw === "values" || raw === "formats" || raw === "formulas" ? raw : "all";
}

function destructiveLevelForRangeCopy(input: AgentRunInput): Extract<ExcelOperation, { kind: "range.copy" }>["destructiveLevel"] {
  const copyType = rangeCopyTypeFromInput(input);
  if (copyType === "formats") {
    return "format";
  }
  if (copyType === "all" || copyType === "formulas") {
    return "structure";
  }
  return "values";
}

function dimensionsFromAddress(address: string): { rows: number; columns: number } | undefined {
  const range = tryParseA1Address(address);
  if (!range) {
    return undefined;
  }
  return {
    rows: range.endRow - range.startRow + 1,
    columns: range.endColumn - range.startColumn + 1
  };
}

function tableReadRowMetadata(address: string | undefined, rowOffset: number, values: CellMatrix): Array<{ rowIndex: number; sheetRowNumber: number; address: string }> {
  const parsed = address ? tryParseA1Address(stripSheetName(address)) : undefined;
  if (!parsed) {
    return [];
  }
  return values.map((_row, rowIndex) => {
    const sheetRowNumber = parsed.startRow + rowOffset + rowIndex;
    return {
      rowIndex,
      sheetRowNumber,
      address: `${numberToColumn(parsed.startColumn)}${sheetRowNumber}:${numberToColumn(parsed.endColumn)}${sheetRowNumber}`
    };
  });
}

function rangeRowMetadata(address: string | undefined, values: CellMatrix, rowIndexes: number[]): Array<{ rowIndex: number; sheetRowNumber: number; address: string }> {
  const parsed = address ? tryParseA1Address(stripSheetName(address)) : undefined;
  if (!parsed || rowIndexes.length === 0) {
    return [];
  }
  return rowIndexes
    .filter((rowIndex) => rowIndex >= 0 && rowIndex < values.length)
    .map((rowIndex) => {
      const sheetRowNumber = parsed.startRow + rowIndex;
      return {
        rowIndex,
        sheetRowNumber,
        address: `${numberToColumn(parsed.startColumn)}${sheetRowNumber}:${numberToColumn(parsed.endColumn)}${sheetRowNumber}`
      };
    });
}

function valuePatchesFromInput(input: AgentRunInput): AgentValuePatch[] {
  const patches = input.values?.patches;
  if (!Array.isArray(patches)) {
    return [];
  }
  return patches.flatMap((patch) => {
    const target = patch?.target;
    if (!target) {
      return [];
    }
    const values = Array.isArray(patch.values)
      ? patch.values
      : Array.isArray(patch.rows)
        ? patch.rows
        : undefined;
    if (!values) {
      return [];
    }
    return [{ target, values: values as CellMatrix, ...(patch.reason ? { reason: patch.reason } : {}) }];
  });
}

function inheritPatchTargetSheet(input: AgentRunInput, target: AgentRunTarget): AgentRunTarget {
  if (target.sheetName || !input.target?.sheetName) {
    return target;
  }
  return { ...target, sheetName: input.target.sheetName };
}

function semanticValuePatchesFromInput(input: AgentRunInput): AgentSemanticValuePatch[] {
  const values = input.values as Record<string, unknown> | undefined;
  const raw = values?.semanticPatches ?? values?.anchorPatches;
  const entries = Array.isArray(raw) ? raw : raw && typeof raw === "object" ? [raw] : [];
  return entries.flatMap((entry) => {
    if (!entry || typeof entry !== "object") {
      return [];
    }
    const patch = entry as Record<string, unknown>;
    const rowMatch = normalizeSemanticRowMatch(patch.rowMatch ?? patch.row ?? patch.match);
    const columnMatch = typeof patch.columnMatch === "string"
      ? patch.columnMatch
      : typeof patch.columnHeader === "string"
        ? patch.columnHeader
        : typeof patch.column === "string"
          ? patch.column
          : undefined;
    const value = patch.value ?? patch.after;
    if (!rowMatch || !columnMatch || value === undefined) {
      return [];
    }
    return [{
      ...(typeof patch.sectionId === "string" ? { sectionId: patch.sectionId } : {}),
      ...(typeof patch.sectionLabel === "string" ? { sectionLabel: patch.sectionLabel } : {}),
      ...(typeof patch.sheetName === "string" ? { sheetName: patch.sheetName } : {}),
      rowMatch,
      columnMatch,
      value,
      ...(typeof patch.reason === "string" ? { reason: patch.reason } : {})
    }];
  });
}

function normalizeSemanticRowMatch(value: unknown): AgentSemanticValuePatch["rowMatch"] | undefined {
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const raw = value as Record<string, unknown>;
  const matchValue = raw.value ?? raw.label ?? raw.text;
  if (matchValue === undefined || matchValue === null || String(matchValue).trim() === "") {
    return undefined;
  }
  return {
    ...(typeof raw.column === "string" ? { column: raw.column } : {}),
    value: matchValue,
    ...(typeof raw.contains === "boolean" ? { contains: raw.contains } : {})
  };
}

function findSectionForSemanticPatch(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  patch: AgentSemanticValuePatch
): { ok: true; section: WorkbookMetadata["sections"][number] } | { ok: false; summary: string; warnings: string[] } {
  const sections = metadata.sections.filter((section) => {
    if (patch.sheetName && normalizeSemanticText(section.sheetName) !== normalizeSemanticText(patch.sheetName)) {
      return false;
    }
    if (input.target?.sheetName && normalizeSemanticText(section.sheetName) !== normalizeSemanticText(input.target.sheetName)) {
      return false;
    }
    return true;
  });
  const exact = sections.find((section) => patch.sectionId && section.id === patch.sectionId);
  if (exact) {
    return { ok: true, section: exact };
  }
  const labelMatch = patch.sectionLabel
    ? sections.filter((section) => semanticTextMatches(section.label, patch.sectionLabel!, true) || section.labels.some((label) => semanticTextMatches(label, patch.sectionLabel!, true)))
    : [];
  if (labelMatch.length === 1) {
    return { ok: true, section: labelMatch[0]! };
  }
  const targetRange = input.target?.range ? stripSheetName(input.target.range) : undefined;
  const rangeMatch = targetRange
    ? sections.filter((section) => section.range === targetRange || section.headerRange === targetRange || dataRangeForSection(section) === targetRange)
    : [];
  if (rangeMatch.length === 1) {
    return { ok: true, section: rangeMatch[0]! };
  }
  if (sections.length === 1) {
    return { ok: true, section: sections[0]! };
  }
  const scored = sections
    .map((section) => ({ section, score: semanticPatchSectionScore(section, input, patch) }))
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score);
  if (scored.length > 0 && scored[0]!.score > (scored[1]?.score ?? 0)) {
    return { ok: true, section: scored[0]!.section };
  }
  return {
    ok: false,
    summary: sections.length === 0 ? "No section metadata is available for the requested sheet." : "Multiple sections are plausible.",
    warnings: [
      sections.length === 0
        ? "Ask for a sheet summary first or provide an explicit range target."
        : `Provide sectionId or sectionLabel. Candidate sections: ${sections.slice(0, 6).map((section) => `${section.id} (${section.range})`).join(", ")}`
    ]
  };
}

function semanticPatchSectionScore(section: WorkbookMetadata["sections"][number], input: AgentRunInput, patch: AgentSemanticValuePatch): number {
  let score = 0;
  const request = normalizeSemanticText(input.request);
  if (request.includes(normalizeSemanticText(section.sheetName))) score += 2;
  if (request.includes(normalizeSemanticText(section.label))) score += 2;
  if (section.columns.some((column) => findSectionColumn({ ...section, columns: [column] }, patch.columnMatch))) score += 3;
  if (patch.rowMatch.column && findSectionColumn(section, patch.rowMatch.column)) score += 2;
  return score;
}

function findSectionColumn(section: WorkbookMetadata["sections"][number], anchor: string): ColumnMetadata | undefined {
  const normalized = normalizeSemanticText(anchor);
  if (!normalized) {
    return undefined;
  }
  return section.columns.find((column) => column.letter.toLowerCase() === anchor.trim().toLowerCase())
    ?? section.columns.find((column) => normalizeSemanticText(column.name) === normalized || normalizeSemanticText(column.normalizedName) === normalized)
    ?? section.columns.find((column) => semanticTextMatches(column.name, anchor, true) || semanticTextMatches(column.normalizedName, anchor, true));
}

function findSemanticPatchRow(
  section: WorkbookMetadata["sections"][number],
  dataRange: string,
  values: CellMatrix,
  rowMatch: AgentSemanticValuePatch["rowMatch"]
): { ok: true; sheetRow: number } | { ok: false; warnings: string[] } {
  const parsed = tryParseA1Address(stripSheetName(dataRange));
  if (!parsed) {
    return { ok: false, warnings: [`Could not parse section data range ${dataRange}.`] };
  }
  const matchColumn = rowMatch.column ? findSectionColumn(section, rowMatch.column) : undefined;
  if (rowMatch.column && !matchColumn) {
    return { ok: false, warnings: [`Could not match row anchor column "${rowMatch.column}" in section ${section.id}.`] };
  }
  const startColumnIndex = parsed.startColumn - 1;
  const candidateColumnIndexes = matchColumn
    ? [matchColumn.index - startColumnIndex]
    : section.columns
      .filter((column) => ["description", "dimension", "identifier", "vendor", "account", "category", "status"].includes(column.role ?? "") || column.importance !== undefined && column.importance >= 0.5)
      .map((column) => column.index - startColumnIndex);
  const boundedCandidateIndexes = candidateColumnIndexes.filter((index) => index >= 0 && index <= parsed.endColumn - parsed.startColumn);
  const indexes = boundedCandidateIndexes.length > 0 ? boundedCandidateIndexes : undefined;
  const rowIndexes = values.flatMap((row, rowIndex) => {
    const cells = indexes ? indexes.map((index) => row[index]) : row;
    return cells.some((cell) => semanticTextMatches(cell, rowMatch.value, rowMatch.contains ?? true)) ? [rowIndex] : [];
  });
  if (rowIndexes.length === 1) {
    return { ok: true, sheetRow: parsed.startRow + rowIndexes[0]! };
  }
  if (rowIndexes.length > 1) {
    return { ok: false, warnings: [`Row anchor matched ${rowIndexes.length} rows. Add another row discriminator or use an explicit range.`] };
  }
  return { ok: false, warnings: [`No data row matched "${String(rowMatch.value)}" in ${section.sheetName}!${dataRange}.`] };
}

function semanticPatchNeedsInput(
  metadata: WorkbookMetadata,
  requestedMode: AgentRunMode,
  summary: string,
  warnings: string[]
): Omit<AgentRunOutput, "telemetry"> {
  return {
    status: "NEEDS_INPUT",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary,
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    warnings
  };
}

function semanticTextMatches(candidate: unknown, needle: unknown, contains: boolean): boolean {
  const left = normalizeSemanticText(candidate);
  const right = normalizeSemanticText(needle);
  if (!left || !right) {
    return false;
  }
  if (contains) {
    return left.includes(right) || right.includes(left);
  }
  return left === right;
}

function normalizeSemanticText(value: unknown): string {
  return String(value ?? "")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{N} ]/gu, "");
}

function matrixShapeIssue(address: string, values: CellMatrix): string | undefined {
  const range = tryParseA1Address(address);
  if (!range) {
    return undefined;
  }
  const expectedRows = range.endRow - range.startRow + 1;
  const expectedColumns = range.endColumn - range.startColumn + 1;
  const actualRows = values.length;
  const actualColumns = values.reduce((max, row) => Math.max(max, row.length), 0);
  if (actualRows !== expectedRows || actualColumns !== expectedColumns || values.some((row) => row.length !== expectedColumns)) {
    return `Expected ${expectedRows} row(s) x ${expectedColumns} column(s), received ${actualRows} row(s) x ${actualColumns} column(s).`;
  }
  return undefined;
}

function isSparseBroadWrite(address: string, values: CellMatrix): boolean {
  const matrixCells = values.reduce((sum, row) => sum + row.length, 0);
  const nonEmpty = values.flat().filter((value) => value !== null && value !== undefined && value !== "").length;
  const touched = Math.max(matrixCells, cellCountFromAddress(address) ?? 0);
  return touched >= 8 && nonEmpty > 0 && nonEmpty / touched <= 0.25 && touched - nonEmpty >= 4;
}

function shouldBuildDetailedPreviewChanges(values: CellMatrix): boolean {
  return matrixCellCount(values) <= AGENT_DETAILED_PREVIEW_CELL_LIMIT;
}

function previewChangesForMatrix(sheetName: string, range: string, values: CellMatrix, before: CellMatrix, forceSummary = false): NonNullable<AgentRunOutput["changes"]> {
  if (forceSummary || !shouldBuildDetailedPreviewChanges(values)) {
    return [{
      sheetName,
      range,
      before: "omitted for large preview",
      after: {
        kind: "range_write_summary",
        rowCount: values.length,
        columnCount: values.reduce((max, row) => Math.max(max, row.length), 0),
        cellCount: matrixCellCount(values),
        sample: values.slice(0, 3).map((row) => row.slice(0, 6))
      }
    }];
  }
  return values.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
    sheetName,
    cell: cellAddressFor(range, rowIndex, columnIndex),
    range,
    before: before[rowIndex]?.[columnIndex],
    after: value
  })));
}

function queryRowsContractOutput(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
  const values = input.values as Record<string, unknown> | undefined;
  const hasWhere = Array.isArray(values?.where);
  const returnColumns = Array.isArray(values?.return) ? values.return.filter((value): value is string => typeof value === "string") : undefined;
  const predicateColumns = hasWhere
    ? (values?.where as Array<Record<string, unknown>>).map((predicate) => predicate.column).filter((column): column is string => typeof column === "string")
    : [];
  const fieldResolutions = resolveSemanticFields(metadata, [...predicateColumns, ...(returnColumns ?? [])], input.target);
  return {
    status: hasWhere ? "NEEDS_INPUT" : "VALIDATION_FAILED",
    mode: requestedMode,
    workbookContextId: metadata.workbookContextId,
    summary: hasWhere
      ? "query_rows is recognized as a read-only row query; execution will use the query_rows pipeline."
      : "query_rows requires values.where predicates before it can run.",
    answer: {
      kind: "query_rows_contract",
      readOnly: true,
      target: input.target ?? {},
      predicates: hasWhere ? values?.where : [],
      returnColumns: returnColumns ?? [],
      fieldCandidates: fieldResolutions.map((resolution) => ({
        term: resolution.term,
        ambiguous: resolution.ambiguous,
        candidates: resolution.candidates
      })),
      supportedOperators: [...AGENT_QUERY_ROW_OPERATORS],
      supportedFormats: ["json_rows", "csv", "summary"],
      requiredShape: {
        intent: { action: "query_rows" },
        values: {
          where: [{ column: "Status", op: "=", value: "Unpaid" }],
          return: ["Date", "Customer", "Amount", "Status"],
          limit: 100,
          format: "json_rows"
        }
      }
    },
    proof: [],
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "ask_user",
    agentInstruction: "Use query_rows for read-only row lookup. Do not call filter_range unless the user explicitly asks to change the visible Excel filter.",
    warnings: hasWhere
      ? ["query_rows execution is scheduled for the next implementation milestone; do not emulate it with a visible filter."]
      : ["Provide values.where with one or more predicates. Do not use filter_range for lookup-only questions."]
  };
}

function resolveQueryRowsTable(metadata: WorkbookMetadata, target: AgentRunTarget | undefined): TableMetadata | undefined {
  const tables = metadata.tables
    .filter((table) => !target?.sheetName || table.sheetName === target.sheetName)
    .filter((table) => !target?.tableName || table.name === target.tableName);
  if (target?.tableName) {
    return tables[0];
  }
  return tables.length === 1 ? tables[0] : undefined;
}

function normalizeQueryRowsPredicates(value: unknown[]): AgentQueryRowsPredicate[] {
  return value
    .filter((entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === "object" && !Array.isArray(entry))
    .map((entry) => ({
      column: String(entry.column ?? ""),
      op: isQueryRowsOperator(entry.op) ? entry.op : "=",
      ...(entry.value !== undefined ? { value: entry.value } : {})
    }))
    .filter((predicate) => predicate.column.trim().length > 0);
}

function isQueryRowsOperator(value: unknown): value is AgentQueryRowsPredicate["op"] {
  return typeof value === "string" && (AGENT_QUERY_ROW_OPERATORS as readonly string[]).includes(value);
}

function clampQueryRowsLimit(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) ? Math.max(1, Math.min(500, Math.floor(value))) : 100;
}

function clampQueryRowsScanLimit(value: unknown, limit: number): number {
  return typeof value === "number" && Number.isFinite(value) ? Math.max(limit, Math.min(5_000, Math.floor(value))) : Math.max(limit, 500);
}

function queryPredicateMatches(value: unknown, predicate: AgentQueryRowsPredicate): boolean {
  const op = predicate.op;
  if (op === "blank") return value === undefined || value === null || String(value).trim() === "";
  if (op === "not_blank") return !(value === undefined || value === null || String(value).trim() === "");
  if (op === "contains") return normalizedCellString(value).includes(normalizedCellString(predicate.value));
  if (op === "starts_with") return normalizedCellString(value).startsWith(normalizedCellString(predicate.value));
  if (op === "ends_with") return normalizedCellString(value).endsWith(normalizedCellString(predicate.value));
  if (op === "in") return Array.isArray(predicate.value) && predicate.value.some((candidate) => cellEquals(value, candidate));
  if (op === "not_in") return Array.isArray(predicate.value) && !predicate.value.some((candidate) => cellEquals(value, candidate));
  if (op === "between" && Array.isArray(predicate.value) && predicate.value.length >= 2) {
    return compareCells(value, predicate.value[0]) >= 0 && compareCells(value, predicate.value[1]) <= 0;
  }
  if (op === ">") return compareCells(value, predicate.value) > 0;
  if (op === ">=") return compareCells(value, predicate.value) >= 0;
  if (op === "<") return compareCells(value, predicate.value) < 0;
  if (op === "<=") return compareCells(value, predicate.value) <= 0;
  if (op === "!=") return !cellEquals(value, predicate.value);
  return cellEquals(value, predicate.value);
}

function cellEquals(left: unknown, right: unknown): boolean {
  const leftNumber = numericCellValue(left);
  const rightNumber = numericCellValue(right);
  if (leftNumber !== undefined && rightNumber !== undefined) {
    return leftNumber === rightNumber;
  }
  return normalizedCellString(left) === normalizedCellString(right);
}

function compareCells(left: unknown, right: unknown): number {
  const leftNumber = numericCellValue(left);
  const rightNumber = numericCellValue(right);
  if (leftNumber !== undefined && rightNumber !== undefined) {
    return leftNumber - rightNumber;
  }
  return normalizedCellString(left).localeCompare(normalizedCellString(right));
}

function normalizedCellString(value: unknown): string {
  return value === undefined || value === null ? "" : String(value).trim().toLowerCase();
}

function queryRowAddress(table: TableMetadata, zeroBasedDataIndex: number): string {
  const parsed = table.dataRange ? tryParseA1Address(stripSheetName(table.dataRange)) : undefined;
  const rowNumber = (parsed?.startRow ?? 2) + zeroBasedDataIndex;
  return `${table.sheetName}!${rowNumber}:${rowNumber}`;
}

function queryRowsSuggestedPatchOperation(
  input: AgentRunInput,
  table: TableMetadata,
  rowIndexes: number[],
  updateColumnLetter: string,
  updateValue: unknown
): Partial<AgentRunInput> {
  const parsed = table.dataRange ? tryParseA1Address(stripSheetName(table.dataRange)) : undefined;
  const startRow = parsed?.startRow ?? 2;
  return {
    request: input.request,
    mode: "preview_update",
    intent: { action: "write_values", reason: "Preview patches generated from read-only query_rows matched row addresses." },
    target: { sheetName: table.sheetName, ...(table.name ? { tableName: table.name } : {}) },
    values: {
      patches: rowIndexes.map((index) => ({
        target: { sheetName: table.sheetName, range: `${updateColumnLetter}${startRow + index}` },
        values: [[updateValue]],
        reason: "query_rows matched row"
      }))
    }
  };
}

function rowsToCsv(columns: string[], rows: Array<Record<string, unknown>>): string {
  return [columns.join(","), ...rows.map((row) => columns.map((column) => csvCell(row[column])).join(","))].join("\n");
}

function csvCell(value: unknown): string {
  const text = value === undefined || value === null ? "" : String(value);
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, "\"\"")}"` : text;
}

function uniqueProofFromChanges(changes: NonNullable<AgentRunOutput["changes"]>): AgentProofReference[] {
  const seen = new Set<string>();
  const proof: AgentProofReference[] = [];
  for (const change of changes) {
    if (!change.range) {
      continue;
    }
    const key = `${change.sheetName}!${change.range}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    proof.push({ sheetName: change.sheetName, range: change.range, label: "preview target" });
  }
  return proof;
}

function overlapsFormulaRegion(metadata: WorkbookMetadata, sheetName: string, address: string): boolean {
  return metadata.formulaRegions.some((region) =>
    region.sheetName === sheetName && (rangesOverlapAddresses(address, region.range) || region.range === address)
  );
}

function cellAddressFor(range: string, rowIndex: number, columnIndex: number): string {
  const match = /^([A-Z]+)(\d+)/i.exec(range);
  const startColumn = match?.[1] ? columnToNumber(match[1]) : 1;
  const startRow = match?.[2] ? Number(match[2]) : 1;
  return `${columnLetter(startColumn + columnIndex - 1)}${startRow + rowIndex}`;
}

function contextResource(workbookContextId: string) {
  return { uri: `excel://agent/contexts/${workbookContextId}`, name: "workbook context", description: "Cached workbook metadata summary.", mimeType: "application/json" as const };
}

function semanticIndexResource(workbookContextId: string) {
  return { uri: `excel://agent/contexts/${workbookContextId}/semantic-index`, name: "semantic workbook index", description: "Role-aware workbook target index.", mimeType: "application/json" as const };
}

function operationResource(operationId: string) {
  return { uri: `excel://agent/operations/${operationId}`, name: "pending operation", description: "Pending previewed workbook update.", mimeType: "application/json" as const };
}

function resultResource(resultId: string, view?: "summary" | "full") {
  const suffix = view ? `?view=${view}` : "";
  return { uri: `excel://agent/results/${resultId}${suffix}`, name: "agent result", description: "Stored agent answer detail.", mimeType: "application/json" as const };
}

function compactResultResource(resultId: string) {
  return { uri: `excel://compact/${resultId}`, name: "compact agent result", description: "Compatibility alias for a stored agent result.", mimeType: "application/json" as const };
}

function compactStoredResult(result: StoredAgentResult): Record<string, unknown> {
  return stripUndefinedRecord({
    resultId: result.resultId,
    resourceUri: result.resourceUri,
    fullResourceUri: result.fullResourceUri,
    workbookContextId: result.workbookContextId,
    freshness: result.freshness,
    kind: result.kind,
    summary: result.summary,
    hash: result.hash,
    createdAt: result.createdAt,
    expiresAt: result.expiresAt,
    answer: compactAnswerForResponseMode(result.answer, "brief", { request: "stored result summary", mode: "answer" }, result.resourceUri, result.fullResourceUri)
  });
}

function enforceStoredResultByteBudget(result: StoredAgentResult, maxBytes: number): StoredAgentResult | Record<string, unknown> {
  const fullBytes = Buffer.byteLength(JSON.stringify(result));
  if (fullBytes <= maxBytes) {
    return result;
  }
  return {
    ...compactStoredResult(result),
    truncated: true,
    fullBytes,
    maxBytes,
    warning: "Full stored result exceeds requested byte budget; use fullResourceUri with a larger budget if required."
  };
}
