import {
  makeId,
  runtimeError,
  type AgentCandidate,
  type AgentOperationId,
  type AgentProofReference,
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
  type RangeSearchRequest,
  type RegionRegisterRequest,
  type RegionSelector,
  type StyleCopyRequest,
  type StyleDimension,
  type TableApplyFiltersRequest,
  type TableAppendRowsRequest,
  type TableCopyStructureRequest,
  type TableCreateRequest,
  type TableReadRequest,
  type TableReorderColumnsRequest,
  type TableResizeRequest,
  type TableSelector,
  type TableSetStyleRequest,
  type TableSetTotalRowRequest,
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
  type SnapshotId,
  type WorkbookId
} from "@components-kit/open-workbook-protocol";
import { columnLetter, type HeaderMetadata, type TableMetadata, type WorkbookMetadata, WorkbookMetadataCache } from "./workbook-metadata-cache.js";
import { AgentOperationStore, type AgentCleanMutationAction, type AgentCleanRequest } from "./agent-operation-store.js";
import { WorkbookMetadataBuilder } from "./workbook-metadata-builder.js";
import { findAgentCandidates, resolveAgentReadTarget, resolveAgentUpdateTarget, type AgentTargetResolution } from "./agent-target-resolver.js";
import type { RuntimeService } from "./runtime-service.js";
import { classifyAgentActionRisk, type AgentOperationRisk } from "./agent-action-policy.js";
import { routeAgentRequest, type IntentRoute } from "./agent-routing.js";
import { isAgentIntentAction, normalizeAgentIntent, type AgentIntentAction, type NormalizedAgentIntent } from "./agent-intent.js";
import { findAgentActionHandler, type AgentActionHandlerDefinition, type AgentActionHandlerId } from "./agent-action-handlers.js";

export class AgentOrchestrator {
  readonly metadataCache = new WorkbookMetadataCache();
  private readonly operations = new AgentOperationStore();
  private readonly metadataBuilder: WorkbookMetadataBuilder;

  constructor(private readonly runtime: RuntimeService) {
    this.metadataBuilder = new WorkbookMetadataBuilder(runtime, this.metadataCache);
  }

  async run(input: AgentRunInput, context?: AgentRunExecutionContext): Promise<AgentRunOutput> {
    const startedAt = Date.now();
    let internalCallCount = 0;
    const intent = normalizeAgentIntent(input);
    const route = routeAgentRequest(input.request, input.mode ?? "auto", intent);
    const runMetrics: AgentRunMetrics = { internalReadCount: 0, fullReadCellCount: 0, validationStatus: "not_run", route, intent };
    const mode = input.mode ?? "auto";
    const finish = (output: Omit<AgentRunOutput, "telemetry">, cacheHit = false): AgentRunOutput => {
      const outputMetrics = output.metrics as Record<string, unknown> | undefined;
      const operationRisk = typeof outputMetrics?.operationRisk === "string" ? outputMetrics.operationRisk : runMetrics.operationRisk;
      const actionHandlerId = typeof outputMetrics?.actionHandlerId === "string" ? outputMetrics.actionHandlerId : runMetrics.actionHandlerId;
      const autoApplyBlockedReason = typeof outputMetrics?.autoApplyBlockedReason === "string" ? outputMetrics.autoApplyBlockedReason : runMetrics.autoApplyBlockedReason;
      const targetFingerprintStatus = isTargetFingerprintStatus(outputMetrics?.targetFingerprintStatus) ? outputMetrics.targetFingerprintStatus : runMetrics.targetFingerprintStatus;
      const preBudgetPayloadBytes = Buffer.byteLength(JSON.stringify(output));
      const budgeted = applyOutputBudget(output, input);
      const targetHintCount = runMetrics.intent.targetHints?.length ?? 0;
      const targetHintUsed = targetHintCount > 0 && outputUsedCallerTargetHint(output);
      const payloadBytes = Buffer.byteLength(JSON.stringify(budgeted));
      return {
        ...budgeted,
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
          candidateCount: budgeted.candidates?.length ?? 0,
          resourceLinkCount: budgeted.resourceLinks.length,
          estimatedTokensSaved: Math.max(0, Math.ceil((preBudgetPayloadBytes - payloadBytes) / 4)),
          routeMode: runMetrics.route.mode,
          routeMatchedRule: runMetrics.route.matchedRule,
          routeConfidence: runMetrics.route.confidence,
          routeReasons: runMetrics.route.reasons,
          ...(operationRisk !== undefined ? { operationRisk } : {}),
          ...(actionHandlerId !== undefined ? { actionHandlerId } : {}),
          ...(autoApplyBlockedReason !== undefined ? { autoApplyBlockedReason } : {}),
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

      if (mode === "rollback") {
        return finish(await this.rollback(input));
      }

      const requestedOverviewIntent = workbookOverviewIntent(input);
      const effectiveMode = route.mode;
      const includeSamples = shouldBuildSampledMetadata(input, effectiveMode, requestedOverviewIntent);
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
      const { metadata, cacheHit } = await this.metadataBuilder.getOrBuild({
        ...(input.workbookContextId ? { workbookContextId: String(input.workbookContextId) } : {}),
        ...(input.target?.workbookId !== undefined ? { workbookId: input.target.workbookId } : {}),
        ...(input.target?.workbookName !== undefined ? { workbookName: input.target.workbookName } : {}),
        includeSamples
      });
      internalCallCount += cacheHit ? 1 : 4;

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
      if (effectiveMode === "find") {
        const sectionAnswer = sectionAnswerOutput(metadata, input, mode);
        return finish(sectionAnswer ?? this.findOutput(metadata, input, mode), cacheHit);
      }
      if (effectiveMode === "preview_update") {
        internalCallCount += 1;
      const preview = await this.previewUpdate(metadata, input, mode);
        if (mode !== "auto" || preview.status !== "PREVIEW_READY") {
          return finish(preview, cacheHit);
        }
        if (process.env.OPEN_WORKBOOK_AGENT_AUTO_APPLY !== "1") {
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
          columns: header.columns.map((column) => ({ name: column.name, normalizedName: column.normalizedName, inferredType: column.inferredType }))
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
        columns: section.columns.map((column) => ({ name: column.name, normalizedName: column.normalizedName, inferredType: column.inferredType })),
        labels: section.labels,
        confidence: section.confidence
      })),
      tables: metadata.tables.map((table) => ({
        name: table.name,
        sheetName: table.sheetName,
        range: table.range,
        columns: table.columns.map((column) => ({ name: column.name, normalizedName: column.normalizedName, inferredType: column.inferredType }))
      })),
      namedRanges: metadata.namedRanges,
      summaryBlocks: metadata.summaryBlocks,
      formulaRegions: metadata.formulaRegions,
      fingerprint: metadata.fingerprint,
      updatedAt: metadata.updatedAt,
      expiresAt: metadata.expiresAt
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
      summary: pending.summary,
      changes: pending.changes,
      createdAt: pending.createdAt
    };
  }

  invalidateWorkbook(workbookId: WorkbookId | string): void {
    this.metadataCache.deleteByWorkbookId(workbookId);
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
    const workbookAnswer = workbookOverviewAnswer(metadata, input, requestedMode);
    if (workbookAnswer) {
      return workbookAnswer;
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
    const templateAnswer = await this.templateAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (templateAnswer) {
      return templateAnswer;
    }
    const styleAnswer = await this.styleAnswerOutput(metadata, input, requestedMode, runMetrics);
    if (styleAnswer) {
      return styleAnswer;
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
    const normalizedRange = normalizeOperationRange(metadata, resolved.sheetName, resolved.range);
    if (isFormulaReadIntentAction(intentAction(input))) {
      return this.formulaAnswerOutput(metadata, input, requestedMode, normalizedRange, resolved, runMetrics);
    }
    const rangeMetadataAnswer = await this.rangeMetadataAnswerOutput(metadata, input, requestedMode, normalizedRange, resolved, runMetrics);
    if (rangeMetadataAnswer) {
      return rangeMetadataAnswer;
    }
    const table = tableFromResolution(metadata, resolved);
    if (table && (resolved.candidate.kind === "table" || input.target?.tableName)) {
      return this.tableCompactAnswerOutput(metadata, input, requestedMode, resolved, table, runMetrics);
    }
    const profile = await this.readAndProfileRange(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, normalizedRange, runMetrics);
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Answered from ${resolved.candidate.label} on ${resolved.sheetName} using a targeted compact read.`,
      answer: profile,
      metrics: profile.metrics,
      candidates: candidates.slice(0, 5),
      proof: [{ sheetName: resolved.sheetName, range: normalizedRange, label: resolved.candidate.label }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: profile.warning ? [profile.warning] : []
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
          }))
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
        headers: headers.map(headerAnswer)
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
    const rowOffset = nonNegativeInteger(values?.rowOffset) ?? 0;
    const rowLimit = compactTableRowLimit(input, table.columns.length);
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
    const totalRows = dimensionsFromAddress(table.dataRange ?? table.range)?.rows ?? profile.shape.rows;
    const projectedColumns = columns.length > 0
      ? table.columns.filter((column) => columns.includes(column.name) || columns.includes(column.index)).map((column) => ({ name: column.name, index: column.index, letter: column.letter }))
      : table.columns.map((column) => ({ name: column.name, index: column.index, letter: column.letter }));
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
        projectedColumns,
        truncated,
        ...(truncated ? { nextPage: { rowOffset: rowOffset + rowLimit } } : {}),
        schema: table.columns.map((column) => ({ name: column.name, index: column.index, letter: column.letter, inferredType: column.inferredType })),
        profile,
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
    const action = intentAction(input);
    const workbookId = metadata.workbook.workbookId as WorkbookId;
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
    const method = rangeMetadataMethodForAction(action);
    if (!method) {
      return undefined;
    }
    const values = input.values as Record<string, unknown> | undefined;
    const request: RangeMetadataRequest | RangeSearchRequest = method === "range.search"
      ? { workbookId, sheetName: resolved.sheetName, address: normalizedRange, text: stringValue(values?.text ?? values?.query ?? input.request) ?? "" }
      : { workbookId, sheetName: resolved.sheetName, address: normalizedRange };
    runMetrics.internalReadCount += 1;
    const result = await this.runtime.readRangeMetadata(method, request);
    return {
      status: (result as { ok?: boolean }).ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Read ${method.replace("range.", "").replace(/_/g, " ")} for ${resolved.sheetName}!${normalizedRange}.`,
      answer: { kind: "range_metadata", method, result },
      metrics: { source: "runtime_range_metadata" },
      proof: [{ sheetName: resolved.sheetName, range: normalizedRange, label: "range metadata" }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: (result as { ok?: boolean }).ok === false ? "manual_review" : "answer_now",
      warnings: []
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
    const patches = valuePatchesFromInput(input);
    if (patches.length > 0) {
      return this.previewPatchUpdate(metadata, input, requestedMode, patches);
    }
    const operationPreview = await this.previewOperationIntent(metadata, input, requestedMode);
    if (operationPreview) {
      return operationPreview;
    }
    const resolved = resolveUpdateTarget(metadata, input);
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
    const matrix = objectToCellMatrix(input.values ?? {});
    if (containsFormulaLikeValue(matrix) && (intentAction(input) === "write_formulas" || isFormulaMutationRequest(input.request))) {
      return this.previewFormulaUpdate(metadata, input, requestedMode, resolved, matrix);
    }
    if ((intentAction(input) === "append_table_rows" || isTableAppendIntent(input.request)) && resolved.candidate.kind === "table") {
      return this.previewTableAppend(metadata, input, requestedMode, resolved, matrix);
    }
    if (isSparseBroadWrite(resolved.range, matrix)) {
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
    if (containsFormulaLikeValue(matrix)) {
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
    if (formulaWarnings.length > 0 && (!input.target?.range || (cellCountFromAddress(resolved.range) ?? 0) > Math.max(1, matrixCellCount(matrix)))) {
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
    const before = await this.readRangeValues(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, resolved.range);
    const operation: ExcelOperation = {
      kind: "range.write_values",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      values: matrix,
      preserveFormats: true
    };
    const changes = matrix.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
      sheetName: resolved.sheetName,
      cell: cellAddressFor(resolved.range, rowIndex, columnIndex),
      range: resolved.range,
      before: before[rowIndex]?.[columnIndex],
      after: value
    })));
    const action = { kind: "batch" as const, operations: [operation] };
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
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "preview target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: formulaWarnings
    };
  }

  private async previewPatchUpdate(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    patches: AgentValuePatch[]
  ): Promise<Omit<AgentRunOutput, "telemetry">> {
    const operations: ExcelOperation[] = [];
    const changes: NonNullable<AgentRunOutput["changes"]> = [];
    const warnings: string[] = [];
    let cellCount = 0;

    for (const [index, patch] of patches.entries()) {
      const resolved = resolveUpdateTarget(metadata, {
        ...input,
        target: patch.target,
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
      const shapeIssue = matrixShapeIssue(resolved.range, patch.values);
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
      if (isSparseBroadWrite(resolved.range, patch.values)) {
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
      if (containsFormulaLikeValue(patch.values)) {
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
      if (formulaWarning && (!patch.target.range || (cellCountFromAddress(resolved.range) ?? 0) > Math.max(1, matrixCellCount(patch.values)))) {
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
      const before = await this.readRangeValues(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, resolved.range);
      operations.push({
        kind: "range.write_values",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "values",
        reason: patch.reason ?? input.request,
        target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
        values: patch.values,
        preserveFormats: true
      });
      cellCount += matrixCellCount(patch.values);
      changes.push(...patch.values.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
        sheetName: resolved.sheetName,
        cell: cellAddressFor(resolved.range, rowIndex, columnIndex),
        range: resolved.range,
        before: before[rowIndex]?.[columnIndex],
        after: value
      }))));
    }

    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
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
        operationCount: operations.length,
        grouped: true
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof,
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings
    };
  }

  private async previewOperationIntent(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry"> | undefined> {
    const action = intentAction(input);
    const workbookLevelHandler = findAgentActionHandler(input, action, false);
    if (workbookLevelHandler) {
      return this.previewActionHandler(metadata, input, requestedMode, workbookLevelHandler);
    }
    const resolved = resolveAgentUpdateTarget(metadata, input);
    if (!resolved.ok) {
      return undefined;
    }
    const normalizedRange = normalizeOperationRange(metadata, resolved.sheetName, resolved.range);
    const normalizedResolved = { ...resolved, range: normalizedRange };
    const targetLevelHandler = findAgentActionHandler(input, action, true);
    if (targetLevelHandler) {
      return this.previewActionHandler(metadata, input, requestedMode, targetLevelHandler, normalizedResolved);
    }
    return undefined;
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
        return resolved ? this.previewTableSelectorMutation(metadata, input, requestedMode, resolved, "clear_filters") : undefined;
      case "filter_range":
        return resolved?.candidate.kind === "table"
          ? this.previewTableApplyFilters(metadata, input, requestedMode, resolved)
          : resolved ? this.previewAutoFilter(metadata, input, requestedMode, resolved) : undefined;
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
      case "write_styles_many":
        return this.previewWriteStylesMany(metadata, input, requestedMode);
      case "insert_rows":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.insert_rows") : undefined;
      case "delete_rows":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.delete_rows") : undefined;
      case "insert_columns":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.insert_columns") : undefined;
      case "delete_columns":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.delete_columns") : undefined;
      case "merge_range":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.merge") : undefined;
      case "unmerge_range":
        return resolved ? this.previewRangeStructuralOperation(metadata, input, requestedMode, resolved, "range.unmerge") : undefined;
      case "write_formulas":
        return resolved ? this.previewFormulaUpdate(metadata, input, requestedMode, resolved, objectToCellMatrix(input.values ?? {})) : undefined;
      case "convert_formulas_to_values":
        return resolved ? this.previewFormulaConvertToValues(metadata, input, requestedMode, resolved) : undefined;
      case "write_number_formats":
        return resolved ? this.previewNumberFormatUpdate(metadata, input, requestedMode, resolved) : undefined;
      case "format_range":
        return resolved ? this.previewStyleUpdate(metadata, input, requestedMode, resolved) : undefined;
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
        ? { kind: "sheet.protect", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName, ...(password ? { password } : {}) }
        : { kind: "sheet.unprotect", operationId: makeId<OperationId>("op"), workbookId, destructiveLevel: "structure", reason: input.request, sheetName: targetSheetName, ...(password ? { password } : {}) };
      summary = `Prepared sheet ${behavior} for ${targetSheetName}.`;
      answer = { ...answer, sheetName: targetSheetName };
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
      return workbookLevelNeedsInput(metadata, requestedMode, `Formula fill ${direction} needs values.source and values.destination ranges on the same sheet.`);
    }
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
      destructiveLevel: "values",
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

  private previewAutofit(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>, dimension: "columns" | "rows"): Omit<AgentRunOutput, "telemetry"> {
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

  private previewAutoFilter(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
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

  private previewNumberFormatUpdate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const numberFormat = numberFormatMatrixFromInput(input, resolved.range);
    if (!numberFormat) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Number-format updates need values.numberFormats, values.numberFormat, or values.formats.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "number-format target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
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
          destructiveLevel: "values",
          reason: input.request,
          source: endpoints.source,
          target: endpoints.destination,
          copyType: rangeCopyTypeFromInput(input)
        }
      : {
          kind: "range.move",
          operationId: makeId<OperationId>("op"),
          workbookId: metadata.workbook.workbookId as WorkbookId,
          destructiveLevel: "values",
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

  private previewRangeStructuralOperation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    kind: "range.clear_values" | "range.insert_rows" | "range.delete_rows" | "range.insert_columns" | "range.delete_columns" | "range.merge" | "range.unmerge"
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
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      [operation],
      [{ sheetName: resolved.sheetName, range: resolved.range, after: actionLabel }],
      `Prepared ${actionLabel} on ${resolved.sheetName}!${resolved.range}.`,
      { kind: `${kind}_preview`, sheetName: resolved.sheetName, range: resolved.range }
    );
  }

  private previewWriteStylesMany(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> {
    const workbookId = metadata.workbook.workbookId as WorkbookId;
    const entries = styleEntriesFromInput(workbookId, input);
    if (entries.length === 0) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Multi-style writes need values.entries with sheetName, range, and style.");
    }
    const operations: ExcelOperation[] = entries.map((entry) => ({
      kind: "range.write_styles",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: entry.target,
      style: entry.style,
      preserveValues: true
    }));
    return this.previewBatchOperation(
      metadata,
      requestedMode,
      operations,
      entries.map((entry) => ({ sheetName: entry.target.sheetName, range: entry.target.address, after: "styles updated" })),
      `Prepared style updates for ${entries.length} range(s).`,
      { kind: "write_styles_many_preview", rangeCount: entries.length }
    );
  }

  private previewBatchOperation(metadata: WorkbookMetadata, requestedMode: AgentRunMode, operations: ExcelOperation[], changes: NonNullable<AgentRunOutput["changes"]>, summary: string, answer: unknown): Omit<AgentRunOutput, "telemetry"> {
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
      changes,
      summary
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary,
      answer,
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
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
    const amountColumn = table.columns.find((column) => /amount/i.test(column.name));
    const key = amountColumn?.index ?? table.columns.findIndex((column) => /amount/i.test(column.name));
    if (key < 0) {
      return undefined;
    }
    const tableName = table.name ?? table.id;
    const request: TableSortRequest = {
      workbookId: metadata.workbook.workbookId as WorkbookId,
      tableName,
      fields: [{ key, ascending: !/\b(highest|descending|desc|largest|lowest to highest)\b/i.test(input.request) }]
    };
    const changes: NonNullable<AgentRunOutput["changes"]> = [{ sheetName: table.sheetName, range: table.range, after: `sorted ${tableName} by ${amountColumn?.name ?? "Amount"}` }];
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
      answer: { kind: "table_sort_preview", tableName, sheetName: table.sheetName, sortField: amountColumn?.name ?? "Amount", ascending: request.fields[0]?.ascending !== false },
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
    const filters = Array.isArray(input.values?.filters) ? input.values.filters as TableApplyFiltersRequest["filters"] : [];
    if (filters.length === 0) {
      return tableNeedsInput(metadata, requestedMode, resolved, "Table filter previews need values.filters.");
    }
    const request: TableApplyFiltersRequest = { workbookId: metadata.workbook.workbookId as WorkbookId, tableName, filters };
    return this.previewTablePendingAction(metadata, requestedMode, {
      action: { kind: "table.apply_filters", request },
      tableName,
      sheetName: resolved.sheetName,
      range: resolved.range,
      summary: `Prepared ${filters.length} table filter(s) for ${tableName}.`,
      answer: { kind: "table_apply_filters_preview", tableName, filterCount: filters.length },
      after: "applied table filters"
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
    const style = styleFromRequest(input.request);
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
      warnings: []
    };
  }

  private previewCleanMutation(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    action: AgentCleanMutationAction
  ): Omit<AgentRunOutput, "telemetry"> {
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
    const request = styleCopyRequestFromInput(metadata, input);
    if (!request) {
      return workbookLevelNeedsInput(metadata, requestedMode, "Style copy needs values.source and values.destination, or source/target sheet names.");
    }
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
    const newSheetName = uniqueSheetName(metadata, `${sourceSheet.name} Template`);
    const operations: ExcelOperation[] = [
      {
        kind: "sheet.copy",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "structure",
        reason: input.request,
        sourceSheetName: sourceSheet.name,
        newSheetName,
        position: "after",
        relativeToSheetName: sourceSheet.name,
        activate: true
      },
      {
        kind: "range.clear_values_keep_format",
        operationId: makeId<OperationId>("op"),
        workbookId: metadata.workbook.workbookId as WorkbookId,
        destructiveLevel: "values",
        reason: input.request,
        target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: newSheetName, address: sourceSheet.usedRange }
      }
    ];
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations },
      changes: [{ sheetName: newSheetName, range: sourceSheet.usedRange, before: "copied values", after: "template formatting only" }],
      summary: `Prepared template copy ${newSheetName} from ${sourceSheet.name} with values cleared.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "template_cleanup_preview", sourceSheetName: sourceSheet.name, newSheetName, clearRange: sourceSheet.usedRange },
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
    const formulas = matrix.map((row) => row.map((value) => typeof value === "string" && value.trim().startsWith("=") ? value : null));
    const formulaCount = formulas.flat().filter((value) => value !== null && value !== undefined && value !== "").length;
    if (formulaCount === 0) {
      return {
        status: "NEEDS_INPUT",
        mode: requestedMode,
        workbookContextId: metadata.workbookContextId,
        summary: "Formula updates need explicit formula values before OpenWorkbook can preview the update.",
        proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula target" }],
        resourceLinks: [contextResource(metadata.workbookContextId)],
        nextAction: "ask_user",
        warnings: []
      };
    }
    const operation: ExcelOperation = {
      kind: "range.write_formulas",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "values",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range },
      formulas,
      preserveFormats: true
    };
    const changes = formulas.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({
      sheetName: resolved.sheetName,
      cell: cellAddressFor(resolved.range, rowIndex, columnIndex),
      range: resolved.range,
      after: value
    })));
    const pending = this.createPendingOperation(metadata, {
      action: { kind: "batch", operations: [operation] },
      changes,
      summary: `Prepared ${changes.length} formula update(s) on ${resolved.sheetName}!${resolved.range}.`
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      answer: { kind: "formula_preview", sheetName: resolved.sheetName, range: resolved.range, formulaCount },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "formula target" }],
      resourceLinks: [operationResource(String(pending.operationId))],
      nextAction: "call_apply_update",
      warnings: []
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

  private async applyUpdate(input: AgentRunInput): Promise<Omit<AgentRunOutput, "telemetry">> {
    const operationId = input.operationId ? String(input.operationId) : "";
    const pending = this.operations.get(operationId);
    if (!pending) {
      return { status: "NOT_FOUND", mode: "apply_update", summary: "No pending previewed operation was found for the supplied operationId.", proof: [], resourceLinks: [], nextAction: "ask_user", warnings: [] };
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
    const result = await this.applyPendingAction(pending, operationId);
    const applyFailed = (result as { ok?: boolean }).ok === false;
    if (!applyFailed) {
      this.operations.delete(operationId);
    }
    const validation = !applyFailed ? await this.runtime.validateWorkbook({ workbookId: pending.workbookId }) : undefined;
    const issueCount = validation?.issues?.length ?? 0;
    const validationFailed = validation?.ok === false;
    return {
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
        transactionId: (result as { transactionId?: string }).transactionId,
        backupIds: (result as { backups?: string[] }).backups ?? [],
        rollbackAvailable: (result as { rollbackAvailable?: boolean }).rollbackAvailable ?? false,
        operationRisk: pending.risk,
        telemetry: (result as { telemetry?: unknown }).telemetry
      },
      metrics: { operationRisk: pending.risk, targetFingerprintStatus: "matched" },
      changes: pending.changes,
      proof: pending.changes.flatMap((change) => change.range ? [{ sheetName: change.sheetName, range: change.range, label: "applied target" }] : []).slice(0, 1),
      resourceLinks: (result as { transactionId?: string }).transactionId ? [{ uri: `excel://transactions/${(result as { transactionId?: string }).transactionId}`, name: "transaction", description: "Applied workbook transaction.", mimeType: "application/json" }] : [],
      nextAction: applyFailed || validationFailed ? "manual_review" : "answer_now",
      warnings: validation?.issues?.slice(0, 5).map((issue) => issue.message) ?? []
    };
  }

  private applyPendingAction(pending: NonNullable<ReturnType<AgentOperationStore["get"]>>, operationId: string) {
    return this.runtime.runWithAgentExecutionContext(
      pending.agentId ? { agentId: pending.agentId, ...(pending.agentName !== undefined ? { agentName: pending.agentName } : {}), clientType: "mcp" } : undefined,
      () => this.applyPendingActionInContext(pending, operationId)
    );
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
        return this.runtime.copyStyleDimensions(pending.action.request);
      case "style.repair_consistency":
        return this.runtime.repairStyleFromTemplate(pending.action.request);
      case "clean.transform":
        return this.applyCleanMutation(pending.action.action, pending.action.request);
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
    const operation: ExcelOperation = {
      kind: "range.read_full",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "none",
      reason: "Agent targeted read",
      target: { workbookId, sheetName, address },
      facets: ["values", "text"]
    };
    const request: BatchRequest = { workbookId, mode: "apply", operations: [operation] };
    const result = await this.runtime.applyBatch(request);
    const values = operationReadSnapshots(result)[0]?.snapshot?.values ?? [];
    if (runMetrics) {
      runMetrics.internalReadCount += 1;
      runMetrics.fullReadCellCount += matrixCellCount(values);
    }
    return values;
  }

  private createPendingOperation(
    metadata: WorkbookMetadata,
    input: {
      action: Parameters<AgentOperationStore["create"]>[0]["action"];
      changes: NonNullable<AgentRunOutput["changes"]>;
      summary: string;
    }
  ) {
    const risk = classifyAgentActionRisk(input.action);
    const agentContext = this.runtime.currentAgentExecutionContext();
    return this.operations.create({
      workbookContextId: metadata.workbookContextId,
      workbookId: metadata.workbook.workbookId as WorkbookId,
      action: input.action,
      changes: input.changes,
      summary: input.summary,
      risk,
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
  route: IntentRoute;
  intent: NormalizedAgentIntent;
  autoApplied?: boolean;
  safetyDecision?: string;
  previewOperationId?: string;
  operationRisk?: AgentOperationRisk;
  actionHandlerId?: AgentActionHandlerId | string;
  autoApplyBlockedReason?: string;
  targetFingerprintStatus?: "matched" | "changed" | "not_applicable";
  validationStatus: "passed" | "failed" | "not_run";
}

interface AgentValuePatch {
  target: AgentRunTarget;
  values: CellMatrix;
  reason?: string;
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
    .filter((region) => targetSheets.has(region.sheetName) || targets.some((target) => target.range && rangesMayOverlap(target.range, region.range)))
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

function shouldBuildSampledMetadata(input: AgentRunInput, effectiveMode: AgentRunMode, overviewIntent: ReturnType<typeof detectWorkbookOverviewIntent>): boolean {
  if (/\b(sections?|blocks?|areas?)\b/i.test(input.request)) {
    return true;
  }
  if (effectiveMode === "prepare") {
    return false;
  }
  if ((effectiveMode === "answer" || effectiveMode === "find") && hasWorkbookOverviewIntent(overviewIntent)) {
    return false;
  }
  return true;
}

function workbookOverviewAnswer(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Omit<AgentRunOutput, "telemetry"> | undefined {
  const intent = workbookOverviewIntent(input);
  if (!hasWorkbookOverviewIntent(intent)) {
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
    metrics: { source: "cached_metadata", sheetCount: metadata.sheets.length, tableCount: metadata.tables.length, namedRangeCount: metadata.namedRanges.length, sectionCount: metadata.sections.length },
    proof: metadata.sheets.slice(0, 5).flatMap((sheet) => sheet.usedRange ? [{ sheetName: sheet.name, range: sheet.usedRange, label: "used range" }] : []),
    resourceLinks: [contextResource(metadata.workbookContextId)],
    nextAction: "answer_now",
    warnings: []
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
  const used = parseA1Range(sheet.usedRange);
  if (!used || used.endColumn < columnToNumber("AJ")) {
    return undefined;
  }
  const summaryColumnSections = metadata.sections
    .filter((section) => section.sheetName === sheet.name)
    .filter((section) => {
      const parsed = parseA1Range(section.range);
      return parsed && parsed.startColumn >= columnToNumber("AG") && parsed.endColumn <= columnToNumber("AJ");
    });
  const merged = mergedSectionColumnRange(summaryColumnSections, summaryColumnSections[0]?.range ?? "");
  if (merged) {
    return merged;
  }
  return `AG1:AJ${Math.min(used.endRow, 60)}`;
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
  const best = parseA1Range(bestRange);
  if (!best) {
    return undefined;
  }
  const sameColumnSections = sections
    .map((section) => ({ section, parsed: parseA1Range(section.range) }))
    .filter((entry): entry is { section: WorkbookMetadata["sections"][number]; parsed: NonNullable<ReturnType<typeof parseA1Range>> } =>
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
  return {
    kind: "range_profile" as const,
    source: "live_read" as const,
    shape,
    metrics: {
      nonEmptyCount: flattened.length,
      numericCount: numeric.length,
      ...(numeric.length > 0 ? { sum, min: Math.min(...numeric), max: Math.max(...numeric), average: sum / numeric.length } : {})
    },
    sample,
    ...(emptySummary.emptyCells > 0 ? { emptySummary } : {}),
    ...(sparseRows ? { sparseRows } : {}),
    ...(includeRows && !sparseRows && nonEmptyRows.length > 0 ? { rows: nonEmptyRows.map((entry) => trimTrailingEmptyCells(entry.row)) } : {}),
    warning: flattened.length === 0 ? "No non-empty cells were found in the targeted range." : undefined
  };
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
  const origin = address ? parseA1Range(address) : undefined;
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

function styleFromRequest(request: string) {
  const lower = request.toLowerCase();
  return {
    ...(lower.includes("header") ? { fontBold: true, fillColor: "#D9EAF7", fontColor: "#1F2937", horizontalAlignment: "center" } : {}),
    ...(/\bbold\b/.test(lower) ? { fontBold: true } : {}),
    ...(/\bitalic\b/.test(lower) ? { fontItalic: true } : {})
  };
}

function styleEntriesFromInput(
  workbookId: WorkbookId,
  input: AgentRunInput
): Array<{ target: A1Range; style: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }> {
  const values = input.values as Record<string, unknown> | undefined;
  const rawEntries = values?.entries;
  const entries: Array<{ target: A1Range; style: Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] }> = [];
  if (Array.isArray(rawEntries)) {
    for (const rawEntry of rawEntries) {
      if (!rawEntry || typeof rawEntry !== "object") {
        continue;
      }
      const entry = rawEntry as Record<string, unknown>;
      const sheetName = stringValue(entry.sheetName);
      const address = stringValue(entry.address ?? entry.range);
      const style = entry.style && typeof entry.style === "object" ? entry.style as Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] : undefined;
      if (sheetName && address && style) {
        entries.push({ target: { workbookId, sheetName, address }, style });
      }
    }
  }
  const sheetName = stringValue(input.target?.sheetName ?? values?.sheetName);
  const address = stringValue(input.target?.range ?? values?.address ?? values?.range);
  const style = values?.style && typeof values.style === "object" ? values.style as Extract<ExcelOperation, { kind: "range.write_styles" }>["style"] : undefined;
  if (entries.length === 0 && sheetName && address && style) {
    entries.push({ target: { workbookId, sheetName, address }, style });
  }
  return entries;
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

function stripSheetName(address: string): string {
  const bang = address.lastIndexOf("!");
  return bang >= 0 ? address.slice(bang + 1).replace(/^'|'$/g, "") : address;
}

function findMentionedSheet(metadata: WorkbookMetadata, input: AgentRunInput): WorkbookMetadata["sheets"][number] | undefined {
  if (input.target?.sheetName) {
    const normalized = normalizeComparableText(input.target.sheetName);
    return metadata.sheets.find((sheet) => normalizeComparableText(sheet.name) === normalized);
  }
  const request = normalizeComparableText(input.request);
  return metadata.sheets.find((sheet) => request.includes(normalizeComparableText(sheet.name)));
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

function tableReadColumnsFromInput(input: AgentRunInput): Array<string | number> {
  const raw = input.values?.columns ?? input.values?.projectedColumns;
  return Array.isArray(raw)
    ? raw.filter((column): column is string | number => typeof column === "string" || typeof column === "number")
    : [];
}

function compactTableRowLimit(input: AgentRunInput, columnCount: number): number {
  const values = input.values as Record<string, unknown> | undefined;
  const requested = nonNegativeInteger(values?.rowLimit ?? values?.maxRows ?? input.budget?.maxExamples);
  if (requested !== undefined && requested > 0) {
    return requested;
  }
  const maxCells = input.budget?.maxPayloadBytes ? Math.max(25, Math.floor(input.budget.maxPayloadBytes / 80)) : undefined;
  if (maxCells !== undefined && columnCount > 0) {
    return Math.max(1, Math.min(50, Math.floor(maxCells / columnCount)));
  }
  return 50;
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
  return value.trim().toLowerCase().replace(/\s+/g, " ").replace(/[^\w ]/g, "");
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

function applyOutputBudget(output: Omit<AgentRunOutput, "telemetry">, input: AgentRunInput): Omit<AgentRunOutput, "telemetry"> {
  const maxExamples = Math.max(0, input.budget?.maxExamples ?? 10);
  const maxPayloadBytes = input.budget?.maxPayloadBytes;
  const maxEstimatedTokens = input.budget?.maxEstimatedTokens;
  const budgeted = stripUndefinedOptionals({
    ...output,
    proof: output.proof.slice(0, Math.min(output.proof.length, Math.max(1, maxExamples))),
    warnings: output.warnings.slice(0, Math.max(5, maxExamples)),
    ...(output.candidates ? { candidates: output.candidates.slice(0, maxExamples) } : {}),
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
    compact = stripUndefinedOptionals({
      ...compact,
      ...(compact.answer && compact.workbookContextId ? { answer: { resource: contextResource(String(compact.workbookContextId)).uri } } : {}),
      ...(compact.candidates ? { candidates: compact.candidates.slice(0, Math.min(compact.candidates.length, 3)).map(compactCandidate) } : {}),
      proof: compact.proof.slice(0, Math.min(compact.proof.length, 3)),
      ...(compact.changes ? { changes: compact.changes.slice(0, Math.min(compact.changes.length, 3)) } : {}),
      warnings: [...compact.warnings, "Agent response was compacted to satisfy the requested payload/token budget."]
    });
  }
  if (byteBudget && Buffer.byteLength(JSON.stringify(compact)) > byteBudget && compact.answer !== undefined) {
    compact = stripUndefinedOptionals({
      ...compact,
      answer: undefined,
      warnings: [...compact.warnings, "Answer details were omitted from the inline response; use resourceLinks for cached context."]
    });
  }
  return stripUndefinedOptionals(compact);
}

function compactCandidate(candidate: AgentCandidate): AgentCandidate {
  const next = { ...candidate };
  delete next.reason;
  delete next.nextRequestHint;
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
  if (next.confirmationToken === undefined) delete next.confirmationToken;
  if (next.operationId === undefined) delete next.operationId;
  if (next.workbookContextId === undefined) delete next.workbookContextId;
  return next;
}

function matrixCellCount(values: CellMatrix): number {
  return values.reduce((total, row) => total + row.length, 0);
}

function resolveUpdateTarget(metadata: WorkbookMetadata, input: AgentRunInput):
  | Extract<AgentTargetResolution, { ok: true }>
  | { ok: false; status: AgentRunOutput["status"]; summary: string; candidates?: AgentCandidate[]; nextAction: AgentRunOutput["nextAction"]; warnings: string[] } {
  if (input.values === undefined) {
    const candidates = findAgentCandidates(metadata, input).slice(0, 5);
    return {
      ok: false,
      status: candidates.length > 0 ? "AMBIGUOUS_TARGET" : "NEEDS_INPUT",
      summary: "Preview needs values and a resolvable workbook target.",
      ...(candidates.length > 0 ? { candidates } : {}),
      nextAction: candidates.length > 0 ? "call_with_target" : "ask_user",
      warnings: []
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
): action is "read_formula_patterns" | "get_formula_dependency_graph" | "trace_formula_precedents" | "trace_formula_dependents" | "find_formula_errors" | "explain_formula" {
  return action === "read_formula_patterns"
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
  const matching = sheet.headers.filter((header) => header.range === resolved.range);
  return matching.length > 0 ? matching : sheet.headers;
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
  const kind: AdvancedMutationKind = /\b(formula|formulas|calculate|calculation|total\s+row)\b/.test(request)
    ? "formula"
    : /\b(template|duplicate|copy)\b/.test(request)
      ? "template"
      : /\b(style|format|formatting|header\s+row)\b/.test(request)
        ? "style"
        : "other";
  if (kind !== "other" || /\b(repair|fix|pivot|chart)\b/.test(request)) {
    return {
      kind,
      safetyDecision: "manual_review:advanced_workflow",
      summary: "This request needs a formula/template/style/report-aware workflow, not a generic value update. Use a dedicated preview or advanced workflow path.",
      warning: "Auto mode blocked advanced mutation intent so formulas, templates, styles, pivots, or charts are not modified as plain values."
    };
  }
  return undefined;
}

function autoApplyDecision(input: AgentRunInput, preview: Omit<AgentRunOutput, "telemetry">): { allow: true; safetyDecision: string } | { allow: false; safetyDecision: string; reason: string } {
  if (preview.status !== "PREVIEW_READY") {
    return { allow: false, safetyDecision: "manual_review:preview_not_ready", reason: "preview did not produce an apply-ready operation" };
  }
  if (!input.values) {
    return { allow: false, safetyDecision: "manual_review:missing_values", reason: "values were not provided explicitly" };
  }
  if (!input.target?.range) {
    return { allow: false, safetyDecision: "manual_review:implicit_range", reason: "the target range was not explicit enough for auto-apply" };
  }
  if (isRiskyMutationRequest(input.request)) {
    return { allow: false, safetyDecision: "manual_review:risky_request", reason: "the request may be structural, destructive, or formula-sensitive" };
  }
  const matrix = objectToCellMatrix(input.values);
  if (matrixCellCount(matrix) > 16) {
    return { allow: false, safetyDecision: "manual_review:large_scope", reason: "the edit touches more than 16 cells" };
  }
  if (containsFormulaLikeValue(matrix)) {
    return { allow: false, safetyDecision: "manual_review:formula_values", reason: "formula writes require a formula-aware workflow" };
  }
  if ((preview.changes?.length ?? 0) === 0 || (preview.changes?.length ?? 0) > 16) {
    return { allow: false, safetyDecision: "manual_review:change_count", reason: "the previewed change count is not safe for auto-apply" };
  }
  return { allow: true, safetyDecision: "auto_apply:scoped_value_edit" };
}

function isRiskyMutationRequest(request: string): boolean {
  return /\b(delete|remove|clear|wipe|reset|drop|resize|rename|move|copy|convert|formula|formulas|template|pivot|chart|style|format|merge|unmerge|sort|filter|append|insert)\b/i.test(request);
}

function containsFormulaLikeValue(values: CellMatrix): boolean {
  return values.flat().some((value) => typeof value === "string" && value.trim().startsWith("="));
}

function isFormulaMutationRequest(request: string): boolean {
  return /\b(formula|formulas|calculate|calculation|total\s+row)\b/i.test(request);
}

function hasCellValues(values: CellMatrix): boolean {
  return values.some((row) => row.some((value) => value !== null && value !== undefined && value !== ""));
}

function firstNonEmptyRows(values: CellMatrix, limit: number): CellMatrix {
  const nonEmpty = values.filter((row) => row.some((value) => value !== null && value !== undefined && value !== ""));
  return (nonEmpty.length > 0 ? nonEmpty : values).slice(0, limit);
}

function objectToCellMatrix(values: Record<string, unknown>): CellMatrix {
  if (Array.isArray(values.rows)) return values.rows as CellMatrix;
  if (Array.isArray(values.values)) return values.values as CellMatrix;
  return [Object.values(values) as CellMatrix[number]];
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
  const source = explicitRangeTarget(metadata, values?.source);
  const destination = explicitRangeTarget(metadata, values?.destination ?? values?.target ?? input.target);
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
    ...(targetAddress ? { targetAddress: normalizeOperationRange(metadata, sheetName, targetAddress) } : {})
  };
  return request;
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

function explicitRangeTarget(metadata: WorkbookMetadata, value: unknown): A1Range | undefined {
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const target = value as Partial<AgentRunTarget> & { address?: unknown };
  const sheetName = typeof target.sheetName === "string" ? target.sheetName : undefined;
  const range = typeof target.range === "string" ? target.range : typeof target.address === "string" ? target.address : undefined;
  if (!sheetName || !range) {
    return undefined;
  }
  const sheet = metadata.sheets.find((candidate) => normalizeComparableText(candidate.name) === normalizeComparableText(sheetName));
  if (!sheet) {
    return undefined;
  }
  return { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: sheet.name, address: normalizeOperationRange(metadata, sheet.name, range) };
}

function rangeCopyTypeFromInput(input: AgentRunInput): NonNullable<Extract<ExcelOperation, { kind: "range.copy" }>["copyType"]> {
  const raw = (input.values as Record<string, unknown> | undefined)?.copyType;
  return raw === "all" || raw === "values" || raw === "formats" || raw === "formulas" ? raw : "all";
}

function dimensionsFromAddress(address: string): { rows: number; columns: number } | undefined {
  const range = parseA1Range(address);
  if (!range) {
    return undefined;
  }
  return {
    rows: range.endRow - range.startRow + 1,
    columns: range.endColumn - range.startColumn + 1
  };
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

function matrixShapeIssue(address: string, values: CellMatrix): string | undefined {
  const range = parseA1Range(address);
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
  const target = parseA1Range(address);
  if (!target) {
    return metadata.formulaRegions.some((region) => region.sheetName === sheetName && region.range === address);
  }
  return metadata.formulaRegions.some((region) => {
    if (region.sheetName !== sheetName) {
      return false;
    }
    const formulaRange = parseA1Range(region.range);
    return formulaRange ? rangesOverlap(target, formulaRange) : region.range === address;
  });
}

function cellAddressFor(range: string, rowIndex: number, columnIndex: number): string {
  const match = /^([A-Z]+)(\d+)/i.exec(range);
  const startColumn = match?.[1] ? columnToNumber(match[1]) : 1;
  const startRow = match?.[2] ? Number(match[2]) : 1;
  return `${columnLetter(startColumn + columnIndex - 1)}${startRow + rowIndex}`;
}

function cellCountFromAddress(address: string): number | undefined {
  const range = parseA1Range(address);
  return range ? Math.max(1, range.endColumn - range.startColumn + 1) * Math.max(1, range.endRow - range.startRow + 1) : undefined;
}

function columnToNumber(column: string): number {
  return column.toUpperCase().split("").reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function numberToColumn(column: number): string {
  let value = column;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function parseA1Range(address: string): { startColumn: number; startRow: number; endColumn: number; endRow: number } | undefined {
  const normalized = address.includes("!") ? address.split("!").pop() ?? address : address;
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(normalized.replace(/\$/g, ""));
  if (!match?.[1] || !match[2]) {
    return undefined;
  }
  const startColumn = columnToNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = match[3] ? columnToNumber(match[3]) : startColumn;
  const endRow = match[4] ? Number(match[4]) : startRow;
  return {
    startColumn: Math.min(startColumn, endColumn),
    startRow: Math.min(startRow, endRow),
    endColumn: Math.max(startColumn, endColumn),
    endRow: Math.max(startRow, endRow)
  };
}

function rangesOverlap(left: { startColumn: number; startRow: number; endColumn: number; endRow: number }, right: { startColumn: number; startRow: number; endColumn: number; endRow: number }): boolean {
  return left.startColumn <= right.endColumn
    && left.endColumn >= right.startColumn
    && left.startRow <= right.endRow
    && left.endRow >= right.startRow;
}

function rangesMayOverlap(leftAddress: string, rightAddress: string): boolean {
  const left = parseA1Range(leftAddress);
  const right = parseA1Range(rightAddress);
  return Boolean(left && right && rangesOverlap(left, right));
}

function contextResource(workbookContextId: string) {
  return { uri: `excel://agent/contexts/${workbookContextId}`, name: "workbook context", description: "Cached workbook metadata summary.", mimeType: "application/json" as const };
}

function operationResource(operationId: string) {
  return { uri: `excel://agent/operations/${operationId}`, name: "pending operation", description: "Pending previewed workbook update.", mimeType: "application/json" as const };
}
