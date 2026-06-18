import {
  makeId,
  runtimeError,
  type AgentCandidate,
  type AgentOperationId,
  type AgentProofReference,
  type AgentRunInput,
  type AgentRunMode,
  type AgentRunOutput,
  type AgentRunTarget,
  type A1Range,
  type BatchRequest,
  type CellMatrix,
  type ExcelOperation,
  type OperationId,
  type TableAppendRowsRequest,
  type TableSortRequest,
  type TransactionId,
  type BackupId,
  type WorkbookId
} from "@components-kit/open-workbook-protocol";
import { columnLetter, type HeaderMetadata, type TableMetadata, type WorkbookMetadata, WorkbookMetadataCache } from "./workbook-metadata-cache.js";
import { AgentOperationStore } from "./agent-operation-store.js";
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

  async run(input: AgentRunInput): Promise<AgentRunOutput> {
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
        if (!readiness.ok) {
          return finish({
            status: "NEEDS_INPUT",
            mode,
            summary: connectionReadinessSummary(readiness.connectionState),
            answer: { ...status, connectionState: readiness.connectionState },
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
          answer: status,
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
        const validation = await this.runtime.validateWorkbook({ workbookId: metadata.workbook.workbookId as WorkbookId });
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
    const workbookAnswer = workbookOverviewAnswer(metadata, input, requestedMode);
    if (workbookAnswer) {
      return workbookAnswer;
    }
    const sectionAnswer = sectionAnswerOutput(metadata, input, requestedMode);
    if (sectionAnswer) {
      return sectionAnswer;
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
      case "copy_template_sheet":
        return this.previewTemplateCleanup(metadata, input, requestedMode);
      case "first_open_reviewed":
        return this.previewFirstStatusMatchUpdate(metadata, input, requestedMode);
      case "sort_table":
        return this.previewTableSort(metadata, input, requestedMode);
      case "filter_range":
        return resolved ? this.previewAutoFilter(metadata, input, requestedMode, resolved) : undefined;
      case "autofit_columns":
        return resolved ? this.previewAutofit(metadata, input, requestedMode, resolved) : undefined;
      case "clear_values":
        return resolved ? this.previewClearValues(metadata, input, requestedMode, resolved) : undefined;
      case "format_range":
        return resolved ? this.previewStyleUpdate(metadata, input, requestedMode, resolved) : undefined;
    }
  }

  private previewWorkbookOperation(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, kind: "workbook.calculate" | "workbook.save"): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind,
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: kind === "workbook.save" ? "workbook" : "none",
      reason: input.request,
      ...(kind === "workbook.calculate" ? { calculationType: "full" as const } : {})
    };
    const label = kind === "workbook.save" ? "save workbook" : "recalculate workbook";
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: metadata.workbook.activeSheet ?? metadata.sheets[0]?.name ?? "", after: label }], `Prepared ${label}.`, { kind: `${kind}_preview` });
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

  private previewAutofit(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode, resolved: Extract<AgentTargetResolution, { ok: true }>): Omit<AgentRunOutput, "telemetry"> {
    const operation: ExcelOperation = {
      kind: "range.autofit_columns",
      operationId: makeId<OperationId>("op"),
      workbookId: metadata.workbook.workbookId as WorkbookId,
      destructiveLevel: "format",
      reason: input.request,
      target: { workbookId: metadata.workbook.workbookId as WorkbookId, sheetName: resolved.sheetName, address: resolved.range }
    };
    return this.previewBatchOperation(metadata, requestedMode, [operation], [{ sheetName: resolved.sheetName, range: resolved.range, after: "autofit columns" }], `Prepared autofit columns on ${resolved.sheetName}!${resolved.range}.`, { kind: "autofit_preview", sheetName: resolved.sheetName, range: resolved.range });
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
    const table = resolveSortTable(metadata, input);
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

  private previewFormulaUpdate(
    metadata: WorkbookMetadata,
    input: AgentRunInput,
    requestedMode: AgentRunMode,
    resolved: Extract<AgentTargetResolution, { ok: true }>,
    matrix: CellMatrix
  ): Omit<AgentRunOutput, "telemetry"> {
    const formulas = matrix.map((row) => row.map((value) => typeof value === "string" && value.trim().startsWith("=") ? value : null));
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
      answer: { kind: "formula_preview", sheetName: resolved.sheetName, range: resolved.range, formulaCount: changes.length },
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
    const result = pending.action.kind === "batch"
      ? await this.runtime.applyBatch({ workbookId: pending.workbookId, operations: pending.action.operations, mode: "apply", idempotencyKey: `agent:${operationId}` })
      : pending.action.kind === "table.append_rows"
        ? await this.runtime.appendTableRows(pending.action.request)
        : await this.runtime.sortTable(pending.action.request);
    const applyFailed = result.ok === false;
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
    return this.operations.create({
      workbookContextId: metadata.workbookContextId,
      workbookId: metadata.workbook.workbookId as WorkbookId,
      action: input.action,
      changes: input.changes,
      summary: input.summary,
      risk,
      sourceFingerprintHash: metadata.fingerprint.structureHash,
      sourceTargetFingerprintHash: targetFingerprintHash(metadata, input.changes)
    });
  }
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

function resolveSortTable(metadata: WorkbookMetadata, input: AgentRunInput): TableMetadata | undefined {
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
