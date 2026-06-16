import {
  makeId,
  runtimeError,
  type AgentCandidate,
  type AgentOperationId,
  type AgentRunInput,
  type AgentRunMode,
  type AgentRunOutput,
  type A1Range,
  type BatchRequest,
  type CellMatrix,
  type ExcelOperation,
  type OperationId,
  type TransactionId,
  type BackupId,
  type WorkbookId
} from "@components-kit/open-workbook-protocol";
import { columnLetter, type WorkbookMetadata, WorkbookMetadataCache } from "./workbook-metadata-cache.js";
import { AgentOperationStore } from "./agent-operation-store.js";
import { WorkbookMetadataBuilder } from "./workbook-metadata-builder.js";
import { findAgentCandidates, resolveAgentReadTarget, resolveAgentUpdateTarget } from "./agent-target-resolver.js";
import type { RuntimeService } from "./runtime-service.js";

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
    const runMetrics: AgentRunMetrics = { internalReadCount: 0, fullReadCellCount: 0, validationStatus: "not_run" };
    const mode = input.mode ?? "auto";
    const finish = (output: Omit<AgentRunOutput, "telemetry">, cacheHit = false): AgentRunOutput => {
      const preBudgetPayloadBytes = Buffer.byteLength(JSON.stringify(output));
      const budgeted = applyOutputBudget(output, input);
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
          estimatedTokensSaved: Math.max(0, Math.ceil((preBudgetPayloadBytes - payloadBytes) / 4))
        }
      };
    };

    try {
      if (mode === "status") {
        internalCallCount += 1;
        const status = this.runtime.getStatus();
        return finish({
          status: "SUCCESS",
          mode,
          summary: status.activeAddinConnected ? "Open Workbook backend is connected to Excel." : "Open Workbook backend is running, but no Excel add-in is connected.",
          answer: status,
          proof: [],
          resourceLinks: [{ uri: "excel://runtime/status", name: "runtime status", description: "Runtime connection and capability status.", mimeType: "application/json" }],
          nextAction: status.activeAddinConnected ? "answer_now" : "manual_review",
          warnings: status.activeAddinConnected ? [] : ["No active Excel add-in session is connected."]
        });
      }

      if (mode === "apply_update") {
        return finish(await this.applyUpdate(input));
      }

      if (mode === "rollback") {
        return finish(await this.rollback(input));
      }

      const { metadata, cacheHit } = await this.metadataBuilder.getOrBuild({
        ...(input.workbookContextId ? { workbookContextId: String(input.workbookContextId) } : {}),
        ...(input.target?.workbookId !== undefined ? { workbookId: input.target.workbookId } : {}),
        ...(input.target?.workbookName !== undefined ? { workbookName: input.target.workbookName } : {})
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

      const effectiveMode = mode === "auto" ? inferAgentMode(input.request) : mode;
      if (mode === "auto" && effectiveMode === "preview_update") {
        const advancedMutation = advancedMutationDecision(input);
        if (advancedMutation) {
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
        return finish(this.findOutput(metadata, input, mode), cacheHit);
      }
      if (effectiveMode === "preview_update") {
        internalCallCount += 1;
        const preview = await this.previewUpdate(metadata, input, mode);
        if (mode !== "auto" || preview.status !== "PREVIEW_READY") {
          return finish(preview, cacheHit);
        }
        const autoDecision = autoApplyDecision(input, preview);
        runMetrics.safetyDecision = autoDecision.safetyDecision;
        if (!autoDecision.allow) {
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
        sheetCount: metadata.sheets.length,
        tableCount: metadata.tables.length,
        namedRangeCount: metadata.namedRanges.length,
        sheets: metadata.sheets.map((sheet) => ({ name: sheet.name, kind: sheet.kind, usedRange: sheet.usedRange }))
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
    const profile = await this.readAndProfileRange(metadata.workbook.workbookId as WorkbookId, resolved.sheetName, resolved.range, runMetrics);
    return {
      status: "SUCCESS",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      summary: `Answered from ${resolved.candidate.label} on ${resolved.sheetName} using a targeted compact read.`,
      answer: profile,
      metrics: profile.metrics,
      candidates: candidates.slice(0, 5),
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: resolved.candidate.label }],
      resourceLinks: [contextResource(metadata.workbookContextId)],
      nextAction: "answer_now",
      warnings: profile.warning ? [profile.warning] : []
    };
  }

  private async previewUpdate(metadata: WorkbookMetadata, input: AgentRunInput, requestedMode: AgentRunMode): Promise<Omit<AgentRunOutput, "telemetry">> {
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
    if (overlapsFormulaRegion(metadata, resolved.sheetName, resolved.range)) {
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
    const pending = this.operations.create({
      workbookContextId: metadata.workbookContextId,
      workbookId: metadata.workbook.workbookId as WorkbookId,
      operations: [operation],
      changes,
      summary: `Prepared ${changes.length} cell update(s) on ${resolved.sheetName}!${resolved.range}.`,
      sourceFingerprintHash: metadata.fingerprint.structureHash
    });
    return {
      status: "PREVIEW_READY",
      mode: requestedMode,
      workbookContextId: metadata.workbookContextId,
      operationId: pending.operationId,
      confirmationToken: pending.confirmationToken,
      summary: pending.summary,
      changes,
      proof: [{ sheetName: resolved.sheetName, range: resolved.range, label: "preview target" }],
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
        changes: pending.changes,
        proof: [],
        resourceLinks: [contextResource(current.metadata.workbookContextId), operationResource(String(pending.operationId))],
        nextAction: "retry_after_refresh",
        warnings: ["Target fingerprints changed after preview."]
      };
    }
    const result = await this.runtime.applyBatch({ workbookId: pending.workbookId, operations: pending.operations, mode: "apply", idempotencyKey: `agent:${operationId}` });
    if (result.ok) {
      this.operations.delete(operationId);
    }
    const validation = result.ok ? await this.runtime.validateWorkbook({ workbookId: pending.workbookId }) : undefined;
    return {
      status: result.ok === false || validation?.ok === false ? "VALIDATION_FAILED" : "SUCCESS",
      mode: "apply_update",
      workbookContextId: pending.workbookContextId,
      operationId: pending.operationId,
      summary: result.ok === false ? "Workbook update failed." : pending.summary.replace("Prepared", "Applied"),
      answer: { result, validation },
      changes: pending.changes,
      proof: pending.changes.flatMap((change) => change.range ? [{ sheetName: change.sheetName, range: change.range, label: "applied target" }] : []).slice(0, 1),
      resourceLinks: [],
      nextAction: result.ok === false || validation?.ok === false ? "manual_review" : "answer_now",
      warnings: validation?.issues?.slice(0, 5).map((issue) => issue.message) ?? []
    };
  }

  private async rollback(input: AgentRunInput): Promise<Omit<AgentRunOutput, "telemetry">> {
    const transactionId = input.target?.entity?.startsWith("tx_") ? input.target.entity : input.operationId;
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
    const flattened = values.flat().filter((value) => value !== null && value !== undefined && value !== "");
    const numeric = flattened.map((value) => typeof value === "number" ? value : typeof value === "string" ? Number(value.replace(/[$,]/g, "")) : NaN).filter(Number.isFinite);
    const sum = numeric.reduce((total, value) => total + value, 0);
    return {
      shape: { rows: values.length, columns: values[0]?.length ?? 0 },
      metrics: {
        nonEmptyCount: flattened.length,
        numericCount: numeric.length,
        ...(numeric.length > 0 ? { sum, min: Math.min(...numeric), max: Math.max(...numeric), average: sum / numeric.length } : {})
      },
      sample: values.slice(0, 5).map((row) => row.slice(0, 8)),
      warning: numeric.length === 0 ? "No numeric values were found in the targeted range." : undefined
    };
  }

  private async readRangeValues(workbookId: WorkbookId, sheetName: string, address: string, runMetrics?: AgentRunMetrics): Promise<CellMatrix> {
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
    const values = (result as { readData?: Array<{ snapshot?: { values?: CellMatrix; text?: string[][] } }> }).readData?.[0]?.snapshot?.values ?? [];
    if (runMetrics) {
      runMetrics.internalReadCount += 1;
      runMetrics.fullReadCellCount += matrixCellCount(values);
    }
    return values;
  }
}

interface AgentRunMetrics {
  internalReadCount: number;
  fullReadCellCount: number;
  autoApplied?: boolean;
  safetyDecision?: string;
  previewOperationId?: string;
  validationStatus: "passed" | "failed" | "not_run";
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
      ...(compact.candidates ? { candidates: compact.candidates.slice(0, Math.min(compact.candidates.length, 3)) } : {}),
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
  | { ok: true; sheetName: string; range: string }
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
  if (resolved.ok && !input.target?.range && (resolved.candidate.kind === "sheet" || resolved.candidate.kind === "table")) {
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

function inferAgentMode(request: string): AgentRunMode {
  return /\b(update|set|change|replace|write|fill|append|edit|fix|repair|create|insert|delete|remove|clear|rename)\b/i.test(request) ? "preview_update" : "answer";
}

function advancedMutationDecision(input: AgentRunInput): { safetyDecision: string; summary: string; warning: string } | undefined {
  const request = input.request.toLowerCase();
  if (/\b(formula|formulas|calculate|calculation|repair|fix|template|pivot|chart|style|format|formatting)\b/.test(request)) {
    return {
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

function objectToCellMatrix(values: Record<string, unknown>): CellMatrix {
  if (Array.isArray(values.rows)) return values.rows as CellMatrix;
  if (Array.isArray(values.values)) return values.values as CellMatrix;
  return [Object.values(values) as CellMatrix[number]];
}

function isSparseBroadWrite(address: string, values: CellMatrix): boolean {
  const matrixCells = values.reduce((sum, row) => sum + row.length, 0);
  const nonEmpty = values.flat().filter((value) => value !== null && value !== undefined && value !== "").length;
  const touched = Math.max(matrixCells, cellCountFromAddress(address) ?? 0);
  return touched >= 8 && nonEmpty > 0 && nonEmpty / touched <= 0.25 && touched - nonEmpty >= 4;
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

function contextResource(workbookContextId: string) {
  return { uri: `excel://agent/contexts/${workbookContextId}`, name: "workbook context", description: "Cached workbook metadata summary.", mimeType: "application/json" as const };
}

function operationResource(operationId: string) {
  return { uri: `excel://agent/operations/${operationId}`, name: "pending operation", description: "Pending previewed workbook update.", mimeType: "application/json" as const };
}
