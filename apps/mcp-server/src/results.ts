import type { AgentRunOutput } from "@components-kit/open-workbook-protocol";

export function jsonResult(value: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(value, null, 2)
      }
    ]
  };
}

export function agentJsonResult(value: AgentRunOutput) {
  const jsonSafeValue = JSON.parse(JSON.stringify(value)) as AgentRunOutput;
  const structuredContent = compactStructuredAgentResult(jsonSafeValue);
  return {
    content: [
      {
        type: "text" as const,
        text: compactAgentResultText(structuredContent)
      }
    ],
    structuredContent
  };
}

function compactStructuredAgentResult(value: AgentRunOutput): AgentRunOutput {
  const responseMode = value.continuation?.responseMode ?? "brief";
  const maxExamples = responseMode === "verbose" ? 10 : responseMode === "standard" ? 5 : 3;
  const compact: AgentRunOutput = {
    ...value,
    ...(value.answer !== undefined ? { answer: compactStructuredAnswer(value.answer, responseMode, value.continuation) } : {}),
    proof: value.proof.slice(0, maxExamples),
    resourceLinks: value.resourceLinks.slice(0, maxExamples),
    warnings: value.warnings.slice(0, Math.max(3, maxExamples)),
    ...(value.candidates ? { candidates: value.candidates.slice(0, maxExamples).map((candidate) => responseMode === "verbose" ? candidate : compactCandidateForMcp(candidate)) } : {}),
    ...(value.changes ? { changes: value.changes.slice(0, maxExamples) } : {}),
    telemetry: compactTelemetry(value.telemetry, responseMode)
  };
  return enforceStructuredContentBudget(compact, responseMode);
}

function compactCandidateForMcp(candidate: NonNullable<AgentRunOutput["candidates"]>[number]): NonNullable<AgentRunOutput["candidates"]>[number] {
  return stripUndefinedObject({
    id: candidate.id,
    kind: candidate.kind,
    label: candidate.label,
    sheetName: candidate.sheetName,
    tableName: candidate.tableName,
    range: candidate.range,
    semanticRole: candidate.semanticRole,
    confidence: candidate.confidence,
    nextRequestHint: candidate.nextRequestHint
  }) as unknown as NonNullable<AgentRunOutput["candidates"]>[number];
}

function compactStructuredAnswer(answer: unknown, responseMode: "brief" | "standard" | "verbose", continuation?: AgentRunOutput["continuation"]): unknown {
  if (!answer || typeof answer !== "object") {
    return answer;
  }
  if (responseMode === "verbose") {
    return answer;
  }
  const compact = stripHeavyAnswerFields(answer);
  if (compact && typeof compact === "object" && !Array.isArray(compact)) {
    const typed = compact as Record<string, unknown>;
    if (typed.resultUri === undefined && continuation?.resultUri) {
      typed.resultUri = continuation.resultUri;
    }
    if (typed.fullResultUri === undefined && continuation?.fullResultUri) {
      typed.fullResultUri = continuation.fullResultUri;
    }
  }
  return compact;
}

function stripHeavyAnswerFields(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.slice(0, 20).map(stripHeavyAnswerFields);
  }
  if (!value || typeof value !== "object") {
    return value;
  }
  const source = value as Record<string, unknown>;
  const preserveFormulaFields = source.kind === "formula_read" || source.kind === "formula_patterns";
  const next: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(source)) {
    if (["sparseRows", "rowMetadata", "debug", "routeReasons", "workflowReasons"].includes(key)) {
      continue;
    }
    if (!preserveFormulaFields && ["formulas", "text", "numberFormat"].includes(key)) {
      continue;
    }
    next[key] = stripHeavyAnswerFields(entry);
  }
  return next;
}

function compactTelemetry(telemetry: AgentRunOutput["telemetry"], responseMode: "brief" | "standard" | "verbose"): AgentRunOutput["telemetry"] {
  if (responseMode === "verbose") {
    return telemetry;
  }
  return {
    internalCallCount: telemetry.internalCallCount,
    payloadBytes: telemetry.payloadBytes,
    estimatedTokens: telemetry.estimatedTokens,
    elapsedMs: telemetry.elapsedMs,
    cacheHit: telemetry.cacheHit,
    ...(telemetry.autoApplied !== undefined ? { autoApplied: telemetry.autoApplied } : {}),
    ...(telemetry.safetyDecision !== undefined ? { safetyDecision: telemetry.safetyDecision } : {}),
    ...(telemetry.validationStatus !== undefined ? { validationStatus: telemetry.validationStatus } : {}),
    ...(telemetry.metadataCacheStatus !== undefined ? { metadataCacheStatus: telemetry.metadataCacheStatus } : {}),
    ...(telemetry.internalReadCount !== undefined ? { internalReadCount: telemetry.internalReadCount } : {}),
    ...(telemetry.fullReadCellCount !== undefined ? { fullReadCellCount: telemetry.fullReadCellCount } : {}),
    ...(telemetry.fullReadUsed !== undefined ? { fullReadUsed: telemetry.fullReadUsed } : {}),
    ...(telemetry.operationRisk !== undefined ? { operationRisk: telemetry.operationRisk } : {}),
    ...(telemetry.targetFingerprintStatus !== undefined ? { targetFingerprintStatus: telemetry.targetFingerprintStatus } : {}),
    ...(telemetry.intentAction !== undefined ? { intentAction: telemetry.intentAction } : {})
  };
}

function enforceStructuredContentBudget(value: AgentRunOutput, responseMode: "brief" | "standard" | "verbose"): AgentRunOutput {
  const maxBytes = responseMode === "verbose" ? 24_000 : responseMode === "standard" ? 12_000 : 6_000;
  if (Buffer.byteLength(JSON.stringify(value)) <= maxBytes) {
    return value;
  }
  const answer = value.answer && typeof value.answer === "object"
    ? minimalStructuredAnswer(value.answer as Record<string, unknown>, value.continuation)
    : value.answer;
  const compact: AgentRunOutput = stripUndefinedObject({
    ...value,
    ...(answer !== undefined ? { answer } : { answer: undefined }),
    proof: value.proof.slice(0, 3),
    resourceLinks: value.resourceLinks.slice(0, 3),
    ...(value.candidates ? { candidates: value.candidates.slice(0, 3) } : {}),
    ...(value.changes ? { changes: value.changes.slice(0, 3) } : {}),
    warnings: [...value.warnings.slice(0, 3), compactedBudgetWarning(value)],
    telemetry: compactTelemetry(value.telemetry, "brief")
  }) as unknown as AgentRunOutput;
  if (Buffer.byteLength(JSON.stringify(compact)) <= maxBytes) {
    return compact;
  }
  return stripUndefinedObject({
    ...compact,
    warnings: [...compact.warnings.slice(0, 3), compactedBudgetWarning(value)]
  }, ["answer", "candidates", "changes"]) as unknown as AgentRunOutput;
}

function minimalStructuredAnswer(answer: Record<string, unknown>, continuation?: AgentRunOutput["continuation"]): Record<string, unknown> {
  if (answer.kind === "data_validation_summary") {
    return minimalDataValidationSummaryAnswer(answer);
  }
  if (answer.kind === "similar_rows") {
    return minimalSimilarRowsAnswer(answer, continuation);
  }
  if (answer.kind === "style_reference_candidates") {
    return minimalStyleReferenceAnswer(answer, continuation);
  }
  return {
    kind: answer.kind ?? "compacted_answer",
    source: answer.source,
    sheetName: answer.sheetName,
    tableName: answer.tableName,
    range: answer.range,
    dataRange: answer.dataRange,
    resultUri: answer.resultUri ?? continuation?.resultUri,
    fullResultUri: answer.fullResultUri ?? continuation?.fullResultUri
  };
}

function compactedBudgetWarning(value: AgentRunOutput): string {
  const answer = value.answer && typeof value.answer === "object" ? value.answer as Record<string, unknown> : undefined;
  if (answer?.kind === "data_validation_summary") {
    return "Compacted for budget; validation metadata/options are complete inline for dropdown inspection. Do not fetch fullResultUri unless the user asks for raw audit metadata.";
  }
  return "Compacted for budget; pass continuation.fullResultUri to excel.agent.run for hidden detail.";
}

function minimalDataValidationSummaryAnswer(answer: Record<string, unknown>): Record<string, unknown> {
  return stripUndefinedObject({
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
    guidance: answer.guidance
  });
}

function minimalSimilarRowsAnswer(answer: Record<string, unknown>, continuation?: AgentRunOutput["continuation"]): Record<string, unknown> {
  const rows = Array.isArray(answer.rows)
    ? answer.rows.slice(0, 5).map((row) => {
        const typed = row && typeof row === "object" ? row as Record<string, unknown> : {};
        return stripUndefinedObject({
          sheetName: typed.sheetName,
          sheetRowNumber: typed.sheetRowNumber,
          range: typed.range,
          columns: minimalReferenceColumns(typed.columns),
          matchedSignals: Array.isArray(typed.matchedSignals) ? typed.matchedSignals.slice(0, 6) : undefined,
          whyMatched: typed.whyMatched
        });
      })
    : undefined;
  return stripUndefinedObject({
    kind: answer.kind,
    source: answer.source,
    sourceMode: answer.sourceMode,
    predicates: Array.isArray(answer.predicates) ? answer.predicates.slice(0, 6) : undefined,
    rows,
    resultUri: answer.resultUri ?? continuation?.resultUri,
    fullResultUri: answer.fullResultUri ?? continuation?.fullResultUri
  });
}

function minimalReferenceColumns(columns: unknown): unknown[] | undefined {
  if (!Array.isArray(columns)) {
    return undefined;
  }
  return columns
    .filter((column) => {
      if (!column || typeof column !== "object") return false;
      const typed = column as Record<string, unknown>;
      const label = `${String(typed.name ?? "")} ${String(typed.role ?? "")}`;
      return /date|description|type|direction|amount|actual|variance|transfer|from|to|truck|job|vendor|customer|account|note|category|status|identifier/i.test(label);
    })
    .slice(0, 8)
    .map((column) => {
      const typed = column as Record<string, unknown>;
      return stripUndefinedObject({
        letter: typed.letter,
        name: typed.name,
        value: typed.value
      });
    });
}

function minimalStyleReferenceAnswer(answer: Record<string, unknown>, continuation?: AgentRunOutput["continuation"]): Record<string, unknown> {
  const candidates = Array.isArray(answer.candidates)
    ? answer.candidates.slice(0, 5).map((candidate) => {
        const typed = candidate && typeof candidate === "object" ? candidate as Record<string, unknown> : {};
        return stripUndefinedObject({
          sheetName: typed.sheetName,
          range: typed.range,
          label: typed.label,
          sourceKind: typed.sourceKind,
          confidence: typed.confidence,
          reason: typed.reason,
          styleSummary: typed.styleSummary
        });
      })
    : undefined;
  return stripUndefinedObject({
    kind: answer.kind,
    source: answer.source,
    candidates,
    resultUri: answer.resultUri ?? continuation?.resultUri,
    fullResultUri: answer.fullResultUri ?? continuation?.fullResultUri
  });
}

function stripUndefinedObject<T extends Record<string, unknown>>(value: T, omitKeys: string[] = []): Record<string, unknown> {
  const omit = new Set(omitKeys);
  const next: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (omit.has(key) || entry === undefined) {
      continue;
    }
    next[key] = entry;
  }
  return next;
}

function compactAgentResultText(value: AgentRunOutput): string {
  const lines = [
    `${value.status} ${value.mode}: ${value.summary}`,
    `nextAction: ${value.nextAction}`
  ];
  if (value.taskOutcome) {
    lines.push(`taskOutcome: ${value.taskOutcome}`);
  }
  if (typeof value.maxRecommendedFollowupCalls === "number") {
    lines.push(`maxRecommendedFollowupCalls: ${value.maxRecommendedFollowupCalls}`);
  }
  if (value.agentInstruction) {
    lines.push(`agentInstruction: ${compactLine(value.agentInstruction, 180)}`);
  }
  if (value.finalAnswer && value.finalAnswer !== value.summary) {
    lines.push(`finalAnswer: ${compactLine(value.finalAnswer, 180)}`);
  }
  if (value.workbookContextId) {
    lines.push(`workbookContextId: ${value.workbookContextId}`);
  }
  if (value.operationId) {
    lines.push(`operationId: ${value.operationId}`);
  }
  if (value.confirmationToken) {
    lines.push("confirmationToken: present in structuredContent");
  }
  const dataAvailability = compactDataAvailabilityText(value);
  if (dataAvailability) {
    lines.push(dataAvailability);
  }
  if (value.continuation?.resultUri) {
    lines.push(`resultUri: ${value.continuation.resultUri}`);
  }
  if (value.continuation?.fullResultUri) {
    lines.push(`fullResultUri: ${value.continuation.fullResultUri}`);
    lines.push("full detail: excel:// not web; never Webfetch/browser. Read via excel.agent.run continuation.fullResultUri.");
  }
  if (value.resourceLinks.length > 0) {
    lines.push(`resources: ${value.resourceLinks.map((resource) => resource.uri).join(", ")}`);
  }
  if (value.invalidatedContextIds?.length) {
    lines.push(`invalidatedContextIds: ${value.invalidatedContextIds.join(", ")}`);
  }
  if (value.invalidatedResourceUris?.length) {
    lines.push(`invalidatedResourceUris: ${value.invalidatedResourceUris.join(", ")}`);
  }
  if (value.warnings.length > 0) {
    lines.push(`warnings: ${value.warnings.slice(0, 3).join(" | ")}`);
  }
  const payloadBytes = value.telemetry?.payloadBytes;
  const estimatedTokens = value.telemetry?.estimatedTokens;
  if (typeof payloadBytes === "number" || typeof estimatedTokens === "number") {
    lines.push(`telemetry: ${payloadBytes ?? "?"} bytes, ~${estimatedTokens ?? "?"} tokens`);
  }
  return lines.join("\n");
}

function compactLine(value: string, maxLength: number): string {
  const normalized = value.replace(/\s+/g, " ").trim();
  return normalized.length <= maxLength ? normalized : `${normalized.slice(0, Math.max(0, maxLength - 1))}…`;
}

function compactDataAvailabilityText(value: AgentRunOutput): string | undefined {
  const answer = value.answer;
  if (!answer || typeof answer !== "object") {
    return undefined;
  }
  const typed = answer as Record<string, unknown>;
  if (typed.kind === "workbook_summary" || typed.kind === "workbook_overview" || typed.kind === "sheet_summary") {
    return "data: summary is complete from cached metadata; do not fetch full detail unless the user asks for all raw rows or exact cell values";
  }
  if (typed.kind === "data_validation_summary") {
    return "data: validation metadata/options complete inline for dropdown inspection; do not fetch full detail unless the user asks for raw audit metadata";
  }
  const exactRows = countMatrixRows(typed.values) ?? countMatrixRows(typed.rows) ?? countSparseRows(typed.sparseRows);
  if (exactRows !== undefined) {
    return `data: ${exactRows} exact row${exactRows === 1 ? "" : "s"} inline in structuredContent`;
  }
  const previewRows = countMatrixRows(typed.valuesPreview) ?? countMatrixRows(typed.sample);
  if (previewRows !== undefined) {
    const truncated = typed.previewTruncated === true || typed.truncated === true ? "; truncated" : "";
    const handle = value.continuation?.fullResultUri ? "; stored detail handle available through excel.agent.run" : "";
    return `data: ${previewRows} preview row${previewRows === 1 ? "" : "s"} inline in structuredContent${truncated}${handle}`;
  }
  if (value.continuation?.fullResultUri) {
    return "data: compact summary inline; exact rows/raw values need excel.agent.run continuation.fullResultUri";
  }
  return undefined;
}

function countMatrixRows(value: unknown): number | undefined {
  return Array.isArray(value) ? value.length : undefined;
}

function countSparseRows(value: unknown): number | undefined {
  return Array.isArray(value) ? value.length : undefined;
}
