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
  const resourceLinks = Array.isArray(jsonSafeValue.resourceLinks) ? jsonSafeValue.resourceLinks : [];
  return {
    content: [
      {
        type: "text" as const,
        text: compactAgentResultText(structuredContent)
      }
    ],
    structuredContent,
    resources: resourceLinks
      .filter((resource) => typeof resource?.uri === "string")
      .map((resource) => ({
        uri: resource.uri,
        name: resource.name ?? resource.uri,
        description: resource.description,
        mimeType: resource.mimeType ?? "application/json"
      }))
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
  const next: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(source)) {
    if (["sparseRows", "rowMetadata", "formulas", "text", "numberFormat", "debug", "routeReasons", "workflowReasons"].includes(key)) {
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
    warnings: [...value.warnings.slice(0, 3), "MCP structuredContent was compacted to stay within response budget; use fullResultUri for hidden detail."],
    telemetry: compactTelemetry(value.telemetry, "brief")
  }) as unknown as AgentRunOutput;
  if (Buffer.byteLength(JSON.stringify(compact)) <= maxBytes) {
    return compact;
  }
  return stripUndefinedObject({
    ...compact,
    warnings: [...compact.warnings.slice(0, 3), "MCP structuredContent answer details were omitted; use resourceLinks/fullResultUri for detail."]
  }, ["answer", "candidates", "changes"]) as unknown as AgentRunOutput;
}

function minimalStructuredAnswer(answer: Record<string, unknown>, continuation?: AgentRunOutput["continuation"]): Record<string, unknown> {
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
    lines.push("full detail: call excel.agent.run with fullResultUri in request/continuation; do not use webfetch for excel:// handles");
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
  const exactRows = countMatrixRows(typed.values) ?? countMatrixRows(typed.rows) ?? countSparseRows(typed.sparseRows);
  if (exactRows !== undefined) {
    return `data: ${exactRows} exact row${exactRows === 1 ? "" : "s"} inline in structuredContent`;
  }
  const previewRows = countMatrixRows(typed.valuesPreview) ?? countMatrixRows(typed.sample);
  if (previewRows !== undefined) {
    const truncated = typed.previewTruncated === true || typed.truncated === true ? "; truncated" : "";
    const handle = value.continuation?.fullResultUri ? "; fullResultUri available" : "";
    return `data: ${previewRows} preview row${previewRows === 1 ? "" : "s"} inline in structuredContent${truncated}${handle}`;
  }
  if (value.continuation?.fullResultUri) {
    return "data: compact summary inline; fullResultUri available for exact rows/raw values";
  }
  return undefined;
}

function countMatrixRows(value: unknown): number | undefined {
  return Array.isArray(value) ? value.length : undefined;
}

function countSparseRows(value: unknown): number | undefined {
  return Array.isArray(value) ? value.length : undefined;
}
