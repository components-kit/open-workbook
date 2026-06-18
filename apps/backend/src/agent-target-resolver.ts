import type { AgentCandidate, AgentRunInput, AgentRunOutput } from "@components-kit/open-workbook-protocol";
import { normalizeHeaderName, type WorkbookMetadata } from "./workbook-metadata-cache.js";

const FILLER_TERMS = new Set([
  "a",
  "an",
  "and",
  "about",
  "analyze",
  "analysis",
  "called",
  "file",
  "for",
  "from",
  "in",
  "me",
  "of",
  "on",
  "please",
  "read",
  "sheet",
  "show",
  "tell",
  "the",
  "this",
  "to",
  "what",
  "workbook",
  "xlsx"
]);

const MONTH_ALIASES = new Map([
  ["january", "jan"],
  ["february", "feb"],
  ["march", "mar"],
  ["april", "apr"],
  ["may", "may"],
  ["june", "jun"],
  ["july", "jul"],
  ["august", "aug"],
  ["september", "sep"],
  ["sept", "sep"],
  ["october", "oct"],
  ["november", "nov"],
  ["december", "dec"]
]);

interface CandidateSource {
  kind: AgentCandidate["kind"];
  id: string;
  label: string;
  sheetName?: string | undefined;
  tableName?: string | undefined;
  range?: string | undefined;
  searchValues: string[];
  baseScore?: number | undefined;
}

const A1_RANGE_PATTERN = /\b\$?[A-Z]{1,3}\$?\d+(?:\s*:\s*\$?[A-Z]{1,3}\$?\d+)?\b/i;

export type AgentTargetResolution = {
  ok: true;
  candidate: AgentCandidate;
  sheetName: string;
  range: string;
} | {
  ok: false;
  status: AgentRunOutput["status"];
  summary: string;
  candidates?: AgentCandidate[];
  nextAction: AgentRunOutput["nextAction"];
  warnings: string[];
};

export function findAgentCandidates(metadata: WorkbookMetadata, input: AgentRunInput): AgentCandidate[] {
  const query = queryText(input);
  return candidateSources(metadata)
    .map((source) => toCandidate(source, scoreSource(query, source)))
    .filter((candidate) => candidate.confidence > 0)
    .sort((left, right) => right.confidence - left.confidence || rankKind(left.kind) - rankKind(right.kind));
}

export function resolveAgentReadTarget(metadata: WorkbookMetadata, input: AgentRunInput): AgentTargetResolution {
  return resolveAgentTarget(metadata, input, {
    requireRange: true,
    ambiguousSummary: "Multiple workbook targets match the request. Choose one candidate before OpenWorkbook reads data.",
    notFoundSummary: "No matching workbook target was found. Provide a more specific sheet, table, header, named range, or region."
  });
}

export function resolveAgentUpdateTarget(metadata: WorkbookMetadata, input: AgentRunInput): AgentTargetResolution {
  if (input.target?.sheetName && input.target.range) {
    const exact = metadata.sheets.find((sheet) => normalizeNatural(sheet.name) === normalizeNatural(input.target!.sheetName!));
    if (exact) {
      return {
        ok: true,
        candidate: toCandidate({
          kind: "sheet",
          id: exact.id,
          label: exact.name,
          sheetName: exact.name,
          range: input.target.range,
          searchValues: [exact.name]
        }, 1),
        sheetName: exact.name,
        range: input.target.range
      };
    }
    const canonical = bestExplicitSheetCandidate(metadata, input.target.sheetName);
    if (canonical?.sheetName) {
      return {
        ok: true,
        candidate: { ...canonical, range: input.target.range },
        sheetName: canonical.sheetName,
        range: input.target.range
      };
    }
    const sheetOnlyInput: AgentRunInput = {
      request: input.target.sheetName,
      target: { sheetName: input.target.sheetName }
    };
    const sheetResolution = resolveAgentTarget(metadata, sheetOnlyInput, {
      requireRange: false,
      kindFilter: new Set(["sheet"]),
      ambiguousSummary: "Multiple sheets match the requested update target. Choose one sheet before OpenWorkbook previews the update.",
      notFoundSummary: "The requested sheet was not found. Provide an exact sheet name or choose a candidate."
    });
    if (!sheetResolution.ok) {
      const canonical = canonicalSheetFromExplicitTarget(sheetResolution.candidates);
      if (canonical?.sheetName) {
        return {
          ok: true,
          candidate: { ...canonical, range: input.target.range },
          sheetName: canonical.sheetName,
          range: input.target.range
        };
      }
      return sheetResolution;
    }
    return {
      ...sheetResolution,
      sheetName: sheetResolution.sheetName,
      range: input.target.range,
      candidate: { ...sheetResolution.candidate, range: input.target.range }
    };
  }
  return resolveAgentTarget(metadata, input, {
    requireRange: true,
    ambiguousSummary: "Preview needs one resolvable sheet/range target. Choose one candidate before OpenWorkbook previews the update.",
    notFoundSummary: "Preview needs a resolvable sheet/range target and values. Provide a clearer target or choose one candidate."
  });
}

function bestExplicitSheetCandidate(metadata: WorkbookMetadata, sheetName: string): AgentCandidate | undefined {
  const candidates = candidateSources(metadata)
    .filter((source) => source.kind === "sheet")
    .map((source) => toCandidate(source, scoreSource(sheetName, source)))
    .filter((candidate) => candidate.confidence > 0)
    .sort((left, right) => right.confidence - left.confidence);
  const best = candidates[0];
  const second = candidates[1];
  if (!best?.sheetName) {
    return undefined;
  }
  const gap = best.confidence - (second?.confidence ?? 0);
  return best.confidence >= 0.6 && gap >= 0.01 ? best : undefined;
}

function canonicalSheetFromExplicitTarget(candidates: AgentCandidate[] | undefined): AgentCandidate | undefined {
  const best = candidates?.[0];
  const second = candidates?.[1];
  if (!best || best.kind !== "sheet" || !best.sheetName) {
    return undefined;
  }
  const gap = best.confidence - (second?.confidence ?? 0);
  return best.confidence >= 0.65 && gap >= 0.03 ? best : undefined;
}

function resolveAgentTarget(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]>; ambiguousSummary: string; notFoundSummary: string }
): AgentTargetResolution {
  const exact = resolveExactAgentTarget(metadata, input, options);
  if (exact) {
    return exact;
  }
  const candidates = findAgentCandidates(metadata, input)
    .filter((candidate) => !options.kindFilter || options.kindFilter.has(candidate.kind))
    .filter((candidate) => candidate.sheetName && (!options.requireRange || candidate.range))
    .slice(0, 10);
  const best = candidates[0];
  if (!best?.sheetName || (options.requireRange && !best.range)) {
    return {
      ok: false,
      status: candidates.length > 0 ? "AMBIGUOUS_TARGET" : "NEEDS_INPUT",
      summary: options.notFoundSummary,
      ...(candidates.length > 0 ? { candidates } : {}),
      nextAction: candidates.length > 0 ? "call_with_target" : "ask_user",
      warnings: []
    };
  }
  const second = candidates[1];
  if (isAmbiguous(best, second)) {
    return {
      ok: false,
      status: "AMBIGUOUS_TARGET",
      summary: options.ambiguousSummary,
      candidates: candidates.slice(0, 5),
      nextAction: "call_with_target",
      warnings: ["Target resolution was intentionally conservative because top candidates were close."]
    };
  }
  return { ok: true, candidate: best, sheetName: best.sheetName, range: best.range ?? "" };
}

function resolveExactAgentTarget(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]> }
): AgentTargetResolution | undefined {
  const sources = candidateSources(metadata);
  if (!input.target?.sheetName && !input.target?.candidateId && requestMentionsActiveSheet(input.request)) {
    const activeSheetName = metadata.workbook.activeSheet;
    const activeSheet = activeSheetName
      ? sources.find((source) => source.kind === "sheet" && normalizeNatural(source.label) === normalizeNatural(activeSheetName))
      : undefined;
    if (activeSheet) {
      return exactSourceResolution(activeSheet, options);
    }
    return {
      ok: false,
      status: "NEEDS_INPUT",
      summary: "The current workbook context does not include an active sheet. Provide a sheet name or choose a candidate.",
      candidates: findAgentCandidates(metadata, input).slice(0, 5),
      nextAction: "call_with_target",
      warnings: ["Active sheet metadata was unavailable."]
    };
  }
  const parsedReference = parseSheetRangeReference(input.request);
  if (parsedReference) {
    const sheetSource = sources.find((source) =>
      source.kind === "sheet" &&
      normalizeNatural(source.label) === normalizeNatural(parsedReference.sheetName)
    );
    if (sheetSource) {
      return exactSourceResolution({ ...sheetSource, range: parsedReference.range }, options);
    }
  }
  const sheetHeaderBlock = resolveExactSheetHeaderBlock(metadata, input.request, options);
  if (sheetHeaderBlock) {
    return sheetHeaderBlock;
  }

  const mentionedTable = resolveMentionedTable(metadata, input.request, options);
  if (mentionedTable) {
    return mentionedTable;
  }

  const mentionedNamedRegion = resolveMentionedNamedRegion(metadata, input.request, options);
  if (mentionedNamedRegion) {
    return mentionedNamedRegion;
  }

  const byCandidateId = input.target?.candidateId
    ? sources.find((source) => source.id === input.target?.candidateId)
    : undefined;
  if (byCandidateId) {
    return exactSourceResolution(byCandidateId, options);
  }
  if (input.target?.candidateId) {
    const { candidateId: _candidateId, ...targetWithoutCandidateId } = input.target;
    return {
      ok: false,
      status: "NOT_FOUND",
      summary: "The requested candidateId was not found in the prepared workbook context.",
      candidates: findAgentCandidates(metadata, { ...input, target: targetWithoutCandidateId }).slice(0, 5),
      nextAction: "call_with_target",
      warnings: ["Candidate ids are scoped to the current prepared workbook context."]
    };
  }

  const tableName = input.target?.tableName;
  if (tableName) {
    const normalizedTableName = normalizeNatural(tableName);
    const sheetName = input.target?.sheetName ? normalizeNatural(input.target.sheetName) : undefined;
    const tableSources = sources.filter((source) =>
      source.kind === "table" &&
      normalizeNatural(source.label) === normalizedTableName &&
      (!sheetName || normalizeNatural(source.sheetName ?? "") === sheetName)
    );
    if (tableSources.length === 1 && tableSources[0]) {
      return exactSourceResolution(tableSources[0], options);
    }
    if (tableSources.length > 1) {
      return {
        ok: false,
        status: "AMBIGUOUS_TARGET",
        summary: "Multiple tables match the requested table name. Choose one candidateId before OpenWorkbook reads data.",
        candidates: tableSources.map((source) => toCandidate(source, 1)).slice(0, 5),
        nextAction: "call_with_target",
        warnings: ["Exact table name matched multiple workbook tables."]
      };
    }
  }

  const sheetName = input.target?.sheetName;
  if (sheetName) {
    const normalizedSheetName = normalizeNatural(sheetName);
    const sheetSource = sources.find((source) =>
      source.kind === "sheet" &&
      normalizeNatural(source.label) === normalizedSheetName
    );
    if (sheetSource) {
      const requestedRange = input.target?.range ?? parseRangeForSheetRequest(input.request, sheetSource.label);
      return exactSourceResolution({ ...sheetSource, range: requestedRange ?? sheetSource.range }, options);
    }
  }

  const selectedTarget = resolveSelectedTarget(metadata, input.request, options);
  if (selectedTarget) {
    return selectedTarget;
  }

  return undefined;
}

function resolveSelectedTarget(
  metadata: WorkbookMetadata,
  request: string,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]> }
): AgentTargetResolution | undefined {
  if (!requestMentionsSelection(request)) {
    return undefined;
  }
  const selection = metadata.selection;
  if (!selection?.sheetName || !selection.address) {
    return {
      ok: false,
      status: "NEEDS_INPUT",
      summary: "The request refers to the selected cell or range, but the current Excel selection is unavailable.",
      nextAction: "ask_user",
      warnings: ["Select a cell/range in Excel, then retry the request."]
    };
  }
  const range = requestMentionsSelectedColumn(request) ? selectedColumnRange(metadata, selection) : selection.address;
  return exactSourceResolution({
    kind: "range",
    id: "selection:active",
    label: requestMentionsSelectedColumn(request) ? "selected column" : selection.isSingleCell ? "selected cell" : "selected range",
    sheetName: selection.sheetName,
    range,
    searchValues: ["selected", "selection", "active cell", "current cell", "selected range", "selected column"]
  }, options);
}

function requestMentionsSelection(request: string): boolean {
  return /\b(selection|highlighted|active cell|current cell|this cell|this range|this column|selected (?:cell|range|col(?:umn)?)|active column|current column)\b/i.test(request);
}

function requestMentionsSelectedColumn(request: string): boolean {
  return /\b(this|selected|active|current)\s+col(?:umn)?\b/i.test(request);
}

function selectedColumnRange(metadata: WorkbookMetadata, selection: NonNullable<WorkbookMetadata["selection"]>): string {
  if (!selection.isSingleCell) {
    return selection.address;
  }
  const sheet = metadata.sheets.find((candidate) => candidate.name === selection.sheetName);
  const endRow = sheet?.usedRange ? endRowFromRange(sheet.usedRange) : Math.max(1, selection.endCell.row);
  const column = numberToColumn(selection.startCell.column);
  return `${column}1:${column}${endRow}`;
}

function exactSourceResolution(
  source: CandidateSource,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]> }
): AgentTargetResolution | undefined {
  if (options.kindFilter && !options.kindFilter.has(source.kind)) {
    return undefined;
  }
  if (!source.sheetName || (options.requireRange && !source.range)) {
    return {
      ok: false,
      status: "NEEDS_INPUT",
      summary: "The selected workbook target does not include a readable range.",
      nextAction: "call_with_target",
      warnings: []
    };
  }
  const candidate = toCandidate(source, 1);
  return { ok: true, candidate, sheetName: source.sheetName, range: source.range ?? "" };
}

function candidateSources(metadata: WorkbookMetadata): CandidateSource[] {
  return [
    ...metadata.sheets.map((sheet) => ({
      kind: "sheet" as const,
      id: sheet.id,
      label: sheet.name,
      sheetName: sheet.name,
      range: sheet.usedRange,
      searchValues: [sheet.name, sheet.kind, ...sheet.headers.flatMap((header) => header.columns.map((column) => column.name))]
    })),
    ...metadata.tables.map((table) => ({
      kind: "table" as const,
      id: table.id,
      label: table.name ?? table.id,
      sheetName: table.sheetName,
      tableName: table.name,
      range: table.range,
      searchValues: [table.name ?? table.id, table.sheetName, ...table.columns.map((column) => column.name)],
      baseScore: 0.03
    })),
    ...metadata.tables.flatMap((table) => table.columns.map((column) => ({
      kind: "column" as const,
      id: `${table.id}:column:${column.normalizedName}`,
      label: column.name,
      sheetName: table.sheetName,
      tableName: table.name,
      range: table.range,
      searchValues: [column.name, column.normalizedName, table.name ?? table.id, table.sheetName],
      baseScore: 0.02
    }))),
    ...metadata.sheets.flatMap((sheet) => sheet.headers.flatMap((header) =>
      headerBlocks(sheet.name, sheet.usedRange, header).map((block, index) => ({
        kind: "range" as const,
        id: `${header.id}:block:${index}`,
        label: block.label,
        sheetName: sheet.name,
        range: block.range,
        searchValues: [sheet.name, block.label, ...block.searchValues],
        baseScore: 0.05
      }))
    )),
    ...metadata.namedRanges.map((name) => ({
      kind: "region" as const,
      id: `name:${name.name}`,
      label: name.name,
      sheetName: name.sheetName,
      range: name.range,
      searchValues: [name.name, name.sheetName ?? "", name.range],
      baseScore: 0.04
    })),
    ...metadata.sections.map((section) => ({
      kind: "region" as const,
      id: section.id,
      label: section.label,
      sheetName: section.sheetName,
      range: section.range,
      searchValues: [
        section.label,
        section.kind,
        ...section.columns.map((column) => column.name)
      ],
      baseScore: 0.06
    })),
    ...metadata.summaryBlocks.map((block) => ({
      kind: "range" as const,
      id: block.id,
      label: block.labels.join(", ") || "summary block",
      sheetName: block.sheetName,
      range: block.range,
      searchValues: [block.sheetName, "summary", "report", ...block.labels],
      baseScore: 0.04
    })),
    ...metadata.formulaRegions.map((region) => ({
      kind: "range" as const,
      id: region.id,
      label: "formula region",
      sheetName: region.sheetName,
      range: region.range,
      searchValues: [region.sheetName, "formula", "calculation"],
      baseScore: 0.01
    }))
  ];
}

function resolveMentionedNamedRegion(
  metadata: WorkbookMetadata,
  request: string,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]> }
): AgentTargetResolution | undefined {
  const requestedKinds = [
    /\binvoice|billed|booking|container|collect\b/i.test(request) ? "invoice" : undefined,
    /\bkpi|summary|metric|revenue|profit\b/i.test(request) ? "kpi" : undefined,
    /\bnotes?|status|owner|action\b/i.test(request) ? "status" : undefined,
    /\breconciliation|variance|formula\b/i.test(request) ? "reconciliation" : undefined
  ].filter((kind): kind is string => Boolean(kind));
  if (requestedKinds.length === 0) {
    return undefined;
  }
  const requestText = normalizeNatural(request);
  const matches = metadata.namedRanges.filter((name) => {
    const label = normalizeNatural(name.name.replace(/([a-z])([A-Z])/g, "$1 $2"));
    const sheetMatch = !name.sheetName || requestText.includes(normalizeNatural(name.sheetName));
    return sheetMatch && requestedKinds.some((kind) => label.includes(kind));
  });
  if (matches.length === 1 && matches[0]) {
    return exactSourceResolution({
      kind: "region",
      id: `name:${matches[0].name}`,
      label: matches[0].name,
      sheetName: matches[0].sheetName,
      range: matches[0].range,
      searchValues: [matches[0].name, matches[0].sheetName ?? "", matches[0].range]
    }, options);
  }
  return undefined;
}

function resolveMentionedTable(
  metadata: WorkbookMetadata,
  request: string,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]> }
): AgentTargetResolution | undefined {
  if (!/\btables?\b/i.test(request)) {
    return undefined;
  }
  const normalizedRequest = normalizeNatural(request);
  const matches = metadata.tables.filter((table) => {
    const name = table.name ?? table.id;
    return normalizeNatural(name) !== "" && normalizedRequest.includes(normalizeNatural(name));
  });
  if (matches.length === 1 && matches[0]) {
    return exactSourceResolution({
      kind: "table",
      id: matches[0].id,
      label: matches[0].name ?? matches[0].id,
      sheetName: matches[0].sheetName,
      tableName: matches[0].name,
      range: matches[0].range,
      searchValues: [matches[0].name ?? matches[0].id, matches[0].sheetName, ...matches[0].columns.map((column) => column.name)]
    }, options);
  }
  if (matches.length > 1) {
    return {
      ok: false,
      status: "AMBIGUOUS_TARGET",
      summary: "Multiple tables match the requested table mention. Choose one candidateId before OpenWorkbook reads data.",
      candidates: matches.map((table) => toCandidate({
        kind: "table",
        id: table.id,
        label: table.name ?? table.id,
        sheetName: table.sheetName,
        tableName: table.name,
        range: table.range,
        searchValues: []
      }, 1)).slice(0, 5),
      nextAction: "call_with_target",
      warnings: ["Table name text matched multiple workbook tables."]
    };
  }
  return undefined;
}

function toCandidate(source: CandidateSource, confidence: number): AgentCandidate {
  const boundedConfidence = Number(Math.min(1, Math.max(0, confidence)).toFixed(3));
  return {
    id: source.id,
    kind: source.kind,
    label: source.label,
    confidence: boundedConfidence,
    ...(source.sheetName !== undefined ? { sheetName: source.sheetName } : {}),
    ...(source.tableName !== undefined ? { tableName: source.tableName } : {}),
    ...(source.range !== undefined ? { range: source.range } : {}),
    reason: candidateReason(source, boundedConfidence),
    nextRequestHint: candidateNextRequestHint(source)
  };
}

function candidateReason(source: CandidateSource, confidence: number): string {
  const location = source.sheetName ? ` on sheet "${source.sheetName}"` : "";
  const scope = source.range ? ` at ${source.range}` : "";
  return `${source.kind} match "${source.label}"${location}${scope} scored ${confidence}.`;
}

function candidateNextRequestHint(source: CandidateSource): string {
  const target = [
    source.sheetName ? `sheetName: "${source.sheetName}"` : undefined,
    source.tableName ? `tableName: "${source.tableName}"` : undefined,
    source.range ? `range: "${source.range}"` : undefined
  ].filter(Boolean).join(", ");
  return target
    ? `Retry with target.candidateId "${source.id}" or target { ${target} }.`
    : `Retry with target.candidateId "${source.id}".`;
}

function scoreSource(query: string, source: CandidateSource): number {
  const queryTokens = meaningfulTokens(query);
  if (queryTokens.length === 0) {
    return 0.1 + (source.baseScore ?? 0);
  }
  let score = 0;
  for (const value of source.searchValues) {
    const candidateTokens = meaningfulTokens(value);
    if (candidateTokens.length === 0) continue;
    const candidateText = candidateTokens.join(" ");
    const queryText = queryTokens.join(" ");
    if (candidateText === queryText) score = Math.max(score, 1);
    if (queryText.includes(candidateText) || candidateText.includes(queryText)) score = Math.max(score, 0.86);
    const overlap = candidateTokens.filter((token) => queryTokens.includes(token)).length;
    if (overlap > 0) {
      const precision = overlap / candidateTokens.length;
      const recall = overlap / queryTokens.length;
      score = Math.max(score, 0.25 + precision * 0.45 + recall * 0.25);
    }
  }
  return Math.min(1, score + (source.baseScore ?? 0));
}

function parseSheetRangeReference(request: string): { sheetName: string; range: string } | undefined {
  const a1Pattern = /(\$?[A-Z]{1,3}\$?\d+(?:\s*:\s*\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}\s*:\s*\$?[A-Z]{1,3})/i;
  const quoted = new RegExp(`['"]([^'"]+)['"]!\\s*${a1Pattern.source}`, "i").exec(request);
  if (quoted?.[1] && quoted[2]) {
    return { sheetName: quoted[1], range: normalizeA1Range(quoted[2]) };
  }
  const sheetRange = new RegExp(`\\b(?:sheet\\s+)?([A-Za-z]{3,9}\\s+\\d{4})\\s*!?\\s+(?:range\\s+)?${a1Pattern.source}`, "i").exec(request);
  if (sheetRange?.[1] && sheetRange[2]) {
    return { sheetName: sheetRange[1], range: normalizeA1Range(sheetRange[2]) };
  }
  const columnsInSheet = /\bcolumns?\s+([A-Z]{1,3})\s+(?:to|through|-)\s+([A-Z]{1,3})\s+in\s+(?:the\s+)?([A-Za-z]{3,9}\s+\d{4})\s+sheet\b/i.exec(request);
  if (columnsInSheet?.[1] && columnsInSheet[2] && columnsInSheet[3]) {
    return { sheetName: columnsInSheet[3], range: `${columnsInSheet[1].toUpperCase()}:${columnsInSheet[2].toUpperCase()}` };
  }
  return undefined;
}

function parseRangeForSheetRequest(request: string, sheetName: string): string | undefined {
  const parsed = parseSheetRangeReference(request);
  if (parsed && normalizeNatural(parsed.sheetName) === normalizeNatural(sheetName)) {
    return parsed.range;
  }
  const explicit = /\brange\s+(\$?[A-Z]{1,3}\$?\d+(?:\s*:\s*\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}\s*:\s*\$?[A-Z]{1,3})/i.exec(request);
  if (explicit?.[1]) {
    return normalizeA1Range(explicit[1]);
  }
  return undefined;
}

function resolveExactSheetHeaderBlock(
  metadata: WorkbookMetadata,
  request: string,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]> }
): AgentTargetResolution | undefined {
  const matchedSheet = metadata.sheets.find((sheet) => requestMentionsSheet(request, sheet.name));
  if (!matchedSheet) {
    return undefined;
  }
  const requestedKind = /\binvoice|billed|booking|container|gross|withholding|collect\b/i.test(request)
    ? "invoice"
    : /\btransaction|cash|payment|truck|proof|tx\b/i.test(request)
      ? "transaction"
      : undefined;
  if (!requestedKind) {
    return undefined;
  }
  const section = isRawMonthlySheet(matchedSheet.name, matchedSheet.usedRange)
    ? undefined
    : metadata.sections.find((candidate) =>
      candidate.sheetName === matchedSheet.name &&
      sectionMatchesRequestedKind(candidate, requestedKind)
    );
  if (section) {
    return exactSourceResolution({
      kind: "region",
      id: section.id,
      label: section.label,
      sheetName: section.sheetName,
      range: section.range,
      searchValues: [section.label, section.kind, ...section.columns.map((column) => column.name)]
    }, options);
  }
  const block = matchedSheet.headers
    .flatMap((header) => headerBlocks(matchedSheet.name, matchedSheet.usedRange, header))
    .find((candidate) => headerBlockKind(candidate.searchValues) === requestedKind);
  const fallback = block ?? fallbackMonthlyBlock(matchedSheet.name, matchedSheet.usedRange, requestedKind);
  if (!fallback) {
    return undefined;
  }
  return exactSourceResolution({
    kind: "range",
    id: `header:${matchedSheet.name}:${requestedKind}`,
    label: fallback.label,
    sheetName: matchedSheet.name,
    range: fallback.range,
    searchValues: [matchedSheet.name, fallback.label, ...fallback.searchValues]
  }, options);
}

function sectionMatchesRequestedKind(
  section: WorkbookMetadata["sections"][number],
  requestedKind: "invoice" | "transaction"
): boolean {
  const text = normalizeHeaderName([
    section.label,
    section.kind,
    ...section.columns.map((column) => column.name)
  ].join(" "));
  if (requestedKind === "invoice") {
    return /invoice|billed|booking|container|gross|withholding|collect/.test(text);
  }
  return /transaction|cash|actual|payment|reconciliation|truck|proof/.test(text);
}

function isRawMonthlySheet(sheetName: string, usedRange: string | undefined): boolean {
  return Boolean(usedRange && /\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{4}\b/i.test(sheetName));
}

function requestMentionsSheet(request: string, sheetName: string): boolean {
  return normalizeNatural(request).includes(normalizeNatural(sheetName));
}

function requestMentionsActiveSheet(request: string): boolean {
  return /\b(active|current|this)\s+sheet\b/i.test(request) || /\bthis\s+sheet\b/i.test(request);
}

function normalizeA1Range(range: string): string {
  return range.replace(/\$/g, "").replace(/\s+/g, "").toUpperCase();
}

function headerBlocks(sheetName: string, usedRange: string | undefined, header: WorkbookMetadata["sheets"][number]["headers"][number]) {
  const blocks: Array<{ label: string; range: string; searchValues: string[] }> = [];
  const contiguous = contiguousColumnGroups(header.columns.map((column) => column.index));
  for (const group of contiguous) {
    const columns = header.columns.filter((column) => column.index >= group.start && column.index <= group.end);
    if (columns.length === 0) continue;
    const rowEnd = usedRange ? endRowFromRange(usedRange) : header.row;
    const range = `${columns[0]?.letter}${header.row}:${columns[columns.length - 1]?.letter}${rowEnd}`;
    const names = columns.map((column) => column.name);
    const label = `${sheetName} ${headerBlockKind(names)} header block`;
    blocks.push({ label, range, searchValues: [headerBlockKind(names), "header", "headers", "raw data", ...names] });
  }
  return blocks;
}

function contiguousColumnGroups(indexes: number[]): Array<{ start: number; end: number }> {
  const sorted = [...new Set(indexes)].sort((left, right) => left - right);
  const groups: Array<{ start: number; end: number }> = [];
  for (const index of sorted) {
    const current = groups[groups.length - 1];
    if (current && index === current.end + 1) {
      current.end = index;
    } else {
      groups.push({ start: index, end: index });
    }
  }
  return groups;
}

function headerBlockKind(names: string[]): string {
  const normalized = names.map(normalizeHeaderName).join(" ");
  if (/\binvoice|billed|booking|container|gross|withholding|collect\b/.test(normalized)) {
    return "invoice";
  }
  if (/\btransaction|cash|actual|payment|reconciliation|truck|proof\b/.test(normalized)) {
    return "transaction";
  }
  return "raw";
}

function fallbackMonthlyBlock(sheetName: string, usedRange: string | undefined, kind: "invoice" | "transaction") {
  if (!usedRange || !/\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{4}\b/i.test(sheetName)) {
    return undefined;
  }
  const endRow = endRowFromRange(usedRange);
  const endColumn = endColumnFromRange(usedRange);
  if (!endColumn) {
    return undefined;
  }
  if (kind === "invoice" && columnIndex(endColumn) >= columnIndex("O")) {
    const end = columnIndex(endColumn) >= columnIndex("AE") ? "AE" : endColumn;
    return { label: `${sheetName} invoice header block`, range: `O1:${end}${endRow}`, searchValues: ["invoice", "header", "headers", "raw data"] };
  }
  if (kind === "transaction") {
    const end = columnIndex(endColumn) >= columnIndex("M") ? "M" : endColumn;
    return { label: `${sheetName} transaction header block`, range: `A1:${end}${endRow}`, searchValues: ["transaction", "header", "headers", "raw data"] };
  }
  return undefined;
}

function endRowFromRange(range: string): number {
  const matches = [...range.matchAll(/\d+/g)].map((match) => Number(match[0])).filter(Number.isFinite);
  return Math.max(...matches, 1);
}

function endColumnFromRange(range: string): string | undefined {
  const cells = [...range.matchAll(/([A-Z]{1,3})\d+/gi)].map((match) => match[1]?.toUpperCase()).filter((value): value is string => Boolean(value));
  return cells[cells.length - 1];
}

function columnIndex(column: string): number {
  return column.toUpperCase().split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function numberToColumn(column: number): string {
  let value = Math.max(1, column);
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters;
}

function queryText(input: AgentRunInput): string {
  return [
    input.request,
    input.target?.candidateId,
    input.target?.entity,
    input.target?.sheetName,
    input.target?.tableName,
    input.target?.column
  ].filter((value): value is string => typeof value === "string" && value.trim() !== "").join(" ");
}

function meaningfulTokens(value: string): string[] {
  return normalizeNatural(value)
    .split(" ")
    .map((token) => MONTH_ALIASES.get(token) ?? token)
    .map((token) => token.endsWith("s") && token.length > 4 ? token.slice(0, -1) : token)
    .filter((token) => token !== "" && !FILLER_TERMS.has(token));
}

function normalizeNatural(value: string): string {
  return normalizeHeaderName(value)
    .replace(/_/g, " ")
    .replace(/\bq([1-4])\b/g, "quarter $1")
    .replace(/\s+/g, " ")
    .trim();
}

function isAmbiguous(best: AgentCandidate, second?: AgentCandidate): boolean {
  if (!second) return false;
  if (best.sheetName === second.sheetName && best.range === second.range) return false;
  if (best.confidence >= 0.86 && best.confidence - second.confidence >= 0.04) return false;
  if (best.confidence >= 0.78 && best.confidence - second.confidence >= 0.12) return false;
  return best.confidence < 0.72 || best.confidence - second.confidence < 0.12;
}

function rankKind(kind: AgentCandidate["kind"]): number {
  switch (kind) {
    case "sheet":
      return 0;
    case "table":
      return 1;
    case "region":
      return 2;
    case "range":
      return 3;
    case "column":
      return 4;
    default:
      return 5;
  }
}
