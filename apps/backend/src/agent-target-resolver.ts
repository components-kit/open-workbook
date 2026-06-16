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
  range?: string | undefined;
  searchValues: string[];
  baseScore?: number | undefined;
}

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

function resolveAgentTarget(
  metadata: WorkbookMetadata,
  input: AgentRunInput,
  options: { requireRange: boolean; kindFilter?: Set<AgentCandidate["kind"]>; ambiguousSummary: string; notFoundSummary: string }
): AgentTargetResolution {
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
      range: table.range,
      searchValues: [table.name ?? table.id, table.sheetName, ...table.columns.map((column) => column.name)],
      baseScore: 0.03
    })),
    ...metadata.tables.flatMap((table) => table.columns.map((column) => ({
      kind: "column" as const,
      id: `${table.id}:column:${column.normalizedName}`,
      label: column.name,
      sheetName: table.sheetName,
      range: table.range,
      searchValues: [column.name, column.normalizedName, table.name ?? table.id, table.sheetName],
      baseScore: 0.02
    }))),
    ...metadata.namedRanges.map((name) => ({
      kind: "region" as const,
      id: `name:${name.name}`,
      label: name.name,
      sheetName: name.sheetName,
      range: name.range,
      searchValues: [name.name, name.sheetName ?? "", name.range],
      baseScore: 0.04
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

function toCandidate(source: CandidateSource, confidence: number): AgentCandidate {
  return {
    id: source.id,
    kind: source.kind,
    label: source.label,
    confidence: Number(Math.min(1, Math.max(0, confidence)).toFixed(3)),
    ...(source.sheetName !== undefined ? { sheetName: source.sheetName } : {}),
    ...(source.range !== undefined ? { range: source.range } : {})
  };
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

function queryText(input: AgentRunInput): string {
  return [
    input.request,
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
