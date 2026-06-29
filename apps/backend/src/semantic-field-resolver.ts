import type { AgentRunTarget } from "@components-kit/open-workbook-protocol";
import { normalizeHeaderName, type ColumnMetadata, type TableMetadata, type WorkbookMetadata } from "./workbook-metadata-cache.js";

export interface SemanticFieldCandidate {
  field: string;
  sheetName: string;
  tableName?: string;
  range: string;
  columnIndex: number;
  columnLetter: string;
  score: number;
  evidence: string[];
}

export interface SemanticFieldResolution {
  term: string;
  candidates: SemanticFieldCandidate[];
  ambiguous: boolean;
  best?: SemanticFieldCandidate;
}

interface ColumnSource {
  sheetName: string;
  tableName?: string;
  range: string;
  column: ColumnMetadata;
}

const FIELD_SYNONYMS: Record<string, string[]> = {
  status: ["status", "state", "stage", "paid", "unpaid", "open", "closed", "paymentstatus", "paymentstate", "invoicestatus"],
  category: ["category", "type", "class", "classification", "group"],
  amount: ["amount", "total", "price", "cost", "value", "revenue", "payment", "cash"],
  date: ["date", "day", "month", "period", "transactiondate", "invoicedate"],
  customer: ["customer", "client", "payer", "payee", "account", "vendor", "supplier", "company"],
  identifier: ["id", "identifier", "code", "number", "invoice", "ref", "reference"],
  note: ["note", "notes", "memo", "comment", "remarks", "description", "detail"]
};

export function resolveSemanticField(metadata: WorkbookMetadata, term: string, target?: AgentRunTarget): SemanticFieldResolution {
  const normalizedTerm = normalizeHeaderName(term);
  const sources = columnSources(metadata, target);
  const candidates = sources
    .map((source) => scoreColumnSource(source, normalizedTerm, term))
    .filter((candidate): candidate is SemanticFieldCandidate => Boolean(candidate))
    .sort((a, b) => b.score - a.score || a.field.localeCompare(b.field))
    .slice(0, 8);
  const [best, second] = candidates;
  return {
    term,
    candidates,
    ambiguous: Boolean(best && second && best.score - second.score < 0.08),
    ...(best ? { best } : {})
  };
}

export function resolveSemanticFields(metadata: WorkbookMetadata, terms: string[], target?: AgentRunTarget): SemanticFieldResolution[] {
  return terms.map((term) => resolveSemanticField(metadata, term, target));
}

function columnSources(metadata: WorkbookMetadata, target?: AgentRunTarget): ColumnSource[] {
  const targetSheetName = target?.sheetName;
  const targetTableName = target?.tableName;
  const tableSources = metadata.tables
    .filter((table) => !targetSheetName || table.sheetName === targetSheetName)
    .filter((table) => !targetTableName || table.name === targetTableName)
    .flatMap((table) => table.columns.map((column) => sourceFromTable(table, column)));
  const headerSources: ColumnSource[] = metadata.sheets
    .filter((sheet) => !targetSheetName || sheet.name === targetSheetName)
    .flatMap((sheet) => sheet.headers.flatMap((header) => header.columns.map((column) => ({
      sheetName: sheet.name,
      range: header.range,
      column
    } satisfies ColumnSource))));
  const seen = new Set<string>();
  const sources: ColumnSource[] = [];
  for (const source of [...tableSources, ...headerSources]) {
    const key = `${source.sheetName}:${source.tableName ?? ""}:${source.column.index}:${source.column.name}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    sources.push(source);
  }
  return sources;
}

function sourceFromTable(table: TableMetadata, column: ColumnMetadata): ColumnSource {
  return {
    sheetName: table.sheetName,
    ...(table.name ? { tableName: table.name } : {}),
    range: table.range,
    column
  };
}

function scoreColumnSource(source: ColumnSource, normalizedTerm: string, rawTerm: string): SemanticFieldCandidate | undefined {
  const normalizedName = source.column.normalizedName || normalizeHeaderName(source.column.name);
  const evidence: string[] = [];
  let score = 0;
  if (normalizedName === normalizedTerm) {
    score += 0.72;
    evidence.push("exact normalized header match");
  } else if (normalizedName.includes(normalizedTerm) || normalizedTerm.includes(normalizedName)) {
    score += 0.5;
    evidence.push("partial header match");
  } else {
    const tokenScore = tokenOverlapScore(normalizedTerm, normalizedName);
    if (tokenScore > 0) {
      score += 0.28 * tokenScore;
      evidence.push("header token overlap");
    }
  }
  const synonymRole = synonymRoleForTerm(normalizedTerm);
  if (synonymRole && (source.column.role === synonymRole || source.column.inferredType === synonymRole)) {
    score += 0.24;
    evidence.push(`${synonymRole} semantic role match`);
  }
  if (synonymRole && FIELD_SYNONYMS[synonymRole]?.includes(normalizedName)) {
    score += 0.18;
    evidence.push(`${synonymRole} synonym match`);
  }
  if (source.tableName) {
    score += 0.03;
    evidence.push("table column");
  }
  if (source.column.importance !== undefined) {
    score += Math.min(0.04, Math.max(0, source.column.importance) * 0.04);
  }
  if (score < 0.18) {
    return undefined;
  }
  return {
    field: source.column.name,
    sheetName: source.sheetName,
    ...(source.tableName ? { tableName: source.tableName } : {}),
    range: source.range,
    columnIndex: source.column.index,
    columnLetter: source.column.letter,
    score: Math.min(1, Number(score.toFixed(3))),
    evidence: evidence.length > 0 ? evidence : [`matched "${rawTerm}" to ${source.column.name}`]
  };
}

function synonymRoleForTerm(normalizedTerm: string): string | undefined {
  for (const [role, terms] of Object.entries(FIELD_SYNONYMS)) {
    if (terms.some((term) => normalizedTerm.includes(term) || term.includes(normalizedTerm))) {
      return role;
    }
  }
  return undefined;
}

function tokenOverlapScore(left: string, right: string): number {
  const leftTokens = tokens(left);
  const rightTokens = tokens(right);
  if (leftTokens.length === 0 || rightTokens.length === 0) {
    return 0;
  }
  const rightSet = new Set(rightTokens);
  const overlap = leftTokens.filter((token) => rightSet.has(token)).length;
  return overlap / Math.max(leftTokens.length, rightTokens.length);
}

function tokens(value: string): string[] {
  return value.split(/[^a-z0-9]+/i).map((token) => token.trim()).filter(Boolean);
}
