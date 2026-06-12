import type {
  FormulaDependencyEdge,
  FormulaDependencyGraph,
  FormulaDependencyNode,
  FormulaPatternResponse,
  FormulaSpillRange,
  OperationWarning,
  TableInfo,
  WorkbookId
} from "@open-workbook/protocol";
import { columnNameToNumber, formatA1Address, parseA1Address } from "./range-address.js";

const REFERENCE_RE = /(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?(?<address>\$?[A-Z]{1,3}\$?\d+#|\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})/g;
const EXTERNAL_RE = /'?\[(?<workbook>[^\]]+)](?<sheet>[^'!]+)'?!?(?<address>\$?[A-Z]{1,3}\$?\d+#|\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})/g;

export interface FormulaDependencyOptions {
  tables?: TableInfo[];
  spillRanges?: FormulaSpillRange[];
}

export function buildFormulaDependencyGraph(patterns: FormulaPatternResponse, options: FormulaDependencyOptions = {}): FormulaDependencyGraph {
  const nodes = new Map<string, FormulaDependencyNode>();
  const edgeKeys = new Set<string>();
  const edges: FormulaDependencyEdge[] = [];
  const warnings: OperationWarning[] = [...patterns.warnings];
  const base = parseA1Address(patterns.address);
  const dependencyOptions: FormulaDependencyOptions = {
    ...options,
    spillRanges: [...(patterns.spillRanges ?? []), ...(options.spillRanges ?? [])]
  };

  for (const cell of patterns.cells) {
    const formulaAddress = formatA1Address({
      startRow: base.startRow + cell.rowIndex,
      endRow: base.startRow + cell.rowIndex,
      startColumn: base.startColumn + cell.columnIndex,
      endColumn: base.startColumn + cell.columnIndex
    });
    const formulaNode = addNode(nodes, {
      workbookId: patterns.workbookId,
      kind: "range",
      sheetName: patterns.sheetName,
      address: formulaAddress,
      formula: cell.formula
    });
    for (const reference of extractFormulaReferences(patterns.workbookId, patterns.sheetName, cell.formula, dependencyOptions)) {
      const referencedNode = addNode(nodes, reference);
      const edgeKey = `${referencedNode.id}->${formulaNode.id}`;
      if (edgeKeys.has(edgeKey)) {
        continue;
      }
      edgeKeys.add(edgeKey);
      edges.push({
        from: referencedNode,
        to: formulaNode,
        kind: "precedent",
        confidence: "parsed"
      });
    }
  }

  return {
    workbookId: patterns.workbookId,
    sheetName: patterns.sheetName,
    address: patterns.address,
    capturedAt: new Date().toISOString(),
    nodes: [...nodes.values()],
    edges,
    warnings: [
      ...warnings,
      ...graphWarnings([...nodes.values()]),
      ...spillWarnings(patterns, dependencyOptions)
    ]
  };
}

export function extractFormulaReferences(workbookId: WorkbookId, defaultSheetName: string, formula: string, options: FormulaDependencyOptions = {}): FormulaDependencyNode[] {
  const sanitized = stripStringLiterals(formula);
  const references: FormulaDependencyNode[] = [];
  const seen = new Set<string>();
  const structuredReferences = extractStructuredReferences(sanitized);
  for (const structured of structuredReferences) {
    pushUnique(references, seen, {
      id: tableNodeId(workbookId, structured.tableName, structured.reference),
      workbookId,
      kind: "table",
      tableName: structured.tableName,
      structuredReference: structured.reference
    });
    const resolved = resolveStructuredReference(workbookId, structured.tableName, structured.reference, options.tables ?? []);
    if (resolved) {
      pushUnique(references, seen, resolved);
    }
  }
  for (const match of sanitized.matchAll(EXTERNAL_RE)) {
    const groups = match.groups;
    if (!groups?.workbook || !groups.address) {
      continue;
    }
    const sheetName = unquoteSheetName(groups.sheet) ?? "";
    const address = normalizeReferenceAddress(groups.address);
    pushUnique(references, seen, {
      id: externalNodeId(workbookId, groups.workbook, sheetName, address),
      workbookId,
      kind: "external",
      externalWorkbook: groups.workbook,
      externalReference: sheetName ? `${sheetName}!${address}` : address
    });
  }
  const withoutStructuredOrExternal = blankSpans(sanitized.replace(EXTERNAL_RE, " "), structuredReferences);
  for (const match of withoutStructuredOrExternal.matchAll(REFERENCE_RE)) {
    const groups = match.groups;
    if (!groups?.address) {
      continue;
    }
    const sheetName = unquoteSheetName(groups.sheet) ?? defaultSheetName;
    const address = normalizeReferenceAddress(groups.address, {
      workbookId,
      sheetName,
      spillRanges: options.spillRanges ?? []
    });
    pushUnique(references, seen, {
      id: rangeNodeId(workbookId, sheetName, address),
      workbookId,
      kind: "range",
      sheetName,
      address
    });
  }
  return references;
}

export function tracePrecedents(graph: FormulaDependencyGraph, sheetName: string, address: string): { nodes: FormulaDependencyNode[]; edges: FormulaDependencyEdge[] } {
  const target = rangeNodeId(graph.workbookId, sheetName, normalizeReferenceAddress(address));
  const edges = graph.edges.filter((edge) => edge.to.id === target);
  return {
    edges,
    nodes: uniqueNodes(edges.map((edge) => edge.from))
  };
}

export function traceDependents(graph: FormulaDependencyGraph, sheetName: string, address: string): { nodes: FormulaDependencyNode[]; edges: FormulaDependencyEdge[] } {
  const targetAddress = normalizeReferenceAddress(address);
  const edges = graph.edges.filter((edge) => edge.from.kind === "range" && edge.from.sheetName === sheetName && edge.from.address !== undefined && rangesIntersect(edge.from.address, targetAddress));
  return {
    edges,
    nodes: uniqueNodes(edges.map((edge) => edge.to))
  };
}

function addNode(nodes: Map<string, FormulaDependencyNode>, input: Omit<FormulaDependencyNode, "id"> & { id?: string }): FormulaDependencyNode {
  const id = input.id ?? nodeId(input);
  const existing = nodes.get(id);
  if (existing) {
    if (existing.formula === undefined && input.formula !== undefined) {
      existing.formula = input.formula;
    }
    return existing;
  }
  const node: FormulaDependencyNode = {
    id,
    workbookId: input.workbookId,
    kind: input.kind
  };
  if (input.sheetName !== undefined) {
    node.sheetName = input.sheetName;
  }
  if (input.address !== undefined) {
    node.address = input.address;
  }
  if (input.tableName !== undefined) {
    node.tableName = input.tableName;
  }
  if (input.structuredReference !== undefined) {
    node.structuredReference = input.structuredReference;
  }
  if (input.externalWorkbook !== undefined) {
    node.externalWorkbook = input.externalWorkbook;
  }
  if (input.externalReference !== undefined) {
    node.externalReference = input.externalReference;
  }
  if (input.formula !== undefined) {
    node.formula = input.formula;
  }
  nodes.set(id, node);
  return node;
}

function nodeId(node: Omit<FormulaDependencyNode, "id">): string {
  if (node.kind === "table") {
    return tableNodeId(node.workbookId, node.tableName ?? "", node.structuredReference ?? "");
  }
  if (node.kind === "external") {
    return externalNodeId(node.workbookId, node.externalWorkbook ?? "", "", node.externalReference ?? "");
  }
  return rangeNodeId(node.workbookId, node.sheetName ?? "", node.address ?? "");
}

function rangeNodeId(workbookId: WorkbookId, sheetName: string, address: string): string {
  return `${workbookId}:range:${sheetName}!${normalizeReferenceAddress(address)}`;
}

function tableNodeId(workbookId: WorkbookId, tableName: string, reference: string): string {
  return `${workbookId}:table:${tableName}[${reference}]`;
}

function externalNodeId(workbookId: WorkbookId, externalWorkbook: string, sheetName: string, reference: string): string {
  return `${workbookId}:external:[${externalWorkbook}]${sheetName ? `${sheetName}!` : ""}${reference}`;
}

function normalizeReferenceAddress(
  address: string,
  spillLookup?: { workbookId: WorkbookId; sheetName: string; spillRanges: FormulaSpillRange[] }
): string {
  const withoutDollar = address.replace(/\$/g, "");
  if (withoutDollar.endsWith("#")) {
    const anchorAddress = formatA1Address(parseA1Address(withoutDollar.slice(0, -1)));
    const spill = findSpillRange(spillLookup, anchorAddress);
    if (spill) {
      return normalizeReferenceAddress(spill.spillAddress);
    }
    return anchorAddress;
  }
  const normalized = withoutDollar;
  if (/^[A-Z]{1,3}:[A-Z]{1,3}$/i.test(normalized)) {
    const [startColumn, endColumn] = normalized.split(":");
    return `${startColumn!.toUpperCase()}:${endColumn!.toUpperCase()}`;
  }
  return formatA1Address(parseA1Address(normalized));
}

function findSpillRange(
  lookup: { workbookId: WorkbookId; sheetName: string; spillRanges: FormulaSpillRange[] } | undefined,
  anchorAddress: string
): FormulaSpillRange | undefined {
  if (!lookup) {
    return undefined;
  }
  return lookup.spillRanges.find((spill) => {
    const spillSheetName = spill.sheetName ?? lookup.sheetName;
    return spillSheetName === lookup.sheetName && normalizeReferenceAddress(spill.anchorAddress) === anchorAddress;
  });
}

function stripStringLiterals(formula: string): string {
  return formula.replace(/"(?:""|[^"])*"/g, "\"\"");
}

function unquoteSheetName(sheetName: string | undefined): string | undefined {
  if (sheetName === undefined) {
    return undefined;
  }
  return sheetName.startsWith("'") && sheetName.endsWith("'") ? sheetName.slice(1, -1).replace(/''/g, "'") : sheetName;
}

function rangesIntersect(left: string, right: string): boolean {
  try {
    if (isColumnRange(left) || isColumnRange(right)) {
      return columnRangesIntersect(left, right);
    }
    const parsedLeft = parseA1Address(left);
    const parsedRight = parseA1Address(right);
    return (
      parsedLeft.startRow <= parsedRight.endRow &&
      parsedLeft.endRow >= parsedRight.startRow &&
      parsedLeft.startColumn <= parsedRight.endColumn &&
      parsedLeft.endColumn >= parsedRight.startColumn
    );
  } catch {
    return left === right;
  }
}

interface StructuredReferenceMatch {
  tableName: string;
  reference: string;
  start: number;
  end: number;
}

function extractStructuredReferences(formula: string): StructuredReferenceMatch[] {
  const matches: StructuredReferenceMatch[] = [];
  const candidateRe = /(?<![A-Za-z0-9_])(?<table>[A-Za-z_][A-Za-z0-9_]*)\[/g;
  for (const match of formula.matchAll(candidateRe)) {
    const tableName = match.groups?.table;
    if (!tableName || match.index === undefined) {
      continue;
    }
    const bracketStart = match.index + tableName.length;
    const end = findBalancedBracketEnd(formula, bracketStart);
    if (end === undefined) {
      continue;
    }
    matches.push({
      tableName,
      reference: formula.slice(bracketStart + 1, end),
      start: match.index,
      end: end + 1
    });
  }
  return matches;
}

function findBalancedBracketEnd(input: string, start: number): number | undefined {
  let depth = 0;
  for (let index = start; index < input.length; index += 1) {
    const char = input[index];
    if (char === "[") {
      depth += 1;
    } else if (char === "]") {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
    }
  }
  return undefined;
}

function blankSpans(input: string, spans: Array<{ start: number; end: number }>): string {
  if (spans.length === 0) {
    return input;
  }
  const chars = [...input];
  for (const span of spans) {
    for (let index = span.start; index < span.end; index += 1) {
      chars[index] = " ";
    }
  }
  return chars.join("");
}

function resolveStructuredReference(workbookId: WorkbookId, tableName: string, reference: string, tables: TableInfo[]): FormulaDependencyNode | undefined {
  const table = tables.find((candidate) => candidate.tableName.toLowerCase() === tableName.toLowerCase());
  if (!table?.sheetName || !table.address) {
    return undefined;
  }
  const parsed = parseA1Address(table.address);
  const parsedReference = parseStructuredReference(reference);
  const rowBounds = structuredRowBounds(parsed.startRow, parsed.endRow, table.showHeaders !== false, table.showTotals === true, parsedReference.qualifiers);
  if (!rowBounds) {
    return undefined;
  }
  const columnBounds = structuredColumnBounds(parsed.startColumn, parsed.endColumn, table.columns, parsedReference.columns);
  const address = formatA1Address({
    startRow: rowBounds.startRow,
    endRow: rowBounds.endRow,
    startColumn: columnBounds.startColumn,
    endColumn: columnBounds.endColumn
  });
  return {
    id: rangeNodeId(workbookId, table.sheetName, address),
    workbookId,
    kind: "range",
    sheetName: table.sheetName,
    address
  };
}

function parseStructuredReference(reference: string): { qualifiers: Set<string>; columns: string[] } {
  const tokens = structuredReferenceTokens(reference);
  const qualifiers = new Set<string>();
  const columns: string[] = [];
  for (const token of tokens) {
    const normalized = normalizeStructuredToken(token);
    if (!normalized) {
      continue;
    }
    if (normalized.startsWith("#")) {
      qualifiers.add(normalized.toLowerCase());
    } else {
      columns.push(normalized);
    }
  }
  return { qualifiers, columns };
}

function structuredReferenceTokens(reference: string): string[] {
  const tokens: string[] = [];
  const trimmed = reference.trim();
  const bracketed = [...trimmed.matchAll(/\[([^\]]+)]/g)].map((match) => match[1]).filter((token): token is string => token !== undefined);
  if (bracketed.length > 0) {
    for (const token of bracketed) {
      for (const part of token.split(/[:,]/g)) {
        tokens.push(part);
      }
    }
    return tokens;
  }
  return trimmed.split(/[:,]/g);
}

function normalizeStructuredToken(token: string): string | undefined {
  const normalized = token.trim().replace(/^@+/, "").replace(/^'+/, "").replace(/''/g, "'");
  return normalized.length > 0 ? normalized : undefined;
}

function structuredRowBounds(
  tableStartRow: number,
  tableEndRow: number,
  showHeaders: boolean,
  showTotals: boolean,
  qualifiers: Set<string>
): { startRow: number; endRow: number } | undefined {
  const dataStartRow = showHeaders ? tableStartRow + 1 : tableStartRow;
  const dataEndRow = showTotals ? tableEndRow - 1 : tableEndRow;
  if (qualifiers.has("#all")) {
    return { startRow: tableStartRow, endRow: tableEndRow };
  }
  if (qualifiers.has("#headers")) {
    return showHeaders ? { startRow: tableStartRow, endRow: tableStartRow } : undefined;
  }
  if (qualifiers.has("#totals")) {
    return showTotals ? { startRow: tableEndRow, endRow: tableEndRow } : undefined;
  }
  if (dataEndRow < dataStartRow) {
    return undefined;
  }
  return { startRow: dataStartRow, endRow: dataEndRow };
}

function structuredColumnBounds(
  tableStartColumn: number,
  tableEndColumn: number,
  columns: TableInfo["columns"],
  referencedColumns: string[]
): { startColumn: number; endColumn: number } {
  const indexes = referencedColumns
    .map((columnName) => columns.find((candidate) => candidate.name.toLowerCase() === columnName.toLowerCase())?.index)
    .filter((index): index is number => index !== undefined);
  if (indexes.length === 0) {
    return { startColumn: tableStartColumn, endColumn: tableEndColumn };
  }
  const startColumn = tableStartColumn + Math.min(...indexes);
  const endColumn = tableStartColumn + Math.max(...indexes);
  return { startColumn, endColumn };
}

function isColumnRange(address: string): boolean {
  return /^[A-Z]{1,3}:[A-Z]{1,3}$/i.test(address);
}

function columnRangesIntersect(left: string, right: string): boolean {
  const leftParsed = columnRangeBounds(left);
  const rightParsed = columnRangeBounds(right);
  return leftParsed.startColumn <= rightParsed.endColumn && leftParsed.endColumn >= rightParsed.startColumn;
}

function columnRangeBounds(address: string): { startColumn: number; endColumn: number } {
  if (isColumnRange(address)) {
    const [start, end] = address.split(":");
    return {
      startColumn: columnNameToNumber(start!),
      endColumn: columnNameToNumber(end!)
    };
  }
  const parsed = parseA1Address(address);
  return {
    startColumn: parsed.startColumn,
    endColumn: parsed.endColumn
  };
}

function uniqueNodes(nodes: FormulaDependencyNode[]): FormulaDependencyNode[] {
  const seen = new Set<string>();
  const unique: FormulaDependencyNode[] = [];
  for (const node of nodes) {
    if (!seen.has(node.id)) {
      seen.add(node.id);
      unique.push(node);
    }
  }
  return unique;
}

function pushUnique(nodes: FormulaDependencyNode[], seen: Set<string>, node: FormulaDependencyNode): void {
  if (seen.has(node.id)) {
    return;
  }
  seen.add(node.id);
  nodes.push(node);
}

function graphWarnings(nodes: FormulaDependencyNode[]): OperationWarning[] {
  const warnings: OperationWarning[] = [];
  if (nodes.some((node) => node.kind === "table")) {
    warnings.push({
      code: "FORMULA_STRUCTURED_REFERENCES_PARSED",
      message: "Structured table references were parsed as table dependency nodes and resolved to table ranges when metadata was available."
    });
  }
  if (nodes.some((node) => node.kind === "external")) {
    warnings.push({
      code: "FORMULA_EXTERNAL_REFERENCES_PARSED",
      message: "External workbook references were parsed as external dependency nodes and cannot be resolved to local workbook ranges."
    });
  }
  return warnings;
}

function spillWarnings(patterns: FormulaPatternResponse, options: FormulaDependencyOptions): OperationWarning[] {
  const warnings: OperationWarning[] = [];
  for (const cell of patterns.cells) {
    for (const match of stripStringLiterals(cell.formula).matchAll(REFERENCE_RE)) {
      const groups = match.groups;
      if (!groups?.address?.endsWith("#")) {
        continue;
      }
      const sheetName = unquoteSheetName(groups.sheet) ?? patterns.sheetName;
      const anchorAddress = normalizeReferenceAddress(groups.address.slice(0, -1));
      const spill = findSpillRange({ workbookId: patterns.workbookId, sheetName, spillRanges: options.spillRanges ?? [] }, anchorAddress);
      if (!spill) {
        warnings.push({
          code: "FORMULA_SPILL_RANGE_UNRESOLVED",
          message: `Dynamic array spill ${sheetName}!${anchorAddress}# was treated as the anchor cell because spill metadata was not available.`,
          details: { sheetName, anchorAddress }
        });
      }
    }
  }
  return warnings;
}
