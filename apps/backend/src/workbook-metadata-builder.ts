import { makeId, type A1Range, type BatchRequest, type CellMatrix, type ExcelOperation, type NameInfo, type OperationId, type RuntimeSelectionResponse, type SelectionInfo, type WorkbookId, type WorkbookRef } from "@components-kit/open-workbook-protocol";
import { stripSheetName } from "@components-kit/open-workbook-excel-core";
import {
  checkMetadataFreshness,
  columnLetter,
  createMetadataFingerprint,
  DEFAULT_METADATA_CACHE_TTL_MS,
  normalizeHeaderName,
  type ColumnMetadata,
  type ColumnRole,
  type ColumnType,
  type FormulaRegionMetadata,
  type HeaderMetadata,
  type SectionKind,
  type SectionMetadata,
  type SheetKind,
  type SheetMetadata,
  type SummaryBlockMetadata,
  type TableMetadata,
  type WorkbookMetadata,
  WorkbookMetadataCache,
  workbookMetadataKey
} from "./workbook-metadata-cache.js";
import type { RuntimeService } from "./runtime-service.js";

export interface MetadataBuildResult {
  metadata: WorkbookMetadata;
  cacheHit: boolean;
  freshnessReason: string;
}

export class WorkbookMetadataBuilder {
  constructor(
    private readonly runtime: RuntimeService,
    private readonly cache: WorkbookMetadataCache
  ) {}

  async getOrBuild(input: { workbookContextId?: string; workbookId?: WorkbookId | string; workbookName?: string; includeSamples?: boolean; targetFreshnessRanges?: A1Range[] }): Promise<MetadataBuildResult> {
    const includeSamples = input.includeSamples === true;
    const existingByContext = input.workbookContextId ? this.cache.getByContextId(input.workbookContextId) : undefined;
    const reusableContextId = existingByContext?.workbookContextId;
    const activeContext = await this.runtime.getActiveContext();
    const activeWorkbook = (activeContext as { activeWorkbook?: WorkbookRef }).activeWorkbook ?? this.runtime.sessions.getActive()?.activeWorkbook;
    if (!activeWorkbook) {
      throw new Error("No active Excel workbook is available. Open Excel and connect the Open Workbook add-in.");
    }
    const selection = await this.readSelection(activeWorkbook.workbookId);
    const contentVersion = typeof (this.runtime as unknown as { getWorkbookContentVersion?: (workbookId: WorkbookId | string) => number }).getWorkbookContentVersion === "function"
      ? (this.runtime as unknown as { getWorkbookContentVersion: (workbookId: WorkbookId | string) => number }).getWorkbookContentVersion(activeWorkbook.workbookId)
      : 0;
    const mapResult = await this.runtime.getWorkbookMap();
    if ((mapResult as { ok?: boolean }).ok === false) {
      throw new Error("Workbook map is unavailable because the Excel add-in is disconnected.");
    }
    const map = (mapResult as { map?: { sheets?: any[]; activeSheet?: string } }).map ?? {};
    const sheets = Array.isArray(map.sheets) ? map.sheets : [];
    const fingerprint = createMetadataFingerprint({
      workbookId: activeWorkbook.workbookId,
      workbook: activeWorkbook,
      sheets
    });
    if (existingByContext) {
      const freshnessContentVersion = this.contentVersionForFreshness(existingByContext, activeWorkbook.workbookId, includeSamples, contentVersion, input.targetFreshnessRanges);
      const freshness = checkMetadataFreshness(existingByContext, fingerprint, { requireSampled: includeSamples, contentVersion: freshnessContentVersion });
      if (freshness.status === "FRESH") {
        return {
          metadata: this.cache.set(withFreshSelection(existingByContext, selection)),
          cacheHit: true,
          freshnessReason: freshnessContentVersion === existingByContext.contentVersion && contentVersion !== existingByContext.contentVersion
            ? "cached metadata is fresh for target; no overlapping changes since context"
            : "cached metadata is fresh"
        };
      }
      this.cache.delete(existingByContext.workbookKey);
    }

    const key = workbookMetadataKey({
      workbookId: activeWorkbook.workbookId,
      workbookName: activeWorkbook.name,
      ...(activeWorkbook.path !== undefined ? { workbookPath: activeWorkbook.path } : {})
    });
    const existing = this.cache.get(key);
    if (existing) {
      const freshnessContentVersion = this.contentVersionForFreshness(existing, activeWorkbook.workbookId, includeSamples, contentVersion, input.targetFreshnessRanges);
      const freshness = checkMetadataFreshness(existing, fingerprint, { requireSampled: includeSamples, contentVersion: freshnessContentVersion });
      if (freshness.status === "FRESH") {
        return {
          metadata: this.cache.set(withFreshSelection(existing, selection)),
          cacheHit: true,
          freshnessReason: freshnessContentVersion === existing.contentVersion && contentVersion !== existing.contentVersion
            ? "cached metadata is fresh for target; no overlapping changes since context"
            : "cached metadata is fresh"
        };
      }
    }

    const tables = await this.buildTableMetadata(activeWorkbook.workbookId, sheets);
    const namedRanges = await this.buildNamedRanges(activeWorkbook.workbookId);
    const registeredRegions = this.runtime.listRegions(activeWorkbook.workbookId).regions.map((region) => ({
      name: region.name,
      sheetName: region.sheetName,
      range: region.address
    }));
    const tableMap = groupBy(tables, (table) => table.sheetName);
    const samples = new Map<string, CellMatrix>();
    if (includeSamples) {
      for (const sheet of sheets) {
        const sample = await this.readSheetSample(activeWorkbook.workbookId, sheet);
        if (sample.length > 0) {
          samples.set(sheet.name, sample);
        }
      }
    }

    const sections = sheets.flatMap((sheet, index) => detectSections(sheet, samples.get(sheet.name) ?? [], index));
    const sectionIdsBySheet = groupBy(sections, (section) => section.sheetName);
    const summaryBlocks = sheets.flatMap((sheet, index) => detectSummaryBlocks(sheet, samples.get(sheet.name) ?? [], index));
    const formulaRegions = sheets.flatMap((sheet, index) => detectFormulaRegions(sheet, samples.get(sheet.name) ?? [], index));
    const summaryIdsBySheet = groupBy(summaryBlocks, (block) => block.sheetName);
    const formulaIdsBySheet = groupBy(formulaRegions, (region) => region.sheetName);
    const now = Date.now();
    const metadata: WorkbookMetadata = {
      workbookContextId: reusableContextId ?? makeId("wbctx"),
      workbookKey: key,
      detailLevel: includeSamples ? "sampled" : "structure",
      contentVersion,
      workbook: {
        workbookId: activeWorkbook.workbookId,
        name: activeWorkbook.name,
        sheetCount: sheets.length,
        ...(activeWorkbook.path !== undefined ? { path: activeWorkbook.path } : {}),
        ...(map.activeSheet !== undefined ? { activeSheet: map.activeSheet } : {})
      },
      ...(selection ? { selection } : {}),
      sheets: sheets.map((sheet, index) =>
        sheetMetadataFromMap(
          sheet,
          index,
          tableMap.get(sheet.name) ?? [],
          samples.get(sheet.name) ?? [],
          sectionIdsBySheet.get(sheet.name) ?? [],
          summaryIdsBySheet.get(sheet.name) ?? [],
          formulaIdsBySheet.get(sheet.name) ?? []
        )
      ),
      tables,
      namedRanges: [...namedRanges, ...registeredRegions],
      sections,
      summaryBlocks,
      formulaRegions,
      fingerprint,
      createdAt: now,
      updatedAt: now,
      expiresAt: now + DEFAULT_METADATA_CACHE_TTL_MS
    };
    return { metadata: this.cache.set(metadata), cacheHit: false, freshnessReason: includeSamples ? "built sampled metadata" : "built structure metadata" };
  }

  private contentVersionForFreshness(
    metadata: WorkbookMetadata,
    workbookId: WorkbookId | string,
    includeSamples: boolean,
    currentContentVersion: number,
    targetFreshnessRanges: A1Range[] | undefined
  ): number {
    if (!includeSamples || metadata.contentVersion === undefined || metadata.contentVersion === currentContentVersion || !targetFreshnessRanges?.length) {
      return currentContentVersion;
    }
    const journalReader = (this.runtime as unknown as {
      getWorkbookChangeJournal?: (input: { workbookId: WorkbookId | string; sinceVersion?: number; ranges?: A1Range[]; limit?: number }) => { ok?: boolean; overlapStatus?: string };
    }).getWorkbookChangeJournal;
    if (typeof journalReader !== "function") {
      return currentContentVersion;
    }
    const journal = journalReader.call(this.runtime, {
      workbookId,
      sinceVersion: metadata.contentVersion,
      ranges: targetFreshnessRanges,
      limit: 1
    });
    return journal?.ok === true && journal.overlapStatus === "no_overlap"
      ? metadata.contentVersion
      : currentContentVersion;
  }

  invalidateWorkbook(workbookId: WorkbookId | string): void {
    this.cache.deleteByWorkbookId(workbookId);
  }

  private async readSelection(workbookId: WorkbookId | string): Promise<SelectionInfo | undefined> {
    try {
      const result = await this.runtime.getSelection() as RuntimeSelectionResponse | { ok?: boolean; selection?: SelectionInfo };
      if ((result as { ok?: boolean }).ok === false) {
        return undefined;
      }
      const selection = result.selection;
      return selection && String(selection.workbookId) === String(workbookId) ? selection : undefined;
    } catch {
      return undefined;
    }
  }

  private async buildTableMetadata(workbookId: WorkbookId, sheets: any[]): Promise<TableMetadata[]> {
    const tableMetadatas: TableMetadata[] = [];
    const listed = await this.runtime.listTables(workbookId);
    const tableNames = ((listed as { tables?: Array<{ name?: string; tableName?: string }> }).tables ?? [])
      .map((table) => table.name ?? table.tableName)
      .filter((name): name is string => typeof name === "string");
    for (const tableName of tableNames) {
      const info = await this.runtime.getTableInfo({ workbookId, tableName });
      const raw = (info as { table?: any; info?: any }).table ?? (info as { table?: any; info?: any }).info ?? info;
      const sheetName = raw.sheetName ?? sheets.find((sheet) => sheet.tables?.some((table: any) => table.name === tableName || table.tableName === tableName))?.name ?? "";
      const columns = Array.isArray(raw.columns)
        ? raw.columns.map((column: any, index: number) => columnMetadata(index, String(column.name ?? column.header ?? column)))
        : [];
      tableMetadatas.push({
        id: `table:${tableName}`,
        sheetName,
        name: tableName,
        range: raw.range?.address ?? raw.address ?? "",
        headerRange: raw.headerRange?.address ?? raw.headerRange,
        dataRange: raw.dataRange?.address ?? raw.dataRange,
        columns
      });
    }
    return tableMetadatas;
  }

  private async buildNamedRanges(workbookId: WorkbookId) {
    const names = await this.runtime.listNames(workbookId);
    return ((names as { names?: NameInfo[] }).names ?? [])
      .filter((name) => name.address)
      .map((name) => ({
        name: name.name,
        ...(name.sheetName !== undefined ? { sheetName: name.sheetName } : {}),
        range: stripSheetName(name.address ?? "")
      }));
  }

  private async readSheetSample(workbookId: WorkbookId, sheet: any): Promise<CellMatrix> {
    const used = sheet.usedRange;
    const address = used?.address ? sampleAddress(used.address, used.rowCount, used.columnCount) : "A1:AD20";
    const operation: ExcelOperation = {
      kind: "range.read_full",
      operationId: makeId<OperationId>("op"),
      workbookId,
      destructiveLevel: "none",
      reason: "Build workbook metadata sample",
      target: { workbookId, sheetName: sheet.name, address },
      facets: ["values", "formulas", "text"]
    };
    const request: BatchRequest = { workbookId, mode: "apply", operations: [operation] };
    const result = await this.runtime.applyBatch(request);
    const snapshot = operationReadSnapshots(result)[0]?.snapshot;
    return mergeFormulaSample(snapshot?.values ?? snapshot?.text ?? [], snapshot?.formulas ?? []);
  }
}

function withFreshSelection(metadata: WorkbookMetadata, selection: SelectionInfo | undefined): WorkbookMetadata {
  const next: WorkbookMetadata = { ...metadata, updatedAt: Date.now() };
  if (selection) {
    next.selection = selection;
  } else {
    delete next.selection;
  }
  return next;
}

function operationReadSnapshots(result: unknown): Array<{ snapshot?: { values?: CellMatrix; formulas?: CellMatrix<string | null>; text?: string[][] } }> {
  const typed = result as {
    data?: Array<{ snapshot?: { values?: CellMatrix; formulas?: CellMatrix<string | null>; text?: string[][] } }>;
    readData?: Array<{ snapshot?: { values?: CellMatrix; formulas?: CellMatrix<string | null>; text?: string[][] } }>;
  };
  return typed.readData ?? typed.data ?? [];
}

function mergeFormulaSample(values: CellMatrix, formulas: CellMatrix<string | null>): CellMatrix {
  if (formulas.length === 0) {
    return values;
  }
  const rowCount = Math.max(values.length, formulas.length);
  return Array.from({ length: rowCount }, (_, rowIndex) => {
    const valueRow = values[rowIndex] ?? [];
    const formulaRow = formulas[rowIndex] ?? [];
    const columnCount = Math.max(valueRow.length, formulaRow.length);
    return Array.from({ length: columnCount }, (_cell, columnIndex) => {
      const formula = formulaRow[columnIndex];
      return typeof formula === "string" && formula.startsWith("=") ? formula : valueRow[columnIndex] ?? null;
    });
  });
}

function sheetMetadataFromMap(
  sheet: any,
  index: number,
  tables: TableMetadata[],
  sample: CellMatrix,
  sections: SectionMetadata[],
  summaryBlocks: SummaryBlockMetadata[],
  formulaRegions: FormulaRegionMetadata[]
): SheetMetadata {
  const usedRange = sheet.usedRange?.address ?? sheet.usedRange;
  const sampleHeaders = detectHeaders(String(sheet.name), sample);
  const tableHeaders: HeaderMetadata[] = tables.flatMap((table) =>
    table.columns.length > 0
      ? [{
          id: `header:${sheet.name}:${table.name ?? table.id}`,
          sheetName: sheet.name,
          row: 1,
          range: table.headerRange ?? table.range,
          columns: table.columns,
          confidence: 0.9
        }]
      : []
  );
  const headers = [...tableHeaders, ...sampleHeaders];
  const tableIds = tables.map((table) => table.id);
  return {
    id: `sheet:${index}`,
    name: sheet.name,
    index,
    ...(typeof usedRange === "string" ? { usedRange } : {}),
    ...(sheet.usedRange?.rowCount !== undefined ? { rowCount: sheet.usedRange.rowCount } : {}),
    ...(sheet.usedRange?.columnCount !== undefined ? { columnCount: sheet.usedRange.columnCount } : {}),
    ...(sheet.isHidden !== undefined ? { isHidden: sheet.isHidden } : {}),
    kind: inferSheetKind(sheet.name, headers.flatMap((header) => header.columns), tableIds.length, summaryBlocks.length),
    headers,
    tableIds,
    sectionIds: sections.map((section) => section.id),
    summaryBlockIds: summaryBlocks.map((block) => block.id),
    formulaRegionIds: formulaRegions.map((region) => region.id)
  };
}

function detectHeaders(sheetName: string, sample: CellMatrix): HeaderMetadata[] {
  const candidates = sample.slice(0, 20).map((row, rowIndex) => {
    const quality = headerRowQuality(row, rowIndex);
    return { row, rowIndex, quality };
  }).filter((candidate) => candidate.quality.accepted && candidate.quality.nonEmptyCount >= 3);
  return candidates
    .sort((left, right) => left.rowIndex - right.rowIndex || right.quality.confidence - left.quality.confidence)
    .map((candidate) => {
      const columns = candidate.row.map((value, index) => ({ value, index }))
        .filter((entry) => entry.value !== null && entry.value !== undefined && String(entry.value).trim() !== "")
        .map((entry) => columnMetadata(entry.index, String(entry.value)));
      return {
        id: `header:${sheetName}:${candidate.rowIndex + 1}`,
        sheetName,
        row: candidate.rowIndex + 1,
        range: `${columnLetter(columns[0]?.index ?? 0)}${candidate.rowIndex + 1}:${columnLetter(columns.at(-1)?.index ?? candidate.row.length - 1)}${candidate.rowIndex + 1}`,
        columns,
        confidence: Number(candidate.quality.confidence.toFixed(3))
      };
    });
}

function headerRowQuality(row: unknown[], rowIndex: number): {
  accepted: boolean;
  confidence: number;
  nonEmptyCount: number;
} {
  const cells = row
    .map((value) => String(value ?? "").trim())
    .filter((value) => value !== "");
  if (cells.length < 2) {
    return { accepted: false, confidence: 0, nonEmptyCount: cells.length };
  }
  const normalized = cells.map(normalizeHeaderName);
  const uniqueRatio = new Set(normalized.filter(Boolean)).size / Math.max(1, normalized.length);
  const formulaRatio = cells.filter((value) => value.startsWith("=")).length / cells.length;
  const numericRatio = cells.filter(isDataLikeNumber).length / cells.length;
  const longTextRatio = cells.filter((value) => value.length > 80).length / cells.length;
  const transactionValueRatio = cells.filter(isTransactionLikeValue).length / cells.length;
  const headerTermRatio = normalized.filter(isHeaderLikeName).length / cells.length;
  const summaryTermRatio = normalized.filter((value) => /summary|metric|amount|view|notes|total|revenue|expense|profit|cash|spend/.test(value)).length / cells.length;
  const topRowBonus = rowIndex <= 1 ? 0.12 : 0;
  const confidence = Math.max(0, Math.min(1,
    headerTermRatio * 0.7 +
    summaryTermRatio * 0.25 +
    uniqueRatio * 0.15 +
    topRowBonus -
    formulaRatio * 0.5 -
    numericRatio * 0.35 -
    longTextRatio * 0.4 -
    transactionValueRatio * 0.3
  ));
  const accepted = confidence >= 0.45
    && uniqueRatio >= 0.7
    && formulaRatio === 0
    && numericRatio <= 0.35
    && longTextRatio <= 0.15
    && transactionValueRatio <= 0.35
    && (headerTermRatio >= 0.2 || (rowIndex <= 1 && headerTermRatio >= 0.12) || summaryTermRatio >= 0.35);
  return { accepted, confidence, nonEmptyCount: cells.length };
}

function isDataLikeNumber(value: string): boolean {
  const normalized = value.replace(/[$,]/g, "");
  return /^\d+(\.\d+)?$/.test(normalized) || /^\d{5}$/.test(normalized);
}

function isTransactionLikeValue(value: string): boolean {
  return /^\d{4}-\d{4,}$/.test(value)
    || /^[A-Z]{3,4}\d{5,}$/i.test(value)
    || /\b(ref|scb|x\d{3,}|paid|transferred|cash transferred|wrong account|driver|refund)\b/i.test(value);
}

function isHeaderLikeName(value: string): boolean {
  return /date|month|year|period|id|no|number|account|customer|client|vendor|description|type|direction|status|amount|price|cost|fee|total|gross|net|tax|collect|payment|variance|note|proof|filename|invoice|booking|job|container|size|billed|lifting|category|metric|view/.test(value);
}

function detectSections(sheet: any, sample: CellMatrix, sheetIndex: number): SectionMetadata[] {
  if (sample.length === 0) {
    return [];
  }
  const sections: SectionMetadata[] = [];
  const rowBands = contiguousIndexGroups(
    sample
      .map((row, rowIndex) => ({ row, rowIndex }))
      .filter(({ row }) => row.some(isNonEmptyCell))
      .map(({ rowIndex }) => rowIndex)
  );
  for (const band of rowBands) {
    const columnIndexes = new Set<number>();
    for (let rowIndex = band.start; rowIndex <= band.end; rowIndex += 1) {
      const row = sample[rowIndex] ?? [];
      row.forEach((value, columnIndex) => {
        if (isNonEmptyCell(value)) {
          columnIndexes.add(columnIndex);
        }
      });
    }
    for (const columnGroup of contiguousIndexGroups([...columnIndexes])) {
      const section = buildSection(sheet, sample, sheetIndex, sections.length, band, columnGroup);
      if (section.nonEmptyCellCount >= 2) {
        sections.push(section);
      }
    }
  }
  return sections;
}

function buildSection(
  sheet: any,
  sample: CellMatrix,
  sheetIndex: number,
  sectionIndex: number,
  rows: { start: number; end: number },
  columns: { start: number; end: number }
): SectionMetadata {
  const matrix = sample.slice(rows.start, rows.end + 1).map((row) => row.slice(columns.start, columns.end + 1));
  const nonEmptyValues = matrix.flat().filter(isNonEmptyCell);
  const header = detectSectionHeader(matrix, rows.start, columns.start);
  const formulaCount = nonEmptyValues.filter((value) => typeof value === "string" && value.startsWith("=")).length;
  const labels = sectionLabels(matrix, header?.columns.map((column) => column.name) ?? []);
  const kind = inferSectionKind({ sheetName: String(sheet.name), labels, columns: header?.columns ?? [], formulaCount, rowStart: rows.start });
  const range = `${columnLetter(columns.start)}${rows.start + 1}:${columnLetter(columns.end)}${rows.end + 1}`;
  return {
    id: `section:${sheetIndex}:${sectionIndex}`,
    sheetName: String(sheet.name),
    label: sectionLabel(kind, labels, header?.columns ?? [], sectionIndex),
    kind,
    range,
    ...(header ? { headerRange: header.range, headerRow: header.row } : {}),
    columns: header?.columns ?? [],
    labels,
    rowCount: rows.end - rows.start + 1,
    columnCount: columns.end - columns.start + 1,
    nonEmptyCellCount: nonEmptyValues.length,
    confidence: sectionConfidence(kind, header?.columns.length ?? 0, nonEmptyValues.length)
  };
}

function detectSectionHeader(matrix: CellMatrix, rowOffset: number, columnOffset: number): { row: number; range: string; columns: ColumnMetadata[] } | undefined {
  const candidates = matrix.slice(0, Math.min(matrix.length, 5)).map((row, localRowIndex) => {
    const nonEmpty = row.map((value, index) => ({ value, index })).filter((entry) => isNonEmptyCell(entry.value));
    const quality = headerRowQuality(row, rowOffset + localRowIndex);
    return { row, localRowIndex, nonEmpty, quality };
  }).filter((candidate) => candidate.nonEmpty.length >= 2 && candidate.quality.accepted);
  const best = candidates.sort((left, right) => right.quality.confidence - left.quality.confidence || left.localRowIndex - right.localRowIndex)[0];
  if (!best) {
    return undefined;
  }
  const columns = best.nonEmpty.map((entry) => columnMetadata(columnOffset + entry.index, String(entry.value)));
  return {
    row: rowOffset + best.localRowIndex + 1,
    range: `${columnLetter(columns[0]?.index ?? columnOffset)}${rowOffset + best.localRowIndex + 1}:${columnLetter(columns.at(-1)?.index ?? columnOffset)}${rowOffset + best.localRowIndex + 1}`,
    columns
  };
}

function sectionLabels(matrix: CellMatrix, headers: string[]): string[] {
  const values = matrix
    .flat()
    .filter((value) => typeof value === "string" && value.trim() !== "" && !value.trim().startsWith("="))
    .map((value) => String(value).trim());
  return [...new Set([...headers, ...values].filter((value) => value.length <= 48))].slice(0, 12);
}

function inferSectionKind(input: { sheetName: string; labels: string[]; columns: ColumnMetadata[]; formulaCount: number; rowStart: number }): SectionKind {
  const text = normalizeHeaderName([input.sheetName, ...input.labels, ...input.columns.map((column) => column.name)].join(" "));
  if (input.formulaCount > 0 || /formula|reconciliation|variance|calc|calculation/.test(text)) return "formula";
  if (/note|comment|owner|status|action|approval|follow_up/.test(text)) return "notes";
  if (/summary|kpi|metric|total|revenue|expense|profit|balance/.test(text)) return "summary";
  if (input.columns.length >= 3) return "table-like";
  if (input.rowStart <= 2 || /report|period|prepared|as_of/.test(text)) return "metadata";
  return "unknown";
}

function sectionLabel(kind: SectionKind, labels: string[], columns: ColumnMetadata[], index: number): string {
  const headerText = normalizeHeaderName(columns.map((column) => column.name).join(" "));
  const text = normalizeHeaderName([...labels, ...columns.map((column) => column.name)].join(" "));
  if (/invoice|billed|booking|container|collect/.test(headerText)) return "invoice section";
  if (/transaction|payment|cash|truck/.test(headerText)) return "transaction section";
  if (/kpi|metric|summary|revenue|expense|profit/.test(text)) return "KPI summary section";
  if (/note|owner|status|action|approval/.test(text)) return "notes/status section";
  if (/formula|reconciliation|variance/.test(text)) return "formula/reconciliation section";
  return `${kind} section ${index + 1}`;
}

function sectionConfidence(kind: SectionKind, columnCount: number, nonEmptyCellCount: number): number {
  const base = kind === "unknown" ? 0.45 : 0.65;
  return Number(Math.min(0.95, base + Math.min(columnCount, 6) * 0.03 + Math.min(nonEmptyCellCount, 20) * 0.005).toFixed(3));
}

function contiguousIndexGroups(indexes: number[]): Array<{ start: number; end: number }> {
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

function isNonEmptyCell(value: unknown): value is string | number | boolean {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function detectSummaryBlocks(sheet: any, sample: CellMatrix, index: number): SummaryBlockMetadata[] {
  const labels = sample.flat().filter((value) => typeof value === "string" && /total|revenue|expense|net|balance|paid|unpaid|profit|loss/i.test(value)).map(String);
  if (labels.length === 0 && !/summary|dashboard|report|p&l|profit|loss/i.test(String(sheet.name))) {
    return [];
  }
  return [{
    id: `summary:${index}`,
    sheetName: String(sheet.name),
    range: sheet.usedRange?.address ?? "A1:AD20",
    labels: labels.slice(0, 20),
    confidence: labels.length > 0 ? 0.75 : 0.55
  }];
}

function detectFormulaRegions(sheet: any, sample: CellMatrix, index: number): FormulaRegionMetadata[] {
  let formulaCount = 0;
  for (const row of sample) {
    for (const value of row) {
      if (typeof value === "string" && value.startsWith("=")) {
        formulaCount += 1;
      }
    }
  }
  if (formulaCount === 0) {
    return [];
  }
  return [{
    id: `formula:${index}`,
    sheetName: String(sheet.name),
    range: sheet.usedRange?.address ?? "A1:AD20",
    formulaCount
  }];
}

export function columnMetadata(index: number, name: string): ColumnMetadata {
  const inferredType = inferColumnType(name);
  const role = inferColumnRole(name, inferredType);
  return {
    index,
    letter: columnLetter(index),
    name,
    normalizedName: normalizeHeaderName(name),
    inferredType,
    role,
    importance: columnImportance(name, inferredType, role)
  };
}

export function inferColumnType(name: string): ColumnType {
  const normalized = normalizeHeaderName(name);
  if (/amount|total|revenue|expense|balance|price|cost/.test(normalized)) return "currency";
  if (/date|month|year|period/.test(normalized)) return "date";
  if (/status|state/.test(normalized)) return "status";
  if (/count|qty|quantity|number|rate/.test(normalized)) return "number";
  if (/formula/.test(normalized)) return "formula";
  return "unknown";
}

export function inferColumnRole(name: string, inferredType = inferColumnType(name)): ColumnRole {
  const normalized = normalizeHeaderName(name);
  if (/date|month|year|period|posted|created|paid_at|invoice_date/.test(normalized)) return "date";
  if (/description|desc|memo|detail|details|narration|particular|purpose|remark/.test(normalized)) return "description";
  if (/vendor|merchant|supplier|payee|customer|client|company|counterparty/.test(normalized)) return "vendor";
  if (/account|bank|wallet|card/.test(normalized)) return "account";
  if (/amount|total|revenue|expense|balance|price|cost|fee|gross|net|tax|collect|paid|payment/.test(normalized)) return "amount";
  if (/status|state|stage|approval|closed|open/.test(normalized)) return "status";
  if (/category|label|class|tag|type|group|kind/.test(normalized)) return "category";
  if (/\bid\b|no|number|ref|reference|invoice|booking|job|container|filename|receipt/.test(normalized)) return "identifier";
  if (/note|comment|remark|reason/.test(normalized)) return "note";
  if (inferredType === "formula" || /formula|calc|variance|delta/.test(normalized)) return "formula";
  if (inferredType === "currency" || inferredType === "number") return "measure";
  if (inferredType === "date") return "date";
  if (inferredType === "status") return "status";
  return "dimension";
}

function columnImportance(name: string, inferredType: ColumnType, role: ColumnRole): number {
  const baseByRole: Record<ColumnRole, number> = {
    date: 0.94,
    description: 0.98,
    vendor: 0.9,
    account: 0.82,
    amount: 0.96,
    status: 0.9,
    category: 0.92,
    identifier: 0.72,
    formula: 0.78,
    note: 0.7,
    dimension: 0.55,
    measure: 0.82,
    unknown: 0.4
  };
  const normalized = normalizeHeaderName(name);
  const typeBonus = inferredType === "currency" || inferredType === "date" || inferredType === "status" ? 0.03 : 0;
  const labelBonus = /label|category|status|amount|date|description|vendor|customer|company/.test(normalized) ? 0.03 : 0;
  return Number(Math.min(1, baseByRole[role] + typeBonus + labelBonus).toFixed(2));
}

export function inferSheetKind(sheetName: string, columns: ColumnMetadata[], tableCount: number, summaryBlockCount = 0): SheetKind {
  const normalized = normalizeHeaderName(sheetName);
  const columnNames = new Set(columns.map((column) => column.normalizedName));
  if (/template/.test(normalized)) return "template";
  if (/summary|dashboard|report|p_l|profit|loss/.test(normalized) || summaryBlockCount > 0) return /dashboard/.test(normalized) ? "dashboard" : "summary";
  if (/lookup|config|setting/.test(normalized)) return "lookup";
  if (tableCount > 0 || (columnNames.has("customer") && (columnNames.has("amount") || columnNames.has("total_amount")))) return "transaction";
  return "unknown";
}

function sampleAddress(address: string, rowCount?: number, columnCount?: number): string {
  const start = /^([A-Z]+)(\d+)/i.exec(stripSheetName(address)) ?? /^([A-Z]+)(\d+)/i.exec("A1");
  const startColumn = columnIndex(start?.[1] ?? "A");
  const startRow = Number.parseInt(start?.[2] ?? "1", 10);
  const rows = Math.max(1, Math.min(rowCount ?? 20, 20));
  const cols = Math.max(1, Math.min(columnCount ?? 40, 40));
  const endColumn = startColumn + cols - 1;
  const endRow = startRow + rows - 1;
  return `${columnLetter(startColumn)}${startRow}:${columnLetter(endColumn)}${endRow}`;
}

function columnIndex(column: string): number {
  let value = 0;
  for (const char of column.toUpperCase()) {
    value = value * 26 + char.charCodeAt(0) - 64;
  }
  return Math.max(0, value - 1);
}

function groupBy<T>(items: T[], keyFn: (item: T) => string): Map<string, T[]> {
  const groups = new Map<string, T[]>();
  for (const item of items) {
    const key = keyFn(item);
    const group = groups.get(key) ?? [];
    group.push(item);
    groups.set(key, group);
  }
  return groups;
}
