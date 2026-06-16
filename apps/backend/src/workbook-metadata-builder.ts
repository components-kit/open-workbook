import { makeId, type BatchRequest, type CellMatrix, type ExcelOperation, type NameInfo, type OperationId, type WorkbookId, type WorkbookRef } from "@components-kit/open-workbook-protocol";
import {
  checkMetadataFreshness,
  columnLetter,
  createMetadataFingerprint,
  DEFAULT_METADATA_CACHE_TTL_MS,
  normalizeHeaderName,
  type ColumnMetadata,
  type ColumnType,
  type FormulaRegionMetadata,
  type HeaderMetadata,
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
}

export class WorkbookMetadataBuilder {
  constructor(
    private readonly runtime: RuntimeService,
    private readonly cache: WorkbookMetadataCache
  ) {}

  async getOrBuild(input: { workbookContextId?: string; workbookId?: WorkbookId | string; workbookName?: string }): Promise<MetadataBuildResult> {
    const existingByContext = input.workbookContextId ? this.cache.getByContextId(input.workbookContextId) : undefined;
    const reusableContextId = existingByContext?.workbookContextId;
    const activeContext = await this.runtime.getActiveContext();
    const activeWorkbook = (activeContext as { activeWorkbook?: WorkbookRef }).activeWorkbook ?? this.runtime.sessions.getActive()?.activeWorkbook;
    if (!activeWorkbook) {
      throw new Error("No active Excel workbook is available. Open Excel and connect the Open Workbook add-in.");
    }
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
      const freshness = checkMetadataFreshness(existingByContext, fingerprint);
      if (freshness.status === "FRESH") {
        return { metadata: existingByContext, cacheHit: true };
      }
      this.cache.delete(existingByContext.workbookKey);
    }

    const key = workbookMetadataKey({
      workbookId: activeWorkbook.workbookId,
      workbookName: activeWorkbook.name,
      ...(activeWorkbook.path !== undefined ? { workbookPath: activeWorkbook.path } : {})
    });
    const existing = this.cache.get(key);
    if (existing && checkMetadataFreshness(existing, fingerprint).status === "FRESH") {
      return { metadata: existing, cacheHit: true };
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
    for (const sheet of sheets) {
      const sample = await this.readSheetSample(activeWorkbook.workbookId, sheet);
      if (sample.length > 0) {
        samples.set(sheet.name, sample);
      }
    }

    const summaryBlocks = sheets.flatMap((sheet, index) => detectSummaryBlocks(sheet, samples.get(sheet.name) ?? [], index));
    const formulaRegions = sheets.flatMap((sheet, index) => detectFormulaRegions(sheet, samples.get(sheet.name) ?? [], index));
    const summaryIdsBySheet = groupBy(summaryBlocks, (block) => block.sheetName);
    const formulaIdsBySheet = groupBy(formulaRegions, (region) => region.sheetName);
    const now = Date.now();
    const metadata: WorkbookMetadata = {
      workbookContextId: reusableContextId ?? makeId("wbctx"),
      workbookKey: key,
      workbook: {
        workbookId: activeWorkbook.workbookId,
        name: activeWorkbook.name,
        sheetCount: sheets.length,
        ...(activeWorkbook.path !== undefined ? { path: activeWorkbook.path } : {}),
        ...(map.activeSheet !== undefined ? { activeSheet: map.activeSheet } : {})
      },
      sheets: sheets.map((sheet, index) =>
        sheetMetadataFromMap(
          sheet,
          index,
          tableMap.get(sheet.name) ?? [],
          samples.get(sheet.name) ?? [],
          summaryIdsBySheet.get(sheet.name) ?? [],
          formulaIdsBySheet.get(sheet.name) ?? []
        )
      ),
      tables,
      namedRanges: [...namedRanges, ...registeredRegions],
      summaryBlocks,
      formulaRegions,
      fingerprint,
      createdAt: now,
      updatedAt: now,
      expiresAt: now + DEFAULT_METADATA_CACHE_TTL_MS
    };
    return { metadata: this.cache.set(metadata), cacheHit: false };
  }

  invalidateWorkbook(workbookId: WorkbookId | string): void {
    this.cache.deleteByWorkbookId(workbookId);
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
    const snapshot = (result as { readData?: Array<{ snapshot?: { values?: CellMatrix; formulas?: CellMatrix<string | null>; text?: string[][] } }> }).readData?.[0]?.snapshot;
    return mergeFormulaSample(snapshot?.values ?? snapshot?.text ?? [], snapshot?.formulas ?? []);
  }
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
    summaryBlockIds: summaryBlocks.map((block) => block.id),
    formulaRegionIds: formulaRegions.map((region) => region.id)
  };
}

function detectHeaders(sheetName: string, sample: CellMatrix): HeaderMetadata[] {
  const candidates = sample.slice(0, 20).map((row, rowIndex) => {
    const nonEmpty = row.filter((value) => value !== null && value !== undefined && String(value).trim() !== "");
    const textLike = nonEmpty.filter((value) => typeof value === "string" && !/^\d+([.,]\d+)?$/.test(value.trim()));
    return { row, rowIndex, nonEmpty, textLike, confidence: nonEmpty.length >= 3 ? textLike.length / nonEmpty.length : 0 };
  }).filter((candidate) => candidate.nonEmpty.length >= 3 && candidate.confidence >= 0.6);
  const best = candidates.sort((left, right) => right.confidence - left.confidence)[0];
  if (!best) {
    return [];
  }
  const columns = best.row.map((value, index) => ({ value, index }))
    .filter((entry) => entry.value !== null && entry.value !== undefined && String(entry.value).trim() !== "")
    .map((entry) => columnMetadata(entry.index, String(entry.value)));
  return [{
    id: `header:${sheetName}:${best.rowIndex + 1}`,
    sheetName,
    row: best.rowIndex + 1,
    range: `A${best.rowIndex + 1}:${columnLetter(best.row.length - 1)}${best.rowIndex + 1}`,
    columns,
    confidence: Number(best.confidence.toFixed(3))
  }];
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
  return {
    index,
    letter: columnLetter(index),
    name,
    normalizedName: normalizeHeaderName(name),
    inferredType: inferColumnType(name)
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
  const start = /^'?[^'!]+(?:'!)?([A-Z]+\d+)/i.exec(address)?.[1] ?? "A1";
  const rows = Math.max(1, Math.min(rowCount ?? 20, 20));
  const cols = Math.max(1, Math.min(columnCount ?? 30, 30));
  return `${start}:${columnLetter(cols - 1)}${rows}`;
}

function stripSheetName(address: string): string {
  return address.includes("!") ? address.split("!").pop() ?? address : address;
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
