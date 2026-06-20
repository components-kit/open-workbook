import { hashStable } from "@components-kit/open-workbook-excel-core";
import type { SelectionInfo, WorkbookFingerprint, WorkbookId } from "@components-kit/open-workbook-protocol";

export type SheetKind = "transaction" | "summary" | "dashboard" | "lookup" | "template" | "unknown";
export type ColumnType = "text" | "number" | "currency" | "date" | "status" | "formula" | "unknown";

export interface ColumnMetadata {
  index: number;
  letter: string;
  name: string;
  normalizedName: string;
  inferredType: ColumnType;
}

export interface HeaderMetadata {
  id: string;
  sheetName: string;
  row: number;
  range: string;
  columns: ColumnMetadata[];
  confidence: number;
}

export interface SheetMetadata {
  id: string;
  name: string;
  index: number;
  usedRange?: string;
  rowCount?: number;
  columnCount?: number;
  isHidden?: boolean;
  kind: SheetKind;
  headers: HeaderMetadata[];
  tableIds: string[];
  sectionIds: string[];
  summaryBlockIds: string[];
  formulaRegionIds: string[];
}

export interface TableMetadata {
  id: string;
  sheetName: string;
  name?: string;
  range: string;
  headerRange?: string;
  dataRange?: string;
  columns: ColumnMetadata[];
}

export interface NamedRangeMetadata {
  name: string;
  sheetName?: string;
  range: string;
}

export interface SummaryBlockMetadata {
  id: string;
  sheetName: string;
  range: string;
  labels: string[];
  confidence: number;
}

export interface FormulaRegionMetadata {
  id: string;
  sheetName: string;
  range: string;
  formulaCount: number;
}

export type SectionKind = "table-like" | "summary" | "notes" | "formula" | "metadata" | "unknown";

export interface SectionMetadata {
  id: string;
  sheetName: string;
  label: string;
  kind: SectionKind;
  range: string;
  headerRange?: string;
  headerRow?: number;
  columns: ColumnMetadata[];
  labels: string[];
  rowCount: number;
  columnCount: number;
  nonEmptyCellCount: number;
  confidence: number;
}

export interface WorkbookMetadata {
  workbookContextId: string;
  workbookKey: string;
  detailLevel: "structure" | "sampled";
  contentVersion?: number;
  workbook: {
    workbookId?: WorkbookId | string;
    name: string;
    path?: string;
    activeSheet?: string;
    sheetCount: number;
  };
  selection?: SelectionInfo;
  sheets: SheetMetadata[];
  tables: TableMetadata[];
  namedRanges: NamedRangeMetadata[];
  sections: SectionMetadata[];
  summaryBlocks: SummaryBlockMetadata[];
  formulaRegions: FormulaRegionMetadata[];
  fingerprint: WorkbookFingerprint;
  createdAt: number;
  updatedAt: number;
  expiresAt: number;
}

export interface WorkbookMetadataFreshness {
  status: "FRESH" | "STALE";
  reason?: string;
}

export const DEFAULT_METADATA_CACHE_TTL_MS = 60 * 60 * 1000;

export class WorkbookMetadataCache {
  private readonly byKey = new Map<string, WorkbookMetadata>();
  private readonly keyByContextId = new Map<string, string>();

  get(workbookKey: string): WorkbookMetadata | undefined {
    const metadata = this.byKey.get(workbookKey);
    if (!metadata) {
      return undefined;
    }
    if (Date.now() > metadata.expiresAt) {
      this.delete(workbookKey);
      return undefined;
    }
    return metadata;
  }

  getByContextId(workbookContextId: string): WorkbookMetadata | undefined {
    const key = this.keyByContextId.get(workbookContextId);
    return key ? this.get(key) : undefined;
  }

  set(metadata: WorkbookMetadata): WorkbookMetadata {
    this.byKey.set(metadata.workbookKey, metadata);
    this.keyByContextId.set(metadata.workbookContextId, metadata.workbookKey);
    return metadata;
  }

  update(workbookKey: string, patch: Partial<WorkbookMetadata>): WorkbookMetadata | undefined {
    const existing = this.get(workbookKey);
    if (!existing) {
      return undefined;
    }
    const updated = {
      ...existing,
      ...patch,
      updatedAt: Date.now()
    };
    return this.set(updated);
  }

  delete(workbookKey: string): void {
    const existing = this.byKey.get(workbookKey);
    if (existing) {
      this.keyByContextId.delete(existing.workbookContextId);
    }
    this.byKey.delete(workbookKey);
  }

  deleteByWorkbookId(workbookId: WorkbookId | string): void {
    for (const [key, metadata] of this.byKey.entries()) {
      if (metadata.workbook.workbookId === workbookId) {
        this.delete(key);
      }
    }
  }

  clearExpired(now = Date.now()): void {
    for (const [key, metadata] of this.byKey.entries()) {
      if (now > metadata.expiresAt) {
        this.delete(key);
      }
    }
  }
}

export function workbookMetadataKey(input: { workbookId?: WorkbookId | string; workbookName: string; workbookPath?: string }): string {
  if (input.workbookPath) {
    return `path:${input.workbookPath}`;
  }
  if (input.workbookId) {
    return `id:${input.workbookId}`;
  }
  return `name:${input.workbookName}`;
}

export function createMetadataFingerprint(input: {
  workbookId: WorkbookId | string;
  workbook: unknown;
  sheets: Array<{ name: string; usedRange?: { address?: string; rowCount?: number; columnCount?: number }; isHidden?: boolean; tables?: unknown[] }>;
}): WorkbookFingerprint {
  const structure = input.sheets.map((sheet) => ({
    name: sheet.name,
    usedRange: sheet.usedRange,
    isHidden: sheet.isHidden,
    tables: sheet.tables
  }));
  return {
    workbookId: input.workbookId as WorkbookId,
    workbookHash: hashStable(input.workbook),
    structureHash: hashStable(structure),
    capturedAt: new Date().toISOString()
  };
}

export function checkMetadataFreshness(
  metadata: WorkbookMetadata,
  latest: WorkbookFingerprint,
  options: { requireSampled?: boolean; contentVersion?: number } = {}
): WorkbookMetadataFreshness {
  if (metadata.fingerprint.structureHash !== latest.structureHash) {
    return { status: "STALE", reason: "Workbook structure fingerprint changed." };
  }
  if (options.requireSampled === true && metadata.detailLevel !== "sampled") {
    return { status: "STALE", reason: "Workbook metadata needs sheet samples." };
  }
  if (
    options.requireSampled === true
    && metadata.detailLevel === "sampled"
    && options.contentVersion !== undefined
    && metadata.contentVersion !== undefined
    && metadata.contentVersion !== options.contentVersion
  ) {
    return { status: "STALE", reason: "Workbook content changed since sampled metadata was captured." };
  }
  return { status: "FRESH" };
}

export function normalizeHeaderName(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, "_").replace(/[^\w]/g, "");
}

export function columnLetter(index: number): string {
  let n = index + 1;
  let value = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    n = Math.floor((n - 1) / 26);
  }
  return value;
}
