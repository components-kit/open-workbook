import { hashStable } from "@components-kit/open-workbook-excel-core";
import type { SelectionInfo, WorkbookFingerprint, WorkbookId } from "@components-kit/open-workbook-protocol";

export type SheetKind = "transaction" | "summary" | "dashboard" | "lookup" | "template" | "unknown";
export type ColumnType = "text" | "number" | "currency" | "date" | "status" | "formula" | "unknown";
export type ColumnRole =
  | "date"
  | "description"
  | "vendor"
  | "account"
  | "amount"
  | "status"
  | "category"
  | "identifier"
  | "formula"
  | "note"
  | "dimension"
  | "measure"
  | "unknown";

export interface ColumnMetadata {
  index: number;
  letter: string;
  name: string;
  normalizedName: string;
  inferredType: ColumnType;
  role?: ColumnRole;
  importance?: number;
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

export type ContextFacet =
  | "metadata"
  | "schema"
  | "headers"
  | "tableDimensions"
  | "regions"
  | "fieldContext"
  | "validation"
  | "formats"
  | "formulas"
  | "formulaResults"
  | "filters"
  | "values"
  | "aggregates"
  | "rowPositions"
  | "selection"
  | "names";

export type FacetFreshnessStatus = "fresh" | "mostly_fresh" | "partially_stale" | "stale";

export interface ContextFreshness {
  status: FacetFreshnessStatus;
  freshFacets: ContextFacet[];
  staleFacets: ContextFacet[];
  staleRanges?: string[];
  confidence: number;
  updatedAt: number;
}

export interface OperationJournalEntry {
  operationId: string;
  workbookContextId: string;
  contextVersion: number;
  appliedAt: number;
  affectedRanges: string[];
  affectedFacets: ContextFacet[];
  invalidatedFacets: ContextFacet[];
  preservedFacets: ContextFacet[];
  changes?: Array<{ sheetName: string; range?: string; cell?: string; columnName?: string; before?: unknown; after?: unknown }>;
  cacheAction: "recorded" | "updated_from_patch" | "invalidated";
}

export interface CacheImpactSummary {
  cacheAction: OperationJournalEntry["cacheAction"];
  contextVersion: number;
  freshness: ContextFreshness;
  journalEntry: OperationJournalEntry;
}

export interface WorkbookContextState {
  workbookContextId: string;
  contextVersion: number;
  lastValidatedAt?: number;
  freshness: ContextFreshness;
  journal: OperationJournalEntry[];
}

export interface ContextFreshnessCheck {
  status: FacetFreshnessStatus;
  requiredFacets: ContextFacet[];
  freshRequiredFacets: ContextFacet[];
  staleRequiredFacets: ContextFacet[];
  staleRanges?: string[];
  requiresRead: boolean;
  reason: string;
  confidence: number;
}

export const DEFAULT_METADATA_CACHE_TTL_MS = 60 * 60 * 1000;

export class WorkbookMetadataCache {
  private readonly byKey = new Map<string, WorkbookMetadata>();
  private readonly keyByContextId = new Map<string, string>();
  private readonly contextStateById = new Map<string, WorkbookContextState>();

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
    if (!this.contextStateById.has(metadata.workbookContextId)) {
      this.contextStateById.set(metadata.workbookContextId, createInitialContextState(metadata.workbookContextId, metadata.updatedAt));
    }
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
      this.contextStateById.delete(existing.workbookContextId);
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

  deleteByContextId(workbookContextId: string): boolean {
    const key = this.keyByContextId.get(workbookContextId);
    if (!key) {
      return false;
    }
    this.delete(key);
    return true;
  }

  clearExpired(now = Date.now()): void {
    for (const [key, metadata] of this.byKey.entries()) {
      if (now > metadata.expiresAt) {
        this.delete(key);
      }
    }
  }

  getContextState(workbookContextId: string): WorkbookContextState | undefined {
    const state = this.contextStateById.get(workbookContextId);
    return state ? cloneContextState(state) : undefined;
  }

  checkFacetFreshness(workbookContextId: string, requiredFacets: ContextFacet[]): ContextFreshnessCheck | undefined {
    const state = this.contextStateById.get(workbookContextId);
    if (!state) {
      return undefined;
    }
    const stale = new Set(state.freshness.staleFacets);
    const uniqueRequired = [...new Set(requiredFacets)];
    const staleRequiredFacets = uniqueRequired.filter((facet) => stale.has(facet));
    const freshRequiredFacets = uniqueRequired.filter((facet) => !stale.has(facet));
    return {
      status: staleRequiredFacets.length === 0 ? "fresh" : freshnessStatus(freshRequiredFacets.length, staleRequiredFacets.length),
      requiredFacets: uniqueRequired,
      freshRequiredFacets,
      staleRequiredFacets,
      ...(state.freshness.staleRanges ? { staleRanges: [...state.freshness.staleRanges] } : {}),
      requiresRead: staleRequiredFacets.length > 0,
      reason: staleRequiredFacets.length === 0
        ? "Required context facets are fresh."
        : `Required context facets are stale: ${staleRequiredFacets.join(", ")}.`,
      confidence: uniqueRequired.length === 0 ? state.freshness.confidence : freshRequiredFacets.length / uniqueRequired.length
    };
  }

  markFacetsStale(workbookContextId: string, facets: ContextFacet[], staleRanges: string[] = [], now = Date.now()): WorkbookContextState | undefined {
    const state = this.contextStateById.get(workbookContextId);
    if (!state) {
      return undefined;
    }
    const stale = new Set<ContextFacet>([...state.freshness.staleFacets, ...facets]);
    const fresh = state.freshness.freshFacets.filter((facet) => !stale.has(facet));
    const ranges = [...new Set([...(state.freshness.staleRanges ?? []), ...staleRanges])];
    const next: WorkbookContextState = {
      ...state,
      contextVersion: state.contextVersion + 1,
      freshness: {
        status: freshnessStatus(fresh.length, stale.size),
        freshFacets: fresh,
        staleFacets: [...stale],
        ...(ranges.length > 0 ? { staleRanges: ranges } : {}),
        confidence: freshnessConfidence(fresh.length, stale.size),
        updatedAt: now
      }
    };
    this.contextStateById.set(workbookContextId, next);
    return cloneContextState(next);
  }

  appendJournalEntry(workbookContextId: string, entry: Omit<OperationJournalEntry, "workbookContextId" | "contextVersion" | "appliedAt">, now = Date.now()): OperationJournalEntry | undefined {
    const state = this.contextStateById.get(workbookContextId);
    if (!state) {
      return undefined;
    }
    const { changes, ...entryWithoutChanges } = entry;
    const journalEntry: OperationJournalEntry = {
      ...entryWithoutChanges,
      workbookContextId,
      contextVersion: state.contextVersion,
      appliedAt: now,
      ...(changes !== undefined ? { changes: changes.map((change) => ({ ...change })) } : {})
    };
    const next = {
      ...state,
      journal: [...state.journal, journalEntry].slice(-100)
    };
    this.contextStateById.set(workbookContextId, next);
    return cloneJournalEntry(journalEntry);
  }
}

const ALL_CONTEXT_FACETS: ContextFacet[] = [
  "metadata",
  "schema",
  "headers",
  "tableDimensions",
  "regions",
  "fieldContext",
  "validation",
  "formats",
  "formulas",
  "formulaResults",
  "filters",
  "values",
  "aggregates",
  "rowPositions",
  "selection",
  "names"
];

function createInitialContextState(workbookContextId: string, now: number): WorkbookContextState {
  return {
    workbookContextId,
    contextVersion: 1,
    lastValidatedAt: now,
    freshness: {
      status: "fresh",
      freshFacets: [...ALL_CONTEXT_FACETS],
      staleFacets: [],
      confidence: 1,
      updatedAt: now
    },
    journal: []
  };
}

function freshnessStatus(freshCount: number, staleCount: number): FacetFreshnessStatus {
  if (staleCount === 0) return "fresh";
  if (freshCount === 0) return "stale";
  return staleCount <= 2 ? "mostly_fresh" : "partially_stale";
}

function freshnessConfidence(freshCount: number, staleCount: number): number {
  const total = freshCount + staleCount;
  return total === 0 ? 0 : Math.max(0, Math.min(1, freshCount / total));
}

function cloneContextState(state: WorkbookContextState): WorkbookContextState {
  return {
    ...state,
    freshness: {
      ...state.freshness,
      freshFacets: [...state.freshness.freshFacets],
      staleFacets: [...state.freshness.staleFacets],
      ...(state.freshness.staleRanges ? { staleRanges: [...state.freshness.staleRanges] } : {})
    },
    journal: state.journal.map(cloneJournalEntry)
  };
}

function cloneJournalEntry(entry: OperationJournalEntry): OperationJournalEntry {
  return {
    ...entry,
    affectedRanges: [...entry.affectedRanges],
    affectedFacets: [...entry.affectedFacets],
    invalidatedFacets: [...entry.invalidatedFacets],
    preservedFacets: [...entry.preservedFacets],
    ...(entry.changes !== undefined ? { changes: entry.changes.map((change) => ({ ...change })) } : {})
  };
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
