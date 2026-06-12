import type { BackupId, OperationId, PlanId, SnapshotId, TemplateId, WorkbookId } from "./ids.js";
import type { A1Range, CellMatrix, RangeFingerprint, WorkbookFingerprint } from "./workbook.js";
import type { ExcelRuntimeError } from "./errors.js";

export type DestructiveLevel = "none" | "values" | "format" | "structure" | "workbook";

export interface OperationTelemetry {
  durationMs?: number;
  syncCount?: number;
  payloadBytes?: number;
  cellsRead?: number;
  cellsWritten?: number;
  rangeCount?: number;
  chunkCount?: number;
  engineName?: string;
  engineVersion?: string;
  warningCount?: number;
}

export interface OperationWarning {
  code: string;
  message: string;
  target?: A1Range;
  details?: Record<string, unknown>;
}

export interface OperationBase {
  operationId: OperationId;
  workbookId: WorkbookId;
  destructiveLevel: DestructiveLevel;
  reason: string;
}

export interface ReadFullOperation extends OperationBase {
  kind: "range.read_full";
  target: A1Range;
  includeStyles?: boolean;
  includeFormulas?: boolean;
  includeValidation?: boolean;
  includeComments?: boolean;
}

export interface WriteValuesOperation extends OperationBase {
  kind: "range.write_values";
  target: A1Range;
  values: CellMatrix;
  preserveFormats: true;
}

export interface WriteFormulasOperation extends OperationBase {
  kind: "range.write_formulas";
  target: A1Range;
  formulas: CellMatrix<string | null>;
  preserveFormats: true;
}

export interface ClearValuesKeepFormatOperation extends OperationBase {
  kind: "range.clear_values_keep_format";
  target: A1Range;
}

export interface CreateSheetFromTemplateOperation extends OperationBase {
  kind: "template.create_sheet_from_template";
  templateId: TemplateId;
  newSheetName: string;
  clearDataRegions: boolean;
}

export type ExcelOperation =
  | ReadFullOperation
  | WriteValuesOperation
  | WriteFormulasOperation
  | ClearValuesKeepFormatOperation
  | CreateSheetFromTemplateOperation;

export interface BatchRequest {
  workbookId: WorkbookId;
  operations: ExcelOperation[];
  mode: "validate" | "dry_run" | "apply";
  confirmationToken?: string;
  expectedTargetFingerprints?: RangeFingerprint[];
  baseSnapshotId?: SnapshotId;
}

export interface CompiledBatch {
  workbookId: WorkbookId;
  operations: ExcelOperation[];
  requiredBackups: Array<"region" | "sheet" | "workbook-copy">;
  targetFingerprints: RangeFingerprint[];
  estimatedCellsTouched: number;
  destructiveLevel: DestructiveLevel;
}

export interface PlanCreateRequest {
  workbookId: WorkbookId;
  goal: string;
  operations: ExcelOperation[];
  baseSnapshotId?: SnapshotId;
}

export interface PlanPreview {
  planId: PlanId;
  workbookId: WorkbookId;
  baseSnapshotId: SnapshotId;
  requiredBackups: BackupId[];
  beforeWorkbookFingerprint: WorkbookFingerprint;
  targetFingerprints: RangeFingerprint[];
  diffSummary: DiffSummary;
  warnings: OperationWarning[];
}

export interface DiffSummary {
  title: string;
  changedRanges: A1Range[];
  cellsChanged: number;
  formulasChanged: number;
  stylesChanged: number;
  tablesChanged: number;
  sheetsChanged: number;
  destructiveLevel: DestructiveLevel;
}

export interface OperationResult {
  ok: boolean;
  operationId?: OperationId;
  planId?: PlanId;
  diffSummary?: DiffSummary;
  data?: unknown;
  rollbackAvailable: boolean;
  backups: BackupId[];
  warnings: OperationWarning[];
  telemetry: OperationTelemetry;
  error?: ExcelRuntimeError;
}

export interface RangeSnapshot {
  fingerprint: RangeFingerprint;
  values?: CellMatrix;
  formulas?: CellMatrix<string | null>;
  numberFormat?: string[][];
  text?: string[][];
  style?: {
    fillColor?: string;
    fontName?: string;
    fontSize?: number;
    fontColor?: string;
    fontBold?: boolean;
    fontItalic?: boolean;
    horizontalAlignment?: string;
    verticalAlignment?: string;
    rowHeight?: number;
    columnWidth?: number;
  };
}

export interface WorkbookSnapshotResponse {
  workbookFingerprint: WorkbookFingerprint;
  rangeSnapshots: RangeSnapshot[];
}

export interface TemplateExecutionSource {
  templateId: TemplateId;
  sourceSheetName: string;
  dataRegions: string[];
}

export interface AddinExecuteBatchRequest {
  request: BatchRequest;
  compiled: CompiledBatch;
  templateSources?: TemplateExecutionSource[];
}
